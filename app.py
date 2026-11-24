import io, os
from datetime import time, timedelta
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'local-template-key')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

def read_excel(file):
    name=file.filename.lower()
    engine='xlrd' if name.endswith('.xls') else 'openpyxl'
    return pd.read_excel(file, engine=engine)
def col(df, hints):
    for c in df.columns:
        s=str(c).strip().lower()
        if any(h in s for h in hints): return c
    return None
def process(files):
    attendance=read_excel(files['attendance'])
    roster=read_excel(files['roster'])
    leave=read_excel(files['leave']) if files.get('leave') else pd.DataFrame()
    trip=read_excel(files['trip']) if files.get('trip') else pd.DataFrame()
    namecol=col(attendance,['姓名','员工姓名','人员姓名','员工'])
    timecol=col(attendance,['事件时间','打卡时间','时间','日期'])
    typcol=col(attendance,['事件类型','打卡类型','考勤类型'])
    rname=col(roster,['姓名','员工姓名','员工'])
    if not namecol or not timecol or not rname: raise ValueError('无法识别姓名或时间列，请确认上传的是原始考勤表和员工名册。')
    df=attendance.copy()
    df['_dt']=pd.to_datetime(df[timecol].astype('string').str.strip(), format='mixed', errors='coerce')
    df=df[df['_dt'].notna() & df[namecol].notna()].copy()
    if df.empty:
        raise ValueError('考勤记录中没有可识别的姓名和打卡时间。')
    start, end = df['_dt'].min().date(), df['_dt'].max().date()
    workdays = [start + timedelta(days=i) for i in range((end-start).days+1)
                if (start + timedelta(days=i)).weekday() < 5]
    df=df[(df['_dt'].dt.date>=start)&(df['_dt'].dt.date<=end)]
    df['_name']=df[namecol].astype(str).str.strip()
    if typcol:
        # Keep ordinary door access and filter obvious non-attendance events only when the source describes event type.
        df['_type']=df[typcol].astype(str)
    daystats={}
    for (person,day),group in df.groupby(['_name',df['_dt'].dt.date]):
        punches=sorted(group['_dt'].tolist()); daystats.setdefault(person,{})[day]=punches
    people=[]
    names = roster[rname].dropna().astype(str).str.strip()
    if not names.ne('').any():
        raise ValueError('员工名册中没有可识别的姓名。')
    for person in names[names.ne('')].unique():
        records=daystats.get(person,{})
        present=sum(1 for d in workdays if records.get(d))
        late=sum(1 for d in workdays if records.get(d) and min(records[d]).time()>time(8,30))
        early=sum(1 for d in workdays if records.get(d) and max(records[d]).time()<time(18,0))
        missing=sum(1 for d in workdays if not records.get(d))
        people.append({'name':person,'present':present,'late':late,'early':early,'missing':missing,'workdays':len(workdays),'punches':sum(len(x) for x in records.values())})
    people.sort(key=lambda x:x['name'])
    return people, len(workdays), len(df), leave.shape[0], trip.shape[0], start, end

@app.route('/', methods=['GET','POST'])
def index():
    return render_template('index.html')

@app.post('/process')
def process_upload():
    try:
        keys = ('attendance', 'roster', 'leave', 'trip')
        files = {key: request.files.get(key) for key in keys}
        if not all(files.values()):
            return jsonify(error='请补齐考勤记录、员工清单、请假审批和出差审批。'), 400
        if any(not f.filename.lower().endswith(('.xls', '.xlsx')) for f in files.values()):
            return jsonify(error='请选择 .xls 或 .xlsx 格式的 Excel 文件。'), 400
        people,wd,events,leave_n,trip_n,start,end=process(files)
        export_df=pd.DataFrame([{'员工姓名':p['name'],'应出勤工作日':p['workdays'],'出勤日':p['present'],'迟到次数':p['late'],'早退次数':p['early'],'无打卡工作日':p['missing'],'打卡事件数':p['punches']} for p in people])
        out=io.BytesIO()
        with pd.ExcelWriter(out,engine='openpyxl') as writer:
            export_df.to_excel(writer,index=False,sheet_name='考勤汇总')
            pd.DataFrame([{'统计起始日期':str(start),'统计结束日期':str(end),'员工人数':len(people),'工作日数':wd,'打卡事件数':events,'审批表记录数':leave_n+trip_n,'统计规则':'工作日 08:30-18:00；周末不排班；审批不改变打卡统计'}]).to_excel(writer,index=False,sheet_name='统计说明')
            for sheet in writer.sheets.values():
                sheet.freeze_panes = 'A2'
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.data_type == 'f':
                            cell.data_type = 's'
        out.seek(0)
        response = send_file(out,as_attachment=True,download_name='考勤汇总表.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response.cache_control.no_store = True
        return response
    except ValueError as error:
        if str(error) in {'考勤记录中没有可识别的姓名和打卡时间。', '员工名册中没有可识别的姓名。', '无法识别姓名或时间列，请确认上传的是原始考勤表和员工名册。'}:
            return jsonify(error=str(error)), 400
        return jsonify(error='无法读取表格，请确认文件格式和内容。'), 400
    except Exception:
        return jsonify(error='文件整理失败，请确认表格可以正常打开后重试。'), 400

@app.errorhandler(413)
def upload_too_large(error):
    return jsonify(error='上传文件合计超过 32 MB，请缩小文件后重试。'), 413

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','8080')),debug=False)
