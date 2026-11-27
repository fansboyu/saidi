import * as XLSX from "xlsx";
import { classifyAttendanceDay } from "./calendar.ts";

const FILES = ["attendance", "roster", "leave", "trip"] as const;
const MAX_UPLOAD = 32 * 1024 * 1024;
const DAY = 86_400_000;
const MINUTE = 60_000;
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const NAME_HEADERS = ["姓名", "员工姓名", "人员姓名", "员工"];
const TIME_HEADERS = ["事件时间", "打卡时间", "时间", "日期"];
type Cell = string | number | boolean | Date | null | undefined;
type InputFiles = Record<(typeof FILES)[number], File>;

class InputError extends Error {}

function errorResponse(message: string, status = 400) {
  return Response.json({ error: message }, { status, headers: { "Cache-Control": "no-store" } });
}

async function firstSheet(file: File): Promise<Cell[][]> {
  try {
    const buffer = await file.arrayBuffer();
    const prefix = new Uint8Array(buffer, 0, Math.min(buffer.byteLength, 8));
    const zip = prefix[0] === 0x50 && prefix[1] === 0x4b && prefix[2] === 0x03 && prefix[3] === 0x04;
    const ole = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((byte, index) => prefix[index] === byte);
    const biff = prefix[0] === 0x09 && [0x00, 0x02, 0x04, 0x08].includes(prefix[1]);
    if (file.name.toLowerCase().endsWith(".xlsx") ? !zip : !ole && !biff) {
      throw new InputError("无法读取表格，请确认文件格式和内容。");
    }
    const workbook = XLSX.read(buffer, {
      sheets: 0,
      cellDates: true,
      cellHTML: false,
      cellText: false,
      dense: true,
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json<Cell[]>(sheet, { header: 1, raw: true, defval: null });
  } catch {
    throw new InputError("无法读取表格，请确认文件格式和内容。");
  }
}

function column(header: Cell[] | undefined, labels: string[]) {
  return header?.findIndex((cell) => labels.some((label) => String(cell ?? "").trim().includes(label))) ?? -1;
}

function timestamp(value: Cell) {
  if (value instanceof Date) return value.getTime();
  if (typeof value !== "string") return NaN;
  // Time values without an offset represent local wall-clock time in the source workbook.
  const match = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/.exec(value.trim());
  if (!match) return NaN;
  const [, year, month, day, hour = "0", minute = "0", second = "0", fraction = "0"] = match;
  const utc = Date.UTC(+year, +month - 1, +day, +hour, +minute, +second, +fraction.padEnd(3, "0"));
  const checked = new Date(utc);
  return checked.getUTCFullYear() === +year && checked.getUTCMonth() + 1 === +month &&
    checked.getUTCDate() === +day && +hour < 24 && +minute < 60 && +second < 60 ? utc : NaN;
}

const CHECK_IN = 510 * MINUTE;
const CHECK_OUT = 1020 * MINUTE;
const SINGLE_PUNCH_SPLIT = 720 * MINUTE;
const MAX_REPORT_DAYS = 366;
const MAX_DETAIL_ROWS = 40_000;

const DAILY = ["姓名", "日期", "上班时间", "下班时间", "正常", "迟到", "早退", "缺勤", "上班缺卡", "下班缺卡", "加班时长", "是否异常"];
const DETAIL = [...DAILY, "请假类型", "出差类型", "月份", "备注", "打卡次数", "日期类型", "节假日", "加班类型"];
const MONTHLY = ["姓名", "正常", "迟到", "早退", "缺勤", "上班缺卡", "下班缺卡", "加班时长", "异常", "请假天数", "出差天数", "工作日", "月份", "出勤日", "打卡事件数", "工作日加班小时", "周末加班小时", "节假日加班小时", "休息日时长待核天数"];

type Punches = { first: number; last: number; count: number };
type SheetRow = (string | number)[];
type ApprovalKind = "leave" | "trip";

const approvalFields = {
  leave: {
    label: "请假",
    name: ["申请人", "请假人", "姓名"],
    type: ["请假类型"],
    actualStart: ["实际休假开始日期", "实际请假开始日期"],
    actualEnd: ["实际休假结束日期", "实际请假结束日期"],
    planStart: ["预计请假开始日期", "请假开始日期"],
    planEnd: ["预计请假结束日期", "请假结束日期"],
  },
  trip: {
    label: "出差",
    name: ["出差人", "姓名", "申请人"],
    type: ["出差类型"],
    actualStart: ["实际出发日期"],
    actualEnd: ["实际返回日期"],
    planStart: ["预计出发日期", "出差开始日期"],
    planEnd: ["预计返回日期", "出差结束日期"],
  },
};

function text(value: Cell) {
  return String(value ?? "").trim();
}

function dateText(day: number) {
  return new Date(day * DAY).toISOString().slice(0, 10);
}

function timeText(time: number | undefined) {
  return time === undefined ? "" : new Date(time).toISOString().slice(11, 19);
}

function approvalRecords(rows: Cell[][], kind: ApprovalKind, start: number, end: number) {
  const fields = approvalFields[kind];
  const header = (rows[0] ?? []).map(text);
  function find(labels: string[]) {
    for (const label of labels) {
      const index = header.findIndex((value) => value === label || value.startsWith(label + "（") || value.startsWith(label + "("));
      if (index >= 0) return index;
    }
    return -1;
  }
  const nameCol = find(fields.name);
  const typeCol = find(fields.type);
  const actualStart = find(fields.actualStart);
  const actualEnd = find(fields.actualEnd);
  const planStart = find(fields.planStart);
  const planEnd = find(fields.planEnd);
  const dateCol = find(["日期", fields.label + "日期"]);
  const statusCol = find(["流程状态", "审批状态", "审批结果"]);
  const cancelCol = find(["是否临时取消出差"]);
  const cancelFlag = header.indexOf("是否临时取消出差*");
  if (nameCol < 0 || typeCol < 0 || ((actualStart < 0 || actualEnd < 0) && (planStart < 0 || planEnd < 0) && dateCol < 0)) {
    throw new InputError("无法识别" + fields.label + "表的姓名、类型或日期列，请上传原始审批表或按日记录表。");
  }

  const byName = new Map<string, Map<number, Set<string>>>();
  let skipped = 0;
  for (const row of rows.slice(1)) {
    const name = text(row[nameCol]);
    if (!name && !row.some((value) => text(value))) continue;
    const state = text(row[statusCol]);
    if ((state && !["归档", "已归档", "已完成", "已通过", "通过", "审批通过", "已审批"].includes(state)) ||
        (kind === "trip" && (text(row[cancelCol]).startsWith("取消") || ["1", "是"].includes(text(row[cancelFlag]))))) {
      skipped++;
      continue;
    }
    const rawType = text(row[typeCol]);
    let from = timestamp(row[actualStart]);
    let to = timestamp(row[actualEnd]);
    if (!Number.isFinite(from) || !Number.isFinite(to)) {
      from = timestamp(row[planStart] ?? row[dateCol]);
      to = timestamp(row[planEnd] ?? row[dateCol]);
    }
    if (!name || !rawType || !Number.isFinite(from) || !Number.isFinite(to) || to < from) {
      skipped++;
      continue;
    }
    const firstDay = Math.max(start, Math.floor(from / DAY));
    const lastDay = Math.min(end, Math.floor(to / DAY));
    if (firstDay > lastDay) continue;
    let days = byName.get(name);
    if (!days) byName.set(name, days = new Map());
    for (let day = firstDay; day <= lastDay; day++) {
      let types = days.get(day);
      if (!types) days.set(day, types = new Set());
      types.add(rawType);
    }
  }

  const records: SheetRow[] = [["姓名", "日期", fields.label + "类型"]];
  for (const [name, days] of byName) {
    for (const [day, types] of days) records.push([name, dateText(day), [...types].join("、")]);
  }
  return { byName, records, skipped };
}

function dayStatus(punches: Punches | undefined, exempt: boolean) {
  const paired = punches !== undefined && punches.first !== punches.last;
  // A single morning punch supplies the start; a single afternoon punch supplies the end.
  const first = punches && (paired || punches.first < SINGLE_PUNCH_SPLIT) ? punches.first : undefined;
  const last = punches && (paired || punches.last >= SINGLE_PUNCH_SPLIT) ? punches.last : undefined;
  const late = !exempt && paired ? Math.ceil(Math.max(0, punches.first - CHECK_IN) / MINUTE) : 0;
  const early = !exempt && paired ? Math.ceil(Math.max(0, CHECK_OUT - punches.last) / MINUTE) : 0;
  const missing = !exempt && !punches;
  const noStart = !exempt && Boolean(punches) && first === undefined;
  const noEnd = !exempt && Boolean(punches) && last === undefined;
  const overtime = paired ? Math.max(0, punches.last - CHECK_OUT) : 0;
  const abnormal = Boolean(late || early || missing || noStart || noEnd);
  return { first, last, normal: !exempt && !abnormal, late, early, missing, noStart, noEnd, overtime, abnormal };
}

function restDayStatus(punches: Punches) {
  const span = punches.last - punches.first;
  return {
    ...dayStatus(undefined, true),
    first: punches.first,
    last: span > 0 ? punches.last : undefined,
    overtime: span,
  };
}

function dailyRow(name: string, date: string, status: ReturnType<typeof dayStatus>): SheetRow {
  return [
    name, date, timeText(status.first), timeText(status.last),
    status.normal ? "是" : "", status.late || "", status.early || "",
    status.missing ? "缺勤" : "", status.noStart ? "是" : "", status.noEnd ? "是" : "",
    status.overtime ? Math.round(status.overtime / 360_000) / 10 : "", status.abnormal ? "是" : "",
  ];
}

function appendSheet(workbook: XLSX.WorkBook, name: string, rows: SheetRow[], widths: number[]) {
  const sheet = XLSX.utils.aoa_to_sheet(rows.map((row) => row.map((value) => value === "" ? undefined : value)), { dense: true });
  sheet["!cols"] = widths.map((wch) => ({ wch }));
  if (name !== "统计说明") sheet["!autofilter"] = { ref: sheet["!ref"]! };
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

async function makeReport(files: InputFiles) {
  const roster = await firstSheet(files.roster);
  const rosterColumn = column(roster[0], NAME_HEADERS);
  const attendance = await firstSheet(files.attendance);
  const nameColumn = column(attendance[0], NAME_HEADERS);
  const timeColumn = column(attendance[0], TIME_HEADERS);
  if (rosterColumn < 0 || nameColumn < 0 || timeColumn < 0) {
    throw new InputError("无法识别姓名或时间列，请确认上传的是原始考勤表和员工名册。");
  }

  const names = [...new Set(roster.slice(1).map((row) => String(row[rosterColumn] ?? "").trim()).filter(Boolean))].sort();
  if (!names.length) throw new InputError("员工名册中没有可识别的姓名。");
  const byName = new Map<string, Map<number, Punches>>();
  let start = Infinity;
  let end = -Infinity;
  let events = 0;
  for (let index = 1; index < attendance.length; index++) {
    const row = attendance[index];
    const name = String(row[nameColumn] ?? "").trim();
    const instant = timestamp(row[timeColumn]);
    if (!name || !Number.isFinite(instant)) continue;
    const day = Math.floor(instant / DAY);
    start = Math.min(start, day);
    end = Math.max(end, day);
    events++;
    let days = byName.get(name);
    if (!days) byName.set(name, days = new Map());
    const time = instant - day * DAY;
    const punches = days.get(day);
    if (punches) {
      punches.first = Math.min(time, punches.first);
      punches.last = Math.max(time, punches.last);
      punches.count++;
    } else {
      days.set(day, { first: time, last: time, count: 1 });
    }
  }
  if (!events) throw new InputError("考勤记录中没有可识别的姓名和打卡时间。");
  const range = end - start + 1;
  if (range > MAX_REPORT_DAYS) throw new InputError("打卡范围超过 366 天，请按月或按季度拆分后上传。");
  const calendar = Array.from({ length: range }, (_, offset) => {
    const day = start + offset;
    const date = dateText(day);
    return { day, date, month: date.slice(0, 7), ...classifyAttendanceDay(date) };
  });
  const workdays = calendar.filter((entry) => entry.isWorkday);
  if (names.length * workdays.length > MAX_DETAIL_ROWS) {
    throw new InputError("逐日明细超过 40,000 行，请缩小日期范围或拆分名册。");
  }

  const leave = approvalRecords(await firstSheet(files.leave), "leave", start, end);
  const trip = approvalRecords(await firstSheet(files.trip), "trip", start, end);
  const middle: SheetRow[] = [DAILY];
  for (const [name, dates] of byName) {
    for (const [day, punches] of dates) {
      if (middle.length > MAX_DETAIL_ROWS) throw new InputError("打卡日明细超过 40,000 行，请分段上传。");
      const entry = calendar[day - start];
      middle.push(dailyRow(name, entry.date, entry.isWorkday ? dayStatus(punches, false) : restDayStatus(punches)));
    }
  }
  middle.splice(1, middle.length - 1, ...middle.slice(1).sort((a, b) =>
    String(a[1]).localeCompare(String(b[1])) || String(a[0]).localeCompare(String(b[0]))));

  const detail: SheetRow[] = [DETAIL];
  const anomalies: SheetRow[] = [DETAIL];
  const restOvertime: SheetRow[] = [DETAIL];
  const monthly: SheetRow[] = [MONTHLY];
  for (const name of names) {
    const punches = byName.get(name);
    const leaveDays = leave.byName.get(name);
    const tripDays = trip.byName.get(name);
    const months = new Map<string, {
      normal: number; late: number; early: number; missing: number; noStart: number; noEnd: number;
      overtime: number; abnormal: number; leave: number; trip: number; workdays: number; present: number; punches: number;
      weekendOvertime: number; holidayOvertime: number; restPending: number;
    }>();
    for (const { day, date, month, type, holiday, isWorkday } of calendar) {
      let totals = months.get(month);
      if (!totals) {
        totals = { normal: 0, late: 0, early: 0, missing: 0, noStart: 0, noEnd: 0, overtime: 0,
          abnormal: 0, leave: 0, trip: 0, workdays: 0, present: 0, punches: 0,
          weekendOvertime: 0, holidayOvertime: 0, restPending: 0 };
        months.set(month, totals);
      }
      const record = punches?.get(day);
      if (!isWorkday && !record) continue;
      if (detail.length > MAX_DETAIL_ROWS) throw new InputError("逐日明细超过 40,000 行，请缩小日期范围或拆分名册。");
      const leaveType = [...(leaveDays?.get(day) ?? [])].join("、");
      const tripType = [...(tripDays?.get(day) ?? [])].join("、");
      if (!isWorkday && record) {
        const status = restDayStatus(record);
        const pending = status.overtime === 0;
        const overtimeType = holiday ? "节假日加班" : "周末加班";
        const row = [...dailyRow(name, date, status), leaveType, tripType, month,
          pending ? "仅一次不同时间打卡，时长待核" : "", record.count, type, holiday, overtimeType];
        detail.push(row);
        restOvertime.push(row);
        if (holiday) totals.holidayOvertime += status.overtime;
        else totals.weekendOvertime += status.overtime;
        totals.restPending += Number(pending);
        continue;
      }
      const status = dayStatus(record, Boolean(leaveType || tripType));
      const row = [...dailyRow(name, date, status), leaveType, tripType, month, "", record?.count ?? 0,
        type, holiday, status.overtime ? "工作日加班" : ""];
      detail.push(row);
      if (status.abnormal) anomalies.push(row);

      totals.normal += Number(status.normal);
      totals.late += Number(status.late > 0);
      totals.early += Number(status.early > 0);
      totals.missing += Number(status.missing);
      totals.noStart += Number(status.noStart);
      totals.noEnd += Number(status.noEnd);
      totals.overtime += status.overtime;
      totals.abnormal += Number(status.abnormal);
      totals.leave += Number(Boolean(leaveType));
      totals.trip += Number(Boolean(tripType));
      totals.workdays++;
      totals.present += Number(Boolean(record));
      totals.punches += record?.count ?? 0;
    }
    for (const [month, totals] of months) {
      monthly.push([name, totals.normal, totals.late, totals.early, totals.missing, totals.noStart, totals.noEnd,
        Math.round((totals.overtime + totals.weekendOvertime + totals.holidayOvertime) / 360_000) / 10,
        totals.abnormal, totals.leave, totals.trip, totals.workdays, month, totals.present, totals.punches,
        Math.round(totals.overtime / 360_000) / 10, Math.round(totals.weekendOvertime / 360_000) / 10,
        Math.round(totals.holidayOvertime / 360_000) / 10, totals.restPending]);
    }
  }

  const report = XLSX.utils.book_new();
  const dailyWidths = [16, 13, 13, 13, 9, 10, 10, 10, 12, 12, 14, 12];
  const detailWidths = [...dailyWidths, 20, 20, 12, 32, 12, 15, 13, 16];
  appendSheet(report, "请假记录", leave.records, [16, 13, 24]);
  appendSheet(report, "出差记录", trip.records, [16, 13, 24]);
  appendSheet(report, "中间考勤", middle, dailyWidths);
  appendSheet(report, "汇总明细", detail, detailWidths);
  appendSheet(report, "考勤汇总统计", monthly, [16, 9, 9, 9, 9, 12, 12, 14, 9, 12, 12, 11, 12, 10, 14, 18, 17, 18, 22]);
  appendSheet(report, "异常考勤", anomalies, detailWidths);
  appendSheet(report, "休息日加班", restOvertime, detailWidths);
  appendSheet(report, "考勤日历", [
    ["日期", "日期类型", "节假日", "是否应出勤", "日历来源"],
    ...calendar.map(({ date, type, holiday, isWorkday, source }) =>
      [date, type, holiday, isWorkday ? "是" : "否", source || "该年度未配置，按普通周历"]),
  ], [13, 16, 14, 15, 86]);
  appendSheet(report, "统计说明", [
    ["项目", "口径或数值"],
    ["统计日期", dateText(start) + " 至 " + dateText(end)],
    ["员工人数", names.length],
    ["统计工作日", workdays.length],
    ["打卡事件数（含周末及名册外人员）", events],
    ["未采用的审批行数（含未完成、取消或资料不齐）", leave.skipped + trip.skipped],
    ["工作日口径", "在打卡首末日期内优先使用国务院办公厅年度放假及补班安排，其余日期按周一至周五。未配置的年份按普通周历，详见考勤日历。"],
    ["班次口径", "08:30–17:00。同日最早、最晚的不同时间作为首末打卡；只有一次时，12:00 前记下班缺卡，12:00 起记上班缺卡。"],
    ["迟到、早退、加班", "逐日迟到和早退为分钟，向上取整。加班为双卡当日 17:00 后的分钟折算小时，保留一位小数；月汇总先累加再四舍五入。"],
    ["审批口径", "完整的实际日期优先，其次预计日期；已取消或明确未完成的流程不采用。无审批状态列的表按已确认记录处理。"],
    ["审批免计异常", "请假或出差覆盖到的工作日保留打卡时间，免计迟到、早退、缺勤和缺卡，不计正常日。请假和出差天数分别按覆盖的工作日去重。"],
    ["休息日口径", "周末和节假日无打卡为正常休息，不计应出勤或异常。有两次不同打卡取首末时间间隔，只有一个时间则标记时长待核。休息日加班时长供复核。"],
    ["工作表关系", "中间考勤仅列原始打卡涉及的人日，未合并审批。汇总明细包含名册工作日及休息日有打卡的记录。月度出勤日、打卡事件数、请假和出差天数只统计工作日。"],
    ["核对备注", "当前来源不含轻推或人工核查结果，备注列供后续填写。"],
  ], [46, 111]);
  return XLSX.write(report, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
}

export default {
  async fetch(request: Request): Promise<Response> {
    if (new URL(request.url).pathname !== "/process") return new Response("Not found", { status: 404 });
    if (request.method !== "POST") return new Response(null, { status: 405, headers: { Allow: "POST" } });
    const tooLarge = () => errorResponse("上传文件合计超过 32 MB，请缩小文件后重试。", 413);
    if (Number(request.headers.get("Content-Length")) > MAX_UPLOAD) return tooLarge();
    try {
      const form = await request.formData();
      const entries = FILES.map((key) => form.get(key));
      if (!entries.every((file): file is File => file !== null && typeof file !== "string" && Boolean(file.name))) {
        return errorResponse("请补齐考勤记录、员工清单、请假审批和出差审批。");
      }
      if (entries.reduce((size, file) => size + file.size, 0) > MAX_UPLOAD) return tooLarge();
      if (!entries.every((file) => /\.xlsx?$/i.test(file.name))) {
        return errorResponse("请选择 .xls 或 .xlsx 格式的 Excel 文件。");
      }
      const files = Object.fromEntries(FILES.map((key, index) => [key, entries[index]])) as InputFiles;
      const body = await makeReport(files);
      return new Response(body, { headers: {
        "Content-Type": XLSX_TYPE,
        "Content-Disposition": "attachment; filename*=UTF-8''" + encodeURIComponent("考勤汇总表.xlsx"),
        "Cache-Control": "no-store",
      } });
    } catch (error) {
      return errorResponse(error instanceof InputError ? error.message : "文件整理失败，请确认表格可以正常打开后重试。");
    }
  },
};
