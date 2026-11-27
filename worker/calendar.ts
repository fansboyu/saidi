// Dates announced by the General Office of the State Council for mainland China.
// To add a year, record the official holiday spans and weekend make-up days here.
type Holiday = readonly [start: string, end: string, name: string];
type AnnualCalendar = { holidays: readonly Holiday[]; makeUp: readonly string[]; source: string };

const schedules: Record<string, AnnualCalendar> = {
  "2026": {
    source: "https://www.beijing.gov.cn/cs/gncs/zcwj/202603/t20260327_4568275.html",
    holidays: [
      ["2026-01-01", "2026-01-03", "元旦"],
      ["2026-02-15", "2026-02-23", "春节"],
      ["2026-04-04", "2026-04-06", "清明节"],
      ["2026-05-01", "2026-05-05", "劳动节"],
      ["2026-06-19", "2026-06-21", "端午节"],
      ["2026-09-25", "2026-09-27", "中秋节"],
      ["2026-10-01", "2026-10-07", "国庆节"],
    ],
    makeUp: ["2026-01-04", "2026-02-14", "2026-02-28", "2026-05-09", "2026-09-20", "2026-10-10"],
  },
  "2027": {
    source: "https://www.forestry.gov.cn/lyj/1/szxx/20261105/695102.html",
    holidays: [
      ["2027-01-01", "2027-01-03", "元旦"],
      ["2027-02-05", "2027-02-13", "春节"],
      ["2027-04-03", "2027-04-05", "清明节"],
      ["2027-05-01", "2027-05-05", "劳动节"],
      ["2027-06-09", "2027-06-09", "端午节"],
      ["2027-09-15", "2027-09-15", "中秋节"],
      ["2027-10-01", "2027-10-07", "国庆节"],
    ],
    makeUp: ["2027-01-31", "2027-02-14", "2027-05-08", "2027-09-26", "2027-10-09"],
  },
};

export function classifyAttendanceDay(date: string) {
  const schedule = schedules[date.slice(0, 4)];
  const holiday = schedule?.holidays.find(([first, last]) => first <= date && date <= last);
  const weekday = new Date(date + "T00:00:00Z").getUTCDay();
  const weekend = weekday === 0 || weekday === 6;
  const makeUp = schedule?.makeUp.includes(date) ?? false;
  return {
    isWorkday: makeUp || (!holiday && !weekend),
    type: makeUp ? "调休补班" : holiday ? "节假日休息" : weekend ? "周末休息" : "工作日",
    holiday: holiday?.[2] ?? "",
    source: schedule?.source ?? "",
  };
}
