import fs from "node:fs";
import assert from "node:assert/strict";

const output = (key, value) => {
  if (process.env.GITHUB_OUTPUT) fs.appendFileSync(process.env.GITHUB_OUTPUT, `${key}=${String(value).replaceAll("\n", " ")}\n`);
};
const log = (message) => fs.appendFileSync(process.env.AUTOMATION_LOG || "automation.log", `${new Date().toISOString()} ${message}\n`);

function dateDetails(value = new Date()) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat("en", {
    timeZone: "Asia/Kuwait", year: "numeric", month: "2-digit", day: "2-digit", weekday: "short",
  }).formatToParts(value).filter(({ type }) => type !== "literal").map(({ type, value: part }) => [type, part]));
  return { date: `${parts.year}-${parts.month}-${parts.day}`, weekday: parts.weekday };
}

function holidayOn(calendar, date) {
  const compactDate = date.replaceAll("-", "");
  return calendar.replace(/\r?\n[ \t]/g, "").split("BEGIN:VEVENT").slice(1).map((event) => ({
    date: event.match(/DTSTART(?:;VALUE=DATE)?:([0-9]{8})/)?.[1],
    name: event.match(/\nSUMMARY(?:;[^:]*)?:(.*)/)?.[1]?.trim() || "Kuwait public holiday",
  })).find((event) => event.date === compactDate);
}

if (process.argv.includes("--self-test")) {
  assert.deepEqual(dateDetails(new Date("2026-08-07T07:00:00Z")), { date: "2026-08-07", weekday: "Fri" });
  assert.equal(holidayOn("BEGIN:VEVENT\nDTSTART;VALUE=DATE:20260225\nSUMMARY:National Day\nEND:VEVENT", "2026-02-25")?.name, "National Day");
  assert.equal(holidayOn("BEGIN:VEVENT\nDTSTART;VALUE=DATE:20260225\nSUMMARY:National Day\nEND:VEVENT", "2026-02-24"), undefined);
  console.log("Kuwait business-day checks passed.");
  process.exit(0);
}

const requestedDate = process.env.CHECK_DATE;
const current = dateDetails(requestedDate ? new Date(`${requestedDate}T00:00:00+03:00`) : new Date());
output("date", current.date);

if (["Fri", "Sat"].includes(current.weekday)) {
  output("decision", "skip");
  output("reason", `${current.date} is a Kuwait weekend (${current.weekday}).`);
  log(`Schedule skipped: Kuwait weekend (${current.date}, ${current.weekday})`);
  process.exit(0);
}

const overrides = (process.env.KUWAIT_HOLIDAY_DATES || "").split(",").map((value) => value.trim()).filter(Boolean);
if (overrides.includes(current.date)) {
  output("decision", "skip");
  output("reason", `${current.date} is listed in KUWAIT_HOLIDAY_DATES.`);
  log(`Schedule skipped: configured Kuwait public holiday (${current.date})`);
  process.exit(0);
}

try {
  const response = await fetch(process.env.KUWAIT_HOLIDAY_CALENDAR_URL, { signal: AbortSignal.timeout(15000) });
  if (!response.ok) throw new Error(`holiday calendar returned HTTP ${response.status}`);
  const calendar = await response.text();
  if (!calendar.includes("BEGIN:VCALENDAR") || !calendar.includes("BEGIN:VEVENT")) {
    throw new Error("holiday calendar did not contain valid iCalendar events");
  }
  fs.writeFileSync("kuwait-holidays.ics", calendar);
  const holiday = holidayOn(calendar, current.date);
  if (holiday) {
    output("decision", "skip");
    output("reason", `${current.date} is ${holiday.name}.`);
    log(`Schedule skipped: ${holiday.name} (${current.date})`);
  } else {
    output("decision", "run");
    output("reason", `${current.date} is a Kuwait business day.`);
    log(`Schedule approved: Kuwait business day (${current.date})`);
  }
} catch (error) {
  output("decision", "error");
  output("error", `Kuwait holiday validation failed: ${error.message}`);
  log(`Schedule blocked: ${error.message}`);
  process.exitCode = 1;
}
