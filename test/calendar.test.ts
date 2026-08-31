import ICAL from "ical.js";
import { describe, expect, it } from "vitest";

import {
  CalendarTransformError,
  transformCalendar as transformCalendarSource,
} from "../src/calendar";

const NOW = new Date("2026-08-31T12:00:00Z");
const HEADER = [
  "BEGIN:VCALENDAR",
  "PRODID:-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN",
  "VERSION:2.0",
  "METHOD:PUBLISH",
];

function calendar(...lines: string[]): string {
  return [...HEADER, ...lines, "END:VCALENDAR", ""].join("\r\n");
}

function timedEvent(
  uid: string,
  busyStatus: string | undefined,
  extra: string[] = [],
  start = "20260826T160000Z",
  end = "20260826T170000Z",
): string[] {
  return [
    "BEGIN:VEVENT",
    `UID:${uid}`,
    "DTSTAMP:20260825T160000Z",
    `DTSTART:${start}`,
    `DTEND:${end}`,
    ...(busyStatus === undefined
      ? []
      : [`X-MICROSOFT-CDO-BUSYSTATUS:${busyStatus}`]),
    ...extra,
    "END:VEVENT",
  ];
}

function transformCalendar(source: string, eventTitle: string): string {
  return transformCalendarSource(source, eventTitle, { now: NOW });
}

describe("transformCalendar", () => {
  it("retains only explicit Busy events and rewrites their summaries", () => {
    const source = calendar(
      ...timedEvent("busy-upper", "BUSY"),
      ...timedEvent(
        "busy-lower",
        "busy",
        [],
        "20260826T180000Z",
        "20260826T190000Z",
      ),
      ...timedEvent("free", "FREE"),
      ...timedEvent("tentative", "TENTATIVE"),
      ...timedEvent("away", "AWAY"),
      ...timedEvent("oof", "OOF"),
      ...timedEvent("unknown", "WORKINGELSEWHERE"),
      ...timedEvent("missing", undefined, ["TRANSP:OPAQUE"]),
    );

    const output = transformCalendar(source, "  Alice Mtg  ");

    expect(output.match(/BEGIN:VEVENT/g)).toHaveLength(2);
    expect(output).toContain("DTSTART:20260826T160000Z");
    expect(output).toContain("DTSTART:20260826T180000Z");
    expect(output).not.toContain("UID:busy-upper");
    expect(output).not.toContain("UID:busy-lower");
    expect(output).not.toContain("UID:free");
    expect(output).not.toContain("UID:tentative");
    expect(output).not.toContain("UID:away");
    expect(output).not.toContain("UID:oof");
    expect(output).not.toContain("UID:unknown");
    expect(output).not.toContain("UID:missing");
    expect(output.match(/SUMMARY:Alice Mtg/g)).toHaveLength(2);
    expect(() => ICAL.parse(output)).not.toThrow();
  });

  it("drops cancelled, date-valued, and Outlook-marked all-day events", () => {
    const source = calendar(
      ...timedEvent("cancelled", "BUSY", ["STATUS:CANCELLED"]),
      "BEGIN:VEVENT",
      "UID:date-valued",
      "DTSTAMP:20260825T160000Z",
      "DTSTART;VALUE=DATE:20260826",
      "DTEND;VALUE=DATE:20260827",
      "X-MICROSOFT-CDO-BUSYSTATUS:BUSY",
      "END:VEVENT",
      ...timedEvent("outlook-all-day", "BUSY", [
        "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE",
      ]),
      "BEGIN:VEVENT",
      "UID:midnight-timed",
      "DTSTAMP:20260825T160000Z",
      "DTSTART:20260827T000000Z",
      "DTEND:20260828T000000Z",
      "X-MICROSOFT-CDO-BUSYSTATUS:BUSY",
      "END:VEVENT",
    );

    const output = transformCalendar(source, "Alice Mtg");

    expect(output.match(/BEGIN:VEVENT/g)).toHaveLength(1);
    expect(output).toContain("DTSTART:20260827T000000Z");
    expect(output).toContain("DTEND:20260828T000000Z");
  });

  it("suppresses excluded recurrence overrides on a retained Busy series", () => {
    const source = calendar(
      ...timedEvent("series", "BUSY", ["RRULE:FREQ=DAILY;COUNT=4"]),
      ...timedEvent("series", "FREE", [
        "RECURRENCE-ID:20260827T160000Z",
      ]),
      ...timedEvent("series", "TENTATIVE", [
        "RECURRENCE-ID:20260828T160000Z",
      ]),
      ...timedEvent("series", "BUSY", [
        "RECURRENCE-ID:20260829T160000Z",
        "STATUS:CANCELLED",
      ]),
    );

    const output = transformCalendar(source, "Alice Mtg");

    expect(output.match(/BEGIN:VEVENT/g)).toHaveLength(1);
    expect(output).toContain("DTSTART:20260826T160000Z");
    expect(output).not.toContain("20260827T160000Z");
    expect(output).not.toContain("20260828T160000Z");
    expect(output).not.toContain("20260829T160000Z");
  });

  it("removes private metadata while expanding recurrence", () => {
    const source = calendar(
      ...timedEvent("recurring", "BUSY", [
        "SUMMARY:Secret roadmap review",
        "DESCRIPTION:Confidential details",
        "LOCATION:Secret room",
        "ORGANIZER;CN=Boss:mailto:boss@example.com",
        "ATTENDEE;CN=Colleague:mailto:colleague@example.com",
        "URL:https://meet.example.com/secret",
        "ATTACH:https://example.com/attachment",
        "CATEGORIES:Confidential",
        "CLASS:PRIVATE",
        "RRULE:FREQ=WEEKLY;COUNT=3",
        "EXDATE:20260902T160000Z",
        "SEQUENCE:4",
        "LAST-MODIFIED:20260825T170000Z",
        "BEGIN:VALARM",
        "ACTION:DISPLAY",
        "TRIGGER:-PT15M",
        "DESCRIPTION:Secret reminder",
        "END:VALARM",
      ]),
      ...timedEvent(
        "recurring",
        "BUSY",
        [
          "RECURRENCE-ID:20260909T160000Z",
          "SUMMARY:Changed secret title",
        ],
        "20260909T180000Z",
        "20260909T190000Z",
      ),
    );

    const output = transformCalendar(source, "Alice Mtg");

    expect(output.match(/BEGIN:VEVENT/g)).toHaveLength(2);
    expect(output).toContain("DTSTART:20260826T160000Z");
    expect(output).toContain("DTSTART:20260909T180000Z");
    for (const privateValue of [
      "Secret roadmap review",
      "Confidential details",
      "Secret room",
      "boss@example.com",
      "colleague@example.com",
      "meet.example.com",
      "attachment",
      "Confidential",
      "VALARM",
      "Secret reminder",
      "X-MICROSOFT-CDO-BUSYSTATUS",
      "RRULE",
      "EXDATE",
      "RECURRENCE-ID",
      "SEQUENCE",
      "LAST-MODIFIED",
    ]) {
      expect(output).not.toContain(privateValue);
    }
  });

  it("resolves timezone-aware events to UTC", () => {
    const source = calendar(
      "BEGIN:VTIMEZONE",
      "TZID:America/Denver",
      "BEGIN:DAYLIGHT",
      "DTSTART:19700308T020000",
      "TZOFFSETFROM:-0700",
      "TZOFFSETTO:-0600",
      "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU",
      "END:DAYLIGHT",
      "BEGIN:STANDARD",
      "DTSTART:19701101T020000",
      "TZOFFSETFROM:-0600",
      "TZOFFSETTO:-0700",
      "RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU",
      "END:STANDARD",
      "END:VTIMEZONE",
      "BEGIN:VTIMEZONE",
      "TZID:America/New_York",
      "BEGIN:STANDARD",
      "DTSTART:19701101T020000",
      "TZOFFSETFROM:-0400",
      "TZOFFSETTO:-0500",
      "END:STANDARD",
      "END:VTIMEZONE",
      "BEGIN:VEVENT",
      "UID:denver",
      "DTSTAMP:20260825T160000Z",
      "DTSTART;TZID=America/Denver:20260826T100000",
      "DTEND;TZID=America/Denver:20260826T110000",
      "X-MICROSOFT-CDO-BUSYSTATUS:BUSY",
      "END:VEVENT",
    );

    const output = transformCalendar(source, "Alice Mtg");

    expect(output).toContain("DTSTART:20260826T160000Z");
    expect(output).toContain("DTEND:20260826T170000Z");
    expect(output).not.toContain("BEGIN:VTIMEZONE");
    expect(output).not.toContain("TZID:America/Denver");
    expect(output).not.toContain("TZID:America/New_York");
    expect(() => ICAL.parse(output)).not.toThrow();
  });

  it("merges overlapping and back-to-back meetings into one block", () => {
    const source = calendar(
      ...timedEvent(
        "first",
        "BUSY",
        [],
        "20260826T160000Z",
        "20260826T170000Z",
      ),
      ...timedEvent(
        "overlapping",
        "BUSY",
        [],
        "20260826T163000Z",
        "20260826T180000Z",
      ),
      ...timedEvent(
        "back-to-back",
        "BUSY",
        [],
        "20260826T180000Z",
        "20260826T183000Z",
      ),
      ...timedEvent(
        "separate",
        "BUSY",
        [],
        "20260826T190000Z",
        "20260826T200000Z",
      ),
    );

    const output = transformCalendar(source, "Alice Mtg");

    expect(output.match(/BEGIN:VEVENT/g)).toHaveLength(2);
    expect(output).toContain("DTSTART:20260826T160000Z");
    expect(output).toContain("DTEND:20260826T183000Z");
    expect(output).toContain("DTSTART:20260826T190000Z");
    expect(output).toContain("DTEND:20260826T200000Z");
  });

  it("merges occurrences from separate recurring series", () => {
    const source = calendar(
      ...timedEvent("first-series", "BUSY", [
        "RRULE:FREQ=DAILY;COUNT=2",
      ]),
      ...timedEvent(
        "second-series",
        "BUSY",
        ["RRULE:FREQ=DAILY;COUNT=2"],
        "20260826T170000Z",
        "20260826T180000Z",
      ),
    );

    const output = transformCalendar(source, "Alice Mtg");

    expect(output.match(/BEGIN:VEVENT/g)).toHaveLength(2);
    expect(output).toContain("DTSTART:20260826T160000Z");
    expect(output).toContain("DTEND:20260826T180000Z");
    expect(output).toContain("DTSTART:20260827T160000Z");
    expect(output).toContain("DTEND:20260827T180000Z");
  });

  it("escapes the configured title and emits a valid empty calendar", () => {
    const namedOutput = transformCalendar(
      calendar(...timedEvent("busy", "BUSY")),
      "Alice, Jr.; Team\nLead Mtg",
    );
    const emptyOutput = transformCalendar(
      calendar(...timedEvent("free", "FREE")),
      "Alice Mtg",
    );

    expect(namedOutput).toContain(
      "SUMMARY:Alice\\, Jr.\\; Team\\nLead Mtg",
    );
    expect(emptyOutput).not.toContain("BEGIN:VEVENT");
    expect(() => ICAL.parse(emptyOutput)).not.toThrow();
  });

  it("preserves Unicode titles through iCalendar serialization", () => {
    const title = "📅Alice Mtg📅";
    const output = transformCalendar(
      calendar(...timedEvent("busy", "BUSY")),
      title,
    );
    const parsed = new ICAL.Component(ICAL.parse(output));
    const event = parsed.getFirstSubcomponent("vevent");

    expect(output).toContain(`SUMMARY:${title}`);
    expect(event?.getFirstPropertyValue("summary")).toBe(title);
  });

  it("rejects malformed input, non-calendars, and blank titles", () => {
    expect(() => transformCalendar("not an ics file", "Alice")).toThrow(
      CalendarTransformError,
    );
    expect(() =>
      transformCalendar("BEGIN:VEVENT\r\nEND:VEVENT\r\n", "Alice"),
    ).toThrow(CalendarTransformError);
    expect(() =>
      transformCalendar(calendar(...timedEvent("busy", "BUSY")), " "),
    ).toThrow(CalendarTransformError);
    expect(() =>
      transformCalendarSource(
        calendar(...timedEvent("busy", "BUSY")),
        "Alice",
        { now: new Date("invalid") },
      ),
    ).toThrow(CalendarTransformError);
  });
});
