import ICAL from "ical.js";

type JCalParameterValue = string | string[];
type JCalParameters = Record<string, JCalParameterValue>;
type JCalProperty = [
  name: string,
  parameters: JCalParameters,
  valueType: string,
  ...values: unknown[],
];
type JCalComponent = [
  name: string,
  properties: JCalProperty[],
  components: JCalComponent[],
];

const DEFAULT_LOOKBACK_DAYS = 30;
const DEFAULT_LOOKAHEAD_DAYS = 366;
const MAX_RECURRENCE_OCCURRENCES = 10_000;
const MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1_000;
const SOURCE_EVENT_PROPERTY_ALLOWLIST = new Set([
  "uid",
  "dtstart",
  "dtend",
  "duration",
  "recurrence-id",
  "rrule",
  "rdate",
  "exdate",
  "sequence",
  "dtstamp",
  "last-modified",
]);

interface BusyInterval {
  start: number;
  end: number;
}

export interface CalendarTransformOptions {
  now?: Date;
  lookbackDays?: number;
  lookaheadDays?: number;
}

export class CalendarTransformError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "CalendarTransformError";
  }
}

export function transformCalendar(
  source: string,
  eventTitle: string,
  options: CalendarTransformOptions = {},
): string {
  const title = eventTitle.trim();
  if (title.length === 0) {
    throw new CalendarTransformError("Event title is required");
  }

  let parsed: unknown;
  try {
    parsed = ICAL.parse(source);
  } catch {
    throw new CalendarTransformError("Source calendar is malformed");
  }

  if (!isJCalComponent(parsed) || parsed[0].toLowerCase() !== "vcalendar") {
    throw new CalendarTransformError("Source is not a VCALENDAR");
  }

  const sourceEvents = parsed[2].filter(
    (component) => component[0].toLowerCase() === "vevent",
  );
  const excludedRecurrenceIds = collectExcludedRecurrenceIds(sourceEvents);
  const sanitizedEvents = sourceEvents
    .filter(shouldKeepEvent)
    .map((component) =>
      sanitizeEvent(component, title, excludedRecurrenceIds),
    );
  const timezones = parsed[2]
    .filter((component) => component[0].toLowerCase() === "vtimezone");
  const intermediateCalendar: JCalComponent = [
    "vcalendar",
    [
      ["prodid", {}, "text", "-//work-outlook-calendar-for-family//EN"],
      ["version", {}, "text", "2.0"],
      ["calscale", {}, "text", "GREGORIAN"],
      ["method", {}, "text", "PUBLISH"],
    ],
    [...timezones, ...sanitizedEvents],
  ];
  const intervals = expandBusyIntervals(intermediateCalendar, options);
  const events = mergeIntervals(intervals).map((interval) =>
    createMergedEvent(interval, title),
  );
  const output: JCalComponent = [
    "vcalendar",
    intermediateCalendar[1],
    events,
  ];

  try {
    return `${ICAL.stringify(output)}\r\n`;
  } catch {
    throw new CalendarTransformError("Calendar could not be serialized");
  }
}

function shouldKeepEvent(component: JCalComponent): boolean {
  const busyStatus = getFirstPropertyText(
    component,
    "x-microsoft-cdo-busystatus",
  );
  if (busyStatus?.trim().toUpperCase() !== "BUSY") {
    return false;
  }

  const status = getFirstPropertyText(component, "status");
  if (status?.trim().toUpperCase() === "CANCELLED") {
    return false;
  }

  const allDayMarker = getFirstPropertyText(
    component,
    "x-microsoft-cdo-alldayevent",
  );
  if (allDayMarker?.trim().toUpperCase() === "TRUE") {
    return false;
  }

  const start = getFirstProperty(component, "dtstart");
  if (start === undefined || start[2].toLowerCase() === "date") {
    return false;
  }

  return getFirstProperty(component, "uid") !== undefined;
}

function sanitizeEvent(
  component: JCalComponent,
  eventTitle: string,
  excludedRecurrenceIds: Map<string, JCalProperty[]>,
): JCalComponent {
  const properties = component[1]
    .filter((property) =>
      SOURCE_EVENT_PROPERTY_ALLOWLIST.has(property[0].toLowerCase()),
    )
    .map(cloneProperty);

  if (getFirstProperty(component, "recurrence-id") === undefined) {
    const uid = getFirstPropertyText(component, "uid");
    for (const recurrenceId of uid === undefined
      ? []
      : (excludedRecurrenceIds.get(uid) ?? [])) {
      const exclusion = cloneProperty(recurrenceId);
      exclusion[0] = "exdate";
      properties.push(exclusion);
    }
  }

  properties.push(["summary", {}, "text", eventTitle]);
  return ["vevent", properties, []];
}

function collectExcludedRecurrenceIds(
  components: JCalComponent[],
): Map<string, JCalProperty[]> {
  const exclusions = new Map<string, JCalProperty[]>();

  for (const component of components) {
    if (shouldKeepEvent(component)) {
      continue;
    }

    const uid = getFirstPropertyText(component, "uid");
    const recurrenceId = getFirstProperty(component, "recurrence-id");
    if (uid === undefined || recurrenceId === undefined) {
      continue;
    }

    const existing = exclusions.get(uid) ?? [];
    existing.push(recurrenceId);
    exclusions.set(uid, existing);
  }

  return exclusions;
}

function expandBusyIntervals(
  calendar: JCalComponent,
  options: CalendarTransformOptions,
): BusyInterval[] {
  const now = options.now ?? new Date();
  const lookbackDays = options.lookbackDays ?? DEFAULT_LOOKBACK_DAYS;
  const lookaheadDays = options.lookaheadDays ?? DEFAULT_LOOKAHEAD_DAYS;
  if (
    !Number.isFinite(now.getTime()) ||
    !Number.isInteger(lookbackDays) ||
    lookbackDays < 0 ||
    !Number.isInteger(lookaheadDays) ||
    lookaheadDays < 1
  ) {
    throw new CalendarTransformError("Calendar window is invalid");
  }

  const windowStart = now.getTime() - lookbackDays * MILLISECONDS_PER_DAY;
  const windowEnd = now.getTime() + lookaheadDays * MILLISECONDS_PER_DAY;

  try {
    const calendarComponent = new ICAL.Component(calendar);
    const intervals: BusyInterval[] = [];

    for (const eventComponent of calendarComponent.getAllSubcomponents(
      "vevent",
    )) {
      const event = new ICAL.Event(eventComponent);
      if (event.isRecurrenceException()) {
        continue;
      }

      if (!event.isRecurring()) {
        addInterval(
          intervals,
          event.startDate.toJSDate().getTime(),
          event.endDate.toJSDate().getTime(),
          windowStart,
          windowEnd,
        );
        continue;
      }

      const iterator = event.iterator();
      let occurrenceCount = 0;
      let expansionFinished = false;
      while (occurrenceCount < MAX_RECURRENCE_OCCURRENCES) {
        const occurrence = iterator.next();
        if (!occurrence) {
          expansionFinished = true;
          break;
        }

        occurrenceCount += 1;
        const details = event.getOccurrenceDetails(occurrence);
        const start = details.startDate.toJSDate().getTime();
        if (start >= windowEnd) {
          expansionFinished = true;
          break;
        }

        addInterval(
          intervals,
          start,
          details.endDate.toJSDate().getTime(),
          windowStart,
          windowEnd,
        );
      }

      if (
        occurrenceCount === MAX_RECURRENCE_OCCURRENCES &&
        !expansionFinished
      ) {
        throw new CalendarTransformError(
          "Calendar recurrence exceeds the supported limit",
        );
      }
    }

    return intervals;
  } catch (error) {
    if (error instanceof CalendarTransformError) {
      throw error;
    }
    throw new CalendarTransformError("Calendar recurrence is invalid");
  }
}

function addInterval(
  intervals: BusyInterval[],
  start: number,
  end: number,
  windowStart: number,
  windowEnd: number,
): void {
  if (
    Number.isFinite(start) &&
    Number.isFinite(end) &&
    end > start &&
    end > windowStart &&
    start < windowEnd
  ) {
    intervals.push({ start, end });
  }
}

function mergeIntervals(intervals: BusyInterval[]): BusyInterval[] {
  const sorted = [...intervals].sort(
    (left, right) => left.start - right.start || left.end - right.end,
  );
  const merged: BusyInterval[] = [];

  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (previous === undefined || interval.start > previous.end) {
      merged.push({ ...interval });
      continue;
    }

    previous.end = Math.max(previous.end, interval.end);
  }

  return merged;
}

function createMergedEvent(
  interval: BusyInterval,
  eventTitle: string,
): JCalComponent {
  const start = toJCalDateTime(interval.start);
  const end = toJCalDateTime(interval.end);
  const uid =
    `busy-${interval.start}-${interval.end}@work-outlook-calendar-for-family`;

  return [
    "vevent",
    [
      ["uid", {}, "text", uid],
      ["dtstamp", {}, "date-time", "1970-01-01T00:00:00Z"],
      ["dtstart", {}, "date-time", start],
      ["dtend", {}, "date-time", end],
      ["summary", {}, "text", eventTitle],
    ],
    [],
  ];
}

function toJCalDateTime(timestamp: number): string {
  return ICAL.Time.fromJSDate(new Date(timestamp), true).toString();
}

function getFirstProperty(
  component: JCalComponent,
  name: string,
): JCalProperty | undefined {
  const normalizedName = name.toLowerCase();
  return component[1].find(
    (property) => property[0].toLowerCase() === normalizedName,
  );
}

function getFirstPropertyText(
  component: JCalComponent,
  name: string,
): string | undefined {
  const property = getFirstProperty(component, name);
  const value = property?.[3];
  if (typeof value === "string" || typeof value === "number") {
    return String(value);
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return undefined;
}

function cloneProperty(property: JCalProperty): JCalProperty {
  const parameters = Object.fromEntries(
    Object.entries(property[1]).map(([name, value]) => [
      name,
      Array.isArray(value) ? [...value] : value,
    ]),
  );
  return [property[0], parameters, property[2], ...property.slice(3)];
}

function isJCalComponent(value: unknown): value is JCalComponent {
  if (
    !Array.isArray(value) ||
    value.length !== 3 ||
    typeof value[0] !== "string" ||
    !Array.isArray(value[1]) ||
    !Array.isArray(value[2])
  ) {
    return false;
  }

  return (
    value[1].every(isJCalProperty) &&
    value[2].every((component) => isJCalComponent(component))
  );
}

function isJCalProperty(value: unknown): value is JCalProperty {
  if (
    !Array.isArray(value) ||
    value.length < 4 ||
    typeof value[0] !== "string" ||
    !isRecord(value[1]) ||
    typeof value[2] !== "string"
  ) {
    return false;
  }

  return Object.values(value[1]).every(
    (parameter) =>
      typeof parameter === "string" ||
      (Array.isArray(parameter) &&
        parameter.every((item) => typeof item === "string")),
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
