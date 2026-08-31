# Work Outlook calendar for family

A Cloudflare Worker that turns an Outlook calendar subscription into a
privacy-minimized family calendar. The published feed contains only timed
events that Outlook explicitly marks as **Busy**, and every event is assigned a
short, configurable title such as `Alice Mtg`. Overlapping and immediately
back-to-back meetings are combined into a single availability block.


<img width="968" height="323" alt="image" src="https://github.com/user-attachments/assets/f370cf78-5bae-4f92-bdb7-ca9ad9dafb4d" />

The Worker removes source titles, descriptions, locations, attendees,
organizers, meeting links, attachments, alarms, and other calendar metadata.
Access requires a shared secret in the subscription URL.


## Requirements

- Node.js and npm
- A Cloudflare account
- An Outlook calendar ICS subscription URL

## Guided setup

For a complete interactive walkthrough—including creating a Cloudflare account,
publishing the Outlook calendar, configuring secrets, testing, deploying, and
verifying the result—run:

```sh
./setup.sh
```

You can also use `npm run setup`. The script explains each step before making a
change, opens the relevant signup and calendar pages when supported, and prints
the final family subscription URL. It never writes the event title, Outlook
URL, or generated access key to the project.

The manual instructions below cover the same process.

To get the Outlook subscription address:

1. Open [Outlook Calendar](https://outlook.cloud.microsoft/calendar).
2. Select the **three-dot menu** in the top bar, then select **Settings**.
3. Select **Shared Calendars** in the side tab of the Settings panel.
4. Find **Publish a Calendar** and select your work calendar.
5. Choose the least-privileged availability option, preferably
   **Can view when I'm busy**.
6. Select **Publish**, then copy the **ICS** link—not the HTML link.

The exact labels may differ slightly depending on the Outlook version and your
organization's configuration. If **Publish a Calendar** is unavailable, your
organization has probably disabled public calendar publishing. This Worker
cannot bypass that policy.

Treat the Outlook ICS URL as a secret. It normally contains a token that grants
access to the source calendar.

## Configure

Install dependencies and authenticate Wrangler:

```sh
npm install
npx wrangler login
```

Set the event title, source calendar URL, and a long random access secret
using Wrangler's interactive secret prompts:

```sh
npm run set-title
npx wrangler secret put OUTLOOK_CALENDAR_URL
npx wrangler secret put ACCESS_SECRET
```

`EVENT_TITLE` is used verbatim after surrounding whitespace is removed. Short
titles such as `Alice Mtg` work well on compact calendar displays. The title
helper UTF-8 encodes the value before sending it to Wrangler, so emoji and other
Unicode characters are preserved. Do not put these values in `wrangler.jsonc`,
a command-line argument, or a committed file.

For local development, copy `.dev.vars.example` to `.dev.vars` and replace all
placeholder values. `.dev.vars` is ignored by Git.

## Develop and test

```sh
npm run dev
npm test
npm run typecheck
```

The local subscription URL is:

```text
http://localhost:8787/calendar.ics?key=<ACCESS_SECRET>
```

Only `GET` and `HEAD` requests to `/calendar.ics` with exactly one correct
`key` parameter can access the feed. All other requests redirect to the
configured decoy page and do not contact Outlook.

## Deploy

Deploy the Worker:

```sh
npm run deploy
```

Subscribe family calendar clients to:

```text
https://<worker-hostname>/calendar.ics?key=<ACCESS_SECRET>
```

Calendar clients choose their own refresh schedule. The Worker additionally
caches a successfully transformed feed at Cloudflare's edge for five minutes.
The cache key does not contain the access secret or the Outlook URL, and
authentication is checked before every cache lookup.

Cloudflare Cache API entries are local to individual data centers, may be
evicted early, and are not durable stale storage. If there is no cache entry and
Outlook is unavailable or returns invalid calendar data, the Worker returns
`502` rather than an empty calendar.

## Change the event title

Update the title without changing code:

```sh
npm run set-title
```

Enter the complete title, such as `Alice Mtg` or `📅Alice Mtg📅`, at the prompt.
The helper transports the title as base64-encoded UTF-8 to prevent command-line
or secret-input encoding damage. Wrangler deploys the updated secret
immediately; no separate deployment is needed. Existing edge caches may show
the previous title for up to five minutes.

## Rotate access

The `key` query parameter is a bearer credential. It can appear in calendar
client configuration, copied URLs, browser history, and infrastructure request
logs. Share it only with intended subscribers.

To revoke the existing family URLs, replace the secret:

```sh
npx wrangler secret put ACCESS_SECRET
```

Then update every family subscription with the new `key` value. Existing URLs
stop authenticating immediately. If the Outlook subscription URL itself is
exposed, revoke or republish it in Outlook and update
`OUTLOOK_CALENDAR_URL` with `wrangler secret put`.

## Filtering behavior

An event is published only when all of these conditions hold:

- `X-MICROSOFT-CDO-BUSYSTATUS` is `BUSY` (case-insensitive).
- `DTSTART` is a date-time rather than an all-day date.
- Outlook's all-day marker is not `TRUE`.
- `STATUS` is not `CANCELLED`.
- The event has both `UID` and `DTSTART`.

`FREE`, `TENTATIVE`, `OOF`, `AWAY`, missing, and unknown busy-state values are
dropped. Standard `TRANSP:OPAQUE` does not override a missing Outlook busy
state. This fail-closed rule avoids publishing events whose availability is
ambiguous.

The Worker resolves recurring and timezone-aware meetings into concrete UTC
blocks covering the previous 30 days through the next 366 days. It then merges
blocks whenever the next meeting starts at or before the current meeting ends.
The output retains only a deterministic identifier, start and end times,
revision timestamp, and rewritten summary. When a recurring Busy series has a
Free, Tentative, Away, or cancelled exception, that occurrence is omitted.
