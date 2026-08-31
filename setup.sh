#!/usr/bin/env bash

set -Eeuo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

DEPLOY_LOG=""

cleanup() {
  if [[ -n "$DEPLOY_LOG" && -f "$DEPLOY_LOG" ]]; then
    rm -f -- "$DEPLOY_LOG"
  fi
}
trap cleanup EXIT

heading() {
  printf '\n============================================================\n'
  printf '%s\n' "$1"
  printf '============================================================\n\n'
}

open_url() {
  local url="$1"

  if [[ "${SETUP_NO_OPEN:-}" == "1" ]]; then
    return
  fi

  case "$(uname -s)" in
    Darwin)
      open "$url" >/dev/null 2>&1 || true
      ;;
    Linux)
      if command -v xdg-open >/dev/null 2>&1; then
        xdg-open "$url" >/dev/null 2>&1 || true
      fi
      ;;
  esac
}

pause() {
  read -r -p "Press Enter when you are ready to continue. "
}

fail() {
  printf '\nSetup stopped: %s\n' "$1" >&2
  printf 'Fix the issue above, then run ./setup.sh again.\n' >&2
  exit 1
}

heading "Outlook Busy Calendar - guided setup"

cat <<'EOF'
This script will:

  1. Check the software needed on this computer.
  2. Help you create or sign in to a free Cloudflare account.
  3. Help you obtain your Outlook calendar subscription link.
  4. Ask for the short event title shown to your family.
  5. Generate a private family access key.
  6. test and deploy the calendar proxy.
  7. Print the subscription URL to share with your family.

Your event title, Outlook calendar URL, and family access key will be sent
directly to Cloudflare as encrypted Worker secrets. They will not be written to
this project or committed to Git.
EOF

pause

heading "Step 1 of 7 - Check required software"

if ! command -v node >/dev/null 2>&1; then
  cat <<'EOF'
Node.js is not installed. This app requires Node.js 22 or newer.

The Node.js download page will open. Install the current LTS release, reopen
your terminal, return to this folder, and run ./setup.sh again.
EOF
  open_url "https://nodejs.org/en/download"
  fail "Node.js is required."
fi

NODE_MAJOR="$(node -p 'Number(process.versions.node.split(".")[0])')"
if [[ "$NODE_MAJOR" -lt 22 ]]; then
  cat <<EOF
Your Node.js version is $(node --version), but this app requires version 22 or
newer. The Node.js download page will open. Install the current LTS release,
reopen your terminal, and run ./setup.sh again.
EOF
  open_url "https://nodejs.org/en/download"
  fail "Node.js 22 or newer is required."
fi

if ! command -v npm >/dev/null 2>&1; then
  fail "npm is missing. Reinstall Node.js from https://nodejs.org/en/download."
fi

printf 'Found Node.js %s and npm %s.\n' "$(node --version)" "$(npm --version)"
printf 'Installing this app'\''s dependencies. This may take a minute...\n'
npm install --no-audit --no-fund ||
  fail "npm could not install the app's dependencies."

heading "Step 2 of 7 - Set up Cloudflare"

cat <<'EOF'
Cloudflare runs this app on its global network. A free Cloudflare account is
enough for normal family-calendar use.

If you do not already have an account:

  1. Create one at https://dash.cloudflare.com/sign-up
  2. Verify your email address if Cloudflare asks you to.
  3. Return here after you can see the Cloudflare dashboard.

The signup page will open now. If it does not, copy the address above into your
browser. You can skip account creation if you already have an account.
EOF

open_url "https://dash.cloudflare.com/sign-up"
pause

cat <<'EOF'
Wrangler, Cloudflare's official command-line tool, will now open a browser so
you can authorize this computer. Sign in to the Cloudflare account you want to
use and approve the request.
EOF

pause
npx wrangler login || fail "Cloudflare login did not complete."
npx wrangler whoami >/dev/null ||
  fail "Wrangler could not confirm your Cloudflare login."
printf 'Cloudflare login confirmed.\n'

heading "Step 3 of 7 - Get the Outlook calendar link"

cat <<'EOF'
The app needs Outlook's private ICS subscription link:

  1. Open https://outlook.cloud.microsoft/calendar in your browser.
  2. Select the three-dot menu in the top bar, then select Settings.
  3. Select Shared Calendars in the side tab of the Settings panel.
  4. Under Publish a Calendar, select your work calendar.
  5. Choose the least-privileged availability option, preferably
     "Can view when I'm busy."
  6. Select Publish, then copy the ICS link (not the HTML link).

If "Publish a calendar" is unavailable, your organization has probably disabled
public calendar publishing. This app cannot bypass that policy; contact your IT
administrator before continuing.

Outlook Calendar will open now. Your organization's Outlook screens may use
slightly different labels.
EOF

open_url "https://outlook.cloud.microsoft/calendar"
pause

OUTLOOK_CALENDAR_URL=""
while [[ "$OUTLOOK_CALENDAR_URL" != https://* ]]; do
  read -r -s -p "Paste the Outlook ICS link, then press Enter: " \
    OUTLOOK_CALENDAR_URL
  printf '\n'
  if [[ "$OUTLOOK_CALENDAR_URL" != https://* ]]; then
    printf 'That does not look like a complete HTTPS link. Please try again.\n'
  fi
done

heading "Step 4 of 7 - Choose the event title"

EVENT_TITLE=""
while [[ -z "${EVENT_TITLE//[[:space:]]/}" ]]; do
  read -r -p "Event title shown to your family (example: Alice Mtg): " \
    EVENT_TITLE

  if [[ -z "${EVENT_TITLE//[[:space:]]/}" ]]; then
    printf 'Please enter at least one visible character.\n'
  fi
done

printf 'Events will be named "%s".\n' "$EVENT_TITLE"

heading "Step 5 of 7 - Generate the family access key"

ACCESS_SECRET="$(
  node --input-type=module -e \
    'import { randomBytes } from "node:crypto"; console.log(randomBytes(32).toString("hex"));'
)"

cat <<EOF
This is your private family access key:

  $ACCESS_SECRET

Save it in your password manager now. Anyone who has the final subscription URL
can see when you are in meetings. The key cannot be recovered from Cloudflare,
but you can replace it later and issue a new subscription URL.
EOF

pause

heading "Step 6 of 7 - Test and deploy"

printf 'Running the automated tests...\n'
npm test || fail "The automated tests failed."
npm run typecheck || fail "The TypeScript check failed."

cat <<'EOF'
The app will now be deployed once to create the Cloudflare Worker. It is safe
while configuration is in progress: without the access secret, calendar
requests are redirected and cannot reach Outlook.
EOF

DEPLOY_LOG="$(mktemp "${TMPDIR:-/tmp}/outlook-calendar-deploy.XXXXXX")"
npm run deploy 2>&1 | tee "$DEPLOY_LOG" ||
  fail "Cloudflare could not deploy the Worker."

WORKER_URL="$(
  grep -Eo 'https://[^[:space:]]+\.workers\.dev' "$DEPLOY_LOG" |
    tail -n 1 || true
)"

printf '\nSending the Outlook link to Cloudflare as an encrypted secret...\n'
printf '%s' "$OUTLOOK_CALENDAR_URL" |
  npx wrangler secret put OUTLOOK_CALENDAR_URL ||
  fail "Cloudflare could not save the Outlook calendar secret."

printf 'Sending the event title to Cloudflare as an encrypted secret...\n'
ENCODED_EVENT_TITLE="$(
  printf '%s' "$EVENT_TITLE" |
    node -e \
      'let value="";process.stdin.on("data",chunk=>value+=chunk);process.stdin.on("end",()=>process.stdout.write(Buffer.from(value,"utf8").toString("base64")));'
)"
printf 'base64:%s' "$ENCODED_EVENT_TITLE" |
  npx wrangler secret put EVENT_TITLE ||
  fail "Cloudflare could not save the event title."

printf 'Sending the family access key to Cloudflare as an encrypted secret...\n'
printf '%s' "$ACCESS_SECRET" |
  npx wrangler secret put ACCESS_SECRET ||
  fail "Cloudflare could not save the family access secret."

heading "Step 7 of 7 - Verify and finish"

if [[ -n "$WORKER_URL" ]]; then
  FEED_URL="${WORKER_URL}/calendar.ics?key=${ACCESS_SECRET}"

  printf 'Checking that the deployed calendar responds correctly...\n'
  if WORKER_URL_VALUE="$WORKER_URL" ACCESS_SECRET_VALUE="$ACCESS_SECRET" \
    node --input-type=module <<'NODE'
const workerUrl = process.env.WORKER_URL_VALUE;
const secret = process.env.ACCESS_SECRET_VALUE;
const url = new URL("/calendar.ics", workerUrl);
url.searchParams.set("key", secret);

let lastStatus = 0;
for (let attempt = 1; attempt <= 6; attempt += 1) {
  try {
    const response = await fetch(url);
    lastStatus = response.status;
    const body = await response.text();
    if (response.ok && body.includes("BEGIN:VCALENDAR")) {
      process.exit(0);
    }
  } catch {
    // A new Worker version can take a few seconds to reach every edge.
  }

  await new Promise((resolve) => setTimeout(resolve, 3000));
}

console.error(`The calendar returned HTTP ${lastStatus || "network error"}.`);
process.exit(1);
NODE
  then
    printf 'The deployed calendar is working.\n'
  else
    cat <<'EOF'
The Worker was deployed, but the calendar check did not succeed. Verify that
the Outlook link is an ICS link and that calendar publishing is enabled, then
run ./setup.sh again. Running setup again safely updates the existing Worker.
EOF
  fi

  cat <<EOF

Setup is complete.

Give this entire subscription URL to your family calendar users:

  $FEED_URL

Keep this URL private. In a calendar app, look for "Add subscribed calendar",
"Subscribe from URL", or a similar option and paste the complete URL.

Cloudflare may serve the previous calendar for up to five minutes after a
configuration change. To change settings later, see README.md.
EOF
else
  cat <<'EOF'
The Worker was deployed, but the script could not identify its public URL in
Wrangler's output.

Open https://dash.cloudflare.com, select Workers & Pages, then select
"outlook-busy-calendar". Copy its workers.dev URL and append:

  /calendar.ics?key=<the family access key shown above>

Keep the complete URL private.
EOF
fi

unset EVENT_TITLE ENCODED_EVENT_TITLE OUTLOOK_CALENDAR_URL ACCESS_SECRET
