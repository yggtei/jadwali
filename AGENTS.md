# AGENTS.md — Jadwali (جدولي) University Schedule PWA

## Overview
A free university timetable PWA that converts Outlook ICS calendars into a clean schedule with lecture reminders (even when closed), grade tracking, GPA calculator, and tasks. Deploys to Cloudflare Workers + Durable Objects.

## Architecture
- **Frontend**: Single-page app in `public/index.html` + `public/ics.js` (ICS parser) + `public/sw.js` (service worker)
- **Backend**: Cloudflare Worker entry at `src/index.mjs` with two Durable Objects in `src/durable.mjs`:
  - `KeyStore` — generates/stores VAPID keys for Web Push
  - `Subscriber` — per-device object storing schedule, settings, alarms; wakes every 15 min + before each lecture to send push
- **Shared utils**: `src/shared.mjs` (allowed hosts, JSON helpers, crypto)

## Key Files
| File | Purpose |
|------|---------|
| `wrangler.toml` | Deploy config: assets from `./public`, worker-first routes `/api/*` and `/proxy`, DO bindings |
| `src/index.mjs` | HTTP entry: serves assets, `/proxy` (fetches ICS from Outlook), `/api/*` (push endpoints) |
| `src/durable.mjs` | Core logic: schedule parsing, diffing, reminder scheduling, VAPID auth, Web Push sending |
| `src/shared.mjs` | Allowed ICS/Push hosts, sanitization, JSON/crypto helpers |
| `public/index.html` | Full app (HTML/CSS/JS inline): schedule, grades, tasks, settings, PWA install flow |
| `public/sw.js` | Service worker: offline caching, push event handler (fetches notification text from `/api/pending`) |
| `public/ics.js` | ICS parser (exported as `ICS.parse`) |
| `public/manifest.webmanifest` | PWA manifest for home-screen install |

## Commands
```bash
# Local dev (requires wrangler login)
npx wrangler dev

# Deploy to Cloudflare
npx wrangler deploy
```
No build step, no lint/typecheck/test scripts.

## Critical Conventions
- **ICS fetching**: Browser cannot fetch Outlook ICS directly (CORS). Worker provides `/proxy?url=...` endpoint; allowed hosts in `ICS_HOSTS` (`src/shared.mjs:4-10`).
- **Push notifications**: Require deployment (Worker + DO). iOS needs "Add to Home Screen" + open from icon (standalone mode) for background alerts.
- **Durable Object identity**: Each device gets a `Subscriber` DO keyed by SHA-256 of its push endpoint (`src/index.mjs:67`).
- **Alarms**: DO uses `setAlarm` to wake at next reminder time or 15-min refresh interval (`src/durable.mjs:297-315`).
- **VAPID keys**: Generated once by `KeyStore` DO on first use (`src/durable.mjs:18-34`).
- **Settings sync**: Client POSTs to `/api/sync` with full config; DO validates/sanitizes (`src/durable.mjs:73-112`).
- **Data storage**: Client uses `localStorage` (key `myschedule.v1`); server stores only ICS URL, push endpoint, reminder settings, tasks, course overrides.

## Adding New ICS Hosts
If a university uses a non-Outlook calendar domain, add to `ICS_HOSTS` in `src/shared.mjs:4-10` and redeploy.

## Environment
- Requires Cloudflare account (free tier sufficient)
- `wrangler.toml` name must match Cloudflare project name (`jadwali`)
- No environment variables needed; VAPID keys auto-generated in DO storage

## Common Tasks
- **Change allowed ICS domains**: Edit `src/shared.mjs` `ICS_HOSTS` array
- **Modify reminder logic**: `src/durable.mjs` `remindersFor()`, `collect()`, `schedule()`
- **Update UI strings**: `public/index.html` `D` object (bilingual ar/en)
- **Adjust refresh interval**: `REFRESH_MS` in `src/durable.mjs:13` (default 15 min)

## Gotchas
- `run_worker_first` in `wrangler.toml` ensures `/proxy` and `/api/*` hit the Worker before static assets
- iOS push only works from home-screen icon (not Safari tab) due to Apple restrictions
- First deploy requires `npx wrangler login` then `npx wrangler deploy`
- `wrangler.toml` migrations (`v1`) created DO classes; changing class names needs new migration tag