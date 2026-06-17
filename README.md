# Outlook Event Inspector — Blazor WebAssembly add-in (POC)

A minimal **Outlook web add-in** written in **Blazor WebAssembly** that uses
**Office.js** to:

- **Read** the selected calendar appointment — recurrence info, master/occurrence
  classification and Exchange IDs; and
- **Create** appointments — a quick pre-filled new-appointment form, plus a
  compose-mode panel that can write a **recurrence pattern**.

No Azure AD app registration is required — it runs purely on the Office.js
item-level API.

The task pane adapts to how it was opened:

| Activation | What you see |
| --- | --- |
| Reading an appointment | Inspector (recurrence / IDs) + "Quick create" form |
| Composing an appointment | "Apply to this appointment" panel with recurrence options |

## What it validates (the POC questions)

| Question | Answered by | Field |
| --- | --- | --- |
| Is the event recurring? | `item.recurrence != null` | **Recurring flag** |
| Is it the master or an occurrence? | `recurrence` + `item.seriesId` | **Classification** |
| Exchange ID of this item? | `item.itemId` | **Exchange ID** |
| Link back to the master? | `item.seriesId` | **Series / master ID** |

Classification logic (in [`wwwroot/js/officeInterop.js`](wwwroot/js/officeInterop.js)):

```
recurrence == null, seriesId == null  ->  Single (non-recurring) appointment
recurrence != null, seriesId == null  ->  Series master (the recurring meeting itself)
recurrence != null, seriesId != null  ->  Occurrence / instance of a series
```

It also pulls subject, organizer, attendees, start/end, recurrence pattern
(type, time zone, series start/end dates, daily start/end times, duration).

### Creating events

| Path | API | Recurrence? | Permission |
| --- | --- | --- | --- |
| Quick create (read pane) | `displayNewAppointmentForm()` | ❌ no | ReadItem |
| Compose panel | `subject/location/start/end/requiredAttendees.setAsync` + `recurrence.setAsync` | ✅ yes (daily/weekly) | ReadWriteItem |

- **Quick create** pops Outlook's new-appointment window pre-filled; the user
  reviews and saves. Single events only — the form API can't accept a recurrence.
- **Compose panel** appears when you open the add-in while *creating/editing* an
  appointment. It sets the fields on that open item and, if "Make recurring" is
  ticked, writes a recurrence pattern via `item.recurrence.setAsync` (built from
  an `Office.SeriesTime`). Then you save in Outlook.

## Honest limitations (read this)

The Office.js **item-level** API only sees the **selected** appointment. It does
**not** provide:

- the full list of occurrences in a series,
- a per-occurrence **index** (1st, 2nd, …),
- a direct **"modified exception"** flag with original-vs-new dates.

Those require **Microsoft Graph** (`GET /me/events` / `calendarView`), where each
event has `type` = `seriesMaster` | `occurrence` | `exception` | `singleInstance`,
a `seriesMasterId`, and `originalStart` (the pre-modification date). That path
needs an Azure AD app registration + SSO/MSAL token — see "Next step" below.

## Project layout

```
manifest.xml                     Add-in manifest (sideload this into Outlook)
wwwroot/index.html               Loads office.js + the Blazor app + interop
wwwroot/js/officeInterop.js      Office.js calls -> JSON for Blazor
wwwroot/assets/icon-*.png        Ribbon/manifest icons
Models/OutlookEventInfo.cs       Typed model of the interop JSON
Pages/Home.razor                 The task-pane UI
```

## Run it

Prerequisites: .NET SDK (9.0+), and a trusted HTTPS dev cert.

```powershell
# one time: trust the local dev certificate
dotnet dev-certs https --trust

# run the dev server on https://localhost:7010
dotnet run --launch-profile https
```

Open `https://localhost:7010/` in a browser to confirm it loads — outside Outlook
it will show "Office.js host not detected", which is expected.

## Sideload into Outlook on the web

1. Keep `dotnet run` running (the manifest points at `https://localhost:7010`).
2. Go to **outlook.office.com**, open **Settings → General → Manage add-ins**
   (or **Get Add-ins → My add-ins → Add a custom add-in → Add from file**).
3. Upload [`manifest.xml`](manifest.xml).
4. **To read:** open any **calendar appointment** (read view). On the ribbon /
   **Apps** menu you'll see **"Inspect Event"** under the *Event POC* group — click
   it to open the task pane. Click a **recurring** meeting vs one of its **single
   occurrences** to watch the classification and IDs change (pane auto-refreshes).
5. **To create:** in the read pane use **Quick create**, or start a **New event**
   in the calendar and open the add-in (**"Event Tools"** button) to use the
   compose panel — tick **Make recurring** to set a daily/weekly pattern, then
   **Save** in Outlook.

> The manifest requires **Mailbox requirement set 1.7** (needed for
> `item.recurrence` and `item.seriesId`). Modern Microsoft 365 mailboxes support this.

## Next step — full occurrence/exception data via Microsoft Graph

To enumerate every occurrence's individual Exchange ID, occurrence index, and
modified exceptions (original vs new date), add a Graph path:

1. Register an Azure AD app; add delegated `Calendars.Read`.
2. Get a token in the add-in via Office SSO (`getAccessTokenAsync`) or MSAL.
3. Call `GET /me/calendar/calendarView?startDateTime=...&endDateTime=...` and read
   `type`, `seriesMasterId`, `originalStart` on each returned event.

The current UI already calls these out in the in-app note so the POC scope is clear.
