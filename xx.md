Here's what's new from this conversation:

**Shared calendar context (important):** The real-world scenario is assistants who support multiple bankers. One assistant has shared access to several bankers' calendars and needs to create and sync meetings on behalf of each banker to Salesforce. This means you'll need the **Calendars.Read.Shared** Microsoft Graph API permission (once the app is working).

**Plugin stability analysis — now more urgent:**

Vipin is asking for a specific deliverable: a document with the **exact changes** needed in the existing VSTO plugin, covering two things:

1. **Load time** — what's causing slow startup, and the precise code changes to fix it so Outlook doesn't auto-disable the plugin
2. **Shared calendar visibility** — what changes are needed so users can see and sync shared calendars

Key constraints he reinforced:
- Must remain a **Windows COM add-in (MSI-deployable)**, not a web app
- Migrating to a newer .NET version is fine, as long as it stays a native Outlook plugin
- **Bare minimum changes** — no large rewrites
- He wants the document shared with him so he can review Claude's reasoning
- This is the "Plan B" — if Microsoft can't fix the web add-in issue, this needs to be ready for a week of testing then packaging

**Your action items:**

1. Get the VSTO plugin code files ready (Roman may have already shared them)
2. Feed them to Claude with clear constraints: minimal changes, must stay as MSI/COM add-in, fix load time, support shared calendars
3. Produce a document with specific, line-level change recommendations
4. Challenge the output — don't accept at face value
5. Share the final document with Vipin

Want me to help draft the prompt you'd use when analyzing the plugin code?
