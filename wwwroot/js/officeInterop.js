// Office.js interop for the Blazor add-in.
// window.officeInterop exposes helpers to READ the selected appointment and to
// CREATE appointments (a quick pre-filled form, and a compose-mode path that can
// set a recurrence pattern). All helpers return JSON strings for Blazor.
(function () {
    let officeReadyPromise = null;
    let pollTimer = null;
    let pollLastId = null;

    function ensureReady() {
        if (officeReadyPromise) return officeReadyPromise;
        officeReadyPromise = new Promise(function (resolve) {
            if (typeof Office === "undefined") {
                resolve({ available: false });
                return;
            }
            Office.onReady(function (info) {
                resolve({
                    available: true,
                    host: info && info.host ? String(info.host) : null,
                    platform: info && info.platform ? String(info.platform) : null
                });
            });
        });
        return officeReadyPromise;
    }

    function safe(fn) { try { return fn(); } catch (e) { return undefined; } }

    function fmtEmail(e) {
        if (!e) return null;
        if (e.emailAddress && e.displayName && e.displayName !== e.emailAddress) {
            return e.displayName + " <" + e.emailAddress + ">";
        }
        return e.displayName || e.emailAddress || null;
    }

    function toIso(v) { return v instanceof Date ? v.toISOString() : (v || null); }

    // "read"  -> selected appointment in read view (fields are plain values)
    // "compose" -> appointment compose view (fields expose getAsync/setAsync)
    // "none" -> no item / not in Outlook
    function detectMode() {
        const item = safe(function () { return Office.context.mailbox.item; });
        if (!item) return "none";
        if (item.subject && typeof item.subject === "object" && typeof item.subject.setAsync === "function") {
            return "compose";
        }
        return "read";
    }

    // Promisify an Office *.setAsync(value, callback) call.
    function setAsyncP(setterFactory) {
        return new Promise(function (resolve) {
            try { setterFactory(function (res) { resolve(res); }); }
            catch (e) { resolve({ status: "failed", error: { message: String(e) } }); }
        });
    }

    // Promisify an Office *.getAsync(callback) call and return its .value (or null).
    function getAsyncValue(getterFactory) {
        return new Promise(function (resolve) {
            try {
                getterFactory(function (res) {
                    resolve(res && res.status === "succeeded" ? res.value : null);
                });
            } catch (e) { resolve(null); }
        });
    }

    function readSeriesTime(rec, result) {
        const st = safe(function () { return rec.seriesTime; });
        if (!st) return;
        result.seriesStartDate = safe(function () { return st.getStartDate(); }) || null;
        result.seriesEndDate = safe(function () { return st.getEndDate(); }) || null;
        result.seriesStartTime = safe(function () { return st.getStartTime(); }) || null;
        result.seriesEndTime = safe(function () { return st.getEndTime(); }) || null;
        result.seriesDurationMinutes = safe(function () { return st.getDuration(); }) || null;
    }

    function classify(result) {
        if (!result.isRecurring) {
            result.classification = result.seriesId
                ? "Occurrence (pattern not returned by host)"
                : "Single (non-recurring) appointment";
        } else if (result.seriesId) {
            result.classification = "Occurrence / instance of a recurring series";
        } else {
            result.classification = "Series master (the recurring meeting itself)";
        }
    }

    // Build a recurrence pattern object for item.recurrence.setAsync.
    // r: { type, interval, days[], startIso, endDateIso, durationMinutes }
    function buildPattern(r) {
        const st = new Office.SeriesTime();
        const s = new Date(r.startIso);
        st.setStartDate(s.getFullYear(), s.getMonth(), s.getDate()); // month is 0-based
        st.setStartTime(s.getHours(), s.getMinutes());
        st.setDuration(r.durationMinutes && r.durationMinutes > 0 ? r.durationMinutes : 30);
        if (r.endDateIso) {
            const e = new Date(r.endDateIso);
            st.setEndDate(e.getFullYear(), e.getMonth(), e.getDate());
        }
        const props = { interval: r.interval && r.interval > 0 ? r.interval : 1 };
        if (r.type === "weekly") {
            props.days = (r.days && r.days.length) ? r.days : ["mon"];
            props.firstDayOfWeek = "sun";
        }
        return { seriesTime: st, recurrenceType: r.type, recurrenceProperties: props };
    }

    window.officeInterop = {
        init: async function () {
            return JSON.stringify(await ensureReady());
        },

        // Host availability + current activation mode (read/compose/none).
        getContext: async function () {
            const ready = await ensureReady();
            // "In Outlook" means a real mailbox exists — NOT merely that office.js
            // loaded (it loads in a plain browser too and fires onReady).
            const inOutlook = ready.available && !!safe(function () { return Office.context.mailbox; });
            const ctx = {
                hostAvailable: inOutlook,
                host: ready.host || null,
                platform: ready.platform || null,
                mode: inOutlook ? detectMode() : "none"
            };
            return JSON.stringify(ctx);
        },

        copyText: async function (text) {
            try { await navigator.clipboard.writeText(text || ""); return true; }
            catch (e) { return false; }
        },

        // Writes a persistent sync-status bar onto the open message/appointment
        // (the area where "This message is from a trusted sender" appears).
        // Read mode on the web only supports informational (text) bars — the
        // action button (insightMessage) is compose-only there.
        setSyncNotification: function (isSynced, detail) {
            try {
                const item = Office.context.mailbox.item;
                if (!item || !item.notificationMessages || !item.notificationMessages.addAsync) return false;
                const msg = isSynced
                    ? ("✓ ALREADY SYNCED TO SALESFORCE" + (detail ? " — " + detail : ""))
                    : "Not synced to Salesforce — open this add-in to sync.";
                item.notificationMessages.addAsync("sfSyncStatus", {
                    type: "informationalMessage",
                    message: String(msg).substring(0, 150),
                    icon: "icon16",
                    persistent: true
                });
                return true;
            } catch (e) { return false; }
        },

        clearSyncNotification: function () {
            try { Office.context.mailbox.item.notificationMessages.removeAsync("sfSyncStatus"); } catch (e) { }
            return true;
        },

        // Shows an informational bar on the current item (Outlook's built-in
        // notification "message box"). Falls back to a browser alert.
        showNotification: function (message) {
            const text = String(message || "").substring(0, 150);
            try {
                const item = Office.context.mailbox.item;
                if (item && item.notificationMessages && item.notificationMessages.addAsync) {
                    item.notificationMessages.addAsync("sfSyncMsg", {
                        type: "informationalMessage",
                        message: text,
                        icon: "icon16",
                        persistent: false
                    });
                    return true;
                }
            } catch (e) { /* fall through */ }
            try { window.alert(text); } catch (e) { }
            return false;
        },

        // READ mode realtime: fires when the user selects a different appointment.
        registerItemChanged: function (dotnetRef) {
            try {
                if (!Office.context || !Office.context.mailbox) return false;
                Office.context.mailbox.addHandlerAsync(
                    Office.EventType.ItemChanged,
                    function () { dotnetRef.invokeMethodAsync("OnItemChanged"); });
                return true;
            } catch (e) { return false; }
        },

        // Fallback for clients where ItemChanged/pinning isn't supported (e.g.
        // Outlook on the web / outlook.com). Polls the selected item's id and
        // invokes OnItemChanged when it changes. Best-effort: only works while the
        // pane stays open and the host updates Office.context.mailbox.item.
        startItemPolling: function (dotnetRef, intervalMs) {
            try {
                if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
                const getId = function () {
                    try {
                        const it = Office.context.mailbox && Office.context.mailbox.item;
                        return it ? (it.itemId || null) : null;
                    } catch (e) { return null; }
                };
                pollLastId = getId();
                pollTimer = setInterval(function () {
                    const cur = getId();
                    if (cur !== pollLastId) {
                        pollLastId = cur;
                        dotnetRef.invokeMethodAsync("OnItemChanged");
                    }
                }, intervalMs && intervalMs > 0 ? intervalMs : 2000);
                return true;
            } catch (e) { return false; }
        },

        stopItemPolling: function () {
            try { if (pollTimer) { clearInterval(pollTimer); pollTimer = null; } } catch (e) { }
            return true;
        },

        // COMPOSE mode realtime: fires as the user edits time / recurrence /
        // location on the open appointment. (Office.js has no subject/attendees
        // change event, so those refresh on manual refresh or one of these fires.)
        registerComposeChanges: function (dotnetRef) {
            try {
                const item = safe(function () { return Office.context.mailbox.item; });
                if (!item || typeof item.addHandlerAsync !== "function" || !Office.EventType) return false;
                const types = [
                    Office.EventType.AppointmentTimeChanged,
                    Office.EventType.RecurrenceChanged,
                    Office.EventType.EnhancedLocationsChanged
                ].filter(Boolean);
                types.forEach(function (t) {
                    item.addHandlerAsync(t, function () {
                        dotnetRef.invokeMethodAsync("OnComposeChanged");
                    });
                });
                return types.length > 0;
            } catch (e) { return false; }
        },

        // ---- READ: pull everything we can from the selected appointment --------
        getSelectedEvent: async function () {
            const ready = await ensureReady();
            const result = { errors: [] };

            if (!ready.available) {
                result.hostAvailable = false;
                result.errors.push("Office.js host not detected. Open this page inside Outlook by sideloading the add-in.");
                return JSON.stringify(result);
            }
            result.hostAvailable = true;
            result.host = ready.host;
            result.platform = ready.platform;

            const mailbox = safe(function () { return Office.context.mailbox; });
            if (!mailbox) {
                // office.js loaded but no mailbox -> running outside Outlook (dev/browser).
                result.hostAvailable = false;
                result.errors.push("Running outside Outlook (no mailbox).");
                return JSON.stringify(result);
            }

            const item = safe(function () { return mailbox.item; });
            if (!item) { result.errors.push("No item selected. Open or click an appointment in your calendar."); return JSON.stringify(result); }

            const apptType = (Office.MailboxEnums && Office.MailboxEnums.ItemType)
                ? Office.MailboxEnums.ItemType.Appointment : "appointment";
            result.itemType = safe(function () { return String(item.itemType); });
            result.isAppointment = result.itemType === apptType;

            result.exchangeId = safe(function () { return item.itemId; }) || null;
            result.seriesId = safe(function () { return item.seriesId; }) || null;

            result.subject = safe(function () { return item.subject; }) || null;
            result.location = safe(function () {
                if (typeof item.location === "string") return item.location;
                return item.location && item.location.displayName;
            }) || null;
            result.organizer = fmtEmail(safe(function () { return item.organizer; }));
            result.requiredAttendees = (safe(function () { return item.requiredAttendees; }) || []).map(fmtEmail).filter(Boolean);
            result.optionalAttendees = (safe(function () { return item.optionalAttendees; }) || []).map(fmtEmail).filter(Boolean);
            result.start = toIso(safe(function () { return item.start; }));
            result.end = toIso(safe(function () { return item.end; }));

            const rec = safe(function () { return item.recurrence; });
            result.isRecurring = !!rec;

            if (!rec) {
                result.classification = result.seriesId
                    ? "Occurrence (pattern not returned by host)"
                    : "Single (non-recurring) appointment";
            } else if (result.seriesId) {
                result.classification = "Occurrence / instance of a recurring series";
            } else {
                result.classification = "Series master (the recurring meeting itself)";
            }

            if (rec) {
                result.recurrenceType = safe(function () { return String(rec.recurrenceType); }) || null;
                result.recurrenceProperties = safe(function () { return rec.recurrenceProperties; }) || null;
                result.recurrenceTimeZone = safe(function () { return rec.recurrenceTimeZone && rec.recurrenceTimeZone.name; }) || null;
                const st = safe(function () { return rec.seriesTime; });
                if (st) {
                    result.seriesStartDate = safe(function () { return st.getStartDate(); }) || null;
                    result.seriesEndDate = safe(function () { return st.getEndDate(); }) || null;
                    result.seriesStartTime = safe(function () { return st.getStartTime(); }) || null;
                    result.seriesEndTime = safe(function () { return st.getEndTime(); }) || null;
                    result.seriesDurationMinutes = safe(function () { return st.getDuration(); }) || null;
                }
            }

            return JSON.stringify(result);
        },

        // ---- READ (compose): pull the OPEN appointment's actual values ---------
        // In compose mode fields are async getters, so this differs from the
        // read-mode path. Used when the add-in is opened on an existing event.
        getComposeEvent: async function () {
            const ready = await ensureReady();
            const result = { errors: [], mode: "compose" };

            if (!ready.available) {
                result.hostAvailable = false;
                result.errors.push("Office.js host not detected.");
                return JSON.stringify(result);
            }
            result.hostAvailable = true;
            result.host = ready.host;
            result.platform = ready.platform;

            const item = safe(function () { return Office.context.mailbox.item; });
            if (!item || typeof item.subject !== "object") {
                result.errors.push("Not in appointment compose mode.");
                return JSON.stringify(result);
            }

            result.itemType = safe(function () { return String(item.itemType); }) || null;
            result.isAppointment = true;

            // IDs: seriesId is synchronous; itemId is usually absent in compose
            // until the item is saved (use saveAndGetExchangeId for that).
            result.exchangeId = safe(function () { return item.itemId; }) || null;
            result.seriesId = safe(function () { return item.seriesId; }) || null;

            // Async getters for the real field values on the open event.
            result.subject = await getAsyncValue(function (cb) { item.subject.getAsync(cb); });
            result.location = await getAsyncValue(function (cb) { item.location.getAsync(cb); });

            const startVal = await getAsyncValue(function (cb) { item.start.getAsync(cb); });
            result.start = startVal instanceof Date ? startVal.toISOString() : (startVal || null);
            const endVal = await getAsyncValue(function (cb) { item.end.getAsync(cb); });
            result.end = endVal instanceof Date ? endVal.toISOString() : (endVal || null);

            const req = await getAsyncValue(function (cb) { item.requiredAttendees.getAsync(cb); }) || [];
            result.requiredAttendees = req.map(fmtEmail).filter(Boolean);
            result.requiredAttendeeEmails = req.map(function (e) { return e && e.emailAddress; }).filter(Boolean);
            const opt = await getAsyncValue(function (cb) { item.optionalAttendees.getAsync(cb); }) || [];
            result.optionalAttendees = opt.map(fmtEmail).filter(Boolean);

            // Recurrence pattern of the open event/series.
            const rec = await getAsyncValue(function (cb) { item.recurrence.getAsync(cb); });
            result.isRecurring = !!rec;
            if (rec) {
                result.recurrenceType = safe(function () { return String(rec.recurrenceType); }) || null;
                result.recurrenceProperties = safe(function () { return rec.recurrenceProperties; }) || null;
                result.recurrenceTimeZone = safe(function () { return rec.recurrenceTimeZone && rec.recurrenceTimeZone.name; }) || null;
                readSeriesTime(rec, result);
            }

            classify(result);
            return JSON.stringify(result);
        },

        // In compose, itemId only exists after a save. This saves the open item
        // and returns its EWS (Exchange) id.
        saveAndGetExchangeId: async function () {
            const item = safe(function () { return Office.context.mailbox.item; });
            if (!item || typeof item.saveAsync !== "function") {
                return JSON.stringify({ ok: false, error: "saveAsync not available." });
            }
            const res = await new Promise(function (resolve) {
                try { item.saveAsync(function (r) { resolve(r); }); }
                catch (e) { resolve({ status: "failed", error: { message: String(e) } }); }
            });
            if (res && res.status === "succeeded") {
                return JSON.stringify({ ok: true, exchangeId: res.value || safe(function () { return item.itemId; }) || null });
            }
            return JSON.stringify({ ok: false, error: res && res.error ? res.error.message : "save failed" });
        },

        // ---- READ: pull info from the selected EMAIL message -------------------
        getSelectedMessage: async function () {
            const ready = await ensureReady();
            const result = { errors: [] };

            if (!ready.available) {
                result.hostAvailable = false;
                result.errors.push("Office.js host not detected. Open this inside Outlook by sideloading the add-in.");
                return JSON.stringify(result);
            }
            result.hostAvailable = true;
            result.host = ready.host;
            result.platform = ready.platform;

            const mailbox = safe(function () { return Office.context.mailbox; });
            if (!mailbox) {
                // office.js loaded but no mailbox -> running outside Outlook (dev/browser).
                result.hostAvailable = false;
                result.errors.push("Running outside Outlook (no mailbox).");
                return JSON.stringify(result);
            }

            const item = safe(function () { return mailbox.item; });
            if (!item) { result.errors.push("No item selected. Open an email message."); return JSON.stringify(result); }

            const msgType = (Office.MailboxEnums && Office.MailboxEnums.ItemType)
                ? Office.MailboxEnums.ItemType.Message : "message";
            result.itemType = safe(function () { return String(item.itemType); }) || null;
            result.isMessage = result.itemType === msgType;
            if (!result.isMessage) {
                result.errors.push("The selected item is not an email (it's a " + result.itemType + ").");
                return JSON.stringify(result);
            }

            const isCompose = typeof item.subject === "object" && typeof item.subject.setAsync === "function";
            result.mode = isCompose ? "compose" : "read";

            if (!isCompose) {
                // Read mode: properties are plain values.
                result.itemId = safe(function () { return item.itemId; }) || null;
                result.subject = safe(function () { return item.subject; }) || null;
                result.normalizedSubject = safe(function () { return item.normalizedSubject; }) || null;
                result.from = fmtEmail(safe(function () { return item.from; }));
                result.sender = fmtEmail(safe(function () { return item.sender; }));
                result.to = (safe(function () { return item.to; }) || []).map(fmtEmail).filter(Boolean);
                result.cc = (safe(function () { return item.cc; }) || []).map(fmtEmail).filter(Boolean);
                result.bcc = (safe(function () { return item.bcc; }) || []).map(fmtEmail).filter(Boolean);
                result.conversationId = safe(function () { return item.conversationId; }) || null;
                result.internetMessageId = safe(function () { return item.internetMessageId; }) || null;
                const created = safe(function () { return item.dateTimeCreated; });
                result.dateTimeCreated = created instanceof Date ? created.toISOString() : (created || null);
                const modified = safe(function () { return item.dateTimeModified; });
                result.dateTimeModified = modified instanceof Date ? modified.toISOString() : (modified || null);
                const atts = safe(function () { return item.attachments; }) || [];
                result.attachments = atts.map(function (a) {
                    return { name: a.name, size: a.size, type: String(a.attachmentType), isInline: !!a.isInline };
                });
            } else {
                // Compose mode: async getters (itemId is null until saved).
                result.itemId = safe(function () { return item.itemId; }) || null;
                result.subject = await getAsyncValue(function (cb) { item.subject.getAsync(cb); });
                result.to = (await getAsyncValue(function (cb) { item.to.getAsync(cb); }) || []).map(fmtEmail).filter(Boolean);
                result.cc = (await getAsyncValue(function (cb) { item.cc.getAsync(cb); }) || []).map(fmtEmail).filter(Boolean);
                result.conversationId = safe(function () { return item.conversationId; }) || null;
            }

            return JSON.stringify(result);
        },

        // ---- CREATE (simple): pop a pre-filled new-appointment form ------------
        // Works from any read context, no extra permissions. Cannot set recurrence.
        displayNewAppointment: async function (paramsJson) {
            const ready = await ensureReady();
            if (!ready.available) return JSON.stringify({ ok: false, error: "No Office host detected." });
            const mailbox = safe(function () { return Office.context.mailbox; });
            if (!mailbox || typeof mailbox.displayNewAppointmentForm !== "function") {
                return JSON.stringify({ ok: false, error: "displayNewAppointmentForm not available on this client." });
            }
            const p = JSON.parse(paramsJson);
            const params = {};
            if (p.subject) params.subject = p.subject;
            if (p.location) params.location = p.location;
            if (p.body) params.body = p.body;
            if (p.requiredAttendees && p.requiredAttendees.length) params.requiredAttendees = p.requiredAttendees;
            if (p.optionalAttendees && p.optionalAttendees.length) params.optionalAttendees = p.optionalAttendees;
            if (p.startIso) params.start = new Date(p.startIso);
            if (p.endIso) params.end = new Date(p.endIso);
            try {
                mailbox.displayNewAppointmentForm(params);
                return JSON.stringify({ ok: true });
            } catch (e) {
                return JSON.stringify({ ok: false, error: String(e) });
            }
        },

        // ---- CREATE (full): set fields on the OPEN compose appointment ---------
        // Requires ReadWriteItem + compose activation. Can set a recurrence pattern.
        setComposeFields: async function (paramsJson) {
            const ready = await ensureReady();
            const out = { results: [], errors: [] };
            if (!ready.available) { out.errors.push("No Office host detected."); return JSON.stringify(out); }

            const item = safe(function () { return Office.context.mailbox.item; });
            if (!item || typeof item.subject !== "object" || typeof item.subject.setAsync !== "function") {
                out.errors.push("Not in appointment compose mode. Create a new appointment, then open this add-in from its ribbon.");
                return JSON.stringify(out);
            }

            const p = JSON.parse(paramsJson);

            async function record(name, setterFactory) {
                const res = await setAsyncP(setterFactory);
                out.results.push({
                    field: name,
                    status: res && res.status ? String(res.status) : "unknown",
                    error: res && res.error ? (res.error.message || String(res.error)) : null
                });
            }

            if (p.subject != null) await record("subject", function (cb) { item.subject.setAsync(p.subject, cb); });
            if (p.location != null) await record("location", function (cb) { item.location.setAsync(p.location, cb); });
            if (p.requiredAttendees && p.requiredAttendees.length) {
                await record("requiredAttendees", function (cb) { item.requiredAttendees.setAsync(p.requiredAttendees, cb); });
            }

            if (p.recurrence) {
                try {
                    const pattern = buildPattern(p.recurrence);
                    await record("recurrence", function (cb) { item.recurrence.setAsync(pattern, cb); });
                } catch (e) {
                    out.errors.push("Recurrence build failed: " + String(e));
                }
            } else {
                if (p.startIso) await record("start", function (cb) { item.start.setAsync(new Date(p.startIso), cb); });
                if (p.endIso) await record("end", function (cb) { item.end.setAsync(new Date(p.endIso), cb); });
            }

            return JSON.stringify(out);
        }
    };
})();
