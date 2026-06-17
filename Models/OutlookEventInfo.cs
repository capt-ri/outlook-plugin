using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookEventPlugin.Models;

/// <summary>
/// Strongly-typed view of the JSON produced by window.officeInterop.getSelectedEvent().
/// </summary>
public class OutlookEventInfo
{
    // Host / availability
    public bool HostAvailable { get; set; }
    public string? Host { get; set; }
    public string? Platform { get; set; }

    // Item identity
    public string? ItemType { get; set; }
    public bool IsAppointment { get; set; }

    /// <summary>EWS (Exchange) id of THIS item — the master id when a series master
    /// is selected, or the occurrence id when an instance is selected.</summary>
    public string? ExchangeId { get; set; }

    /// <summary>EWS id of the parent series master; null when this item is the
    /// master itself or a non-recurring appointment.</summary>
    public string? SeriesId { get; set; }

    // Core fields
    public string? Subject { get; set; }
    public string? Location { get; set; }
    public string? Organizer { get; set; }
    public List<string>? RequiredAttendees { get; set; }
    public List<string>? RequiredAttendeeEmails { get; set; }
    public List<string>? OptionalAttendees { get; set; }
    public string? Start { get; set; }
    public string? End { get; set; }

    // Recurrence
    public bool IsRecurring { get; set; }
    public string? Classification { get; set; }
    public string? RecurrenceType { get; set; }
    public JsonElement? RecurrenceProperties { get; set; }
    public string? RecurrenceTimeZone { get; set; }
    public string? SeriesStartDate { get; set; }
    public string? SeriesEndDate { get; set; }
    public string? SeriesStartTime { get; set; }
    public string? SeriesEndTime { get; set; }
    public int? SeriesDurationMinutes { get; set; }

    public List<string>? Errors { get; set; }

    [JsonIgnore]
    public bool HasErrors => Errors is { Count: > 0 };

    [JsonIgnore]
    public string RecurrencePropertiesText =>
        RecurrenceProperties is { ValueKind: not JsonValueKind.Null and not JsonValueKind.Undefined } el
            ? JsonSerializer.Serialize(el, new JsonSerializerOptions { WriteIndented = true })
            : "—";
}
