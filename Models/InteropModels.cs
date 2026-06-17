using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace OutlookEventPlugin.Models;

/// <summary>Host availability + current activation mode (read/compose/none).</summary>
public class OfficeContext
{
    public bool HostAvailable { get; set; }
    public string? Host { get; set; }
    public string? Platform { get; set; }
    public string Mode { get; set; } = "none";
}

/// <summary>Result of officeInterop.displayNewAppointment.</summary>
public class DisplayResult
{
    public bool Ok { get; set; }
    public string? Error { get; set; }
}

/// <summary>Result of officeInterop.saveAndGetExchangeId.</summary>
public class SaveIdResult
{
    public bool Ok { get; set; }
    public string? ExchangeId { get; set; }
    public string? Error { get; set; }
}

/// <summary>Result of officeInterop.setComposeFields.</summary>
public class ComposeResult
{
    public List<ComposeFieldResult> Results { get; set; } = new();
    public List<string> Errors { get; set; } = new();

    [JsonIgnore]
    public bool AllSucceeded =>
        Errors.Count == 0 && Results.TrueForAll(r => r.Status == "succeeded");
}

public class ComposeFieldResult
{
    public string Field { get; set; } = "";
    public string Status { get; set; } = "";
    public string? Error { get; set; }
}

/// <summary>Shared state for the create form (read &amp; compose panels).</summary>
public class CreateFormModel
{
    public string Subject { get; set; } = "Team sync";
    public string Location { get; set; } = "";
    public string Attendees { get; set; } = "";
    public System.DateTime Start { get; set; }
    public System.DateTime End { get; set; }

    /// <summary>Local-time ISO (no zone) so new Date(iso) in JS parses as local.</summary>
    public string StartIso => Start.ToString("yyyy-MM-ddTHH:mm");
    public string EndIso => End.ToString("yyyy-MM-ddTHH:mm");
}
