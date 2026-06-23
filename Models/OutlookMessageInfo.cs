using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace OutlookEventPlugin.Models;

/// <summary>
/// Strongly-typed view of the JSON produced by window.officeInterop.getSelectedMessage().
/// </summary>
public class OutlookMessageInfo
{
    public bool HostAvailable { get; set; }
    public string? Host { get; set; }
    public string? Platform { get; set; }
    public string? Mode { get; set; }

    public string? ItemType { get; set; }
    public bool IsMessage { get; set; }

    /// <summary>EWS (Exchange) id of the message; null in compose until saved.</summary>
    public string? ItemId { get; set; }
    public string? Subject { get; set; }
    public string? NormalizedSubject { get; set; }
    public string? From { get; set; }
    public string? Sender { get; set; }
    public List<string>? To { get; set; }
    public List<string>? Cc { get; set; }
    public List<string>? Bcc { get; set; }

    public string? ConversationId { get; set; }
    public string? InternetMessageId { get; set; }
    public string? DateTimeCreated { get; set; }
    public string? DateTimeModified { get; set; }

    public List<MessageAttachment>? Attachments { get; set; }

    public List<string>? Errors { get; set; }

    [JsonIgnore]
    public bool HasErrors => Errors is { Count: > 0 };
}

public class MessageAttachment
{
    public string? Name { get; set; }
    public int Size { get; set; }
    public string? Type { get; set; }
    public bool IsInline { get; set; }
}
