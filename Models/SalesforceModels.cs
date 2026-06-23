namespace OutlookEventPlugin.Models;

/// <summary>Sync state of an item as reported by the (mock) Salesforce backend.</summary>
public class SyncStatus
{
    public bool IsSynced { get; set; }
    /// <summary>Salesforce object type, e.g. "Event" or "Task".</summary>
    public string? ObjectType { get; set; }
    /// <summary>Salesforce record id (synthetic in the mock).</summary>
    public string? SalesforceId { get; set; }
    public string? SyncedAtUtc { get; set; }
    public string? RecordUrl { get; set; }
    /// <summary>Set when no usable key (id) was available to check.</summary>
    public bool Unknown { get; set; }
}

/// <summary>Result of attempting to sync an item.</summary>
public class SyncOutcome
{
    public bool Ok { get; set; }
    public bool AlreadySynced { get; set; }
    public string? Message { get; set; }
    public SyncStatus Status { get; set; } = new();
}
