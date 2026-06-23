using System.Collections.Generic;
using System.Threading.Tasks;
using OutlookEventPlugin.Models;

namespace OutlookEventPlugin.Services;

public interface ISalesforceSync
{
    /// <summary>True = the mock backend pretends every known item is already synced.</summary>
    bool SimulateSynced { get; set; }

    Task<SyncStatus> GetStatusAsync(string? id, string kind);
    Task<SyncOutcome> SyncAsync(string? id, string kind, string? subject);
}

/// <summary>
/// In-memory stand-in for the Salesforce backend so the UI can be built and tested
/// without a real integration. Keyed by the item's id (message id / Exchange id).
/// </summary>
public class MockSalesforceSync : ISalesforceSync
{
    // Items "synced" during this session (simulates existing Salesforce records).
    private readonly Dictionary<string, SyncStatus> _store = new();

    /// <summary>Default ON so the synced banner shows immediately for testing.</summary>
    public bool SimulateSynced { get; set; } = true;

    public Task<SyncStatus> GetStatusAsync(string? id, string kind)
    {
        if (string.IsNullOrEmpty(id))
            return Task.FromResult(new SyncStatus { IsSynced = false, Unknown = true });

        if (_store.TryGetValue(id, out var existing))
            return Task.FromResult(existing);

        if (SimulateSynced)
            return Task.FromResult(Synthesize(id, kind));

        return Task.FromResult(new SyncStatus { IsSynced = false });
    }

    public Task<SyncOutcome> SyncAsync(string? id, string kind, string? subject)
    {
        if (string.IsNullOrEmpty(id))
        {
            return Task.FromResult(new SyncOutcome
            {
                Ok = false,
                Message = "No id available to sync (open or save the item first).",
                Status = new SyncStatus { IsSynced = false, Unknown = true }
            });
        }

        // Already synced? (either previously synced this session, or the backend says so)
        var current = _store.TryGetValue(id, out var s) ? s
            : (SimulateSynced ? Synthesize(id, kind) : new SyncStatus { IsSynced = false });

        if (current.IsSynced)
        {
            _store[id] = current;
            return Task.FromResult(new SyncOutcome
            {
                Ok = false,
                AlreadySynced = true,
                Message = $"This {kind} is already synced to Salesforce ({current.ObjectType} {current.SalesforceId}).",
                Status = current
            });
        }

        // Not synced yet -> create a (fake) record and remember it.
        var created = Synthesize(id, kind);
        _store[id] = created;
        return Task.FromResult(new SyncOutcome
        {
            Ok = true,
            AlreadySynced = false,
            Message = $"Synced to Salesforce as {created.ObjectType} {created.SalesforceId}.",
            Status = created
        });
    }

    private static SyncStatus Synthesize(string id, string kind)
    {
        // Deterministic-ish synthetic record derived from the id.
        var h = (uint)id.GetHashCode();
        var sfId = "00U" + h.ToString("X8") + "AB";
        return new SyncStatus
        {
            IsSynced = true,
            ObjectType = kind == "event" ? "Event" : "EmailMessage",
            SalesforceId = sfId,
            SyncedAtUtc = "2026-06-15T09:30:00Z",
            RecordUrl = $"https://example.lightning.force.com/lightning/r/{sfId}/view"
        };
    }
}
