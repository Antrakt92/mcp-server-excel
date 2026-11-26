using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;

namespace Sbroenne.ExcelMcp.Service;

internal sealed partial class ExcelController
{
    internal const int SettledCloseRetention = 512;
    private readonly LinkedList<string> _settledCloseOrder = new();
    private readonly HashSet<string> _settledCloses = new(StringComparer.Ordinal);
    private readonly Dictionary<ExcelWorkerClient, HashSet<string>> _unsettledCloses = new();

    // These methods run under the controller's shutdown lock. A successful
    // close reply can come before a snapshot drops the row. Its tombstone
    // must still suppress that row even if later closes fill the history.
    private void RetainConfirmedClose(ExcelWorkerClient worker, string sessionId)
    {
        if (_settledCloses.Contains(sessionId))
            return;
        if (!_workers.TryGetValue(worker.Registration.PipeName, out var current) || current != worker)
        {
            RetainSettledClose(sessionId);
            return;
        }

        if (!_unsettledCloses.TryGetValue(worker, out var unsettled))
            _unsettledCloses.Add(worker, unsettled = new(StringComparer.Ordinal));
        unsettled.Add(sessionId);
    }

    private void RetainSettledClose(string sessionId)
    {
        if (_settledCloses.Add(sessionId))
        {
            _settledCloseOrder.AddLast(sessionId);
            TrimSettledCloses();
        }
    }

    private void SettleCloseReceipts(ExcelWorkerClient worker, WorkerSnapshot? settledSnapshot)
    {
        if (_unsettledCloses.TryGetValue(worker, out var unsettled))
        {
            var liveIds = settledSnapshot?.Sessions.Select(session => session.SessionId).ToHashSet(StringComparer.Ordinal);
            foreach (var sessionId in unsettled.ToArray())
            {
                if (liveIds is not null && (liveIds.Contains(sessionId) || _sessions.ContainsKey(sessionId)))
                    continue;
                unsettled.Remove(sessionId);
                RetainSettledClose(sessionId);
            }
            if (unsettled.Count == 0)
                _unsettledCloses.Remove(worker);
        }
    }

    private void RecordLostSessionCloseAttempt(string sessionId, ExcelProcessIdentity? excel, ServiceResponse response)
    {
        lock (_shutdownLock)
        {
            // A retry can hold the old result while a concurrent close evicts it.
            if (!_closedSessions.TryGetValue(sessionId, out var current) || current.Excel != excel
                || current.Reason is not (SessionEnd.WorkerExited or SessionEnd.ExcelExited))
                return;
            _closeHistory.Record(new(sessionId, current.Excel), response, onlyIfMissing: true, owned: current.Owned);
            TrimSettledCloses();
        }
    }

    private void TrimSettledCloses()
    {
        if (_settledCloseOrder.Count <= SettledCloseRetention)
            return;

        // Recent diagnostic entries can still report their exact release status.
        var diagnosticIds = _closeHistory.RetainedSessionIds();
        var toRemove = _settledCloseOrder.Count - SettledCloseRetention - diagnosticIds.Count(_settledCloses.Contains);
        var node = _settledCloseOrder.First;
        while (toRemove > 0 && node is not null)
        {
            var next = node.Next;
            if (!diagnosticIds.Contains(node.Value))
            {
                _closedSessions.TryRemove(node.Value, out _);
                _settledCloses.Remove(node.Value);
                _settledCloseOrder.Remove(node);
                toRemove--;
            }
            node = next;
        }
    }
}
