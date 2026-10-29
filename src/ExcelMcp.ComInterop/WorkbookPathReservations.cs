using System.Collections;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>Path reservations shared by controllers and workbook session owners.</summary>
internal sealed class WorkbookPathReservations<T>(Func<T, T, bool>? sameOwner = null) : IEnumerable<KeyValuePair<string, T>>
{
    private sealed class Claim(T owner)
    {
        internal T Owner = owner;
        internal string? FileKey;
    }

    private readonly Lock _sync = new();
    private readonly Dictionary<string, Claim> _claims = new(StringComparer.OrdinalIgnoreCase);
    private readonly Func<T, T, bool> _sameOwner = sameOwner ?? EqualityComparer<T>.Default.Equals;

    internal T this[string key]
    {
        get => TryGetValue(key, out var value) ? value : throw new KeyNotFoundException("No workbook path reservation exists.");
        set
        {
            lock (_sync)
            {
                var identity = Prepare(key);
                if (Matching(identity).Any(entry => !_sameOwner(entry.Value.Owner, value)))
                    throw new InvalidOperationException("The workbook's physical file is reserved by another owner.");
                if (!_claims.TryGetValue(identity.Path, out var claim))
                    _claims.Add(identity.Path, claim = new(value));
                claim.Owner = value;
                Remember(claim, identity);
            }
        }
    }

    internal IReadOnlyList<T> Values => Snapshot().Select(entry => entry.Value).ToArray();
    internal bool ContainsKey(string key) => TryGetValue(key, out _);

    internal bool TryGetValue(string key, out T value)
    {
        lock (_sync)
        {
            var matches = Matching(Prepare(key)).ToArray();
            var match = matches.FirstOrDefault();
            if (match.Value is not null && matches.Any(entry => !_sameOwner(entry.Value.Owner, match.Value.Owner)))
                throw new InvalidOperationException("The workbook's physical file has conflicting ownership records.");
            value = match.Value is null ? default! : match.Value.Owner;
            return match.Value is not null;
        }
    }

    internal bool TryAdd(string key, T value)
    {
        lock (_sync)
        {
            var identity = Prepare(key);
            if (Matching(identity).Any()) return false;
            var claim = new Claim(value);
            Remember(claim, identity);
            _claims.Add(identity.Path, claim);
            return true;
        }
    }

    internal bool TryUpdate(string key, T value, T comparison)
    {
        lock (_sync)
        {
            var identity = Prepare(key);
            if (!_claims.TryGetValue(identity.Path, out var claim)
                || !EqualityComparer<T>.Default.Equals(claim.Owner, comparison)) return false;
            if (Matching(identity).Any(entry => !ReferenceEquals(entry.Value, claim) && !_sameOwner(entry.Value.Owner, value)))
                throw new InvalidOperationException("The workbook's physical file is reserved by another owner.");
            claim.Owner = value;
            Remember(claim, identity);
            return true;
        }
    }

    internal bool TryRemove(string key, out T value)
    {
        lock (_sync)
        {
            // Known stable path keys can be released even if the file was deleted or became inaccessible.
            if (!_claims.ContainsKey(key))
            {
                var identity = Resolve(key);
                key = _claims.ContainsKey(identity.Path) ? identity.Path : Matching(identity).FirstOrDefault().Key;
                if (key is null) { value = default!; return false; }
            }
            if (!_claims.Remove(key, out var claim)) { value = default!; return false; }
            value = claim.Owner;
            return true;
        }
    }

    internal bool TryRemove(KeyValuePair<string, T> entry)
    {
        lock (_sync)
        {
            var identity = _claims.ContainsKey(entry.Key) ? null : Resolve(entry.Key);
            var key = identity is null ? entry.Key : _claims.ContainsKey(identity.Path)
                ? identity.Path : Matching(identity).FirstOrDefault(match => EqualityComparer<T>.Default.Equals(match.Value.Owner, entry.Value)).Key;
            if (key is null) return false;
            return _claims.TryGetValue(key, out var claim) && EqualityComparer<T>.Default.Equals(claim.Owner, entry.Value)
                && _claims.Remove(key);
        }
    }

    internal void Clear() { lock (_sync) _claims.Clear(); }

    private WorkbookFileIdentity Prepare(string key)
    {
        var identity = Resolve(key);
        // A destination may have been created since it was reserved. Keep the stable path
        // claim while learning its file ID before another hard-link name can be admitted.
        foreach (var entry in _claims) Remember(entry.Value, Resolve(entry.Key));
        return identity;
    }

    private IEnumerable<KeyValuePair<string, Claim>> Matching(WorkbookFileIdentity identity) =>
        _claims.Where(entry => StringComparer.OrdinalIgnoreCase.Equals(entry.Key, identity.Path)
            || identity.FileKey is not null && StringComparer.OrdinalIgnoreCase.Equals(entry.Value.FileKey, identity.FileKey));

    private static WorkbookFileIdentity Resolve(string key) => WorkbookLocation.IsRemote(key)
        ? new(key, null) : WorkbookFileIdentity.Read(WorkbookLocation.Normalize(key));

    private static void Remember(Claim claim, WorkbookFileIdentity identity)
    {
        if (identity.FileKey is not null)
        {
            // A Save As can atomically replace a directory entry. The new file belongs
            // to this stable path claim; hard links to its former file become independent.
            claim.FileKey = identity.FileKey;
        }
    }

    private KeyValuePair<string, T>[] Snapshot()
    {
        lock (_sync) return _claims.Select(entry => new KeyValuePair<string, T>(entry.Key, entry.Value.Owner)).ToArray();
    }

    public IEnumerator<KeyValuePair<string, T>> GetEnumerator() => ((IEnumerable<KeyValuePair<string, T>>)Snapshot()).GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
