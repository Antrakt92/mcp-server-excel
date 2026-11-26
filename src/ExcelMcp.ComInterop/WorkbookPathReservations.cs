using System.Collections;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>Path reservations shared by controllers and workbook session owners.</summary>
internal sealed class WorkbookPathReservations<T>(
    Func<T, T, bool>? sameOwner = null, Func<string, WorkbookFileIdentity>? readIdentity = null)
    : IEnumerable<KeyValuePair<string, T>>
{
    private sealed class Claim(T owner)
    {
        internal T Owner = owner;
        internal string? FileKey;
    }

    private readonly Lock _sync = new();
    private readonly Dictionary<string, Claim> _claims = new(StringComparer.OrdinalIgnoreCase);
    private readonly Func<T, T, bool> _sameOwner = sameOwner ?? EqualityComparer<T>.Default.Equals;
    private readonly Func<string, WorkbookFileIdentity> _readIdentity = readIdentity ?? WorkbookFileIdentity.Read;
    private long _claimRevision;

    internal sealed class PublicationSweep(WorkbookPathReservations<T> reservations)
    {
        internal long Revision { get; set; } = -1;

        internal bool TryPublish(string key, T owner, Predicate<T> isProvisionalOwner) =>
            reservations.TryPublish(key, owner, isProvisionalOwner, this);
    }

    internal PublicationSweep BeginPublicationSweep() => new(this);

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
                _claimRevision++;
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
            _claimRevision++;
            return true;
        }
    }

    internal bool TryReserveAlias(string key, T value, Predicate<T> mayShare, out bool added)
    {
        lock (_sync)
        {
            var identity = Prepare(key);
            added = false;
            if (Matching(identity).Any(entry => !mayShare(entry.Value.Owner)))
                return false;
            if (_claims.ContainsKey(identity.Path))
                return true;

            // A native Save As may replace this directory entry rather than the
            // original file. Remember the destination separately before dispatch.
            var claim = new Claim(value);
            Remember(claim, identity);
            _claims.Add(identity.Path, claim);
            _claimRevision++;
            added = true;
            return true;
        }
    }

    internal bool TryPublish(string key, T value, Predicate<T> isProvisionalOwner) =>
        TryPublish(key, value, isProvisionalOwner, null);

    private bool TryPublish(string key, T value, Predicate<T> isProvisionalOwner, PublicationSweep? sweep)
    {
        lock (_sync)
        {
            var previousRevision = sweep?.Revision;
            if (sweep is not null)
                sweep.Revision = -1; // An unreadable candidate invalidates the current sweep.
            var identity = Resolve(key);
            // A background cohort can reuse one scan only for a steady, exact
            // claim. Re-read each reported path; any new owner or file ID must
            // check every physical claim before publication.
            if (identity.FileKey is not null)
            {
                if (sweep is not null && previousRevision == _claimRevision
                    && StringComparer.OrdinalIgnoreCase.Equals(key, identity.Path)
                    && _claims.TryGetValue(identity.Path, out var previous)
                    && EqualityComparer<T>.Default.Equals(previous.Owner, value)
                    && StringComparer.OrdinalIgnoreCase.Equals(previous.FileKey, identity.FileKey))
                    sweep.Revision = _claimRevision;
                else
                    RefreshClaims(sweep);
            }
            else if (sweep is not null && WorkbookLocation.IsRemote(key) && previousRevision == _claimRevision)
                sweep.Revision = _claimRevision;
            var matches = Matching(identity).ToArray();
            if (matches.Any(entry => !_sameOwner(entry.Value.Owner, value)
                && !isProvisionalOwner(entry.Value.Owner)))
                return false;

            // Excel can report another hard-link name for an open reserved file.
            // Transfer the worker's provisional claim without touching a separate
            // in-flight Save As claim that already belongs to the same session.
            foreach (var entry in matches)
                if (isProvisionalOwner(entry.Value.Owner))
                {
                    entry.Value.Owner = value;
                    _claimRevision++;
                }

            if (!_claims.TryGetValue(identity.Path, out var claim))
            {
                _claims.Add(identity.Path, claim = new(value));
                _claimRevision++;
            }
            else if (!EqualityComparer<T>.Default.Equals(claim.Owner, value)
                || identity.FileKey is not null && !StringComparer.OrdinalIgnoreCase.Equals(claim.FileKey, identity.FileKey))
                _claimRevision++;
            claim.Owner = value;
            Remember(claim, identity);
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
            _claimRevision++;
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
            _claimRevision++;
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
            if (!_claims.TryGetValue(key, out var claim) || !EqualityComparer<T>.Default.Equals(claim.Owner, entry.Value)
                || !_claims.Remove(key))
                return false;
            _claimRevision++;
            return true;
        }
    }

    internal void Clear()
    {
        lock (_sync)
        {
            _claims.Clear();
            _claimRevision++;
        }
    }

    private WorkbookFileIdentity Prepare(string key)
    {
        var identity = Resolve(key);
        // A missing destination or a URL can only match a stable path key.
        // An existing file can also match a claim through its physical file ID.
        if (identity.FileKey is null)
            return identity;

        // A destination may have been created since it was reserved. Keep the stable path
        // claim while learning its file ID before another hard-link name can be admitted.
        RefreshClaims(null);
        return identity;
    }

    private void RefreshClaims(PublicationSweep? sweep)
    {
        _claimRevision++;
        var allLocalFilesExist = true;
        foreach (var entry in _claims)
        {
            var identity = Resolve(entry.Key);
            Remember(entry.Value, identity);
            // A pending destination could appear before the next worker reports.
            if (identity.FileKey is null && !WorkbookLocation.IsRemote(entry.Key))
                allLocalFilesExist = false;
        }
        if (sweep is not null && allLocalFilesExist)
            sweep.Revision = _claimRevision;
    }

    private IEnumerable<KeyValuePair<string, Claim>> Matching(WorkbookFileIdentity identity) =>
        _claims.Where(entry => StringComparer.OrdinalIgnoreCase.Equals(entry.Key, identity.Path)
            || identity.FileKey is not null && StringComparer.OrdinalIgnoreCase.Equals(entry.Value.FileKey, identity.FileKey));

    private WorkbookFileIdentity Resolve(string key) => WorkbookLocation.IsRemote(key)
        ? new(key, null) : _readIdentity(WorkbookLocation.Normalize(key));

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
