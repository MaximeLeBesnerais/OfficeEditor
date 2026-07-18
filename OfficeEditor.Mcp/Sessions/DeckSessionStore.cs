namespace OfficeEditor.Mcp.Sessions;

/// <summary>
/// Server-side state for one uploaded deck, mirroring the API's DeckSession semantics
/// (OfficeEditor.Api/Services/DeckSessionStore.cs) without referencing the ASP.NET project.
/// Deck bytes are kept in memory; a session-owned temp directory is reserved for
/// session artifacts and is swept on expiry/eviction.
/// </summary>
public sealed class McpDeckSession
{
    public required Guid Handle { get; init; }
    public required byte[] DeckBytes { get; set; }
    public required string FileName { get; init; }

    /// <summary>Re-derived on every edit (slide ops can change the count).</summary>
    public required int SlideCount { get; set; }

    /// <summary>Monotonic revision, bumped by each applied edit batch.</summary>
    public int Revision { get; set; }

    /// <summary>Sliding-lifetime anchor: refreshed on every access.</summary>
    public DateTimeOffset LastAccessUtc { get; set; }

    /// <summary>Session-owned scratch directory; deleted on expiry/eviction/shutdown.</summary>
    public required string TempDirectory { get; init; }
}

/// <summary>
/// In-memory deck-session map with a 30-minute sliding lifetime, mirroring
/// <c>InMemoryDeckSessionStore</c> (upload → handle; every access slides the expiry;
/// eviction sweeps the session temp dir). Hand-rolled over a Dictionary instead of
/// IMemoryCache to keep the MCP host free of third-party packages.
/// Expired sessions are reclaimed by a periodic sweep timer AND lazily on access;
/// all remaining sessions are swept on <see cref="Dispose"/> (host shutdown).
/// A lock serializes access because the sweep timer fires on a thread-pool thread.
/// </summary>
public sealed class DeckSessionStore : IDisposable
{
    public static readonly TimeSpan DefaultSlidingLifetime = TimeSpan.FromMinutes(30);
    private static readonly TimeSpan SweepInterval = TimeSpan.FromMinutes(5);

    private readonly Dictionary<Guid, McpDeckSession> _sessions = new();
    private readonly object _gate = new();
    private readonly TimeSpan _slidingLifetime;
    private readonly Func<DateTimeOffset> _clock;
    private readonly Timer? _sweepTimer;
    private readonly string _rootDirectory;
    private bool _disposed;

    public DeckSessionStore(
        Func<DateTimeOffset>? clock = null,
        TimeSpan? slidingLifetime = null,
        bool enableSweepTimer = true)
    {
        _clock = clock ?? (() => DateTimeOffset.UtcNow);
        _slidingLifetime = slidingLifetime ?? DefaultSlidingLifetime;
        // Per-process root: stale dirs from dead hosts never collide, and Dispose
        // can remove the whole tree without touching other hosts' sessions.
        _rootDirectory = Path.Combine(
            Path.GetTempPath(), "officeeditor-mcp", Environment.ProcessId.ToString());
        if (enableSweepTimer)
        {
            _sweepTimer = new Timer(_ => SweepExpired(), null, SweepInterval, SweepInterval);
        }
    }

    public int Count
    {
        get { lock (_gate) { return _sessions.Count; } }
    }

    public McpDeckSession Store(byte[] deckBytes, string fileName, int slideCount)
    {
        ArgumentNullException.ThrowIfNull(deckBytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(slideCount);

        var session = new McpDeckSession
        {
            Handle = Guid.NewGuid(),
            DeckBytes = deckBytes,
            FileName = fileName,
            SlideCount = slideCount,
            Revision = 0,
            LastAccessUtc = _clock(),
            TempDirectory = Path.Combine(_rootDirectory, Guid.NewGuid().ToString("N"))
        };

        Directory.CreateDirectory(session.TempDirectory);
        lock (_gate)
        {
            _sessions[session.Handle] = session;
        }
        return session;
    }

    public bool TryGet(Guid handle, out McpDeckSession? session)
    {
        List<string> expiredDirectories;
        lock (_gate)
        {
            expiredDirectories = SweepExpiredLocked();
            if (_sessions.TryGetValue(handle, out session))
            {
                session.LastAccessUtc = _clock();
            }
            else
            {
                session = null;
            }
        }

        // Sweep outside the lock: deletion is slow IO and must not block lookups.
        DeleteDirectories(expiredDirectories);
        return session is not null;
    }

    /// <summary>Replaces the deck bytes after an applied edit batch and bumps the revision.</summary>
    public void UpdateBytes(Guid handle, byte[] deckBytes, int slideCount)
    {
        ArgumentNullException.ThrowIfNull(deckBytes);

        lock (_gate)
        {
            if (!_sessions.TryGetValue(handle, out var session))
            {
                throw new KeyNotFoundException($"Unknown or expired deck handle '{handle}'.");
            }
            session.DeckBytes = deckBytes;
            session.SlideCount = slideCount;
            session.Revision++;
            session.LastAccessUtc = _clock();
        }
    }

    /// <summary>Sweeps expired sessions and their temp directories.</summary>
    public void SweepExpired()
    {
        List<string> directories;
        lock (_gate)
        {
            directories = SweepExpiredLocked();
        }
        DeleteDirectories(directories);
    }

    private List<string> SweepExpiredLocked()
    {
        var now = _clock();
        var directories = new List<string>();
        foreach (var (handle, session) in _sessions.ToArray())
        {
            if (now - session.LastAccessUtc > _slidingLifetime)
            {
                _sessions.Remove(handle);
                directories.Add(session.TempDirectory);
            }
        }
        return directories;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;
        _sweepTimer?.Dispose();

        List<string> directories;
        lock (_gate)
        {
            directories = _sessions.Values.Select(s => s.TempDirectory).ToList();
            _sessions.Clear();
        }
        DeleteDirectories(directories);

        try
        {
            if (Directory.Exists(_rootDirectory) && !Directory.EnumerateFileSystemEntries(_rootDirectory).Any())
            {
                Directory.Delete(_rootDirectory);
            }
        }
        catch (IOException) { /* best-effort cleanup; the OS temp cleaner is the backstop */ }
        catch (UnauthorizedAccessException) { /* best-effort cleanup */ }
    }

    private static void DeleteDirectories(IEnumerable<string> directories)
    {
        foreach (var directory in directories)
        {
            try
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
            catch (IOException) { /* best-effort cleanup; the OS temp cleaner is the backstop */ }
            catch (UnauthorizedAccessException) { /* best-effort cleanup */ }
        }
    }
}
