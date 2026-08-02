using Microsoft.Extensions.Caching.Memory;

namespace OfficeEditor.Api.Services;

public sealed record StoredResult(
    byte[] Bytes,
    string ContentType,
    string FileName);

public interface IConversionResultStore
{
    Guid Store(byte[] bytes, string contentType, string fileName);
    bool TryGet(Guid id, out StoredResult? result);
}

public sealed class InMemoryConversionResultStore : IConversionResultStore
{
    private readonly IMemoryCache _cache;

    public InMemoryConversionResultStore(IMemoryCache cache)
    {
        _cache = cache;
    }

    public Guid Store(byte[] bytes, string contentType, string fileName)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(contentType);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        var id = Guid.NewGuid();
        _cache.Set(id, new StoredResult(bytes, contentType, fileName), CreateCacheOptions(bytes.LongLength));
        return id;
    }

    public bool TryGet(Guid id, out StoredResult? result)
    {
        return _cache.TryGetValue(id, out result);
    }

    internal static MemoryCacheEntryOptions CreateCacheOptions(long byteLength) =>
        new MemoryCacheEntryOptions()
            .SetSlidingExpiration(TimeSpan.FromMinutes(30))
            .SetSize(Math.Max(1, byteLength));
}
