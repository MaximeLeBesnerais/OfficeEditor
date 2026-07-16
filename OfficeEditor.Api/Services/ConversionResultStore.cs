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
    private readonly MemoryCacheEntryOptions _cacheOptions;

    public InMemoryConversionResultStore(IMemoryCache cache)
    {
        _cache = cache;
        _cacheOptions = new MemoryCacheEntryOptions()
            .SetSlidingExpiration(TimeSpan.FromMinutes(30));
    }

    public Guid Store(byte[] bytes, string contentType, string fileName)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(contentType);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        var id = Guid.NewGuid();
        _cache.Set(id, new StoredResult(bytes, contentType, fileName), _cacheOptions);
        return id;
    }

    public bool TryGet(Guid id, out StoredResult? result)
    {
        return _cache.TryGetValue(id, out result);
    }
}
