namespace OfficeEditor.Api.Services;

internal static class ApiResourceLimits
{
    internal const long MaxRequestBodyBytes = 64L * 1024 * 1024;
    internal const long MemoryCacheBytes = 512L * 1024 * 1024;
    internal const long RenderedPagesPerDeckBytes = 32L * 1024 * 1024;
}
