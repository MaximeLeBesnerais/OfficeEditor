namespace OfficeEditor.Api.Services;

public interface IConversionService
{
    Task<ConversionResult> ConvertAsync(ConversionRequest request, CancellationToken ct = default);
}
