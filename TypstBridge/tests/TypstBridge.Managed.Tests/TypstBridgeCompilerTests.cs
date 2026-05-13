using System.Buffers.Binary;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using TypstBridge.Managed;
using TypstBridge.Managed.Models;

namespace TypstBridge.Managed.Tests;

public sealed class TypstBridgeCompilerTests
{
    private const uint Ok = 0;
    private const uint Compile = 2;

    [Fact]
    public void Compile_WithNullRequest_ThrowsArgumentNullExceptionBeforeNativeCalls()
    {
        TypstBridgeCompiler compiler = new();

        var exception = Assert.Throws<ArgumentNullException>(() => compiler.Compile(null!));

        Assert.Equal("request", exception.ParamName);
    }

    [Fact]
    public void CompileRequest_DefaultsAndNullFontPaths_AreStable()
    {
        TypstCompileRequest request = new("hello", "/work");

        Assert.Equal("hello", request.Source);
        Assert.Equal("/work", request.WorkingDirectory);
        Assert.Equal("main.typ", request.RootFileName);
        Assert.Empty(request.FontPaths);
        Assert.Equal(TypstOutputFormat.Pdf, request.OutputFormat);
        Assert.Equal(144.0, request.Ppi);
    }

    [Fact]
    public void CompileRequest_CopiesFontPathsAndPreservesOptionCombinations()
    {
        List<string> fontPaths = ["/fonts/a", "/fonts/b"];

        TypstCompileRequest request = new(
            source: "hello",
            workingDirectory: "/work",
            rootFileName: "slide.typ",
            fontPaths: fontPaths,
            outputFormat: TypstOutputFormat.Png,
            ppi: 192.5);
        fontPaths.Add("/fonts/mutated");

        Assert.Equal("slide.typ", request.RootFileName);
        Assert.Equal(new[] { "/fonts/a", "/fonts/b" }, request.FontPaths);
        Assert.Equal(TypstOutputFormat.Png, request.OutputFormat);
        Assert.Equal(192.5, request.Ppi);
    }

    [Theory]
    [InlineData("source")]
    [InlineData("workingDirectory")]
    [InlineData("rootFileName")]
    public void CompileRequest_NullRequiredStrings_ThrowArgumentNullException(string parameterName)
    {
        var exception = Assert.Throws<ArgumentNullException>(() => parameterName switch
        {
            "source" => new TypstCompileRequest(null!, "/work"),
            "workingDirectory" => new TypstCompileRequest("hello", null!),
            "rootFileName" => new TypstCompileRequest("hello", "/work", rootFileName: null!),
            _ => throw new InvalidOperationException()
        });

        Assert.Equal(parameterName, exception.ParamName);
    }

    [Fact]
    public void CreateNativeRequest_WithNoFonts_UsesZeroFontPointerAndRequestedFormat()
    {
        List<IntPtr> allocations = [];
        try
        {
            TypstCompileRequest request = new("héllo", "/work", outputFormat: TypstOutputFormat.Svg, ppi: 72.25);

            object nativeRequest = InvokeCompilerPrivate("CreateNativeRequest", request, allocations);

            Assert.Equal(TypstBridgeCompiler.SupportedAbiVersion, GetField<uint>(nativeRequest, "AbiVersion"));
            Assert.Equal((uint)TypstOutputFormat.Svg, GetField<uint>(nativeRequest, "OutputFormat"));
            Assert.Equal(72.25, GetField<double>(nativeRequest, "Ppi"));
            Assert.Equal(UIntPtr.Zero, GetField<UIntPtr>(nativeRequest, "FontPathsCount"));
            Assert.Equal(IntPtr.Zero, GetField<IntPtr>(nativeRequest, "FontPaths"));
            Assert.Equal((UIntPtr)Encoding.UTF8.GetByteCount("héllo"), GetField<UIntPtr>(nativeRequest, "SourceLen"));
        }
        finally
        {
            FreeAllocations(allocations);
        }
    }

    [Fact]
    public void CreateNativeRequest_WithFonts_AllocatesFontPathArrayAndLengths()
    {
        List<IntPtr> allocations = [];
        try
        {
            TypstCompileRequest request = new("hello", "/work", fontPaths: ["/fonts/one", "/fonts/deux"], outputFormat: TypstOutputFormat.Png);

            object nativeRequest = InvokeCompilerPrivate("CreateNativeRequest", request, allocations);

            IntPtr fontPathsPtr = GetField<IntPtr>(nativeRequest, "FontPaths");

            Assert.NotEqual(IntPtr.Zero, fontPathsPtr);
            Assert.Equal((UIntPtr)2, GetField<UIntPtr>(nativeRequest, "FontPathsCount"));
            Assert.Equal((uint)TypstOutputFormat.Png, GetField<uint>(nativeRequest, "OutputFormat"));

            Type nativeStringType = NativeType("NativeString");
            object firstFont = Marshal.PtrToStructure(fontPathsPtr, nativeStringType)!;
            Assert.Equal((UIntPtr)Encoding.UTF8.GetByteCount("/fonts/one"), GetField<UIntPtr>(firstFont, "ValueLen"));
        }
        finally
        {
            FreeAllocations(allocations);
        }
    }

    [Fact]
    public void AllocateUtf8_HandlesEmptyStringsAndNullTerminators()
    {
        List<IntPtr> allocations = [];
        try
        {
            object?[] args = ["", true, allocations, 123];

            IntPtr buffer = Assert.IsType<IntPtr>(InvokeCompilerPrivate("AllocateUtf8", args));

            Assert.NotEqual(IntPtr.Zero, buffer);
            Assert.Equal(0, Assert.IsType<int>(args[3]));
            Assert.Equal(0, Marshal.ReadByte(buffer));
        }
        finally
        {
            FreeAllocations(allocations);
        }
    }

    [Fact]
    public void CopyOutputs_WithNullPointerOrZeroCount_ReturnsEmptyOutputs()
    {
        var nullPointerResult = Assert.IsType<TypstOutputFile[]>(InvokeCompilerPrivate("CopyOutputs", IntPtr.Zero, 2));
        var zeroCountResult = Assert.IsType<TypstOutputFile[]>(InvokeCompilerPrivate("CopyOutputs", new IntPtr(1234), 0));

        Assert.Empty(nullPointerResult);
        Assert.Empty(zeroCountResult);
    }

    [Fact]
    public void CopyOutputs_CopiesFileNameAndReturnsEmptyDataForNullBuffer()
    {
        List<IntPtr> allocations = [];
        try
        {
            IntPtr fileName = AllocUtf8("page-001.svg", allocations, out int fileNameLength);
            IntPtr outputsPtr = AllocateNativeArray("NativeOutputItem", 1, allocations);
            object item = CreateNativeStruct("NativeOutputItem");
            SetField(item, "PageIndex", 3u);
            SetField(item, "FileNameUtf8", fileName);
            SetField(item, "FileNameLen", (UIntPtr)fileNameLength);
            SetField(item, "Data", IntPtr.Zero);
            SetField(item, "DataLen", (UIntPtr)42);
            Marshal.StructureToPtr(item, outputsPtr, false);

            var outputs = Assert.IsType<TypstOutputFile[]>(InvokeCompilerPrivate("CopyOutputs", outputsPtr, 1));

            TypstOutputFile output = Assert.Single(outputs);
            Assert.Equal(3u, output.PageIndex);
            Assert.Equal("page-001.svg", output.FileName);
            Assert.Empty(output.Data);
        }
        finally
        {
            FreeAllocations(allocations);
        }
    }

    [Fact]
    public void CopyDiagnostics_CopiesNullAndEmptyFileNames()
    {
        List<IntPtr> allocations = [];
        try
        {
            IntPtr diagnosticsPtr = AllocateNativeArray("NativeDiagnostic", 2, allocations);
            WriteNativeDiagnostic(diagnosticsPtr, 0, TypstDiagnosticSeverity.Error, "bad", null, 10, 20, allocations);
            WriteNativeDiagnostic(diagnosticsPtr, 1, TypstDiagnosticSeverity.Info, "note", string.Empty, 0, 0, allocations);

            var diagnostics = Assert.IsType<TypstDiagnostic[]>(InvokeCompilerPrivate("CopyDiagnostics", diagnosticsPtr, 2));

            Assert.Equal(2, diagnostics.Length);
            Assert.Equal(TypstDiagnosticSeverity.Error, diagnostics[0].Severity);
            Assert.Equal("bad", diagnostics[0].Message);
            Assert.Null(diagnostics[0].File);
            Assert.Equal(10u, diagnostics[0].Line);
            Assert.Equal(TypstDiagnosticSeverity.Info, diagnostics[1].Severity);
            Assert.Equal(string.Empty, diagnostics[1].File);
        }
        finally
        {
            FreeAllocations(allocations);
        }
    }

    [Fact]
    public void CopyResult_CopiesStatusMessageOutputsAndDiagnostics()
    {
        List<IntPtr> allocations = [];
        try
        {
            IntPtr data = AllocBytes([1, 2, 3], allocations);
            IntPtr fileName = AllocUtf8("page-001.png", allocations, out int fileNameLength);
            IntPtr outputsPtr = AllocateNativeArray("NativeOutputItem", 1, allocations);
            object outputItem = CreateNativeStruct("NativeOutputItem");
            SetField(outputItem, "PageIndex", 0u);
            SetField(outputItem, "FileNameUtf8", fileName);
            SetField(outputItem, "FileNameLen", (UIntPtr)fileNameLength);
            SetField(outputItem, "Data", data);
            SetField(outputItem, "DataLen", (UIntPtr)3);
            Marshal.StructureToPtr(outputItem, outputsPtr, false);

            IntPtr diagnosticsPtr = AllocateNativeArray("NativeDiagnostic", 1, allocations);
            WriteNativeDiagnostic(diagnosticsPtr, 0, TypstDiagnosticSeverity.Warning, "warn", "main.typ", 1, 2, allocations);
            IntPtr message = AllocUtf8("render warning", allocations, out int messageLength);
            IntPtr resultPtr = Marshal.AllocHGlobal(Marshal.SizeOf(NativeType("NativeCompileResult")));
            allocations.Add(resultPtr);
            object nativeResult = CreateNativeStruct("NativeCompileResult");
            SetField(nativeResult, "Status", Enum.ToObject(NativeType("TypstBridgeStatus"), 3u));
            SetField(nativeResult, "Outputs", outputsPtr);
            SetField(nativeResult, "OutputsCount", (UIntPtr)1);
            SetField(nativeResult, "Diagnostics", diagnosticsPtr);
            SetField(nativeResult, "DiagnosticsCount", (UIntPtr)1);
            SetField(nativeResult, "MessageUtf8", message);
            SetField(nativeResult, "MessageLen", (UIntPtr)messageLength);
            Marshal.StructureToPtr(nativeResult, resultPtr, false);

            var result = Assert.IsType<TypstCompileResult>(InvokeCompilerPrivate("CopyResult", resultPtr));

            Assert.False(result.Success);
            Assert.Equal(3u, result.Status);
            Assert.Equal("render warning", result.Message);
            Assert.Equal([1, 2, 3], Assert.Single(result.Outputs).Data);
            Assert.Equal("warn", Assert.Single(result.Diagnostics).Message);
        }
        finally
        {
            FreeAllocations(allocations);
        }
    }

    [Fact]
    public void ReadUtf8_HandlesNullPointerAndZeroLengthStrings()
    {
        List<IntPtr> allocations = [];
        try
        {
            IntPtr empty = AllocUtf8(string.Empty, allocations, out _);

            Assert.Null(InvokeCompilerPrivate("ReadUtf8", IntPtr.Zero, (UIntPtr)5));
            Assert.Equal(string.Empty, InvokeCompilerPrivate("ReadUtf8", empty, UIntPtr.Zero));
        }
        finally
        {
            FreeAllocations(allocations);
        }
    }

    [Fact]
    public void ToInt32_WhenNativeCountIsTooLarge_ThrowsBridgeException()
    {
        var exception = Assert.Throws<TargetInvocationException>(() => InvokeCompilerPrivate("ToInt32", (UIntPtr)((ulong)int.MaxValue + 1)));

        TypstBridgeException bridgeException = Assert.IsType<TypstBridgeException>(exception.InnerException);
        Assert.Contains("too large", bridgeException.Message);
    }

    [Fact]
    public void NativeLibraryResolver_RegisterIsIdempotentAndIgnoresOtherLibraries()
    {
        Type resolver = NativeType("NativeLibraryResolver");

        resolver.GetMethod("Register", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, null);
        resolver.GetMethod("Register", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, null);
        object? handle = resolver.GetMethod("Resolve", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null,
            ["not_typst_bridge", typeof(TypstBridgeCompiler).Assembly, null]);

        Assert.Equal(IntPtr.Zero, Assert.IsType<IntPtr>(handle));
    }

    [Fact]
    public void NativeLibraryResolver_CandidatePathsIncludeBaseAndRuntimeNativeDirectories()
    {
        Type resolver = NativeType("NativeLibraryResolver");
        string fileName = Assert.IsType<string>(resolver.GetMethod("GetNativeFileName", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, null));
        string rid = Assert.IsType<string>(resolver.GetMethod("GetRuntimeIdentifier", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, null));

        var candidates = Assert.IsAssignableFrom<IEnumerable<string>>(
            resolver.GetMethod("GetCandidatePaths", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [typeof(TypstBridgeCompiler).Assembly]))
            .ToArray();

        Assert.Equal(ExpectedNativeFileName(), fileName);
        Assert.Equal(ExpectedRuntimeIdentifier(), rid);
        Assert.Contains(Path.Combine(AppContext.BaseDirectory, fileName), candidates);
        Assert.Contains(Path.Combine(AppContext.BaseDirectory, "runtimes", rid, "native", fileName), candidates);
    }

    [Fact]
    public void ProbeAndVersionWorkWhenNativeLibraryIsAvailable()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        Assert.True(compiler.Probe());
        Assert.Equal(2u, TypstBridgeCompiler.SupportedAbiVersion);
        Assert.Equal(2u, compiler.AbiVersion);
        Assert.False(string.IsNullOrWhiteSpace(compiler.Version));
    }

    [Fact]
    public void MinimalPdfCompileReturnsSinglePdfOutput()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "Hello from TypstBridge",
            workingDirectory: Environment.CurrentDirectory,
            rootFileName: "managed-contract.typ");

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        TypstOutputFile output = Assert.Single(result.Outputs);
        Assert.Equal(0u, output.PageIndex);
        Assert.Equal("managed-contract.pdf", output.FileName);
        Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(output.Data, 0, 4));
    }

    [Fact]
    public void InvalidTypstSourceReturnsCompileStatusAndDiagnostic()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "#let =",
            workingDirectory: Environment.CurrentDirectory);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Compile, result.Status);
        Assert.False(result.Success);
        Assert.NotEmpty(result.Message);
        Assert.NotEmpty(result.Diagnostics);
        TypstDiagnostic diagnostic = result.Diagnostics[0];
        Assert.Equal(TypstDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.False(string.IsNullOrWhiteSpace(diagnostic.Message));
    }

    [Fact]
    public void SinglePageSvgCompileReturnsSingleSvgOutput()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "Hello from TypstBridge",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Svg);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        Assert.True(result.Success);
        TypstOutputFile output = Assert.Single(result.Outputs);
        AssertSvgOutput(output, 0u, "page-001.svg");
    }

    [Fact]
    public void MultiPageSvgCompileReturnsOutputsInPageOrder()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "First page\n#pagebreak()\nSecond page",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Svg);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        Assert.True(result.Success);
        Assert.Equal(2, result.Outputs.Count);

        for (int index = 0; index < result.Outputs.Count; index++)
        {
            uint pageIndex = (uint)index;
            AssertSvgOutput(result.Outputs[index], pageIndex, $"page-{index + 1:000}.svg");
        }
    }

    [Fact]
    public void SinglePagePngCompileReturnsSinglePngOutput()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "Hello from TypstBridge",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Png,
            ppi: 96.0);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        Assert.True(result.Success);
        TypstOutputFile output = Assert.Single(result.Outputs);
        AssertPngOutput(output, 0u, "page-001.png");
    }

    [Fact]
    public void MultiPagePngCompileReturnsOutputsInPageOrder()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "First page\n#pagebreak()\nSecond page",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Png,
            ppi: 96.0);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        Assert.True(result.Success);
        Assert.Equal(2, result.Outputs.Count);

        for (int index = 0; index < result.Outputs.Count; index++)
        {
            uint pageIndex = (uint)index;
            AssertPngOutput(result.Outputs[index], pageIndex, $"page-{index + 1:000}.png");
        }
    }

    [Fact]
    public void PngCompileUsesRequestedPpiForRasterDimensions()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        const string source = "#set page(width: 2in, height: 1in)\nHello from TypstBridge";
        TypstCompileRequest lowPpiRequest = new(
            source: source,
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Png,
            ppi: 96.0);
        TypstCompileRequest highPpiRequest = new(
            source: source,
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Png,
            ppi: 192.0);

        TypstCompileResult lowPpiResult = compiler.Compile(lowPpiRequest);
        TypstCompileResult highPpiResult = compiler.Compile(highPpiRequest);

        Assert.Equal(Ok, lowPpiResult.Status);
        Assert.True(lowPpiResult.Success);
        Assert.Equal(Ok, highPpiResult.Status);
        Assert.True(highPpiResult.Success);

        (uint lowWidth, uint lowHeight) = ReadPngDimensions(Assert.Single(lowPpiResult.Outputs).Data);
        (uint highWidth, uint highHeight) = ReadPngDimensions(Assert.Single(highPpiResult.Outputs).Data);
        Assert.True(highWidth > lowWidth, $"Expected higher PPI width ({highWidth}) to exceed lower PPI width ({lowWidth}).");
        Assert.True(highHeight > lowHeight, $"Expected higher PPI height ({highHeight}) to exceed lower PPI height ({lowHeight}).");
    }

    private static object InvokeCompilerPrivate(string methodName, params object?[] args)
    {
        MethodInfo method = typeof(TypstBridgeCompiler).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new MissingMethodException(typeof(TypstBridgeCompiler).FullName, methodName);

        return method.Invoke(null, args)!;
    }

    private static Type NativeType(string typeName)
    {
        return typeof(TypstBridgeCompiler).Assembly.GetType($"TypstBridge.Managed.Native.{typeName}")
            ?? throw new InvalidOperationException($"Native type '{typeName}' was not found.");
    }

    private static object CreateNativeStruct(string typeName)
    {
        return Activator.CreateInstance(NativeType(typeName))
            ?? throw new InvalidOperationException($"Could not create native struct '{typeName}'.");
    }

    private static T GetField<T>(object instance, string fieldName)
    {
        FieldInfo field = instance.GetType().GetField(fieldName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new MissingFieldException(instance.GetType().FullName, fieldName);

        return Assert.IsType<T>(field.GetValue(instance));
    }

    private static void SetField(object instance, string fieldName, object value)
    {
        FieldInfo field = instance.GetType().GetField(fieldName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new MissingFieldException(instance.GetType().FullName, fieldName);

        field.SetValue(instance, value);
    }

    private static IntPtr AllocateNativeArray(string itemTypeName, int count, List<IntPtr> allocations)
    {
        IntPtr array = Marshal.AllocHGlobal(Marshal.SizeOf(NativeType(itemTypeName)) * count);
        allocations.Add(array);
        return array;
    }

    private static IntPtr AllocUtf8(string value, List<IntPtr> allocations, out int byteLength)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        byteLength = bytes.Length;
        IntPtr buffer = Marshal.AllocHGlobal(bytes.Length == 0 ? 1 : bytes.Length);
        allocations.Add(buffer);
        if (bytes.Length > 0)
        {
            Marshal.Copy(bytes, 0, buffer, bytes.Length);
        }

        return buffer;
    }

    private static IntPtr AllocBytes(byte[] bytes, List<IntPtr> allocations)
    {
        IntPtr buffer = Marshal.AllocHGlobal(bytes.Length);
        allocations.Add(buffer);
        Marshal.Copy(bytes, 0, buffer, bytes.Length);
        return buffer;
    }

    private static void WriteNativeDiagnostic(
        IntPtr diagnosticsPtr,
        int index,
        TypstDiagnosticSeverity severity,
        string message,
        string? file,
        uint line,
        uint column,
        List<IntPtr> allocations)
    {
        IntPtr messagePtr = AllocUtf8(message, allocations, out int messageLength);
        IntPtr filePtr = file is null ? IntPtr.Zero : AllocUtf8(file, allocations, out _);
        int fileLength = file is null ? 0 : Encoding.UTF8.GetByteCount(file);
        object diagnostic = CreateNativeStruct("NativeDiagnostic");
        SetField(diagnostic, "Severity", (uint)severity);
        SetField(diagnostic, "MessageUtf8", messagePtr);
        SetField(diagnostic, "MessageLen", (UIntPtr)messageLength);
        SetField(diagnostic, "FileUtf8", filePtr);
        SetField(diagnostic, "FileLen", (UIntPtr)fileLength);
        SetField(diagnostic, "Line", line);
        SetField(diagnostic, "Column", column);

        int itemSize = Marshal.SizeOf(NativeType("NativeDiagnostic"));
        Marshal.StructureToPtr(diagnostic, IntPtr.Add(diagnosticsPtr, index * itemSize), false);
    }

    private static void FreeAllocations(IEnumerable<IntPtr> allocations)
    {
        foreach (IntPtr allocation in allocations)
        {
            Marshal.FreeHGlobal(allocation);
        }
    }

    private static void AssertSvgOutput(TypstOutputFile output, uint pageIndex, string fileName)
    {
        Assert.Equal(pageIndex, output.PageIndex);
        Assert.Equal(fileName, output.FileName);
        Assert.EndsWith(".svg", output.FileName, StringComparison.OrdinalIgnoreCase);

        string svg = System.Text.Encoding.UTF8.GetString(output.Data);
        Assert.True(
            svg.Contains("<svg", StringComparison.OrdinalIgnoreCase) ||
            svg.Contains("<?xml", StringComparison.OrdinalIgnoreCase),
            "SVG output should decode as UTF-8 and contain an SVG marker.");
    }

    private static void AssertPngOutput(TypstOutputFile output, uint pageIndex, string fileName)
    {
        Assert.Equal(pageIndex, output.PageIndex);
        Assert.Equal(fileName, output.FileName);
        Assert.EndsWith(".png", output.FileName, StringComparison.OrdinalIgnoreCase);
        Assert.True(output.Data.Length >= 24, "PNG output should include the signature and IHDR dimensions.");
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, output.Data[..8]);
        Assert.Equal("IHDR"u8.ToArray(), output.Data[12..16]);
    }

    private static (uint Width, uint Height) ReadPngDimensions(byte[] pngData)
    {
        AssertPngOutput(new TypstOutputFile(0u, "dimensions.png", pngData), 0u, "dimensions.png");

        uint width = BinaryPrimitives.ReadUInt32BigEndian(pngData.AsSpan(16, 4));
        uint height = BinaryPrimitives.ReadUInt32BigEndian(pngData.AsSpan(20, 4));
        return (width, height);
    }

    private static TypstBridgeCompiler CreateAvailableCompiler()
    {
        TypstBridgeCompiler compiler = new();
        if (!compiler.Probe())
        {
            throw new InvalidOperationException(
                "TypstBridge native library is not available. Build runtime assets first, e.g. TypstBridge/packaging/build-native.sh linux-x64.");
        }

        return compiler;
    }

    private static string ExpectedNativeFileName()
    {
        if (OperatingSystem.IsWindows())
        {
            return "typst_bridge.dll";
        }

        if (OperatingSystem.IsMacOS())
        {
            return "libtypst_bridge.dylib";
        }

        return "libtypst_bridge.so";
    }

    private static string ExpectedRuntimeIdentifier()
    {
        string os = OperatingSystem.IsWindows() ? "win" : OperatingSystem.IsMacOS() ? "osx" : "linux";
        string arch = RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.X64 => "x64",
            Architecture.Arm64 => "arm64",
            Architecture.X86 => "x86",
            Architecture.Arm => "arm",
            _ => RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant()
        };

        return $"{os}-{arch}";
    }
}
