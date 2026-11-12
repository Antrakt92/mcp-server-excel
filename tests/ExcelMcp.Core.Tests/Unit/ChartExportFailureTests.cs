using System.Reflection;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "false")]
public sealed class ChartExportFailureTests
{
    // This tests our file/success validation, not Excel's image encoder.
    [Theory]
    [InlineData(false, -1)]
    [InlineData(false, 0)]
    [InlineData(false, 3)]
    [InlineData(true, -1)]
    [InlineData(true, 0)]
    [InlineData(true, 3)]
    public void ImageOutput_RequiresNativeSuccessAndNonemptyFile(bool exported, int length)
    {
        var output = Path.Combine(Path.GetTempPath(), $"excel-export-validation-{Guid.NewGuid():N}.png");
        try
        {
            if (length >= 0)
                File.WriteAllBytes(output, new byte[length]);
            var validate = typeof(ChartCommands).GetMethod("ValidateImageOutput", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(validate);
            if (exported && length > 0)
                validate.Invoke(null, [exported, output, ChartImageFormat.Png]);
            else
            {
                var error = Assert.Throws<TargetInvocationException>(() =>
                    validate.Invoke(null, [exported, output, ChartImageFormat.Png]));
                var cause = Assert.IsType<IOException>(error.InnerException);
                Assert.Contains("nonempty Png", cause.Message, StringComparison.Ordinal);
            }
        }
        finally
        {
            File.Delete(output);
        }
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 3)]
    [InlineData(true, 0)]
    [InlineData(true, 3)]
    public void ExportImage_BatchFailurePreservesFileCreatedBeforeTheCallback(bool overwrite, int length)
    {
        var output = Path.Combine(Path.GetTempPath(), $"excel-export-failure-{Guid.NewGuid():N}.png");
        using var batch = DispatchProxy.Create<IExcelBatch, ExportFailureBatch>();
        var failureBatch = Assert.IsAssignableFrom<ExportFailureBatch>(batch);
        failureBatch.OutputPath = output;
        failureBatch.Length = length;
        try
        {
            var error = Assert.Throws<IOException>(() =>
                new ChartCommands().ExportImage(batch, "Chart", output, overwrite: overwrite));
            Assert.Same(failureBatch.Failure, error);
            Assert.Equal(length, new FileInfo(output).Length);
        }
        finally
        {
            File.Delete(output);
        }
    }

    // Simulate another writer creating the destination while the callback never runs.
    public class ExportFailureBatch : DispatchProxy
    {
        public string OutputPath { get; set; } = string.Empty;
        public int Length { get; set; }
        public IOException Failure { get; } = new("Injected chart export failure.");

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod?.Name == nameof(IDisposable.Dispose))
                return null;
            if (targetMethod?.Name == nameof(IExcelBatch.Execute))
            {
                File.WriteAllBytes(OutputPath, new byte[Length]);
                throw Failure;
            }
            throw new NotSupportedException($"Unexpected batch member: {targetMethod?.Name}");
        }
    }
}
