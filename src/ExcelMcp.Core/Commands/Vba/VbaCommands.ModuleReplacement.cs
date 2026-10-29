using System.Globalization;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class VbaCommands
{
    /// <inheritdoc />
    public VbaModuleEditResult ReplaceModule(IExcelBatch batch, string moduleName, string expectedSourceHash, string vbaCode)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(expectedSourceHash);
        return ReplaceModuleCore(batch, moduleName, expectedSourceHash, vbaCode);
    }

    private static VbaModuleEditResult ReplaceModuleCore(IExcelBatch batch, string moduleName, string? expectedSourceHash, string vbaCode)
    {
        ValidateModuleName(moduleName);
        VbaModuleSource.ValidateExpectedHash(expectedSourceHash);
        string normalized = VbaModuleSource.Normalize(vbaCode);
        EnsureVbaFile(batch.WorkbookPath);
        if (!IsVbaTrustEnabled())
            throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage);
        return batch.Execute((ctx, ct) =>
        {
            Vbe.VBProject? project = null;
            Vbe.VBComponents? components = null;
            Vbe.VBComponent? component = null;
            Vbe.CodeModule? module = null;
            try
            {
                project = ctx.Book.VBProject;
                EnsureProjectUnlocked(project);
                if (GetProjectMode(Convert.ToInt32(project.Mode, CultureInfo.InvariantCulture)) != "Design")
                    throw new OperationFailureException(OperationFailureCategory.Conflict, "The VBA project must be in Design mode before replacing source.");
                components = project.VBComponents;
                component = FindComponent(components, moduleName);
                module = component.CodeModule;
                ct.ThrowIfCancellationRequested();
                string stored = VbaModuleSource.Replace(moduleName, expectedSourceHash, normalized,
                    () => ReadLines(module, 1, module.CountOfLines),
                    () => { int count = module.CountOfLines; if (count != 0) module.DeleteLines(1, count); },
                    source => module.InsertLines(1, source));
                return new VbaModuleEditResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    ModuleName = moduleName,
                    Action = expectedSourceHash is null ? "update" : "replace-module",
                    SourceHash = VbaModuleSource.ComputeHash(stored),
                    LineCount = module.CountOfLines,
                    Message = "Module source was replaced and read back. Compilation and runtime behavior were not verified."
                };
            }
            finally
            {
                ComUtilities.Release(ref module);
                ComUtilities.Release(ref component);
                ComUtilities.Release(ref components);
                ComUtilities.Release(ref project);
            }
        });
    }

    private static void ValidateModuleName(string moduleName)
    {
        if (string.IsNullOrWhiteSpace(moduleName) || moduleName.Length > 31 ||
            !Regex.IsMatch(moduleName, @"^[\p{L}][\p{L}\p{Nd}_]*$", RegexOptions.CultureInvariant))
            throw new ArgumentException("moduleName must be a VBA component identifier of at most 31 characters.", nameof(moduleName));
    }
}
