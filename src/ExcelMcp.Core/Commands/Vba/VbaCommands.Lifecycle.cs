using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script lifecycle operations (List, View, Export, Import, Update, Delete)
/// </summary>
public partial class VbaCommands
{
    /// <inheritdoc />
    public VbaListResult List(IExcelBatch batch)
    {
        var result = new VbaListResult { FilePath = batch.WorkbookPath };

        EnsureVbaFile(batch.WorkbookPath);

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
        {
            Vbe.VBProject? vbaProject = null;
            Vbe.VBComponents? vbComponents = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
                EnsureProjectUnlocked(vbaProject);
                vbComponents = vbaProject.VBComponents;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    Vbe.VBComponent? component = null;
                    Vbe.CodeModule? codeModule = null;
                    try
                    {
                        component = vbComponents.Item(i);
                        string name = component.Name;
                        int type = (int)component.Type;

                        string typeStr = GetVbaModuleTypeName(type);

                        codeModule = component.CodeModule;
                        int moduleLineCount = codeModule.CountOfLines;
                        List<VbaProcedureInfo> procedureDetails = GetProcedures(codeModule, ct);

                        result.Scripts.Add(new ScriptInfo
                        {
                            Name = name,
                            Type = typeStr,
                            LineCount = moduleLineCount,
                            Procedures = procedureDetails.Select(procedure => procedure.Name).ToList(),
                            ProcedureDetails = procedureDetails
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref codeModule);
                        ComUtilities.Release(ref component);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (IsVbaTrustError(comEx))
            {
                // Trust was disabled during operation
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == GenericOfficeAutomationError)
            {
                // Trust passed IsVbaTrustError, so this generic 0x800A03EC is a real,
                // non-trust failure. Surface it with COM environment diagnostics (issue #671).
                throw new InvalidOperationException(BuildGenericComErrorMessage(comEx), comEx);
            }
            finally
            {
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public VbaViewResult View(IExcelBatch batch, string moduleName)
    {
        var result = new VbaViewResult { FilePath = batch.WorkbookPath, ModuleName = moduleName };

        EnsureVbaFile(batch.WorkbookPath);

        if (string.IsNullOrWhiteSpace(moduleName))
        {
            throw new ArgumentException("Module name cannot be empty", nameof(moduleName));
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
        {
            Vbe.VBProject? vbaProject = null;
            Vbe.VBComponents? vbComponents = null;
            Vbe.VBComponent? component = null;
            Vbe.CodeModule? codeModule = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
                EnsureProjectUnlocked(vbaProject);
                vbComponents = vbaProject.VBComponents;

                // Find the specified module
                bool found = false;
                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    component = vbComponents.Item(i);
                    string name = component.Name;

                    if (name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        int type = (int)component.Type;
                        result.ModuleType = GetVbaModuleTypeName(type);

                        codeModule = component.CodeModule;
                        result.LineCount = codeModule.CountOfLines;

                        // Read all code lines
                        if (result.LineCount > 0)
                        {
                            result.Code = codeModule.Lines[1, result.LineCount];
                        }

                        result.SourceHash = VbaModuleSource.ComputeHash(result.Code);
                        result.Procedures = GetProcedures(codeModule, ct)
                            .Select(procedure => procedure.Name)
                            .ToList();

                        break;
                    }

                    ComUtilities.Release(ref component);
                    component = null;
                }

                if (!found)
                {
                    throw new OperationFailureException(OperationFailureCategory.NotFound, $"Module '{moduleName}' not found in workbook");
                }

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (IsVbaTrustError(comEx))
            {
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == GenericOfficeAutomationError)
            {
                // Trust passed IsVbaTrustError, so this generic 0x800A03EC is a real,
                // non-trust failure. Surface it with COM environment diagnostics (issue #671).
                throw new InvalidOperationException(BuildGenericComErrorMessage(comEx), comEx);
            }
            finally
            {
                ComUtilities.Release(ref codeModule);
                ComUtilities.Release(ref component);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Import(IExcelBatch batch, string moduleName, string vbaCode)
    {
        ValidateModuleName(moduleName);
        vbaCode = VbaModuleSource.Normalize(vbaCode);
        EnsureVbaFile(batch.WorkbookPath);

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
        {
            Vbe.VBProject? vbaProject = null;
            Vbe.VBComponents? vbComponents = null;
            Vbe.VBComponent? newModule = null;
            Vbe.CodeModule? codeModule = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
                EnsureProjectUnlocked(vbaProject);
                vbComponents = vbaProject.VBComponents;

                var existingNames = new List<string>();
                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    Vbe.VBComponent? component = null;
                    try
                    {
                        component = vbComponents.Item(i);
                        existingNames.Add(component.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref component);
                    }
                }

                VbaModuleSource.Import(moduleName, vbaCode, existingNames,
                    () => newModule = vbComponents.Add(Vbe.vbext_ComponentType.vbext_ct_StdModule),
                    name => newModule!.Name = name,
                    source =>
                    {
                        codeModule = newModule!.CodeModule;
                        codeModule.InsertLines(1, source);
                    },
                    () => newModule?.Name ?? throw new InvalidOperationException("The created component could not be identified."),
                    () =>
                    {
                        codeModule ??= newModule?.CodeModule
                            ?? throw new InvalidOperationException("The created component source could not be inspected.");
                        return ReadLines(codeModule, 1, codeModule.CountOfLines);
                    });

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (COMException comEx) when (IsVbaTrustError(comEx))
            {
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == GenericOfficeAutomationError)
            {
                // Trust passed IsVbaTrustError, so this generic 0x800A03EC is a real,
                // non-trust failure. Surface it with COM environment diagnostics (issue #671).
                throw new InvalidOperationException(BuildGenericComErrorMessage(comEx), comEx);
            }
            finally
            {
                ComUtilities.Release(ref codeModule);
                ComUtilities.Release(ref newModule);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Update(IExcelBatch batch, string moduleName, string vbaCode) =>
        ReplaceModuleCore(batch, moduleName, null, vbaCode);

}
