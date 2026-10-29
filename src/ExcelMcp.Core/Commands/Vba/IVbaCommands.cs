using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA module and procedure operations for macro-enabled workbooks (.xlsm).
///
/// PREREQUISITES:
/// - Workbook must be macro-enabled (.xlsm)
/// - VBA trust must be enabled manually in Excel for project inspection and editing
///
/// SCOPE:
/// - List and view existing VBA components and their procedures
/// - Import creates new standard modules from inline code or a file
/// - Update/delete works on existing VBA components by name
/// - Run executes a procedure by name
///
/// RUN: procedureName format is 'Module.Procedure' (e.g., 'Module1.MySub').
/// Running an existing macro does not require VBA project access.
/// ExcelMcp does not configure VBA trust settings for you.
/// </summary>
[ServiceCategory("Vba")]
[McpTool("vba", Title = "VBA Operations", Destructive = true, Category = "automation",
    Description = "Inspect VBA project status and library references, search source with limited results, list procedures and source ranges, read a bounded part of a module, replace one procedure only when its source fingerprint still matches, replace a full module when its source fingerprint matches, import or update modules, and invoke procedures with typed arguments and function results in .xlsm workbooks. Invoke results allow at most 10,000 array elements and a conservative 4 MiB transport budget; rejected return conversion occurs after macro execution and does not roll back its effects. Project inspection requires Trust Center access; status reports blocked access without enabling it. Replacing source preserves surrounding comments but does not prove it compiles or runs correctly.")]
[McpReadOnlyActions("list", "view", "read", "search", "references", "status")]
public interface IVbaCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    [ServiceAction("list")]
    VbaListResult List(IExcelBatch batch);

    /// <summary>
    /// Reports actual VBA project access, password protection, and execution mode without changing settings.
    /// </summary>
    [ServiceAction("status")]
    VbaProjectStatusResult Status(IExcelBatch batch);

    /// <summary>
    /// Lists VBA library references and flags missing libraries without repairing or changing references.
    /// </summary>
    [ServiceAction("references")]
    VbaReferencesResult References(IExcelBatch batch);

    /// <summary>
    /// Searches VBA source for literal text, returning limited matches with line numbers and excerpts.
    /// </summary>
    /// <param name="searchText">Nonempty, single-line literal text to find; wildcard patterns are not supported</param>
    /// <param name="moduleName">Optional module to search; omitted searches all modules in this workbook</param>
    /// <param name="wholeWord">Match whole words only; default false</param>
    /// <param name="matchCase">Match letter case; default false</param>
    /// <param name="maxMatches">Maximum returned matches, from 1 through 100; default 50. HasMore indicates omitted matches. Excerpts are limited to 200 characters.</param>
    [ServiceAction("search")]
    VbaSearchResult Search(
        IExcelBatch batch,
        [RequiredParameter] string searchText,
        string? moduleName = null,
        bool wholeWord = false,
        bool matchCase = false,
        int maxMatches = 50);

    /// <summary>
    /// Views VBA module code without exporting to file
    /// </summary>
    /// <param name="moduleName">Name of the VBA module</param>
    [ServiceAction("view")]
    VbaViewResult View(IExcelBatch batch, [RequiredParameter] string moduleName);

    /// <summary>
    /// Reads one VBA procedure or a bounded range of module lines.
    /// </summary>
    /// <param name="moduleName">Name of the VBA module</param>
    /// <param name="procedureName">Procedure to read; select this or startLine and lineCount</param>
    /// <param name="procedureKind">Optional kind to distinguish property accessors or same-name procedures</param>
    /// <param name="startLine">First module line to read; select this with lineCount instead of procedureName</param>
    /// <param name="lineCount">Number of module lines to read, from 1 through 500</param>
    [ServiceAction("read")]
    VbaReadResult Read(
        IExcelBatch batch,
        [RequiredParameter] string moduleName,
        string? procedureName,
        string? procedureKind,
        int? startLine,
        int? lineCount);

    /// <summary>
    /// Imports VBA code to create a new standard module
    /// </summary>
    /// <param name="moduleName">Name for the new module</param>
    /// <param name="vbaCode">VBA code. Public callers must supply either inline vbaCode or a readable vbaCodeFile, not both.</param>
    [ServiceAction("import")]
    OperationResult Import(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    /// <param name="moduleName">Name of the module to update</param>
    /// <param name="vbaCode">New VBA code. Public callers must supply either inline vbaCode or a readable vbaCodeFile, not both.</param>
    [ServiceAction("update")]
    OperationResult Update(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>Replaces whole-module source only when its fingerprint still matches; verifies source read back without compiling or rolling back.</summary>
    /// <param name="moduleName">Name of the component whose code is replaced, preserving the component itself</param>
    /// <param name="expectedSourceHash">Expected SourceHash: use vba.view for replace-module, or vba.read for replace-procedure</param>
    /// <param name="vbaCode">Replacement code; inline source or vbaCodeFile, not both. Leading export Attribute VB_* metadata is removed after newline normalization.</param>
    [ServiceAction("replace-module")]
    VbaModuleEditResult ReplaceModule(IExcelBatch batch, [RequiredParameter] string moduleName,
        [RequiredParameter] string expectedSourceHash, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Replaces one VBA procedure only if its source has not changed since it was read, preserving surrounding comments and blank lines.
    /// </summary>
    /// <param name="moduleName">Name of the VBA module</param>
    /// <param name="procedureName">Name of the procedure to replace</param>
    /// <param name="procedureKind">Kind of procedure: Sub, Function, Property Get, Property Let, or Property Set</param>
    /// <param name="expectedSourceHash">Expected SourceHash: use vba.view for replace-module, or vba.read for replace-procedure</param>
    /// <param name="vbaCode">Source for exactly one replacement procedure; saving it does not prove it compiles or runs</param>
    [ServiceAction("replace-procedure")]
    VbaProcedureEditResult ReplaceProcedure(
        IExcelBatch batch,
        [RequiredParameter] string moduleName,
        [RequiredParameter] string procedureName,
        [RequiredParameter] string procedureKind,
        [RequiredParameter] string expectedSourceHash,
        [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Runs a VBA procedure with optional parameters. Normal file opens disable VBA macros.
    /// Running does not enable them or reopen the file: obtain separate approval for
    /// file enable-macros with approved:true and allow_reopen:true
    /// (CLI: session enable-macros --approved --allow-reopen) on a saved owned workbook first.
    /// </summary>
    /// <param name="procedureName">Name of the procedure to run (for example "Module1.MySub")</param>
    /// <param name="timeout">Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch</param>
    /// <param name="parameters">Up to 30 optional parameters to pass to the procedure</param>
    [ServiceAction("run")]
    OperationResult Run(IExcelBatch batch, [RequiredParameter] string procedureName, TimeSpan? timeout, params string[] parameters);

    /// <summary>Runs a procedure with explicit scalar argument types and returns a scalar/null or two-dimensional array of at most 10,000 elements within a conservative 4 MiB transport budget. Oversized results are rejected after execution without cropping or rolling back macro effects. Uses the same macro approval boundary as run; cancellation is not rollback.</summary>
    /// <param name="procedureName">Unqualified Module.Procedure name; the selected session determines the workbook</param>
    /// <param name="timeout">Optional whole-second timeout, matching run; expiry does not undo macro effects</param>
    /// <param name="arguments">Up to 30 objects with type string, number, boolean, null, or missing and a matching scalar value. Null and missing omit value or use null.</param>
    [ServiceAction("invoke")]
    VbaInvokeResult Invoke(IExcelBatch batch, [RequiredParameter] string procedureName, TimeSpan? timeout, List<VbaArgument>? arguments = null);

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    /// <param name="moduleName">Name of the module to delete</param>
    [ServiceAction("delete")]
    OperationResult Delete(IExcelBatch batch, [RequiredParameter] string moduleName);
}
