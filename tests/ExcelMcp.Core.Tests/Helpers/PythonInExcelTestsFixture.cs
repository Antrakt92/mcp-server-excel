// <copyright file="PythonInExcelTestsFixture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE Python in Excel test file per test CLASS.
/// Each test uses a unique target cell for isolation.
/// - Created ONCE before any tests run
/// - Shared by all tests in the class
/// - Each test writes its own source data range and unique PY() target cell
/// </summary>
public class PythonInExcelTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    private int _rowCounter;

    /// <summary>
    /// Path to the shared test file.
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Initializes a new instance of the fixture.
    /// </summary>
    public PythonInExcelTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"PythonInExcelTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Gets a unique row block (10 rows apart) so each test can lay out its own
    /// source data + PY() target cells without colliding with other tests.
    /// </summary>
    public int GetUniqueRowBlockStart()
    {
        // 10 rows per test block: enough room for source data + a couple of PY() targets
        var counter = Interlocked.Increment(ref _rowCounter);
        return 1 + ((counter - 1) * 10);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// Creates a workbook that tests will write source data and PY() formulas to.
    /// </summary>
    public Task InitializeAsync()
    {
        TestFilePath = Path.Join(_tempDir, "PythonInExcelTest.xlsx");
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(TestFilePath, show: false);
        manager.CloseSession(sessionId, save: true);
        return Task.CompletedTask;
    }

    /// <summary>
    /// Called ONCE after all tests in the class complete.
    /// </summary>
    public Task DisposeAsync()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup is best-effort
        }

        return Task.CompletedTask;
    }
}
