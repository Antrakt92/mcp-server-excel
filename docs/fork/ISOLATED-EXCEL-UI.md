# Test Excel UI on a private Windows desktop

Use the test-only [isolated Excel UI runner](../../scripts/tests/native/isolated_excel_ui.py)
for disposable synthetic workbooks. It creates a new Windows desktop, launches
Excel there with `/x` and binds to the workbook through its native document window.
It checks the process creation time and desktop before using a window. It never
switches the input desktop or sends global mouse or keyboard input.

On Windows, use the local Excel testing Python environment with pywin32 available
and provide the absolute path to the installed `EXCEL.EXE`. Supply a new path
under your private test output directory. The runner creates its own fixture and
does not accept a workbook path. Run these scenarios sequentially.

```powershell
python scripts/tests/native/isolated_excel_ui.py --excel $excelExecutable --directory $newOutputDirectory --scenario isolation
```

| Scenario | Evidence |
| --- | --- |
| `isolation` | A new Excel process, its native workbook window and the synthetic startup sheet belong to the private desktop. |
| `sample` | The requested client-area pixels are readable from that window's device context over a bounded interval. |
| `navigation` | The named worksheet shape opens a modeless VBA menu. The addressed menu click moves the bound workbook to the expected sheet on the first attempt after activating each of two windows. Escape stays on the starting sheet. Save As invalidates the old binding and the route works after a new native window bind. |
| `compile` | VBE's compile command completes for the selected synthetic project and remains disabled on a second observation of that same project and code pane. |
| `compile-error` | VBE reports the expected undefined procedure in the selected synthetic project. The runner confirms that the matched dialog closes. |
| `form-packages` | Three saved copies of a synthetic form separate a VBA procedure change from a designer caption change. The runner reopens each read-only with macros disabled and verifies its source and caption independently. |

Navigation, compilation and form package checks require the existing VBA project
object-model access setting. They leave that setting alone. The compile-error
scenario addresses the enabled OK button only after confirming the new dialog's
exact English title, diagnostic and VBE owner on the private desktop. Other
dialog signatures remain untouched and the outcome stays unconfirmed.

Use `form-baseline.xlsm`, `form-code.xlsm` and `form-designer.xlsm` for the
two offline package comparisons. The runner writes the copies only under the
new test directory and records their SHA-256 values. Retire them after checking
the comparison results.

The navigation run also samples the menu's client area during addressed hover.
The receipt counts complete samples in each known color state, other color
states, and invalid device-context reads. It preserves the first and last point
colors and a separate, hashed control frame in `hover-control.bmp`. The
synthetic form briefly paints a known third color to check whether the sampler
can catch an excursion between two expected states. The receipt records the
number of marker samples and the interval between the two addressed messages.
Sampling can miss transients between observations. The check exercises Excel's response
to addressed messages and the pixels exposed by the window's device context.

Read `result.json` after the controller exits. A passing result requires the
worker's assertion, empty cleanup failures, a normal Excel exit and a stable input
desktop. A timeout keeps the partial diagnostic and exact process identity.
Inspect those before starting another run. The runner leaves process
termination to the operator and handles only the exact synthetic compile
dialog described above. Keep the test directory private and retire its synthetic
fixtures after verifying the outcome.

For sessions opened through the CLI, use the recovery guide's example to
[keep operation and cleanup failures](RECOVERY.md#keep-operation-and-cleanup-failures).

The Excel-free checks cover the synthetic package, admission, coordinate bounds,
color classification and preservation of both failures:

```powershell
python -m unittest discover -s scripts/tests/native -p test_isolated_excel_ui.py
```
