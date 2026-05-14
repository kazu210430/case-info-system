# Office / VSTO unload boundary diagnostics - 2026-05-14

## Purpose

This note records the diagnostic context for the Excel crash observed after closing `会計書類セット.xlsx`.
It is an investigation memo, not a root-cause conclusion.

## Observed crash signature

- Flow: accounting workbook close as normal workbook close.
- Close facts: `isManagedClose=False`, `beforeCloseAction=Ignore`.
- AccountingClose marker / post-close follow-up / `Application.Quit` were not involved in the target accounting close.
- The target crashes reached `shutdown-handler-complete`.
- WER event: `OFFICE_MODULE_VERSION_MISMATCH`.
- Fault module: `KERNELBASE.dll`.
- Exception code: `0xe0434352`.
- WER bucket: `2120126298959465185`.
- WER hash: `e2ffb8041e1bafc6ad6c33a2d0714ae1`.
- New Excel PID startup facts after crash: `activeWorkbookPresent=False`, `workbooksCount=0`, `applicationVisible=True`.
- New Excel PID command line showed only `EXCEL.EXE`; no `/restore`, `/automation`, or DDE-like signal was confirmed.

## Version skew facts

- Office / Excel main binaries sampled: `16.0.19929.20136`.
- `AppVIsvSubsystems64.dll`: `16.0.19929.20032`.
- VSTO runtime: `10.0.60910.0`.
- Add-in bundled Office interop assemblies: `15.0.4420.1017`.
- Current deployed Add-in DLL, `bin\Debug`, DebugPackage, and current `dl3` shadow DLL had matching SHA at the time of investigation.
- Old and current `dl3` shadow DLLs both had crash examples, so shadow copy mixing alone is not established as the cause.

These facts are a version-skew risk signal only. They do not prove that Office/VSTO/Interop version mismatch is the root cause.

## Safe mode comparison results

Comparison completed on 2026-05-14:

| Condition | Crash | New empty Excel | WER | CaseInfoSystem trace |
| --- | --- | --- | --- | --- |
| Normal Excel + accounting workbook as last workbook | Yes | Yes | `OFFICE_MODULE_VERSION_MISMATCH` | Updated through `generated-onshutdown-exit` |
| `EXCEL.EXE /safe` + accounting workbook as last workbook | No | No | None observed | No update |
| `EXCEL.EXE /safe` + empty workbook | No | No | None observed | No update |
| Normal Excel + empty state | No | No | None observed | Updated and unloaded cleanly |

The current comparison suggests that neither Excel core alone nor the accounting workbook alone is sufficient to reproduce the target crash.
The crash is currently correlated with normal-mode Office/Add-in extension loading plus closing the accounting workbook as the last workbook.
However, normal Excel startup without the accounting workbook did not crash even with CaseInfoSystem / VSTO / PDFMaker / VSTOExcelAdaptor loaded, so Add-in loading alone is not sufficient either.

Do not treat this as a root-cause conclusion.

## Individual Add-in comparison preparation

The next comparison phase should be reversible and should not permanently disable any Add-in.
Before any temporary Add-in state change, capture the current registry values and loaded-module evidence so the exact pre-test state can be restored.

Safety rules:

- Do not delete Office Resiliency or DocumentRecovery registry entries.
- Do not clear the VSTO cache.
- Do not repair or update Office as part of this comparison.
- Do not change CaseInfoSystem managed close, AccountingClose marker, post-close follow-up, `Application.Quit`, startup guard, or `Application.Visible`.
- Do not keep an Add-in disabled after a test condition finishes.
- Restore the captured value immediately after each temporary-disable condition.
- Record whether Excel was fully closed before changing the next condition.

Recommended read-only inventory before changing anything:

- Excel Add-in registry entries under current user and machine Office Addins keys.
- `LoadBehavior`, `FriendlyName`, `Description`, and `Manifest` for VSTO Add-ins.
- Office COM Add-in candidates observed in normal mode, especially CaseInfoSystem, Adobe PDFMaker, VSTOExcelAdaptor, electronic-signature related Add-ins, and OneDrive/FileCoAuth related modules.
- Runtime loaded modules for each Excel PID.
- CaseInfoSystem loaded assembly path and SHA when the CaseInfoSystem Add-in is active.

Per-condition evidence to capture:

- Test start time and Excel PID.
- Excel command line for the original PID and any new PID.
- Loaded modules matching CaseInfoSystem, VSTO, VSTOExcelAdaptor, PDFMaker, VBE, OneDrive/FileCoAuth, and electronic-signature components.
- CaseInfoSystem trace last write time before and after the test.
- Whether `generated-onshutdown-exit` appears when CaseInfoSystem is loaded.
- Event Viewer `Application Error`, `.NET Runtime`, `Windows Error Reporting`, and `Microsoft Office Alerts` entries since the test start time.
- WER archive/report path, event name, bucket, hash, fault module, and exception code.
- Whether any `EXCEL.EXE` process remains after close.

Suggested comparison order after explicit approval:

1. Reconfirm one normal-mode accounting workbook last-workbook close on the current diagnostic build only if a fresh baseline is needed.
2. Temporarily disable only CaseInfoSystem Excel Add-in, run the accounting last-workbook close, then restore CaseInfoSystem immediately.
3. Temporarily disable only Adobe PDFMaker, run the same condition, then restore it immediately.
4. Temporarily disable only VSTOExcelAdaptor or electronic-signature related Add-ins if present, one at a time, restoring after each condition.
5. Compare a local non-OneDrive copy of the same accounting workbook, because safe mode still loaded FileCoAuth-related modules during the successful no-crash comparison.

Any temporary disable/restore method must be documented with the exact original value and restored value.
If a condition requires a setting that cannot be restored confidently, stop and do not run that condition.

## Next manual comparison checklist

- Close accounting workbook as the last workbook and record whether `generated-onshutdown-boundary phase=base-onshutdown-after-finally` appears.
- Close accounting workbook while another workbook remains open and record whether Excel avoids immediate shutdown/crash.
- Compare form visible vs. no form.
- Compare VBE open vs. closed.
- Compare normal Excel vs. `EXCEL.EXE /safe "<accounting workbook path>"`.
- If explicitly approved later, compare with CaseInfoSystem Add-in temporarily disabled and then restore the setting.
- Do not delete Office Resiliency / DocumentRecovery registry entries as part of this checklist.
- Do not clear VSTO cache or repair/update Office as part of this checklist without a separate decision.

## Added diagnostic boundaries

The runtime diagnostic change adds shutdown/unload boundary logs only.
It does not reintroduce AccountingClose managed close, marker creation, post-close follow-up, startup guard widening, `Application.Visible=False`, or crash avoidance behavior.
