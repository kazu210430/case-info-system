# Accounting Close Lifecycle Current State

## 位置づけ

この文書は、会計書類セットの close / Excelを閉じる / import prompt / accounting form 周辺について、現行コードと実機安定化済みの順序を固定する current-state です。

この領域は今後の helper 化、共通化、close 経路再整理の対象にしません。直 `workbook.Close()` が存在することだけを理由に、未整理、残課題、helper 化予定、将来改善候補へ分類しません。

## 対象範囲

- Excel の × 印からの会計 workbook close
- 会計フォーム上の「Excelを閉じる」ボタン
- 逆算ツールフォーム
- お支払い履歴を取り込む import prompt
- 分割払い予定表入力フォーム
- お支払い履歴入力フォーム
- 会計依頼書の取込対象セル表示
- `workbook.Close()`
- `Application.Quit()`
- `FormClosing` / `FormClosed`
- 黄色セル解除

## 現コードの owner

- `WorkbookLifecycleCoordinator`
  - `WorkbookBeforeClose` で CASE / Kernel の close 処理後に、会計 active form guard を呼びます。
- `AccountingWorkbookLifecycleService`
  - 会計 workbook 判定、active form / prompt close guard、通常 close 時の会計フォーム片付けを調停します。
- `AccountingFormHelperService`
  - 逆算ツール、分割払い予定表入力フォーム、お支払い履歴入力フォームの表示、Excel close button、フォーム close / dispose、直 `workbook.Close()`、必要時の `Application.Quit()` を持ちます。
- `AccountingPaymentHistoryImportService`
  - お支払い履歴 import prompt の表示、Excel close button、黄色セル cleanup、直 `workbook.Close()`、必要時の `Application.Quit()` を持ちます。
- `AccountingWorkbookService`
  - 会計依頼書 `F15:F20` と逆算対象 `F17:F20` の黄色セル表示 / 解除を持ちます。
- 各 Form class
  - UI と `ExcelCloseRequested` などのイベント発火を持ちます。workbook lifecycle の owner ではありません。

## Excel の × 印 close

Excel の × 印による close と、フォーム上の「Excelを閉じる」ボタンによる close は同一視しません。

`WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` は、会計 workbook に対して `AccountingWorkbookLifecycleService.TryCancelWorkbookBeforeCloseForActiveAccountingForm(...)` を呼びます。

- active な会計フォームまたは import prompt が対象 workbook に紐づいている場合、フォームボタン close の allow flag が無ければ close を cancel します。
- cancel 時は「フォームの「Excelを閉じる」ボタンから閉じてください。」を表示します。
- allow flag が一致するフォームボタン経路だけ、会計 close guard を通過します。
- guard を通過した通常 close では、`AccountingPaymentHistoryImportService.HandleWorkbookBeforeClose(...)` と `AccountingFormHelperService.HandleWorkbookBeforeClose(...)` が active prompt / form の片付けを行います。

このため、Excel の × 印 close をフォームボタン close の代替として扱いません。

## フォームボタン経由 close

会計フォーム上の「Excelを閉じる」ボタン経路では、フォーム lifecycle を先に閉じてから workbook を閉じる順序を維持します。

`AccountingFormHelperService.RequestWorkbookCloseFromAccountingForm(...)` の現行順序:

1. 対象 workbook key と form kind を allow flag として記録します。
2. form kind に応じて `CloseActiveInstallmentScheduleInput()`、`CloseActivePaymentHistoryInput()`、`CloseActiveReverseTool()` のいずれかを実行します。
3. active form を `Close()` し、未 dispose なら `Dispose()` します。
4. form references と owner を clear します。
5. その後に直 `workbook.Close()` を呼びます。
6. close 後、同じ `Application` の `Workbooks.Count` を読み、0 の場合だけ `application.Quit()` を呼びます。
7. finally で allow flag を clear します。

この直 `workbook.Close()` は意図的な安定化契約です。`WorkbookCloseInteropHelper` に寄せる対象ではありません。

## import prompt 経由 close

`AccountingPaymentHistoryImportService.RequestWorkbookCloseFromImportPrompt(...)` の現行順序:

1. 対象 workbook key と import prompt form kind を allow flag として記録します。
2. `CloseActivePrompt(clearHighlight: true)` を実行します。
3. prompt の黄色セル cleanup と `Close()` / `Dispose()` を先に完了します。
4. prompt handlers と active references を clear します。
5. その後に直 `workbook.Close()` を呼びます。
6. close 後、同じ `Application` の `Workbooks.Count` を読み、0 の場合だけ `application.Quit()` を呼びます。
7. finally で allow flag を clear します。

この経路も `WorkbookCloseInteropHelper` へ寄せません。import prompt の close 順序は会計依頼書の黄色セル解除と一体の安定化済み契約です。

## FormClosing / FormClosed

Form class 側へ workbook close 判断を持たせません。cleanup は service 側の event handler が持ちます。

- `AccountingPaymentHistoryImportService`
  - `AccountingImportRangePromptForm` に `FormClosing` / `FormClosed` を attach します。
  - `FormClosing` と `FormClosed` の両方で `CleanupActivePromptHighlightOnce(...)` を通し、黄色セル解除を一度だけ行います。
  - `FormClosed` で handlers detach と active references clear を行います。
- `AccountingFormHelperService`
  - `AccountingReverseGoalSeekForm` に `FormClosing` / `FormClosed` を attach します。
  - `FormClosing` と `FormClosed` の両方で `CleanupActiveReverseToolHighlightOnce(...)` を通し、逆算対象黄色セル解除を一度だけ行います。
  - `FormClosed` で handlers detach と active references clear を行います。
  - 分割払い予定表入力フォーム / お支払い履歴入力フォームは、利用者操作で閉じられた場合は `FormClosed` handler が handlers detach と active references clear を行います。
  - service 側から閉じる場合は、`CloseActiveInstallmentScheduleInput()` / `CloseActivePaymentHistoryInput()` が handlers detach、`Close()`、必要時 `Dispose()`、active references clear を同じメソッド内で行います。

`AccountingImportRangePromptForm` と `AccountingReverseGoalSeekForm` には、現コード上 `CloseByCode` や Form class 自身の `OnFormClosing` override はありません。これらを復活させて close gate を Form 側へ戻しません。

## 黄色セル解除

黄色セルの表示と解除は close lifecycle と結合して扱います。

- お支払い履歴 import prompt:
  - `AccountingWorkbookService.HighlightAccountingImportTargets(...)` が会計依頼書 `F15:F20` を黄色表示します。
  - `AccountingWorkbookService.ClearAccountingImportTargetHighlight(...)` が同じ `F15:F20` を解除します。
  - prompt close、FormClosing、FormClosed、sheet activation による prompt close はこの cleanup と接続します。
- 逆算ツール:
  - `AccountingWorkbookService.HighlightReverseToolTargets(...)` が対象 sheet の `F17:F20` を黄色表示します。
  - `AccountingWorkbookService.ClearReverseToolTargets(...)` が同じ `F17:F20` を解除します。
  - reverse tool close、FormClosing、FormClosed はこの cleanup と接続します。

## WorkbookCloseInteropHelper との境界

`WorkbookCloseInteropHelper` は、owner が workbook close mechanics を閉じる managed / owned close 経路で使います。会計系でも `AccountingWorkbookService.CloseWithoutSaving(...)`、`AccountingSetCreateService` の cleanup、`AccountingSetKernelSyncService.CloseWorkbookQuietly(...)` などは helper 経由です。

一方で、会計フォーム / import prompt の「Excelを閉じる」ボタンは、ユーザー操作、フォーム lifecycle、黄色セル cleanup、Excel close guard allow flag、直 `workbook.Close()`、必要時の `Application.Quit()` が一続きの contract です。この経路を helper 経由へ置き換えることは、現時点の安定化済み契約を変更する扱いになります。

## 固定する禁止事項

- 会計フォーム / import prompt の直 `workbook.Close()` を、直 close であることだけを理由に修正対象へ分類しない。
- Excel の × 印 close とフォームボタン経由 close を同一視しない。
- フォームボタン経由 close で、workbook close をフォーム `Close()` / `Dispose()` より前に移動しない。
- `Application.Quit()` の条件を、`Workbooks.Count == 0` 確認後という現行条件から広げない。
- `FormClosing` / `FormClosed` の cleanup owner を Form class 側へ戻さない。
- 黄色セル解除を close lifecycle から切り離さない。
- 逆算ツール / import prompt / 会計書類セットフォーム周辺の close 順序を、helper 化、共通化、責務整理の対象にしない。
