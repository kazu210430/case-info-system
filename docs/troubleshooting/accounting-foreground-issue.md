# Accounting workbook 前面表示問題 調査記録

## 1. 問題概要

- CASE から会計書類セットを開くと、会計書類セット.xlsx が CASE の背後に表示されることがあった。
- 実機では「いったん前面に出た後に奪われる」というより、「最初から前面に出ていない」ように見えた。
- ログ上は Accounting workbook が `ActiveWorkbook` / `ActiveWindow` になる時点があったが、それだけでは OS 前面表示は保証されていなかった。

## 2. 初期仮説（誤り）

- 最初は、CASE の `WorkbookActivate` による前面奪取が主因だと考えた。
- そのため、CASE 側の refresh や foreground guarantee を suppression で止めれば解消できると見た。
- しかし、suppression で吸収できたのは後続の再描画だけであり、実機の「最初から背後に見える」現象は残った。

## 3. 試行と結果

- suppression（1回）  
  第1波の CASE `WorkbookActivate` / `WindowActivate` は吸収できたが、前面表示問題は残った。
- suppression 拡張  
  後続 activation 波の吸収方向は見えたが、本質的に「Accounting workbook が OS 前面に出ていない」問題は解消しなかった。
- wait UI 修正  
  wait form の `TopMost` / activate 挙動を見直しても、実機では再現した。

## 4. 決定的な観察

- Excel 内の active 状態と、OS 上の前面表示は別だった。
- `TaskPaneRefreshCoordinator` の foreground guarantee は、常に OS 前面化を意味していなかった。
- `recovered=True` は「前面化成功」ではなく、recovery 評価処理が完了したことを示すだけだった。
- 実ログでは `bringToFront=true` でも、`appRecovered=False`、`screenUpdatingRecovered=False`、`windowVisibleRecovered=False`、`windowStateRecovered=False` のケースがあった。
- このケースでは `PromoteExcelWindow(...)` が呼ばれず、Accounting workbook は Excel 内では active でも、OS 前面化までは保証されていなかった。

## 5. 根本原因

- `ExcelWindowRecoveryService` では、`bringToFront=true` でも「何かを回復した場合」にしか `PromoteExcelWindow(...)` を呼ばない条件分岐になっていた。
- そのため、対象 workbook window がすでに visible で、Excel 内ではすでに active になっているケースでは、前面化要求があっても promotion が実行されなかった。
- 結果として、「active 状態にはなっているが、OS 前面表示は保証されない」経路が残っていた。

## 6. 修正内容

- `ExcelWindowRecoveryService` の条件分岐を修正した。
- `bringToFront=true` の場合は、回復有無に関係なく `PromoteExcelWindow(...)` を実行するようにした。
- 他の recovery 手順、suppression、render、`TaskPaneManager`、refresh フローには変更を入れていない。

## 7. 結果

- 実機確認で、会計書類セット.xlsx の前面表示は正常化した。
- これにより、今回の問題は suppression ではなく、OS 前面化条件の不足が原因だったことを確認できた。

## 8. 学び（重要）

- `ActiveWorkbook` は前面表示を意味しない。
- foreground guarantee という名前でも、実装上は前面化保証になっていない場合がある。
- suppression は後続処理の吸収には有効でも、前面化不足そのものの解決にはならない。
- UI 問題は、Excel 内部状態と OS レベル表示を分けて考える必要がある。

## 9. 今後の指針

- 前面化が必要な経路では、OS レベルの window promotion を明示的に保証する。
- `bringToFront` の意味は「条件付きで試す」ではなく、「対象 window を前面化する」として扱う。
- `recovered` フラグや active 状態だけをもって、前面表示成功と判断しない。
