# TaskPane Protection / Ready-Show Baseline

## 目的

この文書は、TaskPane protection / ready-show まわりの実装修正前 baseline を残すための記録用 docs です。

今回の記録は [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>) に沿って整理します。

ただし、この Codex 作業では Excel / VSTO Add-in の実機観測を実施していません。したがって、実機観測が必要な項目は推測で埋めず、`未実施` または `ログ未確認` として残します。

## 参照 docs

- [architecture.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/architecture.md>)
- [flows.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/flows.md>)
- [ui-policy.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/ui-policy.md>)
- [taskpane-protection-ready-show-investigation.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-ready-show-investigation.md>)
- [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>)

## 1. 観測日

- `2026-05-01`
- 補足: 実機観測は未実施です。この日は baseline 記録 docs の作成日です。

## 2. 観測対象 commit hash

- `11bb00850b78bd9bb946fe9ce12910448bece95e`

## 3. 観測環境

- git 基準点: `main`
- 実施環境: Codex shell session
- Excel 実機観測: 未実施
- Excel プロセス確認: 未検出
- 実行時ログファイル確認: ログ未確認

## 4. CASE 新規作成直後の TaskPane 表示

- 観測結果: 未実施
- 記録欄:
  - CASE workbook 表示後に TaskPane が自然に出るか: 未実施
  - ちらつきの有無: 未実施
  - HOME セル移動や初期カーソル位置調整との競合有無: 未実施
  - ログ確認: ログ未確認

## 5. CASE を開き直した直後の TaskPane 表示

- 観測結果: 未実施
- 記録欄:
  - `WorkbookActivate` / `WindowActivate` による二重 refresh の有無: 未実施
  - protection が効きすぎて TaskPane が出ない状態の有無: 未実施
  - ログ確認: ログ未確認

## 6. 既に visible pane がある場合

- 観測結果: 未実施
- 記録欄:
  - visible pane early-complete により余計な refresh が走らないか: 未実施
  - 既存 pane が維持されるか: 未実施
  - ログ確認: ログ未確認

## 7. fallback timer が関係しそうな場面

- 観測結果: 未実施
- 記録欄:
  - 対象 workbook / window が一時的に解決できない場面の再現: 未実施
  - active CASE context による補完 refresh の有無: 未実施
  - `fallback timer` 系ログ確認: ログ未確認

## 8. `WindowActivate` / `WorkbookActivate` の再入抑止

- 観測結果: 未実施
- 記録欄:
  - CASE 切替時の再入抑止: 未実施
  - Kernel / CASE 間切替時の再入抑止: 未実施
  - 複数ウィンドウ時の再入抑止: 未実施
  - ログ確認: ログ未確認

## 9. NG 症状の有無

- TaskPane が出ない: 未実施
- TaskPane が二重に出る: 未実施
- CASE 表示直後にちらつく: 未実施
- Kernel HOME に戻される: 未実施
- 操作後 refresh が遅れる: 未実施
- Excel スタート画面が出る: 未実施
- `WindowActivate` / `WorkbookActivate` が連鎖する: 未実施

## 10. ログで確認できた語句

- ログ確認結果: ログ未確認
- 補足: [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>) に記載した語句群は、今回の Codex 作業では runtime ログ上で照合していません。

## 11. 未実施項目

- CASE 新規作成直後の TaskPane 表示観測
- CASE を開き直した直後の TaskPane 表示観測
- visible pane early-complete 観測
- fallback timer 関連シナリオ観測
- `WindowActivate` / `WorkbookActivate` の再入抑止観測
- runtime ログ照合

## 12. 未確認事項

- 現行 `main` を実機で観測するための Excel / VSTO 実行環境が、この Codex 作業では未確認です。
- `80ms` / `400ms` / `3 attempts` の正式な仕様根拠
- protection 5 秒失効の正式な設計根拠
- fallback timer が必要になる代表ケースの固定的な再現条件
- `WindowActivate` 側を単独で追いやすいログの有無

## 13. 今後コード修正時に比較すべき観点

- CASE workbook 表示後に TaskPane が出るまでの順序
- CASE 表示直後のちらつき有無
- CASE 表示直後のカーソル位置調整と TaskPane 表示の競合有無
- `WorkbookActivate` / `WindowActivate` による二重 refresh の有無
- visible pane early-complete により追加 refresh を避けられているか
- fallback timer が必要になった場合でも最終表示が失われないか
- protection が効きすぎて TaskPane が出ない状態にならないか
- `TaskPane wait-ready retry scheduled.`
- `TaskPane timer fallback scheduled.`
- `action=ignore-during-protection`

## 次回の追記方針

- 実機観測が可能な環境で、この文書の `未実施` を実測結果へ置き換える。
- 観測できなかった項目は `OK` にせず、引き続き `未実施` または `未確認` のまま残す。
