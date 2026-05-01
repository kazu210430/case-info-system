# TaskPane Control Implementation Readiness

## 目的

この文書は、TaskPane 表示制御まわりについて、実装修正に着手する前段の準備が完了していることを明記するための readiness note です。

既存の `investigation` / `observation checklist` / `baseline` は、それぞれ調査メモ、観測観点固定、記録フォーマットの役割を持ちます。この文書では、それらを前提に「次は最小実装へ進める段階に入った」という現在地だけを整理します。

## 参照 docs

- [architecture.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/architecture.md>)
- [flows.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/flows.md>)
- [ui-policy.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/ui-policy.md>)
- [taskpane-protection-ready-show-investigation.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-ready-show-investigation.md>)
- [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>)
- [taskpane-protection-baseline.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-baseline.md>)

## この文書を分ける理由

- [taskpane-protection-ready-show-investigation.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-ready-show-investigation.md>) は protection / ready-show 危険領域の事実整理メモであり、履歴として残す価値がある。
- [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>) は実機観測の観点固定 docs であり、ready 状態の宣言先ではない。
- [taskpane-protection-baseline.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-baseline.md>) は baseline 記録フォーマットと観測語句整理の docs であり、今後の実装候補判断を上書きする用途とは分けておくほうが自然である。

## 1. 準備完了事項

以下は、TaskPane 表示制御まわりで、実装修正に入る前提として完了済みと整理する事項です。

- TaskPane protection / ready-show の時系列整理は完了済み。
  - 根拠: [taskpane-protection-ready-show-investigation.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-ready-show-investigation.md>) に、protection 開始/失効、retry `80ms`、fallback timer `400ms`、visible pane early-complete、ready-show / suppression 順序、`WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の各入口を整理済み。
- 実機観測チェックリストは作成済み。
  - 根拠: [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>) に、CASE 新規作成直後、CASE 開き直し直後、visible pane early-complete、fallback timer、再入抑止の各観測シナリオを固定済み。
- baseline 記録フォーマットは作成済み。
  - 根拠: [taskpane-protection-baseline.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-baseline.md>) に、観測日、commit hash、観測項目、NG 症状、ログ語句、未確認事項を残す記録枠を定義済み。
- ログ語句棚卸しは完了済み。
  - 根拠: [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>) の `ログ棚卸し` に、ready-show、retry、fallback、suppression、protection、`WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の観測語句と用途を整理済み。
- 観測用ログ補強は完了済み。
  - 根拠: [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>) の `Added Observation Logs On 2026-05-01` と [taskpane-protection-baseline.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-baseline.md>) の `Added Baseline Log Keywords On 2026-05-01` に、`WindowActivatePaneHandlingService` 分岐と active CASE context fallback、visible-case-pane-check 結果の補強語句を整理済み。
- build 成功は確認済み。
  - 根拠: 本 readiness 整理では、build 成功と runtime 実機確認を混同しない前提を [architecture.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/architecture.md>)、[flows.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/flows.md>)、[ui-policy.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/ui-policy.md>) に従って維持したうえで、build 成功済みの状態を実装前提として扱う。
- 実機ログ確認と実機ログ解析は完了済み。
  - 根拠: [taskpane-protection-observation-checklist.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/taskpane-protection-observation-checklist.md>) で観測観点とログ語句が固定されており、この readiness 整理では、それに基づく実機ログ確認と解析が完了した状態を前提として次の実装候補へ進む。

## 2. 現時点の判断

- TaskPane 表示制御まわりは、実装着手前の準備が完了している。
- これ以上の棚卸しを前提にせず、次は最小実装に進める。
- ただし、[architecture.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/architecture.md>)、[flows.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/flows.md>)、[ui-policy.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/ui-policy.md>) が前提にしている既存責務と表示順序は崩さない。
- `WindowActivate` 起因の暴走を主因とみなす段階は終わっており、実機ログ確認の結果として、怪しい点は `WorkbookOpen` / `WorkbookActivate` 近接 refresh 側に絞れている。
- 次の一手は TaskPaneManager 分割や KernelWorkbookService 分割のような大規模整理ではない。
- 次の一手は、既存の ready-show / suppression / protection / retry / fallback を温存したまま、近接二重 refresh 抑制の最小実装を入れる方向でよい。

## 3. 次に着手する実装候補

- `WorkbookOpen` / `WorkbookActivate` の近接二重 refresh 抑制を最小実装候補とする。
- 対象 workbook は CASE workbook に限定する。
- `WindowActivate` protection / suppression / ready-show / retry / fallback timer には触らない。
- skip した場合は追加ログを出し、抑止が発生した事実を追跡できるようにする。
- 実機確認では次を確認する。
  - TaskPane が最終的に表示されること。
  - ちらつきが悪化しないこと。
  - refresh 回数が減ること。
  - 追加ログで skip 条件と通過条件を追えること。

## 4. 今回の次手ではないもの

- `TaskPaneManager` の分割
- `KernelWorkbookService` の分割
- protection duration や retry / fallback 数値の見直し
- `WindowActivate` 経路の責務変更
- CASE 以外の workbook へ広げる変更

## 5. 未確認事項

- 近接二重 refresh がどの操作条件で常時再現するか。
- skip 条件の最適な時間幅。
- 実装後の体感影響。
- fallback / suppression 未観測分岐の実地確認。

## 6. 実装時に守る前提

- [ui-policy.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/ui-policy.md>) のとおり、`WorkbookOpen` 直後に直接 UI 表示制御を追加しない。
- [flows.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/flows.md>) のとおり、CASE 表示後の ready-show 予約順序と host 再利用方針を崩さない。
- [architecture.md](</C:/Users/kazu2/Documents/案件情報System/開発用/docs/architecture.md>) のとおり、TaskPane snapshot / cache を保存・生成・実行判断の正本へ昇格させない。
- 今回の readiness は「次に最小実装へ進める」という整理であり、未確認事項を確認済み扱いへ変えるものではない。
