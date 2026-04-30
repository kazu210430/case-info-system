# TaskPane Protection / Ready-Show Observation Checklist

## 目的

この文書は、TaskPane protection / ready-show まわりの実装修正に入る前後で、同じ観点を実機観測できるようにするためのチェックリストです。

この文書は、既存 docs と現行コードから確認できる事実をもとに、観測観点と確認順を固定することを目的とします。retry 値や protection 条件の妥当性をこの文書で仕様化するものではありません。

## 前提

- `docs/flows.md` の CASE 表示フローに従って観測する。
- `docs/ui-policy.md` のとおり、TaskPane は遅延表示前提で観測する。
- `docs/taskpane-protection-ready-show-investigation.md` に整理した protection / ready-show / retry / fallback の事実を前提にする。
- CASE 表示順序、retry 値、timer 値、suppression 条件、protection 条件は変更しない。

## 参照 docs

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/a-priority-service-responsibility-inventory.md`
- `docs/taskpane-protection-ready-show-investigation.md`

## 参照コード

- `dev/CaseInfoSystem.ExcelAddIn/App/KernelCasePresentationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookLifecycleCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WindowActivatePaneHandlingService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`

## 共通記録項目

- 実施日
- 実施ブランチ / commit hash
- 観測シナリオ名
- 対象 workbook の種類
- CASE workbook / Kernel workbook の表示状態
- TaskPane の表示 / 非表示
- TaskPane が出るまでの見た目の変化
- NG 症状の有無
- 取得できたログ語句
- 未確認事項

## ログで見てよい語句

現行コードで確認できた範囲では、次の語句を観測補助に使える。

- `[KernelFlickerTrace]`
- `Excel WorkbookOpen fired.`
- `Excel WorkbookActivate fired.`
- `ShowCreatedCase task pane ready-show requested.`
- `ShowCreatedCase cursor positioned.`
- `TaskPane wait-ready start.`
- `TaskPane wait-ready attempt start.`
- `TaskPane wait-ready retry scheduled.`
- `TaskPane wait-ready retry firing.`
- `TaskPane wait-ready early-complete because visible CASE pane is already shown.`
- `TaskPane timer fallback prepare.`
- `TaskPane timer fallback scheduled.`
- `TaskPane timer fallback immediate refresh succeeded.`
- `TaskPane timer retry start.`
- `TaskPane timer retry result.`
- `action=ignore-during-protection`
- `NewCaseDefault timing. segment=hiddenOpenToWindowVisible`
- `NewCaseDefault timing. segment=taskPaneReadyWaitToRefreshCompleted`

`WindowActivatePaneHandlingService` 固有のログ語句は、今回参照したコード断面では未確認です。

## 観測シナリオ 1: CASE 新規作成直後の TaskPane 表示

### 観測の狙い

- CASE workbook 表示後に TaskPane が自然に出るか確認する。
- CASE 表示直後にちらつきがないか確認する。
- HOME セル移動や初期カーソル位置調整と TaskPane 表示が競合しないか確認する。

### 手順

1. `docs/flows.md` の CASE 新規作成フローに沿って CASE workbook を作成する。
2. CASE workbook が visible になった直後から、TaskPane が表示されるまでの画面変化を観測する。
3. CASE workbook 表示直後に、HOME セル移動または初期カーソル位置調整に見えるフォーカス移動があっても、TaskPane 表示が欠落しないか確認する。
4. 必要ならログ上で `ShowCreatedCase task pane ready-show requested.` と `ShowCreatedCase cursor positioned.` の前後関係を確認する。
5. 必要ならログ上で `TaskPane wait-ready start.`、`TaskPane wait-ready attempt start.`、`NewCaseDefault timing. segment=taskPaneReadyWaitToRefreshCompleted` を確認する。

### 確認項目

- Workbook Window 可視化後に TaskPane が出る。
- CASE 表示直後に TaskPane の表示 / 非表示が短時間で往復しない。
- 初期カーソル位置調整後も TaskPane が出たままである。
- `TaskPane wait-ready retry scheduled.` が出た場合でも、最終的に TaskPane が表示される。

## 観測シナリオ 2: CASE を開き直した直後の TaskPane 表示

### 観測の狙い

- `WorkbookActivate` / `WindowActivate` による二重 refresh が起きないか確認する。
- protection が効きすぎて TaskPane が出ない状態にならないか確認する。

### 手順

1. 既存の CASE workbook を開き、表示直後の TaskPane 挙動を観測する。
2. 開き直し直後に `WorkbookActivate` / `WindowActivate` が連続して起きそうな場面で、TaskPane の表示回数と見た目を確認する。
3. 必要ならログ上で `Excel WorkbookActivate fired.` と `action=ignore-during-protection` を確認する。
4. `WorkbookActivate` が出ても TaskPane が出ないまま止まっていないか確認する。

### 確認項目

- TaskPane refresh が二重に見えない。
- protection 中に不要な refresh は抑止されるが、TaskPane 表示自体は失われない。
- CASE workbook を開き直したあと、TaskPane が無表示のまま固まらない。

## 観測シナリオ 3: 既に visible pane がある場合

### 観測の狙い

- visible pane early-complete により余計な refresh が走らないか確認する。
- 既存 pane が維持されるか確認する。

### 手順

1. すでに CASE pane が visible な状態を作る。
2. 同じ workbook / window に対して ready-show が再度走りうる操作を行う。
3. TaskPane が消えてから出直す挙動にならないか観測する。
4. 必要ならログ上で `TaskPane wait-ready early-complete because visible CASE pane is already shown.` を確認する。

### 確認項目

- 既存の visible CASE pane がそのまま維持される。
- TaskPane の再生成や再点滅が見えない。
- `early-complete` が出た場合、追加の refresh 成功ログが不要に重ならない。

## 観測シナリオ 4: fallback timer が必要になりそうな場面

### 観測の狙い

- 対象 workbook / window が一時的に解決できない場面でも、active CASE context による補完 refresh が期待どおりか確認する。

### 手順

1. CASE workbook 表示直後で、TaskPane が即時には出ないケースを観測対象にする。
2. 可能ならログ上で `TaskPane wait-ready retry scheduled.` の後に `TaskPane timer fallback prepare.` または `TaskPane timer fallback scheduled.` が出るか確認する。
3. 対象 workbook 解決に失敗したままでも、最終的に active CASE context から TaskPane が表示されるかを観測する。
4. `TaskPane timer fallback immediate refresh succeeded.` が出た場合は、timer 開始前に回復したケースとして記録する。

### 確認項目

- ready-show 即時成功しない場合でも、最終的に TaskPane が表示される。
- fallback 系ログが出た場合、TaskPane が出ないまま終わらない。
- fallback に入ったことで CASE workbook の表示順序や操作感が大きく崩れない。

### 未確認

- fallback timer が確実に起きる再現条件は、既存 docs と今回参照コードだけでは確定しない。

## 観測シナリオ 5: `WindowActivate` / `WorkbookActivate` の再入抑止

### 観測の狙い

- CASE 切替、Kernel / CASE 切替、複数ウィンドウ時に、再入抑止が効きすぎず弱すぎず動くか確認する。

### 手順

1. CASE から別 CASE に切り替える。
2. Kernel と CASE の間を切り替える。
3. 複数ウィンドウがある場合は window 切替を行う。
4. 各操作で、TaskPane の表示 / 非表示、refresh の見た目、無反応時間の有無を観測する。
5. 必要ならログ上で `Excel WorkbookActivate fired.`、`action=ignore-during-protection`、`TaskPane timer retry start.` を確認する。

### 確認項目

- 切替のたびに TaskPane が二重表示されない。
- protection が効いても、必要な最終表示は失われない。
- CASE から CASE、Kernel から CASE、CASE から Kernel の各切替で表示順序が極端に崩れない。

### 未確認

- `WindowActivate` 側の固有ログが薄いため、再入抑止の一部は見た目と周辺ログからの確認になる。

## 実装修正前後で比較する観点

- CASE workbook 表示後に TaskPane が出るまでの順序
- TaskPane refresh の見た目の回数
- TaskPane の表示 / 非表示の安定性
- CASE 表示直後のちらつき有無
- CASE 表示直後のカーソル位置調整との競合有無
- `WorkbookActivate` / `WindowActivate` が連鎖したときの操作感
- `TaskPane wait-ready retry scheduled.`、`TaskPane timer fallback scheduled.`、`action=ignore-during-protection` の発生有無

## NG 症状リスト

- TaskPane が出ない
- TaskPane が二重に出る
- CASE 表示直後にちらつく
- Kernel HOME に戻される
- 操作後 refresh が遅れる
- Excel スタート画面が出る
- `WindowActivate` / `WorkbookActivate` が連鎖して見える
- visible pane があるのに再描画で不安定になる
- CASE workbook は見えているのに TaskPane だけ遅れて出る、または最後まで出ない

## まだ実装着手しない方がよい理由

- protection 判定は `WorkbookActivate`、`WindowActivate`、`TaskPaneRefresh` の 3 入口にまたがる。
- ready-show は複数サービスにまたがり、retry、fallback、early-complete、suppression 順序が相互依存している。
- 現行 docs でも、既存表示順序を壊さないことが前提になっている。

このため、実装修正前後で同じ観測を回せる状態を先に固定しないと、変更影響を切り分けにくい。

## 次に着手するなら最小単位

事実ベースで言える範囲では、次の最小単位は「このチェックリストを使って、3 入口の protection 判定と CASE 表示直後 ready-show の観測結果を先に集めること」です。

その先の実装修正単位は、実機結果なしにはこの文書だけで確定できません。

## 未確認事項

- `80ms` / `400ms` / `3 attempts` の正式な仕様根拠
- protection 5 秒失効の正式な設計根拠
- fallback timer が必要になる代表ケースの固定的な再現条件
- `WindowActivate` 側を単独で追いやすいログの有無
