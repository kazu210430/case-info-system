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

## ログ棚卸し

| ログ語句 | 出力元ファイル / 処理 | 観測できる内容 | 観測できない内容 | baseline記録時の使い方 |
| --- | --- | --- | --- | --- |
| `TaskPane wait-ready start.` | `App/TaskPaneRefreshOrchestrationService.cs` `ShowWorkbookTaskPaneWhenReady(...)` | ready-show 開始、対象 workbook、active workbook、active window | その後に成功したか、retry に入ったか | CASE 新規作成直後や CASE 開き直し直後の ready-show 開始点として記録する |
| `TaskPane wait-ready attempt start.` | `App/TaskPaneRefreshOrchestrationService.cs` `TryShowWorkbookTaskPaneOnce(...)` | wait-ready の attempt 番号 | この attempt がどの分岐で終わったかの全体像 | 1回目成功か retry 入りかを見る起点にする |
| `TaskPane wait-ready retry scheduled.` | `App/TaskPaneRefreshOrchestrationService.cs` `ScheduleTaskPaneReadyRetry(...)` | wait-ready retry が予約されたこと、attempt 番号、`80ms` 遅延 | retry 後に最終成功したか | 即時成功しなかったケースの印として記録する |
| `TaskPane wait-ready retry firing.` | `App/TaskPaneRefreshOrchestrationService.cs` retry timer tick | retry timer が実際に発火したこと | 発火後の refresh 成否 | retry 実発火の有無を記録する |
| `TaskPane wait-ready early-complete because visible CASE pane is already shown.` | `App/TaskPaneRefreshOrchestrationService.cs` early-complete 判定 | visible pane early-complete が成立したこと | host がどの経路で visible になっていたか | 「既に visible pane がある場合」の主要判定として使う |
| `TaskPane wait-ready attempt refresh skipped because visible CASE pane is already shown.` | `App/TaskPaneRefreshOrchestrationService.cs` early-complete 後 | early-complete により refresh をスキップしたこと | 実際の pane 内容差分の有無 | 追加 refresh を避けたかの確認に使う |
| `TaskPane wait-ready attempt refresh.` | `App/TaskPaneRefreshOrchestrationService.cs` attempt refresh 実行後 | attempt ごとの refresh 実行有無、`refreshed=true/false` | refresh 内部で何が起きたか | attempt 成否の記録に使う |
| `TaskPane wait-ready attempt window.` | `App/TaskPaneRefreshOrchestrationService.cs` window 解決後 | attempt 時点での resolved window、active workbook 一致 | pane 表示の最終成否 | visible window 解決状況の記録に使う |
| `TaskPane wait-ready pre-visibility timing.` | `App/TaskPaneRefreshOrchestrationService.cs` `EnsureWorkbookWindowVisibleForTaskPaneDisplay(...)` | ready-show 前の visible 化補助の timing 詳細 | 体感ちらつきの有無 | 可視化補助が走ったかの深掘り時だけ使う |
| `TaskPane timer fallback prepare.` | `App/TaskPaneRefreshOrchestrationService.cs` `ScheduleWorkbookTaskPaneRefresh(...)` | fallback 準備に入ったこと、対象 workbook、resolved window | fallback timer が最後まで必要だったか | fallback 入口の有無を記録する |
| `TaskPane timer fallback immediate refresh succeeded.` | `App/TaskPaneRefreshOrchestrationService.cs` fallback 前即時 refresh 成功 | fallback 準備に入ったが timer 開始前に回復したこと | なぜ即時回復したか | fallback に入る前に復帰したケースとして記録する |
| `TaskPane timer fallback scheduled.` | `App/TaskPaneRefreshOrchestrationService.cs` pending timer 開始 | fallback timer が実際に予約されたこと、残 attempt 数 | timer 後の最終表示成否 | fallback 発動の有無として記録する |
| `TaskPane timer retry start.` | `App/TaskPaneRefreshOrchestrationService.cs` `PendingPaneRefreshTimer_Tick(...)` | fallback timer tick が始まったこと、残 attempt 数 | active CASE context fallback 分岐に入ったかの明示 | fallback 実発火の有無として記録する |
| `TaskPane timer retry result.` | `App/TaskPaneRefreshOrchestrationService.cs` `PendingPaneRefreshTimer_Tick(...)` | timer tick ごとの refresh 成否 | workbook 解決失敗時の active context fallback 分岐そのもの | fallback 後の結果記録に使う |
| `source=TaskPaneRefreshOrchestrationService action=try-refresh-start` | `App/TaskPaneRefreshOrchestrationService.cs` `TryRefreshTaskPane(...)` | refresh 試行開始、`reason`、入力 window、active state | coordinator 内部の文脈採用結果 | `WorkbookActivate` / `WindowActivate` / post-action のどこから refresh が来たかの起点にする |
| `source=TaskPaneRefreshOrchestrationService action=ignore-during-protection` | `App/TaskPaneRefreshOrchestrationService.cs` protection 判定 | TaskPaneRefresh が protection で無視されたこと | `WorkbookActivate` / `WindowActivate` の ignore 有無 | protection による refresh 抑止の証跡として使う |
| `source=TaskPaneRefreshCoordinator action=start` | `App/TaskPaneRefreshCoordinator.cs` `TryRefreshTaskPane(...)` | coordinator 内部処理開始、suppression count、kernelHomeVisible | pane 表示の最終成否 | refresh 処理本体に入ったことの確認に使う |
| `source=TaskPaneRefreshCoordinator action=context-resolved` | `App/TaskPaneRefreshCoordinator.cs` context 解決後 | role、resolved window、context 解決結果 | host reuse / render の詳細 | Kernel / CASE 切替時にどの context が採用されたかを見る |
| `source=TaskPaneRefreshCoordinator action=refresh-pane-complete` | `App/TaskPaneRefreshCoordinator.cs` `_taskPaneManager.RefreshPane(...)` 後 | pane refresh 成否、CASE の場合 warm-up 予約有無 | host reuse か render かの詳細 | refresh 完了判定として記録する |
| `source=TaskPaneRefreshCoordinator action=final-foreground-guarantee-start` / `...end` | `App/TaskPaneRefreshCoordinator.cs` `GuaranteeFinalForegroundAfterRefresh(...)` | refresh 後 foreground 回復の開始 / 終了 | protection 継続中の後続イベント連鎖 | CASE 表示後の最終前面化の観測に使う |
| `TryRefreshTaskPane suppressed.` | `App/TaskPaneRefreshCoordinator.cs` suppression count 判定 | suppression count により refresh が止まったこと | どの suppressor が count を積んだか | suppression count が効いたケースの確認に使う |
| `Case pane activation suppression prepared.` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `SuppressUpcomingCasePaneActivationRefresh(...)` | CASE pane activation suppression 準備、対象 workbook、失効時刻 | 実際に WorkbookActivate / WindowActivate で消費されたか | CASE 新規作成直後の suppression 設定確認に使う |
| `Case pane refresh suppressed.` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `ShouldSuppressCasePaneRefresh(...)` | `WorkbookActivate` / `WindowActivate` 側 suppress 発生、残回数 | suppress 後に最終表示が回復したか | suppress が実際に消費されたかを記録する |
| `Case pane activation suppression cleared.` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `ClearCasePaneSuppression(...)` | suppression が `Consumed` または `Expired` で解除されたこと | 解除前に何回 refresh が抑止されたかの全量 | suppression の終了点として記録する |
| `source=WorkbookActivateProtection action=start` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `BeginCaseWorkbookActivateProtection(...)` | foreground protection 開始、対象 workbook/window、失効時刻 | どの後続イベントを最終的に止めたか | CASE refresh 後 protection 開始の証跡に使う |
| `source=WorkbookActivateProtection action=ignore event=WorkbookActivate` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `ShouldIgnoreWorkbookActivateDuringProtection(...)` | `WorkbookActivate` が protection で無視されたこと | `WindowActivate` 側 ignore の有無 | 開き直し直後の二重 refresh 抑止確認に使う |
| `source=WindowActivateProtection action=ignore event=WindowActivate` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `ShouldIgnoreWindowActivateDuringProtection(...)` | `WindowActivate` が protection で無視されたこと | Handle 内で suppress return したケース全量 | `WindowActivate` 側 protection ignore の確認に使う |
| `source=TaskPaneRefreshProtection action=ignore` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `ShouldIgnoreTaskPaneRefreshDuringProtection(...)` | active window 基準で TaskPaneRefresh が無視されたこと | なぜ active window が protected target になったか | protection による refresh 抑止確認に使う |
| `source=WorkbookActivateProtection action=clear` | `App/KernelHomeCasePaneSuppressionCoordinator.cs` `ClearCaseWorkbookActivateProtection(...)` | protection 解除、`Expired` を含む解除理由 | 解除前の個々のイベント数 | protection 終了点の確認に使う |
| `Excel WorkbookOpen fired.` / `Excel WorkbookActivate fired.` | `App/WorkbookLifecycleCoordinator.cs` | `WorkbookOpen` / `WorkbookActivate` の発火 | `WindowActivate` の発火順 | activate 系イベント列の基準点として使う |
| `TaskPane event entry. event=WorkbookOpen` / `...WorkbookActivate` | `App/WorkbookLifecycleCoordinator.cs` | TaskPane 系入口へ渡った時点の active workbook / window | その後の suppression / refresh の採否 | 入口時点の workbook/window 状態記録に使う |
| `source=WorkbookEventCoordinator action=ignore-reentrant-activate event=WorkbookActivate` | `App/WorkbookLifecycleCoordinator.cs` | `WorkbookActivate` が protection で再入抑止されたこと | suppress count による抑止か protection かの完全分離 | `WorkbookActivate` 側の二重 refresh 抑止確認に使う |
| `source=WorkbookEventCoordinator action=suppress-refresh event=WorkbookActivate` | `App/WorkbookLifecycleCoordinator.cs` | `WorkbookActivate` で CASE pane refresh suppression が効いたこと | その後の `WindowActivate` で何が起きたか | suppression 消費確認に使う |
| `Excel WindowActivate fired.` | `ThisAddIn.cs` `Application_WindowActivate(...)` | `WindowActivate` の発火、window hwnd | Handle 内で suppression return したか refresh したか | `WindowActivate` 発火の基準点に使う |
| `source=WorkbookEventCoordinator action=enter event=WindowActivate` | `ThisAddIn.cs` `HandleWindowActivateEvent(...)` | `WindowActivate` が coordinator 入口まで到達したこと | 入口後に ignore / suppress / refresh のどれで終わったか | `WindowActivate` の入口確認に使う |
| `TaskPane event entry. event=WindowActivate` | `ThisAddIn.cs` `HandleWindowActivateEvent(...)` | `WindowActivate` 時点の workbook / event window / active window | `WindowActivatePaneHandlingService` 内の分岐そのもの | `WindowActivate` と active window の食い違い確認に使う |
| `source=ThisAddIn action=refresh-call-start` / `...end` | `ThisAddIn.cs` `RefreshTaskPane(...)` | VSTO 境界から refresh 呼び出しへ入ったこと、`reason`、入力 window、最終 result | coordinator 内の各分岐詳細 | `WindowActivate` や post-action から refresh へ進んだかの確認に使う |
| `ShowCreatedCase post-release activation suppression prepared.` | `App/KernelCasePresentationService.cs` | ready-show 前に suppression 準備が完了したこと | suppression が実際に消費されたか | CASE 新規作成直後の suppression 順序確認に使う |
| `ShowCreatedCase task pane ready-show requested.` | `App/KernelCasePresentationService.cs` | ready-show 要求発行 | その後 retry / fallback に入ったか | CASE 新規作成直後の ready-show 要求時点として記録する |
| `ShowCreatedCase cursor positioned.` | `App/KernelCasePresentationService.cs` | 初期カーソル位置調整まで到達したこと | TaskPane 表示と体感競合したか | CASE 表示順序の比較に使う |
| `NewCaseDefault timing. segment=hiddenOpenToWindowVisible` | `App/KernelCasePresentationService.cs` `NewCaseDefaultTimingLogHelper` | hidden open から window visible までの timing | TaskPane 表示成否 | CASE 新規作成直後の timing 記録に使う |
| `NewCaseDefault timing detail. segment=...` | `App/KernelCasePresentationService.cs` `NewCaseDefaultTimingLogHelper` | CASE 新規作成直後の timing 詳細 phase | phase 間の UI 見え方 | 体感差分の深掘り時だけ使う |
| `NewCaseDefault timing. segment=taskPaneReadyWaitToRefreshCompleted` | `App/KernelCasePresentationService.cs` / `App/TaskPaneRefreshCoordinator.cs` | ready-show 待機から refresh 完了まで、`completion` と `refreshed` | どのイベントで完了したかの全量 | CASE 新規作成直後の完成点として使う |
| `TaskPane existing host shown.` / `TaskPane reused.` / `TaskPane refreshed.` | `App/TaskPaneManager.cs` | 既存 host 再表示、CASE host 再利用、通常 refresh 完了 | visible pane early-complete 判定そのもの | visible pane 維持や host 再利用の観測補助に使う |
| `TaskPane host created.` | `App/TaskPaneManager.cs` | 新規 host 作成、pane role、windowKey | create 後の表示成否 | 再利用されず再作成されたかの確認に使う |

### 現行ログで追跡困難な点

- `WindowActivatePaneHandlingService` 自体には固有ログがなく、`Handle(...)` 内で protection で return したのか、suppression で return したのか、refresh へ進んだのかは周辺ログの突き合わせが必要です。
- `PendingPaneRefreshTimer_Tick(...)` の active CASE context fallback 分岐には専用ログがなく、workbook 解決失敗後に `TryRefreshTaskPane(_pendingPaneRefreshReason, null, null)` へ落ちたことを直接示す語句は現行コードでは見当たりません。
- `HasVisibleCasePaneForWorkbookWindow(...)` 自体にはログがなく、visible pane early-complete は `TaskPane wait-ready early-complete...` 側から間接的に判断します。
- `WindowActivate` 側だけを単独で完全追跡するログは不足しており、`Excel WindowActivate fired.`、`TaskPane event entry. event=WindowActivate`、`source=WindowActivateProtection action=ignore event=WindowActivate`、その後の `refresh-call-start` を組み合わせて読む必要があります。

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
