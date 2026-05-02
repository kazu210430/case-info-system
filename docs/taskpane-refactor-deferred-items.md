# TaskPane Refactor Deferred Items

この文書は、現行 `main` の TaskPane refactor 系コードと関連 docs を基準に、今回の一連の整理で意図的に見送った項目だけを記録する。

参照前提:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-architecture.md`
- `docs/taskpane-refactor-current-state.md`
- `docs/taskpane-manager-responsibility-inventory.md`
- `docs/a-priority-service-responsibility-inventory.md`

## Dispatcher の route policy 切り出し

### 概要
`TaskPaneActionDispatcher` の `actionKind` 判定を専用 policy に切り出し、dispatcher 本体から `"doc"` / `"accounting"` の直接比較を外すことを検討した。

### 現状
`TaskPaneActionDispatcher.TryRouteSeparatedActionKind(...)` が `"doc"` と `"accounting"` を直接比較して分岐している。該当しない case action は `HandleFrozenFallbackActionEntry(...)` に集約される。`DocumentCommandService` 側には `DocumentCommandActionRoutePolicy` があるが、dispatcher 側には同等の route policy は存在しない。

### 見送った理由
- 現コードのコメントで、dispatcher は separated action kinds の route と frozen fallback path の単一入口を持つ thin shell として固定されている。
- fallback 側は target resolution、例外処理、post-action refresh 順序をまとめて保持しており、route 境界だけ先に分離すると責務境界が二重化しやすい。
- `docs/taskpane-architecture.md` は action dispatch 分離済みを現在地として固定しており、今回の到達点に dispatcher 専用 route policy 追加までは含めていない。

### 将来やる条件
fallback path の責務分解が先に確定し、separated route と fallback route の境界を専用テストで固定できる場合に限り再検討する。

### 関連箇所
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneActionDispatcher.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneCaseFallbackActionExecutor.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/DocumentCommandActionRoutePolicy.cs`
- `dev/CaseInfoSystem.Tests/TaskPaneActionDispatcherTests.cs`

## document / accounting handler の共通シェル統合

### 概要
`TaskPaneCaseDocumentActionHandler` と `TaskPaneCaseAccountingActionHandler` の重複しているシェル処理を共通化することを検討した。

### 現状
両 handler は、host / workbook 解決、workbook 未解決時の state 描画、fallback executor 呼出、post-action refresh 呼出、例外処理の構造が同じで、差分は action kind 定数だけである。共通 base class や共通 helper は導入していない。

### 見送った理由
- 今回の到達点は、dispatcher から action kind ごとの handler を明示的に分けるところまでであり、その上に新しい共通 abstraction は追加していない。
- `doc` 側の prompt 準備順序は `TaskPaneBusinessActionLauncher` が `DocumentNamePromptService.TryPrepare(...)` を先に呼ぶ構造で固定されており、共通 shell を先に作ると action kind ごとの差分境界が見えにくくなる。
- 現在の 2 handler はどちらも薄く、まずは分離後の責務境界を固定することを優先した。

### 将来やる条件
handler ごとの差分が action kind 以外にも増えず、prompt 順序と post-action refresh 契約を崩さない共通化単位をテスト込みで固定できる場合に限り再検討する。

### 関連箇所
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneCaseDocumentActionHandler.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneCaseAccountingActionHandler.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneBusinessActionLauncher.cs`
- `dev/CaseInfoSystem.Tests/TaskPaneActionDispatcherTests.cs`

## CASE fallback path の追加分離

### 概要
dispatcher に残っている fallback path をさらに分解し、`caselist` / unsupported を含む残経路を別 service 群へ外出しすることを検討した。

### 現状
`TaskPaneActionDispatcher` は `HandleFrozenFallbackActionEntry(...)` と `HandlePostActionRefresh(...)` を持ち、fallback path の target resolution、例外処理、post-action refresh 順序をまとめて保持している。`TaskPaneCaseFallbackActionExecutor` は `TaskPaneBusinessActionLauncher` への薄い委譲であり、fallback path 自体の policy は持たない。

### 見送った理由
- 現コードのコメントで、fallback path は intentionally split further するまで単一入口を維持し、post-action refresh / render / show の既存順序を保つ前提になっている。
- `caselist` は `TaskPanePostActionRefreshPolicy` で defer と signature invalidation を伴い、`doc` / `accounting` とは前景維持の扱いが異なる。
- `docs/taskpane-refactor-current-state.md` と `docs/a-priority-service-responsibility-inventory.md` は、TaskPane refresh 本線と host 再利用まわりを未着手・保留として固定している。

### 将来やる条件
fallback path の route、post-action refresh、display request の順序を別 service 契約として固定でき、`caselist` の遅延 refresh 挙動を専用テストで保持できる場合に限り再検討する。

### 関連箇所
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneActionDispatcher.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneCaseFallbackActionExecutor.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPanePostActionRefreshPolicy.cs`
- `dev/CaseInfoSystem.Tests/TaskPaneActionDispatcherTests.cs`

## TaskPaneManager 本体の追加分割

### 概要
`TaskPaneManager` に残っている host 管理中心責務と refresh flow 中心責務をさらに分割することを検討した。

### 現状
`TaskPaneDisplayCoordinator`、`TaskPaneActionDispatcher`、`CasePaneSnapshotRenderService`、`CasePaneCacheRefreshNotificationService`、`TaskPaneHostRegistry` などは既に分離済みである。一方で `TaskPaneManager` には host 選択、role 別 render 切替、CASE host 再利用、`TaskPaneRefreshFlowCoordinator`、`RemoveStaleKernelHosts(...)` などが残っている。

### 見送った理由
- `docs/taskpane-manager-responsibility-inventory.md` は host 再利用、visible pane early-complete、Workbook / Window 境界を危険度 `A` として扱っている。
- `docs/taskpane-refactor-current-state.md` は `TaskPaneManager` を未着手・保留領域に残しており、完了済みへ移していない。
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の境界に近い責務が残っているため、今回の dispatcher 周辺整理と同時に触る範囲ではない。

### 将来やる条件
ready-show / protection / retry を含む表示本線の観測結果と既存契約が固定され、1 回の変更で 1 責務だけを外へ出せる状態になった場合に限り再検討する。

### 関連箇所
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `docs/taskpane-manager-responsibility-inventory.md`
- `docs/taskpane-refactor-current-state.md`

## TaskPaneHostRegistry / ThisAddIn 境界の追加整理

### 概要
`TaskPaneHostRegistry` と `ThisAddIn` の境界をさらに薄くし、VSTO `CustomTaskPane` 生成・表示要求まわりの直接依存を縮小することを検討した。

### 現状
`TaskPaneHostRegistry` は独立クラス化済みだが、`TaskPaneManager` は `ThisAddIn` を constructor で受け取り、`TaskPaneActionDispatcher.RefreshCaseHostAfterAction(...)` は `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)` を使う。`docs/taskpane-architecture.md` と `docs/thisaddin-boundary-inventory.md` はこの境界を今後課題として残している。

### 見送った理由
- `TaskPaneHostRegistry` は action event 配線と VSTO `CustomTaskPane` 生成境界を持ち、`ThisAddIn` 側の lifecycle と密接に結び付いている。
- `docs/taskpane-refactor-current-state.md` は `TaskPaneHostRegistry` と `ThisAddIn` 境界を、action dispatch や refresh 本線とは分けて慎重に扱う対象として固定している。
- dispatcher 周辺の責務整理と同時にこの境界まで動かすと、起動、終了、pane 表示の波及範囲が広がる。

### 将来やる条件
VSTO 境界を触る対象を `TaskPaneHostRegistry` または `ThisAddIn` のどちらか一方に限定し、startup / shutdown / pane 表示順序を別途検証できる状態になった場合に限り再検討する。

### 関連箇所
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneHostRegistry.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneActionDispatcher.cs`
- `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`
- `docs/taskpane-architecture.md`
- `docs/thisaddin-boundary-inventory.md`

## ready-show / protection / retry 本線への着手

### 概要
`KernelCasePresentationService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator`、`WorkbookLifecycleCoordinator` にまたがる ready-show / protection / retry 本線の整理を今回同時に進めることを検討した。

### 現状
`docs/taskpane-refactor-current-state.md` と `docs/a-priority-service-responsibility-inventory.md` は、この領域を未着手・保留として固定している。retry `80ms`、fallback timer `400ms`、`3 attempts`、visible pane early-complete、protection `5 秒` 失効はコード上の事実として整理済みだが、正式な仕様根拠は未確認として残っている。

### 見送った理由
- `docs/taskpane-refactor-current-state.md` は protection / ready-show / retry / suppression を含む危険領域を、完了済みへ移さないと明記している。
- `docs/a-priority-service-responsibility-inventory.md` は、この本線を実機観測と別途調査が必要な領域として残している。
- dispatcher 周辺の責務整理と同一変更で扱うと、UI 表示順序、foreground 維持、window 解決の挙動差分を切り分けにくい。

### 将来やる条件
retry 値、fallback 条件、protection 適用範囲、visible pane early-complete の観測結果と維持契約を docs とテストで先に固定できた場合に限り再検討する。

### 関連箇所
- `dev/CaseInfoSystem.ExcelAddIn/App/KernelCasePresentationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookLifecycleCoordinator.cs`
- `docs/taskpane-refactor-current-state.md`
- `docs/a-priority-service-responsibility-inventory.md`
