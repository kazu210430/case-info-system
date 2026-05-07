# TaskPaneManager Responsibility Inventory

## 目的

`TaskPaneManager` 周辺の current-state を、現行 `main` のコードと既存 docs を前提に再棚卸しする。

今回の目的は「巨大クラスを小さく見せること」ではなく、次フェーズで安全に切れる単位を決めるために、

- ownership
- 変更理由
- runtime state owner
- composition root 候補
- UX / WorkbookOpen / visibility 依存

を事実ベースで整理することです。

## 参照した docs

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-architecture.md`
- `docs/taskpane-refresh-policy.md`
- `docs/taskpane-refactor-current-state.md`
- `docs/taskpane-refactor-deferred-items.md`
- `docs/workbook-window-activation-notes.md`
- `docs/taskpane-protection-ready-show-investigation.md`
- `docs/thisaddin-boundary-inventory.md`
- `docs/a4-c2-current-state.md`
- `docs/a-priority-service-responsibility-inventory.md`

## 調査対象コード

- `dev/CaseInfoSystem.ExcelAddIn/AddInCompositionRoot.cs`
- `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneHostRegistry.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneHostFactory.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneHostLifecycleService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneHostFlowService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneDisplayCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneActionDispatcher.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneNonCaseActionHandler.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookTaskPaneReadyShowAttemptWorker.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/CasePaneCacheRefreshNotificationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/UI/TaskPaneHost.cs`

## current-state 補足

- 現行コードでは `TaskPaneRefreshFlowCoordinator` は存在せず、`RefreshPane(...)` 主経路の owner は `TaskPaneHostFlowService` です。
- `TryReuseCaseHostForRefresh(...)` は `TaskPaneManager` ではなく `TaskPaneHostFlowService` にあります。
- `RemoveStaleKernelHosts(...)` は `TaskPaneManager` ではなく `TaskPaneHostLifecycleService` にあります。
- したがって、旧 docs に残る `TaskPaneRefreshFlowCoordinator` / `TaskPaneManager.TryReuseCaseHostForRefresh(...)` / `TaskPaneManager.RemoveStaleKernelHosts(...)` は current-state とはずれています。

## 対象フロー要約

- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` は `WorkbookLifecycleCoordinator` と `WindowActivatePaneHandlingService` から入り、TaskPane 本線は `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` を経由して `TaskPaneManager.RefreshPane(...)` に到達する。
- `WorkbookOpen` は workbook-only 境界であり、window-dependent な pane 対象決定は `ResolveWorkbookPaneWindow(...)` と後続イベントへ委ねる。
- `TaskPaneManager` 自体は event 入口を持たず、host lifecycle / display / role 別 render / action dispatch の facade 面として振る舞う。
- CASE ready-show は `WorkbookTaskPaneReadyShowAttemptWorker` が担当し、既存 visible CASE pane があれば `HasVisibleCasePaneForWorkbookWindow(...)` を使って success 相当で early-complete する。

## 責務棚卸し

| 分類 | 現在の owner | runtime state owner | 変更理由 | AddInCompositionRoot 側候補 | manager / facade に残すべきか | 危険度 |
| --- | --- | --- | --- | --- | --- | --- |
| runtime composition / wiring | `AddInTaskPaneCompositionFactory` と `TaskPaneManagerRuntimeGraphFactory` に集約し、manager attach は runtime-consumed collaborator のみ | なし | helper / handler / registry / display / flow の wiring を変える時に変わる | はい。最終的には `AddInTaskPaneCompositionFactory` 側へ寄せるのが自然 | いいえ。facade の変更理由ではない | `Safe extraction` |
| facade entry surface | `TaskPaneManager` | `TaskPaneManager` | `ThisAddIn` / bridge / test から呼ぶ surface を変える時に変わる | いいえ | はい。`RefreshPane`、`Hide*`、`Has*`、`DisposeAll` などの薄い入口は残してよい | `Runtime-sensitive` |
| host registry data | 実装上は `TaskPaneManager` が `_hostsByWindowKey` を所有し、`TaskPaneHostRegistry` / `TaskPaneHostLifecycleService` / `TaskPaneDisplayCoordinator` / action resolver 群が共有利用 | 実装上は `TaskPaneManager` | host 集合の持ち方や lookup 方式を変える時に変わる | 条件付き。まず owner を docs で固定しないと危険 | facade 直下に state を持つのは薄い facade と相性が悪い | `Runtime-sensitive` |
| host lifecycle primitive | `TaskPaneHostLifecycleService` | registry 共有 host map | get-or-replace / remove / dispose / workbook 単位 cleanup の semantics を変える時に変わる | composition だけ root 側へ寄せられる | facade には残さない | `Runtime-sensitive` |
| stale kernel host cleanup | `TaskPaneHostLifecycleService` | registry 共有 host map | Kernel host cleanup 条件を変える時に変わる | composition だけ root 側へ寄せられる | facade には残さない | `WorkbookOpen-sensitive` |
| refresh-time host flow | `TaskPaneHostFlowService` | registry 共有 host map と host metadata | host 選択 / reuse / render 要否 / show 順序を変える時に変わる | composition だけ root 側へ寄せられる | facade には残さない | `Runtime-sensitive` |
| display coordinator | `TaskPaneDisplayCoordinator` | registry 共有 host map と host visibility | hide/show / prepare-before-show / visible pane 判定を変える時に変わる | composition だけ root 側へ寄せられる | facade には残さない | `UX-sensitive` |
| pane creation / VSTO 実体生成 | `TaskPaneHostRegistry` -> `TaskPaneHostFactory` -> `TaskPaneHost` -> `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)` | `TaskPaneHost` / `CustomTaskPane` 実体 | VSTO pane の create / dispose / dock / action event 配線を変える時に変わる | すぐには寄せない方がよい | facade に残さない。VSTO 境界として別扱い | `今は触るべきでない` |
| control binding / UI event subscription | `TaskPaneHostFactory` | host / control 実体 | control 種別と `ActionInvoked` 配線を変える時に変わる | すぐには寄せない方がよい | facade に残さない | `Runtime-sensitive` |
| CASE action dispatch | `TaskPaneActionDispatcher` と handler 群 | host lookup は registry 共有 host map、post-action refresh は `ThisAddIn` 再入あり | action route / post-action refresh order / error handling を変える時に変わる | はい。dispatcher subtree の composition は root 側候補 | facade には残さない | `UX-sensitive` |
| non-case action dispatch | `TaskPaneNonCaseActionHandler` | host lookup は registry 共有 host map | Kernel / Accounting action handling を変える時に変わる | composition だけ root 側へ寄せられる | facade には残さない | `Runtime-sensitive` |
| role 別 render 最終切替 | `TaskPaneManager.RenderHost(...)` | host metadata (`WorkbookFullName`, `LastRenderSignature`) | role ごとの render seam を変える時に変わる | 条件付き。まず render seam を facade の外側へ出す根拠が必要 | 当面は yes。facade の最後の role switch として残してよい | `Runtime-sensitive` |
| CASE snapshot render / view state | `CasePaneSnapshotRenderService`、`ICaseTaskPaneSnapshotReader`、`CaseTaskPaneViewStateBuilder` | CASE workbook / snapshot build result | CASE pane content の生成方法を変える時に変わる | 既に root 側で一部 compose 済み | facade には残さない | `Safe extraction` |
| CASE render 後副作用 | `CasePaneCacheRefreshNotificationService` | workbook `Saved` 状態と build result | cache 更新通知、`WorkbookOpen` / `WorkbookActivate` timing、`Saved` restore を変える時に変わる | composition だけ root 側へ寄せられる | manager 本体に残さない方が筋は良いが、意味的には lifecycle 寄り | `WorkbookOpen-sensitive` |
| WorkbookOpen / Activate / WindowActivate reaction | `WorkbookLifecycleCoordinator`、`WindowActivatePaneHandlingService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator` | refresh retry / suppression / protection state | event 境界や defer 条件を変える時に変わる | いいえ。既に root で compose 済み | manager に残さない | `WorkbookOpen-sensitive` |
| visible pane early-complete | `WorkbookTaskPaneReadyShowAttemptWorker` + `ICasePaneHostBridge` + `TaskPaneManager.HasVisibleCasePaneForWorkbookWindow(...)` | host metadata (`WorkbookFullName`, `IsVisible`, windowKey) | ready-show success 判定を変える時に変わる | いいえ | manager facade には `HasVisibleCasePaneForWorkbookWindow(...)` surface だけ残る余地あり | `visibility/foreground-sensitive` |
| foreground / protection / final recovery | `TaskPaneRefreshCoordinator`、`KernelHomeCasePaneSuppressionCoordinator` | protection state, active window state | foreground 保証や protection 入口を変える時に変わる | いいえ | manager に残さない | `visibility/foreground-sensitive` |
| CASE hidden create handoff 後の pane 表示 | owner は `KernelCasePresentationService` 側。TaskPane 側は ready-show の downstream | hidden create session 自体は TaskPane owner ではない | hidden create -> shared app handoff 後 UX を変える時に変わる | いいえ | manager に残さない | `hidden-session-sensitive` |
| logging / diagnostic helper | `TaskPaneManagerDiagnosticHelper` と各 coordinator | なし | trace 形式や descriptor を変える時に変わる | 任意 | 当面はどちらでもよい | `C` 相当 |

## ownership 混在

### 1. runtime composition owner が分裂している

- `AddInTaskPaneCompositionFactory.Compose(...)` が `TaskPaneManager`、`WindowActivatePaneHandlingService`、`TaskPaneRefreshCoordinator`、`TaskPaneRefreshOrchestrationService` を組み立てる。
- 同時に `TaskPaneManager` constructor 自身が `CasePaneCacheRefreshNotificationService`、`TaskPaneHostRegistry`、`TaskPaneHostLifecycleService`、`TaskPaneDisplayCoordinator`、`TaskPaneActionDispatcher`、`TaskPaneHostFlowService` などを `new` している。
- さらに `TaskPaneHostRegistry` constructor も `TaskPaneHostFactory` を内部で `new` している。

事実として、TaskPane 周辺には

- composition root
- secondary composition root
- VSTO host wiring root

が 3 層で存在します。

### 2. state owner と behavior owner がずれている

- host 集合 `_hostsByWindowKey` は `TaskPaneManager` が生成・所有している。
- しかし mutate / lookup / cleanup / visibility 判定は `TaskPaneHostRegistry`、`TaskPaneHostLifecycleService`、`TaskPaneDisplayCoordinator`、action resolver 群に分散している。

これは「runtime state owner は manager、変更理由 owner は周辺 service 群」というねじれです。

### 3. CASE / non-case 境界は handler で分かれているが wiring では混ざっている

- CASE action は `TaskPaneActionDispatcher` と handler 群へ流れる。
- Kernel / Accounting action は `TaskPaneNonCaseActionHandler` へ流れる。
- ただし control 作成と `ActionInvoked` 配線は `TaskPaneHostFactory` が一括で持ち、compose-time delegate supply は `TaskPaneManagerRuntimeGraphFactory` で閉じている。

したがって「処理 owner」と「compose-time delegate supply」は分離済みでも「binding owner」は未分離です。

### 4. CASE post-action refresh は dispatcher から VSTO display 入口へ直接戻る

- `TaskPaneActionDispatcher.RefreshCaseHostAfterAction(...)` は `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)` を直接呼ぶ。
- CASE dispatch subtree は action dispatch だけで閉じず、display 入口まで再入しています。

これは CASE action dispatch と display surface の ownership 混在です。

### 5. notification service が notification だけでは閉じていない

- `CasePaneCacheRefreshNotificationService` は message 表示だけでなく `workbook.Saved` restore と `WorkbookOpen` / `WorkbookActivate` timing 依存を持つ。
- 名前よりも実体は render 後 side effect / lifecycle adjacency です。

## composition root 化している箇所

### 実装上の composition root 候補

1. `AddInTaskPaneCompositionFactory`
   - すでに TaskPane 周辺の top-level compose owner です。
   - `TaskPaneManager` 内部 `new` を引き上げる自然な受け皿になります。

2. `TaskPaneManager` constructor
   - 現状では subgraph の大半を内部 compose しており、secondary composition root になっています。
   - facade / orchestration owner と wiring owner が同居しています。

3. `TaskPaneHostRegistry` constructor
   - `TaskPaneHostFactory` を内部 compose しており、VSTO host wiring root になっています。
   - ただしここは create / dispose / control event 配線の危険領域なので、今すぐ root 側へ押し戻す対象とは言い切れません。

## facade / orchestration / composition の境界

### facade に残すべきもの

- `TaskPaneManager` の公開 surface
  - `RefreshPane(...)`
  - `TryShowExistingPane*`
  - `HasManagedPaneForWindow(...)`
  - `HasVisibleCasePaneForWorkbookWindow(...)`
  - `Hide*`
  - `RemoveWorkbookPanes(...)`
  - `DisposeAll()`
  - `PrepareTargetWindowForForcedRefresh(...)`
- role 別 render の最終切替
  - `RenderHost(...)` と `RenderKernelHost(...)` / `RenderAccountingHost(...)` / `RenderCaseHost(...)`
  - host metadata 更新 (`WorkbookFullName`, `LastRenderSignature`) と近接しているため、現時点では facade の内側に残すほうが安全です。

### orchestration に残すべきもの

- refresh-time host flow: `TaskPaneHostFlowService`
- visibility / show-hide 調停: `TaskPaneDisplayCoordinator`
- CASE action dispatch と post-action refresh order: `TaskPaneActionDispatcher`
- non-case action handling: `TaskPaneNonCaseActionHandler`

### composition に寄せる候補

- `TaskPaneManager` constructor 内の `new` 群
  - `CasePaneCacheRefreshNotificationService`
  - `TaskPaneCaseFallbackActionExecutor`
  - `TaskPaneCaseActionTargetResolver`
  - `TaskPaneCaseDocumentActionHandler`
  - `TaskPaneCaseAccountingActionHandler`
  - `TaskPaneActionDispatcher`
  - `TaskPaneDisplayCoordinator`
  - `TaskPaneHostLifecycleService`
  - `TaskPaneHostFlowService`
- これらは「どの helper / handler 実装を組み合わせるか」という wiring change に反応するため、facade の変更理由ではありません。

## AddInCompositionRoot 側へ寄せる候補

### 優先度が高い候補

- `TaskPaneActionDispatcher` subtree の compose
  - CASE action handler / resolver / fallback executor / dispatcher は runtime wiring change に反応する。
  - `WorkbookOpen` / ready-show / protection 本線には直接触れない。

- `TaskPaneDisplayCoordinator` / `TaskPaneHostLifecycleService` / `TaskPaneHostFlowService` の compose
  - これらは current-state では既に独立 owner であり、manager 内に `new` で残す理由が薄い。
  - ただし shared host state の owner をどう持つかは先に固定が必要です。

- `CasePaneCacheRefreshNotificationService` の compose
  - leaf に近いが、意味的には lifecycle adjacency を持つ。
  - wiring change と side effect policy の境界を分けて扱うためにも、manager constructor からは外したほうが読みやすい。

### まだ寄せない方がよい候補

- `TaskPaneHostRegistry` / `TaskPaneHostFactory`
  - `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)` と control event 配線を持つ。
  - startup / shutdown / pane dispose 順序と密結合しており、別タスク化すべきです。

- `TaskPaneHost` / `ThisAddIn` の VSTO boundary
  - `CustomTaskPane` 実体生成・破棄です。
  - facade / composition cleanup と同時に触るべきではありません。

## CASE pane / non-case pane の責務線

### 既に分かれている線

- CASE action 実行
  - `TaskPaneActionDispatcher`
  - `TaskPaneCaseDocumentActionHandler`
  - `TaskPaneCaseAccountingActionHandler`
  - `TaskPaneCaseFallbackActionExecutor`

- Kernel / Accounting action 実行
  - `TaskPaneNonCaseActionHandler`

- CASE pane content render
  - `CasePaneSnapshotRenderService`
  - `CasePaneCacheRefreshNotificationService`

### まだ曖昧な線

- control 作成と event 配線
  - `TaskPaneHostFactory` が CASE / Kernel / Accounting 全部をまとめて配線する。
- host store
  - CASE / non-case 共通の `_hostsByWindowKey` を manager が所有する。
- display 入口
  - `RequestTaskPaneDisplayForTargetWindow(...)` は CASE post-action refresh と `WindowActivate` を同じ入口で受ける。

## hidden assumptions

- visible pane early-complete は host metadata に依存する
  - `WorkbookFullName`
  - `IsVisible`
  - `windowKey`
- CASE host reuse は `TaskPaneHostReusePolicy` の reason 文字列に依存する
  - `WorkbookActivate`
  - `WindowActivate`
  - `KernelHomeForm.FormClosed`
- display request 判定は host render signature の current 判定に依存する
- `CasePaneCacheRefreshNotificationService` は `WorkbookOpen` / `WorkbookActivate` だけで通知する前提を持つ
- post-action refresh は `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)` 経由で display policy に再入する
- `WorkbookOpen` 自体は manager owner ではないが、post-render side effect と ready-show early-complete が event timing に間接依存する

## 危険度仕分け

| 区分 | 対象 | 事実ベースの理由 |
| --- | --- | --- |
| `Safe extraction` | `TaskPaneManager` constructor 内の leaf / handler wiring の root 側移動 | runtime semantics を変えずに owner だけ整理できる可能性がある。対象は dispatcher subtree や notification service compose など |
| `Runtime-sensitive` | shared host map owner、`TaskPaneHostLifecycleService` / `TaskPaneHostFlowService` compose、control binding delegate 受け渡し | host lookup / replace / render/show 直列順序に触れるため |
| `UX-sensitive` | `TaskPaneDisplayCoordinator`、post-action refresh 再表示、`PaneDisplayPolicy` 入口 | 見え方、再表示、前面維持に直結するため |
| `WorkbookOpen-sensitive` | `TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator`、`CasePaneCacheRefreshNotificationService` timing、window resolve 近辺 | `WorkbookOpen` を window 安定境界にしない契約を壊しやすいため |
| `visibility/foreground-sensitive` | visible pane early-complete、ready-show、final foreground guarantee、protection 3入口 | pane visible 判定と foreground 回復が連動しているため |
| `hidden-session-sensitive` | CASE hidden create handoff 後の ready-show / final foreground 連鎖 | TaskPane owner 自体は hidden session を持たないが、shared app handoff 後 UX に依存しているため |
| `今は触るべきでない` | `TaskPaneHostRegistry` / `TaskPaneHostFactory` / `TaskPaneHost` / `ThisAddIn.CreateTaskPane` / `RemoveTaskPane`、ready-show / protection / retry 本線 | VSTO 実体境界と実 UX 危険領域が重なるため |

## 次に切るべき安全単位

### 第1候補

`TaskPaneManager` の internal composition を `AddInTaskPaneCompositionFactory` へ寄せるための docs / constructor graph 整理

- 対象は runtime wiring だけに限定する
- まずは dispatcher subtree と notification service から入る
- host map の owner と VSTO create / dispose 境界はまだ変えない

### 第2候補

`TaskPaneManager` が内部 `new` している orchestration services の compose owner を root 側へ寄せる

- `TaskPaneDisplayCoordinator`
- `TaskPaneHostLifecycleService`
- `TaskPaneHostFlowService`

ただし、この段階では

- `_hostsByWindowKey` owner
- `TaskPaneHostRegistry`
- `TaskPaneHostFactory`
- `TaskPaneHost`

には同時に触れないほうが安全です。

### 第3候補

`TaskPaneHostRegistry` / `TaskPaneHostFactory` / `ThisAddIn` の VSTO boundary 整理を独立タスク化する

- create / remove
- control event binding
- startup / shutdown / dispose order

を別検証に分けるべきです。

## 推奨 refactor 順序

1. docs current-state をコードへ同期する
2. `TaskPaneManager` の internal composition owner を整理する
3. shared host map の owner をどこに置くかを明文化する
4. dispatcher subtree の compose を root 側へ寄せる
5. display / lifecycle / host flow の compose を root 側へ寄せる
6. VSTO host boundary は最後に独立して扱う
7. ready-show / protection / retry は別フェーズのまま維持する

## 今回の結論

- 見つかった ownership 混在
  - `TaskPaneManager` が facade でありながら shared host map owner と secondary composition root を兼ねている
  - CASE action dispatch が post-action refresh で `ThisAddIn` の display 入口へ再入している
  - `CasePaneCacheRefreshNotificationService` が notification 以上の lifecycle timing を持っている

- composition root 化している箇所
  - `AddInTaskPaneCompositionFactory`
  - `TaskPaneManager` constructor
  - `TaskPaneHostRegistry` constructor

- facade に残すべきもの
  - `TaskPaneManager` の外部 surface
  - role 別 render の最終切替
  - visible pane / existing pane 参照の facade 入口

- AddInCompositionRoot へ寄せる候補
  - dispatcher subtree
  - notification service
  - display / lifecycle / host flow の composition-time wiring

- “今はまだ触るべきでない” 箇所
  - `TaskPaneHostRegistry` / `TaskPaneHostFactory` / `TaskPaneHost` / `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)`
  - visible pane early-complete
  - ready-show / retry / protection
  - `WorkbookOpen` window-safe boundary

- 実UXリスク
  - foreground 回復
  - visible pane early-complete
  - CASE create hidden route handoff 後の ready-show
  - `WorkbookActivate` / `WindowActivate` 再利用経路

- 事実として言える次フェーズ候補
  - まずは runtime behavior を変えずに「composition owner の移動」だけを切る
  - VSTO boundary と ready-show 本線は次の別タスクに分ける

## 不明点

- retry `80ms` / pending retry `400ms` / `3 attempts` / protection `5秒` の正式な仕様根拠
- visible pane early-complete の最終 UX 定義
- shared host map を manager 外へ出すときの最小 state holder が何か
- `TaskPaneHostRegistry` / `TaskPaneHostFactory` / `ThisAddIn` 境界をどの形で整理するのが最小差分か

これらは current docs とコードだけでは断定しません。
## B1 Status (2026-05-06)

- Production compose owner for the TaskPaneManager constructor graph moved from `TaskPaneManager` to `AddInTaskPaneCompositionFactory`.
- `TaskPaneManager` still owns `_hostsByWindowKey`, host existence state, and the facade/orchestration entry surface.
- The extracted compose set is: notification wiring, host registry/lifecycle/display wiring, dispatcher subtree wiring, and host-flow wiring.
- `WorkbookOpen`, ready-show retry, protection flow, visibility/foreground handling, `ThisAddIn.CreateTaskPane`, `RemoveTaskPane`, and `TaskPaneHostRegistry` ownership redesign remain untouched.
- Test paths now use the same external runtime-graph factory after manager construction; this is the only remaining manager-adjacent compose path in this phase.

## B1.1 Status (2026-05-06)

- Compose entry is reduced to one bootstrap boundary: `TaskPaneManagerRuntimeBootstrap.CreateAttached(...)`.
- Attach timing is now explicit at the bootstrap boundary instead of being repeated at each caller.
- Graph composition now reads a passive compose context, so the manager no longer exposes most runtime dependencies just to feed graph assembly.
- The manager still owns host state, render callbacks, and facade/orchestration behavior; this phase did not redesign `TaskPaneHostRegistry`, VSTO lifecycle, or any visibility / ready-show / retry flow.

## B1.2 Status (2026-05-06)

- Production runtime entry is now bootstrap-only in name and access pattern: `CreateAttached(...)` for production, `CreateAttachedForTests(...)` / `CreateThinAttachedForTests(...)` for harness paths.
- Unused convenience constructors were removed, and raw construction/attach now sit behind `TaskPaneManager.RuntimeBootstrapAccess` instead of remaining general internal seams.
- Test and snapshot ergonomics remain supported through explicit harness entrypoints that preserve the same build/attach order as production.

## B1.3 Status (2026-05-06)

- `RuntimeBootstrapAccess` is documented and scoped as a bootstrap-only bridge to private manager construction and private attach.
- The compose input is split into entry-time and graph-compose-time contexts so the graph factory no longer receives manager-construction-only dependencies.
- `TaskPaneManagerRuntimeGraph` remains a passive bundle and still does not own runtime state, lifecycle, or orchestration.

## B2 Prep Status: VSTO Boundary Inventory (2026-05-06)

### Current owner map

| Concern | Current owner | Runtime state owner | Notes |
| --- | --- | --- | --- |
| pane creation request | `TaskPaneHostRegistry` -> `TaskPaneHostFactory` -> `TaskPaneHost` | `_hostsByWindowKey` is still owned by `TaskPaneManager`; `CustomTaskPane` instance is owned by `TaskPaneHost` | `TaskPaneHostRegistry.GetOrReplaceHost(...)` decides reuse vs replace, then `TaskPaneHostFactory.CreateHost(...)` constructs the host. |
| VSTO `CustomTaskPane` create | `ThisAddIn.CreateTaskPane(...)` | `TaskPaneHost` | `TaskPaneHost` constructor calls into `ThisAddIn` immediately, so registry/factory cleanup and VSTO create timing are coupled today. |
| pane remove request | `TaskPaneHostRegistry.RemoveHost(...)` / `RemoveWorkbookPanes(...)` / `DisposeAll()` | `_hostsByWindowKey` is still owned by `TaskPaneManager`; `CustomTaskPane` instance is still owned by `TaskPaneHost` until dispose | Remove paths are orchestrated by registry and lifecycle service, but the actual VSTO remove is still delegated to `TaskPaneHost.Dispose()`. |
| VSTO `CustomTaskPane` remove | `ThisAddIn.RemoveTaskPane(...)` via `TaskPaneHost.Dispose()` | `TaskPaneHost` | Current-state remove order is `Hide()` -> `RemoveTaskPane(...)` -> null out `_pane`. |
| host existence state | `TaskPaneManager` | `TaskPaneManager` | `_hostsByWindowKey` remains the shared state holder and should not move during VSTO boundary work. |
| control creation | `TaskPaneHostFactory` | control instance lives under `TaskPaneHost` / WinForms control tree | Kernel, Accounting, and Case controls are all created here. |
| `ActionInvoked` binding | `TaskPaneHostFactory` | bound handler lifetime is implicit in control / host lifetime | Factory wires `windowKey`-capturing handlers inline during control creation. |
| event unbinding | implicit dispose path only | implicit in control / pane teardown | There is no explicit unbinding owner today; this is one reason create/remove and binding should not be split casually. |
| host metadata storage | `TaskPaneHost` | `TaskPaneHost` | `WorkbookFullName` and `LastRenderSignature` live on the host, but write timing is owned elsewhere. |
| host metadata write timing | `TaskPaneManager.RenderHost(...)` and `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)` | `TaskPaneHost` | `WorkbookFullName` is written before role render, and `LastRenderSignature` is written after refresh-time render evaluation succeeds. |
| visible pane / early-complete evaluation | `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` | reads `TaskPaneHost.WorkbookFullName`, `TaskPaneHost.IsVisible`, `windowKey` | This is the direct dependency point for ready-show early-complete. |
| ready-show retry touchpoint | `WorkbookTaskPaneReadyShowAttemptWorker` | external retry state is outside the host map | The worker does not own host state, but it depends on the display coordinator result staying semantically stable. |
| foreground recovery touchpoint | `TaskPaneRefreshCoordinator` / `ExcelWindowRecoveryService` | outside host map | Foreground recovery is downstream UX stabilization and should not be bundled with host creation cleanup. |

### Why `_hostsByWindowKey` should not move yet

- Host create, replace, remove, hide, reuse, stale cleanup, visible-pane lookup, and action-target resolution all still observe the same map.
- Moving the map owner now would mix a state-owner refactor with VSTO create/remove timing, which would make regressions hard to localize.
- Ready-show early-complete currently depends on `windowKey` lookup plus host metadata plus visibility state; that chain should stay stable while the VSTO boundary is only being inventoried.

### Risk classification

| Category | Scope | Why |
| --- | --- | --- |
| `Safe documentation only` | owner map, metadata timing map, create/remove call chain, early-complete touchpoint map | These are fact-finding updates with no runtime impact. |
| `Safe naming/comment cleanup` | comments around `TaskPaneHostFactory`, `TaskPaneHost`, `ThisAddIn.CreateTaskPane(...)`, `RemoveTaskPane(...)` | Safe only if limited to intent clarification and no behavior change. |
| `Low-risk extraction` | none recommended yet beyond doc-backed helper naming | Even small code moves are likely to cross control binding and dispose timing. |
| `Runtime-sensitive` | `TaskPaneHostRegistry` replace/create/remove chain, `TaskPaneHostFactory` binding path, host metadata write timing | Create/remove and event binding are coupled to actual pane existence and render reuse. |
| `UX-sensitive` | `TaskPaneDisplayCoordinator` show/hide behavior, post-action display re-entry | User-visible pane behavior can regress even if state ownership stays unchanged. |
| `Visibility-sensitive` | `TaskPaneHost.IsVisible`, visible pane early-complete, display coordinator visibility checks | Any timing drift here can cause duplicate refresh or hidden pane failures. |
| `WorkbookOpen-sensitive` | host availability timing observed from `WorkbookOpen` downstream flows | `WorkbookOpen` itself is not the owner, but downstream window-safe timing is sensitive to host existence. |
| `VSTO lifecycle-sensitive` | `ThisAddIn.CreateTaskPane(...)`, `RemoveTaskPane(...)`, `TaskPaneHost.Dispose()` | These are the real COM/VSTO create-remove boundaries. |
| `Do not touch yet` | ready-show retry, protection-ready-show, foreground recovery, `_hostsByWindowKey` owner move, event unbinding behavior changes | These combine VSTO timing with existing UX stabilization logic and should stay frozen during the next boundary cut. |

### Next safe unit

- Keep the next implementation phase narrower than "VSTO boundary redesign".
- The safest next unit is an owner-preserving inventory-backed cleanup that isolates only the create/remove adapter story:
  - document `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)` as the VSTO adapter boundary,
  - document `TaskPaneHost` as the wrapper that owns the concrete `CustomTaskPane` instance lifetime,
  - document `TaskPaneHostFactory` as the control creation and event binding owner,
  - document `TaskPaneHostRegistry` as the replace/register/remove orchestration owner.
- Do not combine create/remove cleanup with event unbinding changes, metadata timing changes, ready-show retry changes, or foreground recovery changes in the same implementation step.

## B2.1 Status: VSTO Adapter Boundary Clarification (2026-05-06)

- Code comments now mirror the current owner map at the actual code seams:
  - `ThisAddIn` marks concrete VSTO create/remove.
  - `TaskPaneHost` marks concrete pane lifetime ownership.
  - `TaskPaneHostFactory` marks control creation and `ActionInvoked` binding ownership.
  - `TaskPaneHostRegistry` marks replace/register/remove orchestration over the shared map.
- Event unbinding is still not given an explicit owner. The code now treats that as a documented current-state ambiguity, not as something silently "fixed".
- This clarification phase does not move `_hostsByWindowKey`, change metadata timing, or alter ready-show / retry / visibility / foreground behavior.

## B2.2 Status: Host Factory Compose Owner Shift (2026-05-06)

- `TaskPaneHostRegistry` constructor no longer composes `TaskPaneHostFactory`.
- `TaskPaneHostFactory` compose owner moved outward to `TaskPaneManagerRuntimeGraphFactory.Compose(...)`, so the registry now receives a composed factory and stays focused on replace/register/remove orchestration.
- `TaskPaneHostFactory` still owns control creation and `ActionInvoked` binding behavior exactly as before.
- `TaskPaneHost`, `ThisAddIn.CreateTaskPane(...)`, `RemoveTaskPane(...)`, `_hostsByWindowKey`, metadata timing, ready-show / retry / visibility / foreground, and `WorkbookOpen` downstream behavior remain unchanged.

## B2.3 Status: Registry Logging and Metadata Mini Inventory (2026-05-06)

### Constructor surface

| Dependency | Current owner | Why registry sees it | Risk |
| --- | --- | --- | --- |
| `Dictionary<string, TaskPaneHost>` | `TaskPaneManager` | registry orchestrates replace/register/remove against the shared host map | `Runtime-sensitive` |
| `Logger` | registry log emission timing is owned by `TaskPaneHostRegistry` | registry logs registration and remove-host timing itself | `Safe docs/comment only` |
| `Func<TaskPaneHost, string> formatHostDescriptor` | descriptor source owner is `TaskPaneManager.FormatHostDescriptor(...)` -> `TaskPaneManagerDiagnosticHelper` | registry needs a host descriptor only for remove-host trace output | `Metadata-timing-sensitive` |
| `TaskPaneHostFactory` | compose owner is `TaskPaneManagerRuntimeGraphFactory.Compose(...)` | registry needs a composed factory to create a replacement host after decision-making | `Runtime-sensitive` |

### Mini inventory

| Concern | Current owner | What registry does today | Risk |
| --- | --- | --- | --- |
| logging owner | `TaskPaneHostRegistry` for `TaskPane host registered...` and `action=remove-host`; `TaskPaneHostFactory` for `action=create-host` | registry owns log timing for register/remove, but not descriptor construction | `Safe docs/comment only` |
| descriptor source owner | `TaskPaneManager.FormatHostDescriptor(...)` -> `TaskPaneManagerDiagnosticHelper.FormatHostDescriptor(...)` | registry receives a formatter and does not build host identity strings itself | `Metadata-timing-sensitive` |
| metadata write timing owner | `TaskPaneManager.RenderHost(...)` for `WorkbookFullName`; `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)` and `TaskPaneActionDispatcher.RefreshCaseHostAfterAction(...)` for `LastRenderSignature`; `TaskPaneDisplayCoordinator` invalidates `LastRenderSignature` for forced refresh | registry does not write host metadata | `Metadata-timing-sensitive` |
| metadata read timing owner | `TaskPaneHostRegistry.CollectWindowKeysForWorkbook(...)` reads `WorkbookFullName`; `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` reads `WorkbookFullName` and `IsVisible`; `TaskPaneRenderStateEvaluator` reads `WorkbookFullName` and `LastRenderSignature` | registry only reads `WorkbookFullName` for workbook-scope remove selection | `Metadata-timing-sensitive` |
| host identity owner | upstream `SafeGetWindowKey(...)` / `TaskPaneHost.WindowKey` and control type on the host | registry keys by `windowKey` and treats control type compatibility as role identity | `Runtime-sensitive` |
| replace/remove decision owner | `TaskPaneHostRegistry` | `TryGetReusableHost(...)`, `IsHostCompatibleWithRole(...)`, `RemoveExistingHostForReplacement(...)`, `RemoveHost(...)`, `CollectWindowKeysForWorkbook(...)` | `Runtime-sensitive` |
| visible pane early-complete touchpoint | direct owner is `TaskPaneDisplayCoordinator` / `WorkbookTaskPaneReadyShowAttemptWorker` | registry is only an indirect prerequisite through host-map contents and `WorkbookFullName` retention | `Visibility-sensitive` |
| WorkbookOpen downstream touchpoint | direct owners are `TaskPaneRefreshOrchestrationService` / `TaskPaneRefreshCoordinator` / `TaskPaneHostFlowService` | registry is reached indirectly via `TaskPaneHostLifecycleService.GetOrReplaceHost(...)` after downstream refresh dispatch, and via lifecycle cleanup through `RemoveWorkbookPanes(...)` | `WorkbookOpen-sensitive` |
| retry / foreground touchpoint | direct owners are ready-show worker, display coordinator, and foreground recovery services | registry has no direct retry/foreground logic; it is reached only if downstream flows create/remove hosts or remove a failed host | `Do not touch yet` |

### Current-state notes

- `TaskPaneHostRegistry` does not read `LastRenderSignature` or `IsVisible`.
- `TaskPaneHostRegistry` does read `WorkbookFullName`, so it is already adjacent to metadata timing even though it does not own metadata writes.
- `FormatHostDescriptor(...)` includes `windowKey`, pane role, `WorkbookFullName`, and window descriptor. That makes descriptor logs diagnostic-only in intent, but metadata-adjacent in content.
- The remove-host trace is emitted before the host is removed from the map, so it still sees the current host descriptor state at removal time.

### Risk classification

| Category | Scope | Why |
| --- | --- | --- |
| `Safe docs/comment only` | constructor-surface inventory, logging owner map, metadata touchpoint map | Fact-finding only. |
| `Safe naming clarification` | clarifying that `formatHostDescriptor` is a diagnostic dependency, not an identity owner | Safe only if behavior and logging timing stay unchanged. |
| `Runtime-sensitive` | `windowKey` reuse logic, control-type compatibility, replace/remove orchestration | These directly affect which host instance survives. |
| `Metadata-timing-sensitive` | `WorkbookFullName` reads in registry, descriptor content, `LastRenderSignature` adjacency in neighboring services | These couple diagnostics and workbook-scope remove behavior to upstream write timing. |
| `Visibility-sensitive` | indirect dependency on visible-pane early-complete through map contents and host metadata | Changing registry timing can alter the preconditions observed by visibility logic. |
| `WorkbookOpen-sensitive` | indirect host creation/removal after downstream refresh dispatch | `WorkbookOpen` itself does not call registry directly, but downstream refresh host availability depends on it. |
| `Do not touch yet` | logging timing, metadata write timing, visible-pane early-complete prerequisites, retry/foreground-adjacent host removal timing | These sit too close to current UX stabilization logic. |

### Next safe unit

- The next safe unit is still not a behavior change.
- If we clarify code further, the narrowest safe step is to document or lightly name the registry's descriptor formatter dependency as a diagnostic-only input without changing:
  - when logs are emitted,
  - what metadata is written,
  - how replace/remove is decided,
  - or how visible pane early-complete sees the host map.

## B2.4 Status: Diagnostic-only Descriptor Dependency Clarification (2026-05-06)

- `TaskPaneHostRegistry` now names its formatter dependency as a diagnostic-only input instead of a generic host-descriptor dependency.
- The clarified intent is:
  - registry consumes descriptor output only for logging,
  - registry does not become a host identity owner through that dependency,
  - registry does not become a metadata timing owner through that dependency,
  - replace/remove behavior does not consult descriptor output.
- Descriptor content, descriptor creation timing, registration/remove log timing, metadata read/write timing, and all runtime behavior remain unchanged.

## B2.5 Status: Registry Registration/Remove Logging Inventory (2026-05-06)

### Ownership and timing map

| Concern | Current owner | Current-state timing / touchpoint | Risk |
| --- | --- | --- | --- |
| registration logging timing | `TaskPaneHostRegistry` | emitted after `TaskPaneHostFactory.CreateHost(...)` returns and after `_hostsByWindowKey.Add(windowKey, host)` completes | `Runtime-sensitive` |
| remove logging timing | `TaskPaneHostRegistry` | emitted in `RemoveHost(...)` before `_hostsByWindowKey.Remove(windowKey)` and before host dispose completes | `Runtime-sensitive` |
| create-host logging timing | `TaskPaneHostFactory` | emitted during concrete host creation path, outside registry-owned register/remove logging | `Safe docs/comment only` |
| descriptor input source | `TaskPaneManager.FormatHostDescriptor(...)` -> `TaskPaneManagerDiagnosticHelper` | registry consumes upstream formatter output only at remove-host log time | `Diagnostic-only` |
| descriptor consume timing | `TaskPaneHostRegistry.LogHostRemoval(...)` | formatter reads the host while it is still in the shared map and before teardown nulls the VSTO pane | `Metadata-timing-sensitive` |
| metadata read timing | `TaskPaneHostRegistry.CollectWindowKeysForWorkbook(...)` | registry reads `WorkbookFullName` only for workbook-scope remove selection | `Metadata consumer` |
| replace/remove decision relation | `TaskPaneHostRegistry` | reuse / compatibility / replacement decisions are based on `windowKey`, control type, and workbook-scope selection, not descriptor output | `Orchestration-only` |
| visible pane early-complete indirect dependency | `TaskPaneDisplayCoordinator` / `WorkbookTaskPaneReadyShowAttemptWorker` | early-complete observes host-map contents plus `WorkbookFullName` and `IsVisible`; registry logging shares the same host state at removal time | `Visibility-sensitive` |
| WorkbookOpen downstream indirect dependency | refresh/lifecycle coordinators downstream of `WorkbookOpen` | registry register/remove timing affects when downstream flows can observe a host, even though `WorkbookOpen` is not the direct caller | `WorkbookOpen-sensitive` |
| retry / foreground indirect dependency | ready-show worker / display coordinator / foreground recovery services | registry has no direct retry or foreground ownership, but host removal timing can still be observed downstream | `Do not touch yet` |

### Current-state clarifications

- `TaskPaneHostRegistry` is a diagnostic consumer, not a descriptor source owner.
- `TaskPaneHostRegistry` is a metadata consumer for `WorkbookFullName` only; it is not a metadata owner and does not write `WorkbookFullName`, `LastRenderSignature`, or `IsVisible`.
- registration logging is orchestration-adjacent because it happens after the host has been inserted into the shared map.
- remove logging is metadata-timing-adjacent because it formats the host before removal from the shared map and before dispose completes.
- descriptor output is not consulted for reuse, compatibility, replacement, or workbook-scope remove decisions.

### Risk classification

| Category | Scope | Why |
| --- | --- | --- |
| `Diagnostic-only` | formatter dependency and remove-host descriptor output | The dependency exists only to describe host state for logs. |
| `Orchestration-only` | register/remove log emission timing inside registry | Registry owns when these logs are emitted, alongside register/remove orchestration. |
| `Metadata consumer` | `WorkbookFullName` read for workbook-scope remove selection | Registry consumes metadata but does not own write timing. |
| `Metadata owner ではない` | `WorkbookFullName`, `LastRenderSignature`, `IsVisible` writes | Those writes are owned upstream and downstream of registry. |
| `Runtime-sensitive` | registration/remove timing around map insertion/removal | Moving log timing casually could hide or reorder state transitions around host existence. |
| `Metadata-timing-sensitive` | remove-host descriptor formatting and workbook-scope metadata read | Both are adjacent to the timing of `WorkbookFullName` retention and teardown. |
| `Visibility-sensitive` | indirect dependency with visible pane early-complete | Shared host-map contents and metadata are observed by visibility logic. |
| `WorkbookOpen-sensitive` | indirect dependency through downstream host availability | Downstream refresh/lifecycle flows rely on stable host register/remove timing. |
| `Do not touch yet` | retry / foreground-adjacent removal timing and any log cleanup that reorders register/remove | These sit too close to current UX stabilization logic. |

### Next safe unit

- The next safe unit remains clarification-only unless a later phase explicitly accepts runtime-sensitive work.
- If code comments are added later, keep them limited to:
  - registration/remove logging as registry-owned orchestration timing,
  - descriptor formatting as diagnostic-only input,
  - metadata reads as consumer-only touchpoints.
- Do not combine logging cleanup with metadata timing changes, host identity changes, visible pane early-complete changes, or `WorkbookOpen` downstream behavior changes.

## B2.6 Status: Create/Remove Adapter Timing Inventory (2026-05-06)

### Scope

- This is a docs-only current-state inventory for create/remove adapter timing before any runtime surgery.
- The target boundary is limited to `ThisAddIn`, `TaskPaneHost`, `TaskPaneHostRegistry`, `TaskPaneHostFactory`, `TaskPaneManager`, and runtime-graph-adjacent lifecycle/display flows.
- This section does not change create/remove timing, metadata timing, visibility retention, `WorkbookOpen` flow, ready-show / retry, foreground recovery, visible pane early-complete, event unbinding order, or `_hostsByWindowKey` ownership.

### Create timing inventory

| Concern | Current owner | Current-state timing / sequence | Risk |
| --- | --- | --- | --- |
| display entry to refresh | `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)` -> `TaskPaneRefreshOrchestrationService` -> `TaskPaneRefreshCoordinator` | when `PaneDisplayPolicy` does not resolve to `ShowExisting` or `Hide`, the request re-enters `RefreshTaskPane(...)` and reaches `TaskPaneManager.RefreshPane(...)` | `WorkbookOpen-sensitive` |
| refresh-time create decision | `TaskPaneHostFlowService.RefreshPane(...)` | `RemoveStaleKernelHostsForRefresh(...)` runs first, then `TaskPaneHostLifecycleService.GetOrReplaceHost(...)` decides reuse vs replacement for the target `windowKey` | `Runtime-sensitive` |
| reusable host short-circuit | `TaskPaneHostRegistry.TryGetReusableHost(...)` | if the existing host on the same `windowKey` is compatible with the requested role, no new host and no new `CustomTaskPane` are created | `Runtime-sensitive` |
| replacement-before-create path | `TaskPaneHostRegistry.RemoveExistingHostForReplacement(...)` | if the existing host is incompatible, `host.Dispose()` runs before the old map entry is removed, and only after that does the flow continue to new host creation | `Runtime-sensitive` |
| concrete control create + binding | `TaskPaneHostFactory.CreateHost(...)` | the factory creates the role-specific control and wires `ActionInvoked`; for Case the host is constructed before the inline event subscription, while Kernel / Accounting bind before `TaskPaneHost` construction | `Runtime-sensitive` |
| concrete `CustomTaskPane` create | `TaskPaneHost` constructor -> `ThisAddIn.CreateTaskPane(...)` | host construction crosses the VSTO boundary immediately; `ThisAddIn.CreateTaskPane(...)` calls `CustomTaskPanes.Add(control, TaskPaneTitle, window)` and sets the dock position | `VSTO lifecycle-sensitive` |
| registry registration timing | `TaskPaneHostRegistry.CreateAndRegisterHost(...)` | the new host is inserted into `_hostsByWindowKey` only after `TaskPaneHostFactory.CreateHost(...)` returns; registration logging happens after `_hostsByWindowKey.Add(...)` | `Runtime-sensitive` |
| first visibility after create | `TaskPaneDisplayCoordinator.TryShowHost(...)` | pane creation and registry registration happen before `host.Show()`; visibility becomes true only later on the refresh/display path | `Visibility-sensitive` |
| first metadata write after create | `TaskPaneManager.RenderHost(...)` and `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)` | `WorkbookFullName` is written during render, and `LastRenderSignature` is written after render; concrete pane creation itself does not write those fields | `Metadata-timing-sensitive` |

### Remove timing inventory

| Remove trigger | Current owner | Current-state timing / sequence | Risk |
| --- | --- | --- | --- |
| incompatible replacement | `TaskPaneHostRegistry.RemoveExistingHostForReplacement(...)` | replacement dispose runs first, then the old `windowKey` entry is removed, and only then can the new host be created | `Runtime-sensitive` |
| workbook close cleanup | `WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` -> `TaskPaneManager.RemoveWorkbookPanes(...)` -> `TaskPaneHostLifecycleService.RemoveWorkbookPanes(...)` -> `TaskPaneHostRegistry.RemoveWorkbookPanes(...)` | registry selects `windowKey` values by `WorkbookFullName`, then removes each host | `WorkbookOpen-sensitive` |
| stale Kernel host cleanup during refresh | `TaskPaneHostLifecycleService.RemoveStaleKernelHostsForRefresh(...)` | before create/reuse for the active target window, stale Kernel hosts for the same workbook are removed through registry-owned remove calls | `Runtime-sensitive` |
| explicit remove-by-window | `TaskPaneHostLifecycleService.RemoveHost(...)` -> `TaskPaneHostRegistry.RemoveHost(...)` | the normal remove path logs while the host is still present, removes the map entry, then disposes the host | `Runtime-sensitive` |
| show/hide failure fallback | `TaskPaneDisplayCoordinator.TryShowHost(...)` / `SafeHideHost(...)` | if `host.Show()` or `host.Hide()` throws, the coordinator removes the host through the lifecycle callback instead of trying a partial recovery in place | `visibility/foreground-sensitive` |
| shutdown cleanup | `ThisAddIn_Shutdown(...)` -> `TaskPaneManager.DisposeAll(...)` -> `TaskPaneHostLifecycleService.DisposeAll(...)` -> `TaskPaneHostRegistry.DisposeAll(...)` | shutdown snapshots all current hosts, disposes them, then clears the shared map | `VSTO lifecycle-sensitive` |

### Current-state teardown ownership

- Shared registry state:
  - `TaskPaneHostRegistry.RemoveHost(...)` logs first, removes the `windowKey` entry second, and disposes the host last.
  - `TaskPaneHostRegistry.RemoveExistingHostForReplacement(...)` disposes first and removes the old `windowKey` entry second.
  - `TaskPaneHostRegistry.DisposeAll(...)` snapshots hosts, disposes the snapshot, and clears `_hostsByWindowKey` after disposal completes.
- Concrete pane lifetime:
  - `TaskPaneHost.Dispose()` is still the concrete remove wrapper for a live host.
  - The current remove order inside `TaskPaneHost.Dispose()` is `Hide()` -> `ThisAddIn.RemoveTaskPane(_pane)` -> `_pane = null`.
  - `ThisAddIn.RemoveTaskPane(...)` remains the only concrete VSTO `CustomTaskPanes.Remove(...)` caller.
- Control and event binding lifetime:
  - `TaskPaneHostFactory` owns control creation and inline `ActionInvoked` subscription timing.
  - An explicit event-unbinding owner is still not present in current state.
  - Current-state teardown therefore remains implicit and dispose-driven; this inventory does not infer a separate event-unbinding phase.

### Adapter boundary inventory

| Boundary | Current owner | Current-state contract | Not owner of | Risk |
| --- | --- | --- | --- | --- |
| `ThisAddIn` | VSTO adapter boundary | owns concrete `CustomTaskPane` create/remove entrypoints and display-request entry surface | shared host map, control binding ownership, host metadata timing | `VSTO lifecycle-sensitive` |
| `TaskPaneHost` | concrete pane lifetime holder | holds `window`, `control`, `view`, and concrete pane instance; crosses into `ThisAddIn.CreateTaskPane(...)` at construction and `RemoveTaskPane(...)` at dispose | shared host map, create/reuse decision, metadata write timing | `Runtime-sensitive` |
| `TaskPaneHostFactory` | control creation + `ActionInvoked` binding owner | builds role-specific controls, wires action delegates, then returns a constructed host | shared host map, workbook-scope remove selection, visibility policy | `Runtime-sensitive` |
| `TaskPaneHostRegistry` | replace/register/remove orchestration owner over the shared map | decides reuse vs replacement, performs registration, and drives concrete host teardown by calling `TaskPaneHost.Dispose()` | `_hostsByWindowKey` ownership, control binding ownership, metadata writes | `Runtime-sensitive` |
| `TaskPaneManagerRuntimeGraphFactory.Compose(...)` | runtime graph compose owner | wires factory, registry, lifecycle, display coordinator, and callbacks around the shared map | runtime state ownership, create/remove timing ownership after composition | `Runtime-sensitive` |
| `TaskPaneManager` | facade + `_hostsByWindowKey` owner | owns the shared host map, render seam, and downstream facade entrypoints | concrete VSTO `CustomTaskPane` API calls, inline control binding ownership | `Metadata-timing-sensitive` |

### Runtime-sensitive danger boundaries

| Touchpoint | Why it is dangerous | Current-state dependency |
| --- | --- | --- |
| `WorkbookOpen` downstream host availability | `WorkbookOpen` itself is workbook-only, so moving create/remove timing earlier or later can violate the documented `WorkbookActivate` / `WindowActivate` window-safe boundary | create/reuse/remove must stay observable only in downstream refresh/lifecycle flows |
| ready-show early-complete | `WorkbookTaskPaneReadyShowAttemptWorker` can short-circuit if `HasVisibleCasePaneForWorkbookWindow(...)` sees a visible Case host for the workbook window | host-map contents plus `WorkbookFullName` and `IsVisible` must stay stable at the observed timing |
| visibility / hide vs remove | `PaneDisplayPolicy` and `TaskPaneDisplayCoordinator` distinguish hiding an existing host from replacing or removing it | a created host is not yet a visible host, and hide failure can escalate to remove |
| foreground recovery adjacency | show/hide failure paths already remove hosts through lifecycle callbacks | changing remove timing here can leak into foreground-recovery behavior even without touching that code directly |
| metadata timing | `WorkbookFullName` and `LastRenderSignature` are written after create and are read by render-state, workbook-scope remove, and visible-pane checks | create/remove cleanup must not silently move metadata read/write timing |
| event unbinding ambiguity | current teardown does not expose a separate unbind phase | making unbinding explicit would change teardown order and is therefore outside this inventory |

### GO conditions for the next code phase

- GO only if the task isolates one runtime-sensitive boundary and keeps the other frozen.
- GO only if the owner map stays intact:
  - `TaskPaneManager` keeps `_hostsByWindowKey` ownership,
  - `TaskPaneHostRegistry` keeps replace/register/remove orchestration ownership,
  - `TaskPaneHostFactory` keeps control creation and `ActionInvoked` binding ownership,
  - `TaskPaneHost` keeps concrete pane lifetime ownership,
  - `ThisAddIn` remains the concrete VSTO adapter boundary unless that single boundary is the explicit subject of the runtime-surgery task.
- GO only if the planned diff can explain the before/after timing for:
  - incompatible replacement remove,
  - standard remove-by-window,
  - shutdown dispose,
  - ready-show visible-pane observation,
  - and `WorkbookOpen` downstream host availability.
- GO only if validation distinguishes compile/build success from runtime `Addins\` reflection and human-side smoke.

### STOP conditions for the next code phase

- STOP if the change starts to touch metadata timing, visibility retention, `WorkbookOpen` flow, ready-show / retry, foreground recovery, visible pane early-complete, or event unbinding order.
- STOP if the change requires moving `_hostsByWindowKey` ownership, redesigning `TaskPaneManager` / `TaskPaneHostRegistry` / `TaskPaneHost` / `TaskPaneHostFactory`, adding services, or introducing abstraction-first refactoring.
- STOP if the change needs to move create timing and remove timing together instead of isolating one boundary.
- STOP if the change cannot state which layer owns `CustomTaskPane`, control binding, registry state, and metadata timing both before and after the diff.

## B2 Checkpoint Status (2026-05-06)

### Final owner map at the end of B2

| Component | Current responsibility at checkpoint | Notes |
| --- | --- | --- |
| `TaskPaneManagerRuntimeBootstrap` | runtime entry + attach-order owner | runtime graph knowledge is localized here |
| `TaskPaneManagerRuntimeGraphFactory` | passive graph builder + compose owner | does not own runtime state or orchestration |
| `TaskPaneManager` | facade/orchestration boundary + `_hostsByWindowKey` owner | state owner intentionally not moved |
| `TaskPaneHostRegistry` | replace/register/remove orchestration + registration/remove logging timing | consumes diagnostics and metadata, but does not own them |
| `TaskPaneHostFactory` | control creation + `ActionInvoked` binding + create-host logging timing | role unchanged through B2 |
| `TaskPaneHost` | concrete `CustomTaskPane` lifetime holder | current remove path stays `Hide -> RemoveTaskPane -> teardown` |
| `ThisAddIn` | VSTO adapter boundary | concrete create/remove entrypoint remains here |

### Runtime-sensitive and metadata-sensitive boundary

| Boundary | Why it is sensitive | B2 treatment |
| --- | --- | --- |
| host-map existence timing | observed by reuse, remove, refresh, and visible-pane checks | frozen |
| metadata timing | `WorkbookFullName`, `LastRenderSignature`, `IsVisible` participate in downstream decisions | frozen |
| remove timing | logging, map removal, and dispose are tightly adjacent | frozen |
| visible pane early-complete | depends on map contents, `WorkbookFullName`, and `IsVisible` | frozen |
| `WorkbookOpen` downstream host availability | host creation/removal is observed after downstream refresh/lifecycle dispatch | frozen |
| ready-show / retry / foreground | downstream UX stabilization depends on stable host state and visibility semantics | frozen |
| event unbinding behavior | still implicit in dispose-driven teardown | frozen as debt |

### Do not touch yet

- `_hostsByWindowKey` ownership move
- `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)` timing cleanup
- `TaskPaneHost.Dispose()` behavior cleanup
- metadata timing cleanup
- registration/remove logging reordering
- visible pane early-complete cleanup
- `WorkbookOpen` / `WindowActivate` downstream cleanup
- ready-show / retry / foreground cleanup
- event unbinding redesign

### Runtime surgery warning

- B2 ended at the point where owner maps are explicit but runtime-sensitive timings are still shared across multiple layers.
- The next phase should be treated as runtime surgery because even "small" cleanups can change:
  - which host exists at a given point,
  - which metadata is still readable at log / remove time,
  - whether visible-pane early-complete short-circuits,
  - whether downstream `WorkbookOpen` refresh sees a host.
- Any later implementation phase should name one sensitive boundary up front and keep the others frozen.

## B2.7 Status: Metadata Timing Inventory (2026-05-07)

### Scope

- This is a docs-only current-state inventory for metadata timing before any runtime surgery.
- The target boundary is limited to `TaskPaneHost.WorkbookFullName`, `TaskPaneHost.LastRenderSignature`, and the adjacent visible-pane observation inputs that are read with that metadata: `windowKey` and `TaskPaneHost.IsVisible`.
- This section does not change metadata timing, create/remove timing, `WorkbookOpen` downstream behavior, ready-show / retry, visible pane early-complete, visibility retention, foreground recovery, event unbinding behavior, or `_hostsByWindowKey` ownership.

### Metadata and adjacent runtime inputs

| State / input | Current owner | Current write owner | Current readers | Notes |
| --- | --- | --- | --- | --- |
| `TaskPaneHost.WorkbookFullName` | `TaskPaneHost` | `TaskPaneManager.RenderHost(...)` | render-state checks, visible-pane check, workbook-scope remove selection, stale Kernel cleanup, action target resolution | current-state workbook identity join key for host-side decisions |
| `TaskPaneHost.LastRenderSignature` | `TaskPaneHost` | `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)`, `TaskPaneDisplayCoordinator.InvalidateHostRenderStateForForcedRefresh(...)`, `TaskPaneActionDispatcher` fallback post-action rerender | render-state checks, display-request render-current checks, CASE host reuse without render | current-state render cache key for the host |
| `TaskPaneHost.IsVisible` | concrete pane state below `TaskPaneHost` | no separate metadata write; computed from `_pane.Visible` at read time | `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` | not host metadata, but part of the same early-complete precondition |
| `windowKey` | `_hostsByWindowKey` / `TaskPaneManager` | window-bound host registration path | visible-pane lookup, display-request lookup, remove/reuse lookup | not metadata, but the first lookup gate before metadata can be read |

### Metadata write inventory

| State | Current write owner | Current-state timing | Downstream dependency | Risk |
| --- | --- | --- | --- | --- |
| `WorkbookFullName` initial write | `TaskPaneManager.RenderHost(...)` | written at the start of render, before role-specific `RenderKernelHost(...)`, `RenderAccountingHost(...)`, or `RenderCaseHost(...)` | makes the host joinable to the current workbook before later visible-pane, remove-selection, and action-target reads | `Metadata-timing-sensitive` |
| `LastRenderSignature` refresh-time write | `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)` | `TaskPaneRenderStateEvaluator.EvaluateRenderState(...)` runs first; if render is required, host render runs and the signature is written afterward | render-current checks and CASE host reuse depend on this write already having happened | `Metadata-timing-sensitive` |
| `LastRenderSignature` forced-refresh invalidation | `TaskPaneDisplayCoordinator.InvalidateHostRenderStateForForcedRefresh(...)` | clears the signature before forced refresh / target-window rerender is attempted | forces later render-state evaluation to observe the host as stale | `Metadata-timing-sensitive` |
| `LastRenderSignature` fallback post-action rewrite | `TaskPaneActionDispatcher` CASE post-action fallback path | when the fallback rerenders a CASE host locally, it rebuilds the signature after render and before `TryShowHost(...)` completes | keeps post-action refresh behavior aligned with later render-current checks even without a full add-in refresh round-trip | `Metadata-timing-sensitive` |
| host metadata teardown | no field-level clear; host removal drops the whole host from `_hostsByWindowKey` | remove / replacement / failure cleanup stop future readers by removing the host, not by clearing `WorkbookFullName` or `LastRenderSignature` first | later consumers either still see the old host or do not see a host at all; there is no intermediate metadata-only cleanup phase | `Runtime-sensitive` |

### Metadata read / consumer inventory

| Consumer | What it reads | Current-state expectation | Why it matters |
| --- | --- | --- | --- |
| `TaskPaneRenderStateEvaluator.EvaluateDisplayRequestPaneState(...)` | `windowKey`, `WorkbookFullName`, `LastRenderSignature` | display-request show/reuse decisions assume the host for that window is already bound to the same workbook and that a non-empty signature means "render-current" is testable | drives `PaneDisplayPolicy` decisions for show-existing vs rerender |
| `TaskPaneRenderStateEvaluator.EvaluateRenderState(...)` | `LastRenderSignature` | refresh-time render skip assumes signature equality means the host is already current for the resolved `WorkbookContext` | direct render/no-render decision |
| `TaskPaneHostFlowService.ShouldReuseCaseHostWithoutRender(...)` | `WorkbookFullName`, `LastRenderSignature` | CASE host reuse without render requires same workbook, non-empty signature, Case control type, and a reuse-allowed reason | moving either write changes reuse semantics |
| `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` | `windowKey`, `WorkbookFullName`, `IsVisible`, hosted role | visible Case pane exists only when lookup succeeds, workbook matches, host role is Case, and the concrete pane is visible | direct early-complete predicate |
| `WorkbookTaskPaneReadyShowAttemptWorker` | consumes `HasVisibleCasePaneForWorkbookWindow(...)` result | ready-show attempt can short-circuit successfully without refresh only if the visible-pane predicate is already true | this is the runtime-sensitive bridge from metadata to ready-show behavior |
| `TaskPaneHostRegistry.RemoveWorkbookPanes(...)` | `WorkbookFullName` | workbook-close cleanup expects the host metadata already identifies which registered windows belong to that workbook | workbook-scope remove selection is metadata-driven |
| `TaskPaneHostLifecycleService.RemoveStaleKernelHostsForRefresh(...)` | `WorkbookFullName` | stale Kernel cleanup expects the active `WorkbookContext` and stale host metadata to join on workbook identity before removal | refresh-time cleanup is metadata-driven |
| `TaskPaneCaseActionTargetResolver` and `TaskPaneNonCaseActionHandler` | `WorkbookFullName` | action dispatch assumes `FindOpenWorkbook(host.WorkbookFullName)` can resolve the workbook tied to the host | action routing depends on metadata already being populated |

### Render-signature dependency inventory

- `TaskPaneRenderStateEvaluator.BuildRenderSignature(...)` currently derives the signature from:
  - `WorkbookContext.Role`
  - `WorkbookContext.WorkbookFullName`
  - `WorkbookContext.ActiveSheetCodeName`
  - CASE-only document property `CASELIST_REGISTERED`
  - CASE-only document property `TASKPANE_SNAPSHOT_CACHE_COUNT`
- That means `LastRenderSignature` is not just a render marker. It is the cache key for workbook identity, active sheet identity, and selected CASE workbook document-property state.
- `WorkbookFullName` is therefore both:
  - an input into signature generation,
  - and a separate host identity key for visible-pane checks, remove selection, stale cleanup, and action-target resolution.

### `WorkbookOpen` downstream expectation

- `WorkbookLifecycleCoordinator.OnWorkbookOpen(...)` currently does workbook-side lifecycle/setup work only:
  - external workbook detection,
  - Kernel / Accounting / Case workbook lifecycle handlers,
  - accounting sheet control ensure,
  - Kernel home availability notification.
- `OnWorkbookOpen(...)` does not directly call `_refreshTaskPane(...)`.
- `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` keeps `WorkbookOpen` with `workbook != null` and `window == null` on the skip/defer side of the boundary.
- Current-state implication:
  - pure `WorkbookOpen` does not guarantee that host metadata already exists,
  - the first reliable metadata population point is still downstream render on a window-safe path such as `WorkbookActivate`, `WindowActivate`, explicit display refresh, ready-show refresh, or pending retry fallback.

### Ready-show / retry / foreground coupling

| Flow | Metadata dependency | Current-state behavior |
| --- | --- | --- |
| ready-show attempt | direct | `WorkbookTaskPaneReadyShowAttemptWorker` attempt 1 can early-complete only after workbook-window resolve succeeds and `HasVisibleCasePaneForWorkbookWindow(...)` sees `windowKey` hit + `WorkbookFullName` match + Case role + `IsVisible == true`. |
| ready retry `80ms` | direct but read-only | retry scheduling does not write metadata; it re-observes the same visible-pane predicate and otherwise falls back to refresh. |
| pending retry `400ms` | indirect | `PendingPaneRefreshRetryService` tracks workbook full name separately from host metadata, but it depends on refresh-time metadata semantics staying stable once refresh succeeds and the host becomes observable again. |
| final foreground recovery | indirect | `TaskPaneRefreshCoordinator.GuaranteeFinalForegroundAfterRefresh(...)` runs only after refresh success. It does not read host metadata directly, but metadata-sensitive early-complete and refresh-success boundaries determine whether foreground recovery runs at all. |

### Stale cleanup / remove / replacement coupling

| Flow | Metadata dependency | Current-state coupling |
| --- | --- | --- |
| stale Kernel cleanup before refresh | direct `WorkbookFullName` read | `TaskPaneHostLifecycleService.RemoveStaleKernelHostsForRefresh(...)` removes other Kernel hosts for the same workbook before the target window host is reused or replaced. |
| workbook-scope remove | direct `WorkbookFullName` read | `TaskPaneHostRegistry.RemoveWorkbookPanes(...)` first selects target `windowKey` values by workbook metadata, then delegates to normal remove-by-window behavior. |
| standard remove-by-window | no new metadata read after target selection | once `RemoveHost(windowKey)` starts, the shared-map entry is removed and later metadata readers no longer see that host through the map. |
| replacement remove | indirect observability coupling | incompatible replacement does not use metadata to decide compatibility, but it disposes the old host before removing the old map entry, so metadata observability remains tied to this frozen ordering. |
| show/hide failure fallback remove | indirect observability coupling | `TaskPaneDisplayCoordinator` failure recovery removes the host instead of clearing metadata in place, so later visible-pane and render-state readers switch from "old metadata" to "no host" rather than to a partially cleared host. |

### Runtime-sensitive coupling summary

- metadata timing is coupled to visible-pane timing because `WorkbookFullName` and `IsVisible` are read together under the same `windowKey` lookup.
- metadata timing is coupled to render-cache semantics because `LastRenderSignature` is both a refresh skip key and a display-request render-current key.
- metadata timing is coupled to workbook cleanup because workbook-scope remove and stale Kernel cleanup both depend on `WorkbookFullName` already being written on the host they are about to remove.
- metadata timing is coupled to action dispatch because both Case and non-Case action target resolution locate the workbook through `FindOpenWorkbook(host.WorkbookFullName)`.
- metadata timing is coupled to `WorkbookOpen` downstream availability because `WorkbookOpen` does not populate host metadata itself; later window-safe flows are the first consumers that can rely on the metadata existing.

### GO conditions for a later metadata-timing surgery phase

- GO only if the task is scoped to metadata timing itself and keeps create/remove timing, visibility retention, `WorkbookOpen` flow, ready-show / retry, foreground recovery, visible pane early-complete, event unbinding behavior, and `_hostsByWindowKey` ownership frozen.
- GO only if the planned diff can explain the before/after read timing for:
  - `EvaluateDisplayRequestPaneState(...)`
  - `EvaluateRenderState(...)`
  - `ShouldReuseCaseHostWithoutRender(...)`
  - `HasVisibleCasePaneForWorkbookWindow(...)`
  - workbook-scope remove selection
  - stale Kernel cleanup
  - action target resolution
- GO only if the planned diff can explain the before/after write timing for:
  - `TaskPaneManager.RenderHost(...)`
  - `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)`
  - forced-refresh invalidation
  - post-action fallback signature rewrite
- GO only if validation keeps compile/build confirmation separate from runtime `Addins\` reflection and human-side smoke.

### STOP conditions for a later metadata-timing surgery phase

- STOP if the change starts to move create timing, remove timing, `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` boundaries, ready-show retry timing, pending retry timing, foreground recovery timing, or visible pane early-complete semantics in the same diff.
- STOP if the change requires moving `_hostsByWindowKey`, redesigning `TaskPaneManager` / `TaskPaneHostFlowService` / `TaskPaneDisplayCoordinator` / `TaskPaneHostRegistry`, adding services, or introducing abstraction-first refactoring.
- STOP if the change cannot preserve and explain how `WorkbookFullName` remains observable for workbook-scope remove, stale Kernel cleanup, and action dispatch.
- STOP if the change cannot preserve and explain how `LastRenderSignature` remains coherent for render skip, display-request render-current checks, and CASE host reuse.

## B2.8 Status: Ready-Show / Visibility / Foreground Dependency Inventory (2026-05-07)

### Scope

- This is a docs-only current-state inventory before any runtime surgery.
- The target boundary is limited to:
  - ready-show entry / ready retry / pending retry fallback,
  - visibility retention over existing hosts,
  - final foreground recovery and CASE protection start,
  - visible-pane early-complete preconditions,
  - the connection back to `WorkbookOpen` downstream host availability.
- This section does not change ready-show / retry, visibility retention, foreground recovery, visible pane early-complete, `WorkbookOpen` flow, metadata timing, create/remove timing, event unbinding behavior, or `_hostsByWindowKey` ownership.

### Ready-show entry inventory

| Entry | Current owner | Dependency state before the request | Current-state note |
| --- | --- | --- | --- |
| created CASE ready-show | `KernelCasePresentationService.ExecuteDeferredPresentationEnhancements(...)` | transient suppression released, workbook non-null, workbook window visibility ensure attempted, one-shot CASE activation suppression prepared | current order is `ReleaseWorkbook(...) -> EnsureWorkbookWindowVisibleBeforeReadyShow(...) -> SuppressUpcomingCasePaneActivationRefresh(...) -> ShowWorkbookTaskPaneWhenReady(...)` |
| accounting-set ready-show | `AccountingSetCreateService.Execute(...)` | transient suppression released, workbook non-null, invoice-entry activation already requested | current path does not introduce a separate TaskPane-only visibility policy before the ready-show request |
| TaskPane ready-show entry | `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)` | workbook non-null | entry is workbook-targeted, not window-targeted; it passes scheduling/fallback delegates into `WorkbookTaskPaneReadyShowAttemptWorker` |

### Ready-show / retry execution inventory

| Step | Current owner | Dependency state | Current-state behavior |
| --- | --- | --- | --- |
| attempt orchestration | `TaskPaneDisplayRetryCoordinator.ShowWhenReady(...)` | workbook, reason, `maxAttempts=2` | attempt 1 runs immediately, attempt 2 is the only scheduled retry, and later fallback starts only after attempts are exhausted |
| attempt 1 pre-visibility ensure | `WorkbookTaskPaneReadyShowAttemptWorker.EnsureWorkbookWindowVisibleForTaskPaneDisplay(...)` | `attemptNumber == 1`, workbook non-null | `WorkbookWindowVisibilityService.EnsureVisible(...)` is called only on attempt 1 |
| workbook-window resolve | `WorkbookTaskPaneReadyShowAttemptWorker` -> `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` | workbook non-null; either workbook visible window exists, or active workbook matches and active window exists | ready-show uses the same resolver boundary as refresh, but asks it to activate the workbook first |
| ready retry scheduling | `TaskPaneRefreshOrchestrationService.ScheduleTaskPaneReadyRetry(...)` | retry action non-null | schedules a WinForms timer with `WorkbookPaneWindowResolveDelayMs = 80` |
| fallback prepare | `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` | workbook full name trackable, `WorkbookOpen` skip policy not triggered | fallback records workbook target first, tries one immediate refresh, and only then starts pending retry |
| pending retry | `PendingPaneRefreshRetryService` | `PendingPaneRefreshIntervalMs = 400`, `PendingPaneRefreshMaxAttempts = 3` | workbook-target retry resolves the workbook again by full name; if the workbook is gone, active CASE context fallback can still call `TryRefreshTaskPane(reason, null, null)` |

### Visibility retention dependency inventory

- current-state visibility retention is not backed by a standalone `VisibilityRetentionState`.
- It is the combined behavior of:
  - keeping the host registered in `_hostsByWindowKey`,
  - keeping host-side workbook/render metadata on that same host instance,
  - toggling concrete pane visibility through `TaskPaneHost.Show()` / `Hide()` instead of removing the host when reuse is still intended.

| State / input | Current write owner | Current readers / consumers | Current-state meaning |
| --- | --- | --- | --- |
| host registration in `_hostsByWindowKey` | `TaskPaneHostRegistry` register/remove/replace/dispose-all paths | `PaneDisplayPolicy`, `TaskPaneDisplayCoordinator`, `TaskPaneHostFlowService`, `TaskPaneHostLifecycleService`, visible-pane early-complete | retention starts with host registration and ends with registry removal, not with metadata clearing |
| concrete pane visible bit (`_pane.Visible`) | `TaskPaneHost.Show()` / `Hide()` / `Dispose()` via `TaskPaneDisplayCoordinator` or outer teardown | `TaskPaneHost.IsVisible` -> `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` | visibility is live VSTO state, not persisted metadata |
| `TaskPaneHost.WorkbookFullName` | `TaskPaneManager.RenderHost(...)` | `TryShowExistingPane(...)`, visible-pane check, workbook-scope remove selection, stale Kernel cleanup, action target resolution | workbook identity join key for retained-host reuse and cleanup |
| `TaskPaneHost.LastRenderSignature` | `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)`, forced-refresh invalidation, CASE post-action fallback rerender | display-request render-current checks, refresh-time render skip, CASE host reuse without render | render-current marker for retained hosts |
| `KernelCaseInteractionState.IsKernelCaseCreationFlowActive` | outside this inventory | `TaskPaneHostPreparationPolicy` / `TaskPaneDisplayCoordinator.PrepareHostsBeforeShow(...)` | when true, CASE hosts keep CASE visibility while hiding non-CASE hosts; non-CASE hosts hide all except the active window |
| workbook role / hosted control type | resolved at request time and on host instance | `PaneDisplayPolicy.ShouldDisplayPane(...)`, `GetHostedWorkbookRole(...)`, CASE-only early-complete | visibility retention is role-sensitive and CASE-only shortcuts stay CASE-only |

### Visibility retention consumers

| Consumer | What it depends on | Current-state expectation |
| --- | --- | --- |
| `PaneDisplayPolicy` + `TaskPaneDisplayCoordinator.TryShowExistingPaneForDisplayRequest(...)` | accepted request, non-null window, resolvable `windowKey`, managed host existence, same workbook, render signature current | explicit display requests prefer showing an already-retained host over rerendering |
| `TaskPaneHostFlowService.ShouldReuseCaseHostWithoutRender(...)` | CASE role, Case host control type, non-empty `LastRenderSignature`, same `WorkbookFullName`, reason in `WorkbookActivate` / `WindowActivate` / `KernelHomeForm.FormClosed` | refresh-time CASE reuse is narrower than general host retention and is reason-sensitive |
| `TaskPaneDisplayCoordinator.PrepareHostsBeforeShow(...)` | kernel-case-creation flow flag, active host role | retained hosts can still be hidden before another host is shown; retention and visibility are separate concerns |
| show/hide failure fallback | `TaskPaneDisplayCoordinator.TryShowHost(...)` / `SafeHideHost(...)` exception path | a visibility failure escalates to `_removeHost(windowKey)`, so retention is intentionally lost rather than left partially inconsistent |

### Foreground recovery and protection inventory

| Step | Current owner | Required state | Current-state behavior |
| --- | --- | --- | --- |
| post-refresh foreground trigger | `TaskPaneRefreshCoordinator.TryRefreshTaskPane(...)` | `refreshed == true`, `window != null`, `_excelWindowRecoveryService != null` | `GuaranteeFinalForegroundAfterRefresh(...)` runs only after a successful refresh with a resolved pane window |
| final foreground recovery | `TaskPaneRefreshCoordinator.GuaranteeFinalForegroundAfterRefresh(...)` | resolved `WorkbookContext` and/or target workbook | calls `TryRecoverWorkbookWindow(..., bringToFront: true)` for the target workbook, or `TryRecoverActiveWorkbookWindow(..., bringToFront: true)` when workbook input is null |
| CASE protection start | `TaskPaneRefreshCoordinator` -> `ICasePaneHostBridge.BeginCaseWorkbookActivateProtection(...)` -> `KernelHomeCasePaneSuppressionCoordinator.BeginCaseWorkbookActivateProtection(...)` | context role is `Case`, protected workbook non-null, protected window non-null, workbook role re-resolves to `Case`, workbook full name non-empty, window hwnd non-empty | writes protected workbook/window target and `SuppressionDuration = 5 seconds` |
| `WorkbookActivate` protection ignore | `WorkbookLifecycleCoordinator` -> `ShouldIgnoreWorkbookActivateDuringCaseProtection(...)` | protected workbook full name matches input workbook, active window hwnd matches protected window hwnd | ignore is narrower than general workbook match because active window must also match |
| `WindowActivate` protection ignore | `WindowActivatePaneHandlingService` -> `ShouldIgnoreWindowActivateDuringCaseProtection(...)` | protected workbook full name matches input workbook, event window hwnd matches protected window hwnd | window-event path uses event window identity directly |
| `TaskPaneRefresh` protection ignore | `TaskPaneRefreshOrchestrationService.RefreshPreconditionEvaluator` -> `ShouldIgnoreTaskPaneRefreshDuringCaseProtection(...)` | protection active and current active window hwnd matches protected window hwnd | refresh-side ignore is keyed by active window, not by input workbook/window equality |

### Visible-pane early-complete conditions

| Condition | Current source | Why it matters |
| --- | --- | --- |
| workbook is non-null | `WorkbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(...)` | no ready-show path exists without a workbook target |
| resolved pane window exists | `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` | early-complete does not run until workbook-window resolution succeeds |
| `windowKey` resolves and the host is still registered | `_safeGetWindowKey(window)` + `_hostsByWindowKey.TryGetValue(...)` | visibility retention must still expose the host through the shared map |
| target workbook full name matches `host.WorkbookFullName` | `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` | protects against reusing a visible pane that belongs to another workbook |
| hosted role is `Case` | `GetHostedWorkbookRole(host)` | current shortcut is CASE-only |
| `host.IsVisible == true` | `TaskPaneHost.IsVisible` | early-complete is a visible-pane shortcut, not a hidden-host shortcut |

### Visible-pane early-complete danger points

- early-complete is looser than display-request show-existing.
- `HasVisibleCasePaneForWorkbookWindow(...)` does not read `LastRenderSignature`, so the shortcut does not prove that the host is render-current.
- therefore early-complete also does not compare:
  - active sheet identity,
  - `CASELIST_REGISTERED`,
  - `TASKPANE_SNAPSHOT_CACHE_COUNT`,
  because those inputs exist only inside the render signature.
- when early-complete succeeds, the worker returns success before `_tryRefreshTaskPane(...)` runs.
- that means the current attempt does not perform:
  - refresh-time rerender,
  - `LastRenderSignature` rewrite,
  - `GuaranteeFinalForegroundAfterRefresh(...)`,
  - a new `BeginCaseWorkbookActivateProtection(...)`.
- if show/hide failure has already escalated to `_removeHost(windowKey)`, early-complete observes `NoHost` rather than a partially cleared host.
- this shortcut is current-state CASE-only and must not be widened to accounting by adjacency.

### `WorkbookOpen` downstream host availability connection

| Boundary | Current owner | Current-state implication |
| --- | --- | --- |
| pure `WorkbookOpen` event | `WorkbookLifecycleCoordinator.OnWorkbookOpen(...)` | runs workbook-side lifecycle/setup only and does not directly call `_refreshTaskPane(...)` |
| shared `WorkbookOpen` skip/defer boundary | `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` | keeps `WorkbookOpen` with `workbook != null` and `window == null` on the non-window-safe side of the boundary |
| first window-safe refresh path | `WorkbookLifecycleCoordinator.OnWorkbookActivate(...)` | `WorkbookActivate` is the first built-in path here that can call `_refreshTaskPane("WorkbookActivate", workbook, null)` |
| first display-entry path | `WindowActivatePaneHandlingService.Handle(...)` / post-action callers | `RequestTaskPaneDisplayForTargetWindow(...)` decides show-existing / hide / render only after a window is already available |
| first create/reuse host path | `TaskPaneHostFlowService.RefreshPane(...)` | host creation or reuse does not happen until context is accepted, `windowKey` is non-empty, and the host lifecycle path is entered |
| concrete create boundary | `TaskPaneHostLifecycleService` -> `TaskPaneHostRegistry` -> `TaskPaneHostFactory` -> `TaskPaneHost` -> `ThisAddIn.CreateTaskPane(...)` | host availability is downstream of refresh/display flow, not a direct `WorkbookOpen` side effect |

- Current-state reading:
  - pure `WorkbookOpen` does not guarantee host existence,
  - pure `WorkbookOpen` does not guarantee visible-pane observability,
  - pure `WorkbookOpen` does not guarantee foreground recovery eligibility,
  - ready-show / display-request / retry paths are all downstream consumers of the later window-safe boundary.

### GO conditions for a later ready-show / visibility / foreground phase

- GO only if the task isolates this boundary and keeps metadata timing, create/remove timing, `WorkbookOpen` flow, event unbinding behavior, and `_hostsByWindowKey` ownership frozen.
- GO only if the planned diff can explain the before/after behavior for:
  - ready-show entry order in `KernelCasePresentationService` and `AccountingSetCreateService`,
  - `80ms` ready retry and `2` ready-show attempts,
  - `400ms` pending retry, `3` attempts, and active CASE fallback,
  - visible-pane early-complete success and failure paths,
  - final foreground recovery trigger,
  - protection start plus the 3 ignore readers,
  - show/hide failure fallback remove.
- GO only if the planned diff can preserve and explain the current distinction between:
  - retained host existence,
  - live pane visibility,
  - render-current state.
- GO only if validation keeps compile/build confirmation separate from runtime `Addins\` reflection and human-side smoke.

### STOP conditions for a later ready-show / visibility / foreground phase

- STOP if the change starts to move metadata timing, create timing, remove timing, `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` boundaries, or event unbinding order in the same diff.
- STOP if the change requires moving `_hostsByWindowKey`, redesigning `TaskPaneManager` / `TaskPaneDisplayCoordinator` / `TaskPaneHostFlowService` / `TaskPaneHostRegistry` / `ThisAddIn`, adding services, or introducing abstraction-first refactoring.
- STOP if the change cannot state whether visible-pane early-complete continues to bypass:
  - render-current checks,
  - refresh-time rerender,
  - final foreground recovery,
  - protection start.
- STOP if the change cannot preserve and explain the current `KernelCasePresentationService` handoff order:
  - transient suppression release,
  - workbook-window visibility ensure,
  - CASE activation suppression preparation,
  - ready-show request.

## B2.9 Status: Event Unbinding Behavior Inventory (2026-05-07)

### Scope

- This is a docs-only current-state inventory before any runtime surgery.
- The target boundary is limited to:
  - `ActionInvoked` binding timing and owner,
  - explicit unbinding owner presence/absence,
  - dispose-driven teardown convergence,
  - retained/reused host event-state retention,
  - workbook-close / stale cleanup / replacement interaction with event lifetime,
  - ready-show / retry / foreground paths that can keep or observe an already-bound host.
- This section does not change event binding order, event unbinding behavior, create/remove timing, metadata timing, `WorkbookOpen` downstream flow, ready-show / retry, visibility retention, foreground recovery, visible pane early-complete, or `_hostsByWindowKey` ownership.

### Binding inventory

| Binding point | Current owner | Current-state timing / sequence | Current consumer |
| --- | --- | --- | --- |
| compose-time delegate supply | `TaskPaneManagerRuntimeGraphFactory.CreateTaskPaneHostFactory(...)` | compose-time only; it passes delegates into the factory but does not subscribe to control events itself | `TaskPaneNonCaseActionHandler` and `TaskPaneActionDispatcher` callback surfaces |
| Kernel `ActionInvoked` bind | `TaskPaneHostFactory.CreateKernelHost(...)` | `KernelNavigationControl.ActionInvoked += ...` runs before `new TaskPaneHost(...)` | `TaskPaneNonCaseActionHandler.HandleKernelActionInvoked(...)` through a `windowKey`-capturing lambda |
| Accounting `ActionInvoked` bind | `TaskPaneHostFactory.CreateAccountingHost(...)` | `AccountingNavigationControl.ActionInvoked += ...` runs before `new TaskPaneHost(...)` | `TaskPaneNonCaseActionHandler.HandleAccountingActionInvoked(...)` through a `windowKey`-capturing lambda |
| CASE `ActionInvoked` bind | `TaskPaneHostFactory.CreateCaseHost(...)` | `new TaskPaneHost(...)` runs first, then `DocumentButtonsControl.ActionInvoked += ...` runs | `TaskPaneActionDispatcher.HandleCaseControlActionInvoked(...)` through a `windowKey`- and `caseControl`-capturing lambda |
| retained host reuse / show-existing | no new owner | compatible host reuse, display-request show-existing, and CASE no-render reuse do not recreate the control and do not rebind the event | existing inline delegate remains on the retained control instance |

### Explicit unbinding inventory

| Surface | Explicit unbinding owner in current state? | Evidence | Current-state meaning |
| --- | --- | --- | --- |
| factory-installed TaskPane control handlers | no | repo code shows no `ActionInvoked -= ...` callsite outside `DocumentButtonsControl`'s event accessor | binding is one-way at create time and is not explicitly reversed by the TaskPane runtime |
| `TaskPaneHost.Dispose()` | no | current code is `Hide()` -> `ThisAddIn.RemoveTaskPane(_pane)` -> `_pane = null`; there is no `ActionInvoked -= ...` and no explicit `control.Dispose()` call | host teardown depends on outer host disposal plus lower-layer pane/control teardown |
| `ThisAddIn.RemoveTaskPane(...)` | no | current code only calls `CustomTaskPanes.Remove(pane)` | VSTO adapter removal is not the explicit `ActionInvoked` unbind owner |
| Excel Application event wiring | yes, but separate boundary | `ThisAddIn.UnhookApplicationEvents()` -> `ApplicationEventSubscriptionService.Unsubscribe()` | explicit Excel event unwiring exists, but it is not TaskPane control-event unwiring |

- `DocumentButtonsControl` exposes `add/remove` forwarding accessors for `ActionInvoked`, but the TaskPane runtime does not call the `remove` side for the factory-installed handler.
- `KernelNavigationControl`, `AccountingNavigationControl`, `DocumentButtonsControl`, and `DocTaskPaneControl` do not add a repo-local custom dispose/unbind phase for `ActionInvoked`.

### Teardown convergence inventory

| Trigger | Current teardown path | Event-lifetime implication |
| --- | --- | --- |
| incompatible replacement | `TaskPaneHostRegistry.GetOrReplaceHost(...)` -> `DisposeThenUnregisterHostForReplacement(...)` -> `TaskPaneHost.Dispose()` -> old map entry removal -> new host creation/bind | old inline delegate stays on the old control until dispose-driven teardown; replacement gets a fresh control and a fresh inline bind |
| standard remove-by-window | `TaskPaneHostRegistry.RemoveHost(...)` -> log -> shared-map removal -> `TaskPaneHost.Dispose()` | the host becomes unobservable through `_hostsByWindowKey` before dispose completes; there is no separate unbind phase between removal and dispose |
| workbook close cleanup | `WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` -> `TaskPaneManager.RemoveWorkbookPanes(...)` -> `TaskPaneHostLifecycleService.RemoveWorkbookPanes(...)` -> `TaskPaneHostRegistry.RemoveWorkbookPanes(...)` -> per-window `RemoveHost(...)` | event lifetime ends only for hosts selected by `WorkbookFullName`; workbook close is an explicit Excel event boundary that converges onto the same dispose-driven host teardown |
| stale Kernel cleanup | `TaskPaneHostLifecycleService.RemoveStaleKernelHostsForRefresh(...)` -> `TaskPaneHostRegistry.RemoveHost(...)` | stale host event state is dropped only through normal registry remove/dispose flow |
| show/hide failure fallback | `TaskPaneDisplayCoordinator.TryShowHost(...)` / `SafeHideHost(...)` -> lifecycle remove callback -> registry remove/dispose | visibility failure escalates to full host teardown rather than to detached event cleanup |
| shutdown cleanup | `ThisAddIn_Shutdown(...)` -> `TaskPaneManager.DisposeAll()` -> `TaskPaneHostLifecycleService.DisposeAll()` -> `TaskPaneHostRegistry.DisposeAll()` | retained hosts are snapshotted, disposed, and only then removed from the live shared map by the final clear step |

### Retained / reused host event-state inventory

| Flow | Host recreated? | Binding recreated? | Current-state meaning |
| --- | --- | --- | --- |
| compatible same-window reuse in `GetOrReplaceHost(...)` | no | no | the same control instance and its original `windowKey`-capturing delegate remain attached |
| CASE no-render reuse in `TaskPaneHostFlowService.TryReuseCaseHostForRefresh(...)` | no | no | host show is retried on the retained instance without a new bind step |
| display-request show-existing in `TaskPaneDisplayCoordinator.TryShowExistingPaneForDisplayRequest(...)` | no | no | explicit display requests can keep the original binding alive without re-entering create flow |
| CASE post-action fallback rerender in `TaskPaneActionDispatcher.RefreshCaseHostAfterAction(...)` | no | no | local rerender/signature rewrite happens on the same control instance, so event state stays as-is |
| incompatible replacement | yes | yes | replacement is the only current-state path here that guarantees a fresh control and a fresh inline binding |

### Implicit teardown dependency and unknowns

- Repo-local current state stops the explicit teardown chain at `TaskPaneHost.Dispose()` -> `ThisAddIn.RemoveTaskPane(...)` -> `CustomTaskPanes.Remove(pane)`.
- No repo-local callsite explicitly disposes the TaskPane control object, and no repo-local control class in this boundary adds a custom `Dispose` override for `ActionInvoked` cleanup.
- Therefore this inventory fixes only the facts that:
  - explicit TaskPane-side unbinding ownership is absent,
  - teardown is dispose-driven from the host side,
  - the exact lower-layer moment when VSTO/WinForms releases the control and delegate graph is not asserted from repository code here.

### `WorkbookOpen` / close boundary relation

- `WorkbookLifecycleCoordinator.OnWorkbookOpen(...)` currently does workbook-side lifecycle/setup only and does not create a TaskPane host or bind `ActionInvoked` by itself.
- The first TaskPane control-event bind opportunity remains downstream of the window-safe refresh/display path:
  - `RequestTaskPaneDisplayForTargetWindow(...)`
  - refresh orchestration
  - `TaskPaneHostRegistry.GetOrReplaceHost(...)`
  - `TaskPaneHostFactory.CreateHost(...)`
- The inverse boundary is explicit workbook close:
  - `WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` does not unbind TaskPane control events directly,
  - it selects hosts for teardown through workbook metadata and converges onto registry remove/dispose instead.

### Ready-show / retry / foreground coupling

| Flow | Direct `ActionInvoked` involvement | Current-state implicit dependency |
| --- | --- | --- |
| ready-show early-complete | none | `WorkbookTaskPaneReadyShowAttemptWorker` can complete on a retained visible CASE host, which bypasses `GetOrReplaceHost(...)` / `CreateHost(...)` and therefore bypasses any recreate/rebind opportunity on that path |
| ready retry `80ms` | none | retry scheduling re-observes the same retained host state until refresh or fallback is needed; it does not create a new binding by itself |
| pending retry `400ms` | none | only a later successful refresh re-enters host lifecycle and can create/rebind a host; pending retry itself does not unbind or rebind |
| final foreground recovery / CASE protection start | none | both remain downstream of refresh success in `TaskPaneRefreshCoordinator`; if ready-show early-complete succeeds on a retained host, those downstream steps and any recreate/rebind opportunity are skipped together |

### Runtime-sensitive coupling summary

- binding timing is asymmetric by role:
  - Kernel / Accounting bind before `TaskPaneHost` construction,
  - CASE binds after `TaskPaneHost` construction.
- event lifetime is coupled to host lifetime, not to an explicit unbind phase.
- retained host reuse is coupled to current binding state because reuse/show-existing paths do not recreate the control.
- workbook-close cleanup and stale Kernel cleanup are coupled to metadata timing because they choose which already-bound hosts to tear down by `WorkbookFullName`.
- ready-show early-complete is coupled to retained event state because it can choose "keep the existing visible host" before any recreate/rebind path runs.
- lower-layer control/event release timing below `CustomTaskPanes.Remove(...)` remains unasserted from repository code and must stay in the "unknown" bucket.

### GO conditions for a later event-unbinding phase

- GO only if the task is scoped to event unbinding behavior itself and keeps create/remove timing, metadata timing, `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` boundaries, ready-show / retry, visibility retention, foreground recovery, visible pane early-complete, and `_hostsByWindowKey` ownership frozen.
- GO only if the planned diff can explain the before/after behavior for:
  - the 3 factory bind points,
  - incompatible replacement teardown,
  - standard remove-by-window teardown,
  - workbook-close cleanup,
  - stale Kernel cleanup,
  - shutdown cleanup,
  - retained host reuse / show-existing,
  - ready-show early-complete bypass.
- GO only if validation keeps compile/build confirmation separate from runtime `Addins\` reflection and human-side smoke.

### STOP conditions for a later event-unbinding phase

- STOP if the change starts to move create timing, remove timing, metadata timing, `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` boundaries, ready-show / retry behavior, visibility retention, foreground recovery, or visible pane early-complete semantics in the same diff.
- STOP if the change requires moving `_hostsByWindowKey`, redesigning `TaskPaneManager` / `TaskPaneHostRegistry` / `TaskPaneHostFactory` / `TaskPaneHost` / `ThisAddIn`, adding services, or introducing abstraction-first refactoring.
- STOP if the change assumes a lower-layer control disposal guarantee below `CustomTaskPanes.Remove(...)` without proving it separately.
- STOP if the change cannot state whether retained-host reuse still preserves the original binding or starts to force recreate/rebind on paths that are currently reuse-only.
