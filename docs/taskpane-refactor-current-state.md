# TaskPane Refactor Current State

## 位置づけ

この文書は、TaskPane 側の優先度Aリファクタについて、現行 `main` で確認できる到達点を固定するための現在地文書です。

- TaskPane 設計正本: `docs/taskpane-architecture.md`
- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- A4 / C2 checkpoint: `docs/a4-c2-current-state.md`
- Startup / TaskPane 初期表示の実機チェック: `docs/thisaddin-startup-test-checklist.md`
- 優先度A棚卸し: `docs/a-priority-service-responsibility-inventory.md`
- TaskPane refresh policy 正本: `docs/taskpane-refresh-policy.md`
- protection / ready-show 危険領域の補足:
  - `docs/taskpane-protection-ready-show-investigation.md`
  - `docs/taskpane-protection-baseline.md`
  - `docs/taskpane-protection-observation-checklist.md`

この文書は設計正本を置き換えるものではありません。TaskPane 優先度Aで「どこまで main に固定済みか」「どこが helper 分離・bridge 化まで完了し、どこが未確定・実機未確認として残るか」を明示するための補助文書です。B1 / B2-1 / A4 / C1 / C2 をまたぐ広い checkpoint は `docs/a4-c2-current-state.md` に分けて記録します。

## 今回固定する到達点

現行 `main` に対して、TaskPane 側の優先度A到達点は次の整理で固定します。

1. TaskPane の runtime 設計正本は `docs/taskpane-architecture.md` とする。
2. 文書ボタン定義の正本、Base 埋込 snapshot、CASE cache、prompt / resolver の責務分離は、`docs/taskpane-architecture.md` の記述を現行到達点として扱う。
3. 優先度Aのうち、production code 変更なしで完了確認できた棚卸し結果は `docs/a-priority-service-responsibility-inventory.md` を基準に読む。
4. protection / ready-show / retry / suppression を含む危険領域は、policy 正本化と helper 分離までは完了済みとして扱う。
5. ただし、数値根拠、dead route 判定、実機 UX、visible pane early-complete の単純化可否は未確定のまま残し、コードだけでは断定しない。

## 完了済みとして固定する事項

### 1. TaskPane 設計正本の固定

- TaskPane の正本は Kernel `雛形一覧` と Kernel `TASKPANE_MASTER_VERSION` である。
- Base 埋込 snapshot と CASE snapshot cache は、いずれも派生 cache であり正本ではない。
- TaskPane 表示の解決順は `CASE cache -> Base cache -> Master rebuild` である。
- 開いている CASE は、後から成功した雛形登録・更新へ自動追随しない。
- `DocumentNamePromptService` は CASE cache だけを参照し、master fallback しない。
- `DocumentTemplateResolver` は CASE cache 優先で解決し、miss 時のみ master fallback する。

### 1-1. Master access / snapshot read path の到達点

- `MasterWorkbookReadAccessService` が、Master workbook path 解決、read-only open、所有 workbook の close、window 非表示化の共有境界です。
- `MasterTemplateCatalogService` と `TaskPaneSnapshotBuilderService` は、どちらも `MasterWorkbookReadAccessService.ResolveMasterPath(...)` と `OpenReadOnly(...)` を使う構成に揃っています。
- `MasterWorkbookReadAccessResult.CloseIfOwned()` により、既に開いていた workbook と自前で開いた workbook の close 責務が分離されています。
- `Master workbook read access` は shared access service へ集約済みであり、個別サービス側に open / close / hidden window 副作用を戻しません。

### 1-2. TaskPaneManager リファクタの到達点

- `TaskPaneManager` は、もはや TaskPane 側の全責務を抱える単一巨大クラスではありません。
- 現在の `TaskPaneManager` は、主に host 管理、role 別 render 切替、render/show orchestration、host 再利用調停の中心です。
- 次の主責務は分離済みとして固定します。
  - 表示・非表示: `TaskPaneDisplayCoordinator`
  - refresh 入口調停: `TaskPaneRefreshOrchestrationService` / `TaskPaneRefreshCoordinator`
  - WindowActivate 境界処理: `WindowActivatePaneHandlingService`
  - snapshot / view state: `CasePaneSnapshotRenderService`、`CaseTaskPaneViewStateBuilder`、`TaskPaneSnapshotParser`
  - doc prompt / business action: `TaskPaneBusinessActionLauncher`
  - render 後副作用: `CasePaneCacheRefreshNotificationService`
  - CASE pane UIイベント dispatch: `TaskPaneActionDispatcher`
- refresh-time orchestration は `TaskPaneHostFlowService` へ外出し済みです。
- 軽量 helper / policy として、`TaskPaneManagerDiagnosticHelper`、`TaskPaneHostReusePolicy`、`TaskPaneRenderStateEvaluator`、`TaskPaneShowExistingPolicy`、`TaskPaneShowWithRenderPolicy` が main に反映済みです。
- `TaskPaneHostRegistry` は外出し済みで、host 生成、差し替え、破棄、workbook 単位 cleanup の内部整理が main に反映済みです。
- `TaskPaneManager` には render seam と facade entry に必要な collaborator だけが attach され、`TaskPaneHostRegistry` を含む registration / handler compose は `TaskPaneManagerRuntimeGraphFactory` 側へ残す current-state に揃いました。

### 1-3. TaskPane refresh orchestration の到達点

- `TaskPaneRefreshOrchestrationService` は、いまは refresh 挙動を全部抱え込むよりも、順序調停に寄った役割として読めます。
- `RefreshPreconditionEvaluator` により precondition 判定が整理済みです。
- `RefreshDispatchShell` により coordinator 呼び出し shell が整理済みです。
- `PendingPaneRefreshRetryState` により pending retry state が整理済みです。
- `WorkbookPaneWindowResolver` により window resolver が整理済みです。
- `TaskPaneRefreshPreconditionPolicy` は `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` の shared skip policy 正本です。

### 2. TaskPane 周辺で完了済みとして扱う bridge / 境界整理

`docs/a-priority-service-responsibility-inventory.md` を基準に、現行 `main` で完了済みとして扱うのは次です。

- `DocumentCommandService`
  - `ScreenUpdating`、TaskPane refresh suppression、active refresh、Kernel sheet refresh は bridge 経由へ整理済み。
- `WindowActivatePaneHandlingService`
  - `ShouldIgnoreWindowActivateDuringCaseProtection(...)` 判定は bridge 経由へ整理済み。
- `KernelCasePresentationService` / `TaskPaneRefreshOrchestrationService` / `TaskPaneRefreshCoordinator` / `WorkbookLifecycleCoordinator`
  - suppression、ready-show、protection、visible pane 判定の case-pane 系 `ThisAddIn` 依存は `ICasePaneHostBridge` 経由へ整理済み。
- `TaskPaneRefreshCoordinator`
  - `KernelFlickerTrace` の structured trace は維持され、`04150a7` で obsolete route に付随していた duplicate plain log が削除済み。
- 補助境界として確認済みの事項
  - `TaskPaneHost` は `Globals.ThisAddIn` ではなく constructor 注入の `ThisAddIn` を VSTO `CustomTaskPane` の生成・破棄境界として使う。
  - `TaskPaneHost` 自体は表示判断を持たない薄い host ラッパーとして扱う。

### 3. docs 側で固定済みの危険領域棚卸し

次の論点は、すでに docs 上で危険領域として棚卸し済みであることを到達点に含めます。

- ready-show / suppression の順序を壊してはいけないこと
- `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の protection 判定が連動していること
- retry `80ms`、fallback timer `400ms`、`3 attempts` はコード上の事実として確認できるが、仕様根拠は未確認であること
- visible pane early-complete が既存 CASE pane の不要な refresh 回避に使われること

## 未確定・実機未確認として残す事項

次は優先度Aに含まれるが、現時点では「helper 分離や bridge 化は main 済みでも、挙動変更や簡素化はまだ固定しない」領域です。

- `KernelCasePresentationService`
  - ready-show 要求前後の suppression / release / workbook window 可視化の順序を含む危険領域
- `TaskPaneRefreshOrchestrationService`
  - retry state / window resolver / precondition / dispatch shell は整理済みだが、retry route の削減、active fallback の必要条件、dead route 判定は未確定
- `TaskPaneRefreshCoordinator`
  - CASE refresh 完了後の foreground 保証と protection 開始は main に残る危険領域であり、数値根拠と最終 UX は未確認
- `WorkbookLifecycleCoordinator`
  - `WorkbookActivate` 再入抑止の判定境界
- `TaskPaneManager`
  - host 再利用、visible pane early-complete、VSTO 境界を含む pane 制御本体

## 今後課題として固定する事項

### 次タスク候補

- `ScheduleActiveTaskPaneRefresh` が production route か dead route 候補かを調査する
- active CASE context fallback の必要条件を整理する
- visible pane early-complete 条件の単純化可否を整理する
- protection 3入口判定差を現行タイミングを崩さず説明できる形へ寄せる
- `80ms` / `400ms` / `3 attempts` / `5秒` の仕様根拠を整理する

### `TaskPaneHostRegistry`

- `TaskPaneManager` 周辺に残る主要責務です。
- host 生成、差し替え、破棄、workbook 単位の掃除を担います。
- 独立クラス化済みだが、VSTO `TaskPaneHost` 生成境界と action event 配線に関わるため、引き続き分離リスクが高いです。
- 次に触る場合は `TaskPaneHostRegistry` だけを対象にし、action dispatch や refresh 本線には触れないほうが安全です。

### `ThisAddIn` 境界

- `ThisAddIn` は VSTO lifecycle、application event、custom task pane 生成、TaskPane 表示要求の入口です。
- application event wiring / unwiring は `ApplicationEventSubscriptionService` へ分離済みだが、handler 本体と lifecycle 呼び出し位置は `ThisAddIn` に残しています。
- Startup 周辺は呼び出し順を変えずに private helper で見通し整理するに留め、`HookApplicationEvents()`、`TryShowKernelHomeFormOnStartup()`、`RefreshTaskPane("Startup", null, null)` の位置は維持します。
- Startup 順序固定メモは `docs/thisaddin-boundary-inventory.md` を参照し、`InitializeApplicationEventSubscriptionService()` -> `HookApplicationEvents()` -> `TryShowKernelHomeFormOnStartup()` -> `RefreshTaskPane("Startup", null, null)` の並びを現行契約として維持します。
- `TaskPaneManager` / `TaskPaneHostRegistry` との依存境界を急に変えると起動、終了、pane 表示に波及しやすいです。
- `ThisAddIn` 整理は HostRegistry 分離よりさらに慎重に扱い、先に現状メモと依存関係棚卸しを行い、コード変更は後回しにする判断を固定します。
- 詳細な棚卸しは `docs/thisaddin-boundary-inventory.md` を参照します。

## 今回の到達点に含めない事項

次は現行 docs / code だけでは確定しないため、到達点として固定しません。

- retry 値や protection 5 秒の正式な仕様根拠
- Pane 再利用判定の全条件
- 実機でのちらつき、二重表示、出遅れの最終観測結果
- `WindowActivate` 固有の体感挙動の完全な期待仕様

## 次の実装着手時に守ること

- `docs/taskpane-architecture.md` を設計正本として維持する
- `WorkbookOpen` 直後に直接 UI 表示制御を追加しない
- snapshot / cache を保存・生成・実行判断の正本へ戻さない
- ready-show / suppression / protection の順序を変える変更は、危険領域として別途確認してから扱う
- host 再利用経路と visible pane early-complete を安易に単純化しない
- `TaskPaneHostRegistry` と `ThisAddIn` 境界の変更は、安定化後に小単位で扱う

## 一言まとめ

TaskPane 側の優先度Aは、設計正本・責務棚卸し・危険領域の事実整理に加え、Master access の一本化、`TaskPaneManager` の helper 分離、`TaskPaneHostRegistry` の外出し、`TaskPaneRefreshOrchestrationService` の順序調停化、refresh policy 正本化までは `main` に固定済みです。

一方で、ready-show / protection / retry / host 再利用を含む本線ロジックの簡素化、実機未確認事項の確定、`TaskPaneHostRegistry` / `ThisAddIn` の VSTO 境界整理は、まだ完了済みとは扱わず、安定化後に慎重に進める課題として残します。

## TaskPaneManager 周辺の最終棚卸し判断 (2026-05-08)

- 現行 `main` `6e367db0e6865a35d8bda422c151f4b9faa26689` を基準に再棚卸しした結果、compose-side / bootstrap-side で安全に削れる ownership unit は、B1.7 までで概ね出し切ったと扱います。

### A/B/C/D

- `A. まだ安全単位で削る価値がある`
  - なし。
- `B. 残っているが runtime-sensitive に近いため見送る`
  - `TaskPaneManager` の `_hostsByWindowKey` shared state owner。
  - `TaskPaneHostFlowService` の refresh-time host flow。
  - `TaskPaneActionDispatcher` の post-action display re-entry。
  - いずれも metadata timing、ready-show、visibility retention、foreground / protection、`WorkbookOpen -> WorkbookActivate -> WindowActivate` 境界と近接しているため、この棚卸しでは見送ります。
- `C. 既に docs の責務線と概ね一致している`
  - `TaskPaneManagerRuntimeBootstrap` / `TaskPaneManagerRuntimeGraphFactory` の compose / attach 境界。
  - `TaskPaneNonCaseActionHandler`、`TaskPaneRefreshPreconditionPolicy`、`TaskPaneHostReusePolicy`、`PaneDisplayPolicy`、`TaskPaneRenderStateEvaluator` の責務線。
- `D. モニター配布前に触るより、実運用で観察すべき`
  - `TaskPaneHostRegistry` / `TaskPaneHostFactory` / `TaskPaneHost` / `ThisAddIn` の VSTO create-remove と event lifetime。
  - `TaskPaneDisplayCoordinator` の visible-pane early-complete / show-hide failure remove。
  - これらは runtime-sensitive boundary に近いため、ここで削るより実運用観察を優先します。

- したがって、TaskPaneManager 周辺 refactor はこの時点で「一旦完了扱い」とします。
- 再開条件は、実運用観察で追加根拠が出るか、または runtime-sensitive unit を 1 boundary だけ単独で扱う必要が明確になった場合に限定します。
- この追記は docs-only です。known issue の白 Excel フラッシュ、ready-show / retry / foreground sequencing、VSTO create-remove timing、既存 docs の大枠は変更しません。
## B1 Update (2026-05-06)

- Production runtime composition owner for the TaskPaneManager constructor graph moved to `AddInTaskPaneCompositionFactory`.
- `TaskPaneManager` still owns `_hostsByWindowKey` and remains the facade / orchestration boundary.
- The moved graph includes `CasePaneCacheRefreshNotificationService`, `TaskPaneHostRegistry`, `TaskPaneHostLifecycleService`, `TaskPaneDisplayCoordinator`, `TaskPaneActionDispatcher`, and `TaskPaneHostFlowService`.
- `WorkbookOpen`, `WindowActivatePaneHandlingService`, ready-show retry, protection flow, and VSTO create/remove boundaries were not changed in this phase.
- Remaining composition inside `TaskPaneManager` is limited to test-only construction paths that call the same external runtime-graph factory after manager creation.

## B1.1 Update (2026-05-06)

- `TaskPaneManager` runtime graph entry is now fixed behind `TaskPaneManagerRuntimeBootstrap.CreateAttached(...)`.
- The old split entrypoints (`new TaskPaneManager(...)`, `TaskPaneManagerRuntimeGraphFactory.Compose(...)`, `AttachRuntimeGraph(...)`) no longer appear at production or test callsites.
- `TaskPaneManagerRuntimeGraphFactory` now receives a passive compose context instead of reading the manager as a dependency bag.
- `TaskPaneManager` still owns `_hostsByWindowKey` and the facade/orchestration surface; this phase did not move state ownership, `WorkbookOpen`, visibility, ready-show, retry, or VSTO lifecycle boundaries.

## B1.2 Update (2026-05-06)

- `TaskPaneManagerRuntimeBootstrap.CreateAttached(...)` is now the production runtime entrypoint, while test harnesses use explicit `CreateAttachedForTests(...)` or `CreateThinAttachedForTests(...)`.
- Raw `TaskPaneManager` constructors and graph attach are no longer visible as general internal entrypoints; bootstrap reaches them through manager-local bootstrap access only.
- The removed convenience constructors were unused after the bootstrap shift, so cleanup did not change runtime behavior or attach timing.
- `WorkbookOpen`, visibility, ready-show, retry, foreground recovery, VSTO lifecycle, and host metadata timing remain untouched in this phase.

## B1.3 Update (2026-05-06)

- The old bootstrap compose context is now split into `TaskPaneManagerRuntimeEntryContext` and a smaller `TaskPaneManagerRuntimeGraphComposeContext`.
- Entry context owns only bootstrap-time construction input, while graph compose context carries only the dependencies needed to build the runtime graph.
- As of 2026-05-08, `TaskPaneManagerRuntimeEntryContext` no longer carries `ICaseTaskPaneSnapshotReader`; the snapshot reader stays rooted in `CasePaneSnapshotRenderService` and is not treated as an attach-orchestration input.
- `TaskPaneManager.RuntimeBootstrapAccess` remains the sole bridge to private constructors and private attach, and is explicitly documented as bootstrap-only.

## B1.4 Update (2026-05-07)

- `TaskPaneManagerRuntimeBootstrap` now creates an explicit `TaskPaneManagerRuntimeGraphComposeSurface` for manager-owned compose-time input.
- `TaskPaneManagerRuntimeGraphFactory.Compose(...)` no longer reads `TaskPaneManager` directly as a compose-time dependency bag; it consumes the explicit host-map / formatter / render-seam surface instead.
- This phase does not move `_hostsByWindowKey` ownership, does not change manager attach payload, and does not alter `WorkbookOpen`, ready-show, retry, foreground recovery, or VSTO create/remove timing.

## B1.5 Update (2026-05-08)

- `TaskPaneManagerRuntimeGraphFactory` now narrows the create-side adapter helper payloads to explicit `TaskPaneHostFactoryComposeContext` and `TaskPaneHostRegistryComposeContext`.
- `TaskPaneHostFactory` helper compose no longer receives the full runtime graph compose context / surface when it only needs `ThisAddIn`, `Logger`, and the host descriptor formatter.
- `TaskPaneHostRegistry` helper compose no longer receives the full runtime graph compose context / surface when it only needs the shared host map, `Logger`, and the diagnostic formatter.
- This phase does not change `_hostsByWindowKey` ownership, `ActionInvoked` bind timing, `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)`, `WorkbookOpen`, ready-show, retry, foreground recovery, or visible-pane behavior.

## B1.6 Update (2026-05-08)

- `TaskPaneManagerRuntimeGraphFactory` now narrows the non-case action helper payload to an explicit `TaskPaneNonCaseActionHandlerComposeContext`.
- `TaskPaneNonCaseActionHandler` helper compose no longer reads the full runtime graph compose context / surface when it only needs the non-case action collaborators plus resolved host/render/show delegates.
- This phase does not change `TaskPaneActionDispatcher` compose, `_hostsByWindowKey` ownership, `ActionInvoked` bind timing, `WorkbookOpen`, ready-show, retry, foreground recovery, or VSTO create/remove timing.

## B1.7 Update (2026-05-08)

- `TaskPaneManagerRuntimeGraphFactory` now narrows the CASE dispatcher subtree compose to explicit `TaskPaneCaseActionTargetResolverComposeContext`, `TaskPaneCaseActionHandlerComposeContext`, and `TaskPaneActionDispatcherComposeContext`.
- CASE dispatcher compose no longer reads the full runtime graph compose context after it crosses into the CASE action subtree; target resolver, separated action handlers, and dispatcher each receive only the collaborators they consume.
- This phase does not change `TaskPaneActionDispatcher` runtime order, post-action refresh fallback behavior, `_hostsByWindowKey` ownership, `ActionInvoked` bind timing, `WorkbookOpen`, ready-show, retry, foreground recovery, or VSTO create/remove timing.

## B2 Prep Update: VSTO Boundary Inventory (2026-05-06)

- Current VSTO create/remove ownership is still split across four layers:
  - `TaskPaneManager` owns `_hostsByWindowKey` and remains the host existence state owner.
  - `TaskPaneHostRegistry` owns replace/register/remove orchestration over that shared map.
  - `TaskPaneHostFactory` owns control creation and `ActionInvoked` binding for Case / Kernel / Accounting hosts.
  - `TaskPaneHost` owns the concrete `CustomTaskPane` instance lifetime and reaches the actual VSTO boundary through `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)`.
- Event unbinding does not have an explicit owner today; handler lifetime is still coupled to control disposal and host teardown.
- Host metadata timing is also split intentionally in current-state:
  - `TaskPaneManager.RenderHost(...)` writes `host.WorkbookFullName` before role render.
  - `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)` writes `host.LastRenderSignature` after refresh-time render succeeds.
  - `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` and `TaskPaneRenderStateEvaluator` both consume that metadata.
- visible pane early-complete is therefore not just a display concern. It depends on `windowKey`, `host.WorkbookFullName`, and `host.IsVisible`, and is consumed by `WorkbookTaskPaneReadyShowAttemptWorker`.
- This phase does not move `_hostsByWindowKey`, does not redesign registry/factory/host ownership, and does not alter `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)`.

### Implications for the next implementation phase

- Do not cut create/remove, event binding, metadata timing, and ready-show / retry in one task.
- The smallest safe next unit is a VSTO adapter boundary clarification that preserves all current owners while making the create/remove call chain explicit in code comments and docs.
- Any future extraction that touches `TaskPaneHostRegistry` or `TaskPaneHostFactory` should keep `_hostsByWindowKey` where it is and avoid changing visible pane early-complete semantics.

## B2.1 Update: VSTO Adapter Boundary Clarification (2026-05-06)

- Code comments now align with the current owner map without changing behavior:
  - `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)` are clarified as the concrete VSTO adapter boundary.
  - `TaskPaneHost` is clarified as the concrete `CustomTaskPane` lifetime holder.
  - `TaskPaneHostFactory` is clarified as the control creation and `ActionInvoked` binding owner.
  - `TaskPaneHostRegistry` is clarified as the replace/register/remove orchestration owner over the shared host map.
- Event unbinding remains intentionally unchanged. Current-state still relies on dispose-driven teardown, and the ambiguity is documented as debt rather than redesigned here.
- `_hostsByWindowKey`, metadata timing, ready-show / retry, visibility, foreground recovery, and `WorkbookOpen` downstream behavior all remain untouched in this clarification phase.

## B2.2 Update: Host Factory Compose Owner Shift (2026-05-06)

- `TaskPaneHostRegistry` no longer creates `TaskPaneHostFactory` inside its constructor.
- `TaskPaneManagerRuntimeGraphFactory.Compose(...)` now composes `TaskPaneHostFactory` and passes it into the registry, which reduces one layer of secondary composition-root behavior without changing registry orchestration semantics.
- This phase does not change:
  - factory role ownership,
  - `ActionInvoked` binding behavior,
  - event unbinding ambiguity,
  - `TaskPaneHost` lifetime ownership,
  - `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)`,
  - `_hostsByWindowKey`,
  - metadata timing,
  - visible pane early-complete,
  - ready-show / retry / protection / foreground,
  - `WorkbookOpen` / `WindowActivate` downstream behavior.

## B2.3 Update: Registry Logging and Metadata Mini Inventory (2026-05-06)

- `TaskPaneHostRegistry` constructor surface is now small but still runtime-adjacent:
  - shared host map handle
  - logger
  - host descriptor formatter
  - composed `TaskPaneHostFactory`
- current-state owner split around logging and metadata is:
  - registry owns registration/remove log timing,
  - factory owns create-host log timing,
  - descriptor source stays in `TaskPaneManagerDiagnosticHelper`,
  - registry reads `WorkbookFullName` but does not write host metadata,
  - visible pane early-complete and render-current checks remain outside the registry.
- registry's direct decision surface is still limited to:
  - `windowKey` reuse lookup,
  - control-type compatibility by role,
  - replacement disposal,
  - workbook-scope remove selection by `WorkbookFullName`.
- Because descriptor content includes `WorkbookFullName`, registry logging is diagnostic-only in purpose but metadata-timing-adjacent in content. That is why constructor cleanup and logging cleanup should not be mixed with metadata timing changes.
- This phase remains docs-only. No runtime behavior, logging timing, metadata timing, visible pane early-complete, retry, foreground, `WorkbookOpen`, or `WindowActivate` behavior changed.

## B2.4 Update: Diagnostic-only Descriptor Dependency Clarification (2026-05-06)

- `TaskPaneHostRegistry` now makes the formatter dependency intent explicit in code:
  - the dependency is diagnostic-only,
  - it is not a host identity source,
  - it is not a metadata timing owner,
  - it is not part of replace/remove decisions.
- `TaskPaneManagerRuntimeGraphFactory.Compose(...)` also labels the callsite the same way, so the dependency meaning is visible at both compose time and consume time.
- This phase does not change descriptor content, call order, logging timing, metadata timing, visible pane early-complete, retry, foreground, `WorkbookOpen`, or `WindowActivate` behavior.

## B2.5 Update: Registry Registration/Remove Logging Inventory (2026-05-06)

- `TaskPaneHostRegistry` current-state around logging is now fixed as:
  - registration logging timing owner for `TaskPane host registered...`,
  - remove logging timing owner for `action=remove-host`,
  - diagnostic consumer of upstream descriptor output,
  - metadata consumer of `WorkbookFullName` only for workbook-scope remove selection.
- registration logging occurs after concrete host creation and after insertion into the shared host map.
- remove logging occurs before removal from the shared host map and before host dispose completes.
- descriptor consume timing remains removal-time only. It is diagnostic-only in purpose, but metadata-timing-adjacent because descriptor content includes `WorkbookFullName`.
- replace/remove decisions remain descriptor-independent and continue to rely on `windowKey`, control-type compatibility, and workbook-scope selection.
- visible pane early-complete, `WorkbookOpen` downstream flows, and retry/foreground paths remain indirect dependencies only. They were documented as sensitive touchpoints, not changed behaviors.
- This phase is docs-only. No logging timing, metadata timing, visibility behavior, or lifecycle behavior changed.

## B2 Checkpoint Update (2026-05-06)

### Architecture snapshot

- `TaskPaneManagerRuntimeBootstrap`
  - production runtime entry owner
  - attach-order owner
  - bridge between raw manager construction and attached runtime graph
- `TaskPaneManagerRuntimeGraphFactory`
  - passive runtime-graph builder
  - compose owner for runtime-only dependencies such as `TaskPaneHostFactory`
- `TaskPaneManager`
  - facade / orchestration boundary
  - shared host-map state owner via `_hostsByWindowKey`
  - final render seam owner for Case / Kernel / Accounting panes
- `TaskPaneHostRegistry`
  - replace/register/remove orchestration owner
  - registration/remove logging timing owner
  - diagnostic consumer and `WorkbookFullName` metadata consumer only
- `TaskPaneHostFactory`
  - control creation owner
  - `ActionInvoked` binding owner
  - create-host logging timing owner
- `TaskPaneHost`
  - concrete `CustomTaskPane` lifetime holder
- `ThisAddIn`
  - concrete VSTO adapter boundary for `CreateTaskPane(...)` / `RemoveTaskPane(...)`

### Runtime-sensitive boundary

- The next step beyond B2 is no longer composition cleanup. It is runtime surgery.
- The risk boundary starts where a change would affect any of:
  - host existence timing,
  - metadata write/read timing,
  - visible pane early-complete preconditions,
  - ready-show / retry behavior,
  - foreground recovery timing,
  - VSTO create/remove timing,
  - event unbinding behavior.
- B2 intentionally stopped before those changes. Current-state is now fixed so later work can isolate one runtime-sensitive unit at a time.

### Intentionally untouched in B2

- `WorkbookOpen` downstream flow
- ready-show / retry
- foreground recovery
- visibility retention and visible pane early-complete behavior
- metadata timing
- remove timing
- event unbinding behavior
- `_hostsByWindowKey` ownership

### Manual smoke checkpoint

- Human-side manual smoke for the B2 checkpoint was reported as OK.
- This checkpoint assumes external/manual validation for the previously tracked pane scenarios:
  - CASE pane
  - non-case pane
  - post-action refresh
  - visibility
  - `WorkbookOpen`
  - ready-show
  - pane remove / recreate
- As of the `taskpane-bootstrap-entry-shrink` merge on `main` (`98aa2eb735638a4e805fd16e7e257ef30c7f7607`), normal CASE / TaskPane operation regressions were not reported in human-side verification.
- A white-Excel flash can still recur during new CASE creation. This is a known unresolved issue that predates the merge and is not treated as a blocker or rollback trigger for `taskpane-bootstrap-entry-shrink`.
- Handle that issue separately under the window activation / ready-show / isolated Excel visibility / CASE create foreground ordering context rather than mixing it into this merge checkpoint.
- Observation-only logging was added around new CASE hidden create / hidden-for-display / WorkbookOpen-WindowActivate / ready-show / TaskPane show checkpoints for white-flash investigation. No cause assertion or behavior change is included in this note.

### Next-phase warning

- Do not treat the next phase as "small cleanup". The remaining work sits on runtime-sensitive boundaries.
- If runtime surgery starts, keep the unit narrower than "VSTO boundary redesign". Examples of acceptable future slicing are:
  - create/remove adapter timing only,
  - event unbinding debt only,
  - metadata timing only,
  - ready-show / visibility only.
- Do not mix those units in the same change.

### Metadata timing pointer (2026-05-07)

- Detailed current-state write/read/consumer/dependency inventory now lives in `docs/taskpane-manager-responsibility-inventory.md` under `B2.7 Status: Metadata Timing Inventory (2026-05-07)`.
- Current-state metadata write owners remain:
  - `TaskPaneManager.RenderHost(...)` -> `TaskPaneHost.WorkbookFullName`
  - `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)` -> `TaskPaneHost.LastRenderSignature`
  - forced-refresh invalidation and CASE post-action fallback rerender only rewrite `LastRenderSignature`
- Current-state direct consumers remain:
  - `TaskPaneRenderStateEvaluator` render-state / display-request checks
  - `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` -> ready-show early-complete
  - `TaskPaneHostRegistry.RemoveWorkbookPanes(...)`
  - `TaskPaneHostLifecycleService.RemoveStaleKernelHostsForRefresh(...)`
  - Case / non-Case action target resolution through `FindOpenWorkbook(host.WorkbookFullName)`
- `WorkbookOpen` still does not guarantee host metadata availability by itself. The window-safe downstream paths remain the first reliable metadata population/consumption boundary.
- This update is docs-only. No runtime timing, metadata timing, ready-show, retry, foreground, visibility, or lifecycle behavior changed.

### Ready-show / visibility / foreground dependency pointer (2026-05-07)

- Detailed current-state dependency inventory now lives in `docs/taskpane-manager-responsibility-inventory.md` under `B2.8 Status: Ready-Show / Visibility / Foreground Dependency Inventory (2026-05-07)`.
- Current-state ready-show entry still depends on downstream presentation handoff:
  - `KernelCasePresentationService` keeps the order `transient suppression release -> workbook-window visibility ensure -> CASE activation suppression preparation -> ready-show request`.
  - `AccountingSetCreateService` still hands off to the same ready-show entry after workbook activation-side preparation.
- Current-state visible-pane early-complete still depends on:
  - resolved `windowKey`,
  - retained host existence in `_hostsByWindowKey`,
  - `TaskPaneHost.WorkbookFullName` match,
  - hosted role being `Case`,
  - `TaskPaneHost.IsVisible`.
- Current-state foreground recovery still remains downstream of refresh success:
  - `TaskPaneRefreshCoordinator.GuaranteeFinalForegroundAfterRefresh(...)` runs only after refresh success with a resolved pane window.
  - CASE protection start remains downstream of that foreground recovery path.
- `WorkbookOpen` still does not guarantee host availability, visible-pane observability, or foreground recovery eligibility by itself. Those remain downstream, window-safe behaviors.
- This update is docs-only. No ready-show, visibility retention, foreground recovery, visible-pane early-complete, or lifecycle behavior changed.

### Event unbinding behavior pointer (2026-05-07)

- Detailed current-state inventory now lives in `docs/taskpane-manager-responsibility-inventory.md` under `B2.9 Status: Event Unbinding Behavior Inventory (2026-05-07)`.
- Current-state bind owner remains `TaskPaneHostFactory`:
  - Kernel / Accounting bind `ActionInvoked` before `TaskPaneHost` construction.
  - CASE binds `ActionInvoked` after `TaskPaneHost` construction.
- Explicit TaskPane-side unbinding ownership is still absent:
  - `TaskPaneHost.Dispose()` remains `Hide() -> ThisAddIn.RemoveTaskPane(...) -> _pane = null`.
  - `ThisAddIn.RemoveTaskPane(...)` still only calls `CustomTaskPanes.Remove(...)`.
  - repo-local code does not add a separate `ActionInvoked -= ...` or explicit control-dispose phase for these hosts.
- Compatible host reuse, display-request show-existing, and ready-show early-complete can all keep an already-bound host alive without re-entering create/bind.
- This update is docs-only. No event binding order, event unbinding behavior, create/remove timing, ready-show / retry, visibility, foreground, or lifecycle behavior changed.

## B2.10 Runtime-Sensitive Current-State Summary (2026-05-07)

### Fixed baseline

- This section fixes the resumable current-state on `main` commit `8ce9fed49b7f6a924a74b624f7c79d098e7f6a04`.
- Read this summary together with:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/taskpane-architecture.md`
  - `docs/taskpane-manager-responsibility-inventory.md`
  - `docs/taskpane-refresh-policy.md`
  - `docs/workbook-window-activation-notes.md`
  - `docs/thisaddin-boundary-inventory.md`
- This section is a summary layer only. Detailed inventories remain fixed in the linked docs.

### 1. 現在の設計到達点

#### Runtime-sensitive inventory fixed on `main`

- create/remove adapter ownership
  - create/remove adapter boundary remains `TaskPaneHostRegistry` / `TaskPaneHostFactory` / `TaskPaneHost` / `ThisAddIn.CreateTaskPane(...)` / `RemoveTaskPane(...)`.
  - detailed inventory pointer: `docs/taskpane-manager-responsibility-inventory.md` `B2.6 Status: Create/Remove Adapter Timing Inventory (2026-05-06)` and `docs/thisaddin-boundary-inventory.md` `9.2 Create/Remove Timing Pointer (2026-05-06)`.
- metadata timing
  - host metadata write/read timing is fixed as current-state and remains downstream of render/reuse/remove semantics.
  - detailed inventory pointer: `docs/taskpane-manager-responsibility-inventory.md` `B2.7 Status: Metadata Timing Inventory (2026-05-07)`.
- ready-show / visibility / foreground dependency
  - ready-show attempt, visible-pane early-complete, fallback refresh, final foreground recovery, and protection start are fixed as one dependency surface.
  - detailed inventory pointer: `docs/taskpane-manager-responsibility-inventory.md` `B2.8 Status: Ready-Show / Visibility / Foreground Dependency Inventory (2026-05-07)`.
- event unbinding behavior
  - explicit TaskPane-side unbinding ownership is still absent; teardown remains dispose-driven.
  - detailed inventory pointer: `docs/taskpane-manager-responsibility-inventory.md` `B2.9 Status: Event Unbinding Behavior Inventory (2026-05-07)` and `docs/thisaddin-boundary-inventory.md` `9.3 Event Unbinding Pointer (2026-05-07)`.
- `WorkbookOpen` downstream availability
  - `WorkbookOpen` still does not guarantee host creation, metadata availability, visible-pane observability, or foreground-recovery eligibility by itself.
  - the first reliable consumers remain downstream window-safe paths such as `WorkbookActivate`, `WindowActivate`, explicit display refresh, ready-show refresh, or pending retry fallback.
- visible-pane early-complete coupling
  - current early-complete still depends on `windowKey` lookup, retained host existence, `WorkbookFullName` match, hosted `Case` role, and `host.IsVisible`.
  - it still bypasses render-current checks, rerender, final foreground recovery, and new protection start on the success path.
- `_hostsByWindowKey` ownership
  - `TaskPaneManager` still owns `_hostsByWindowKey` and remains the shared host existence state owner.
  - ownership move is not part of the current safe boundary.

#### Diagnostics-only observability fixed on `main`

- visible-case-pane early-complete trace
  - `TaskPaneDisplayCoordinator` records `visible-case-pane-check` and `WorkbookTaskPaneReadyShowAttemptWorker` records `visibleCasePaneEarlyComplete=true` with `earlyCompleteBasis=retainedHost+metadataJoin+visibilityRetention`.
- ready-show / retry / fallback trace
  - `WorkbookTaskPaneReadyShowAttemptWorker` records `wait-ready-entry`, attempt start, window resolution, early-complete, refresh result, and attempts-exhausted.
  - `TaskPaneRefreshOrchestrationService` records retry scheduling/firing plus fallback handoff / prepare / immediate-success / deferred-active-context fallback.
- foreground recovery / protection skip trace
  - `TaskPaneRefreshCoordinator` records `foreground-recovery-decision`, `final-foreground-guarantee-start`, `final-foreground-guarantee-end`, and `protection-decision`.
  - precondition-side protection skip remains observable as `ignore-during-protection`.
- post-action metadata rewrite trace
  - `TaskPaneActionDispatcher` records `post-action-metadata` for both invalidation-only and fallback rewrite paths.
- show-existing / show-with-render decision trace
  - `TaskPaneDisplayCoordinator` records `display-entry-decision` with `decision=ShowExisting` or `decision=ShowWithRender`, and separately records `show-existing-pane`.
- `ExcelWindowRecoveryService` mutation trace
  - `ExcelWindowRecoveryService` records step-by-step mutation trace for `promote-*`, application foreground, and recovery flow.
- `WINDOWPLACEMENT` trace
  - mutation trace now includes `showCmd`, `rcNormalPosition`, `ptMinPosition`, `ptMaxPosition`, minimized/maximized/normal flags, `restoreSkipped`, `restoreSkipReason`, and `changedFields`.
- All of the above are observability additions or frozen current-state observations.
  - they are not a license to change behavior by adjacency.

### 2. Window placement 問題の調査結果

#### 症状

- Human-side observation on current `main` reported the following recurring symptom before the latest minimal fix:
  - CASE① を左半分に snap 配置
  - CASE② を右半分に snap 配置
  - re-activate 時に中央中サイズへ戻る
  - CASE close 後 follow-up activate でも再発する

#### trace で確定したこと

- rect が最初に変化した step は `promote-showwindow-restore-after` です。
- `showCmd` は restore 前から `SW_SHOWNORMAL` でした。
- minimized / maximized 状態ではありませんでした。
- 変化後の rect は `rcNormalPosition` と整合していました。
- `WindowState=xlNormal`、`Activate`、`SetForegroundWindow` は近接する処理ですが、現 trace では rect を最初に変えた step としては弱いです。

#### 原因仮説

- current best-fit hypothesis is a restore / normalize side effect caused by `ShowWindow(SW_RESTORE)` on a visible snapped window that is already `SW_SHOWNORMAL`.
- current docs do not assert when or why Excel / Windows rewrites `rcNormalPosition`; that part remains unknown.

### 3. 実施した最小修正

- current `main` keeps the foreground promotion path but changes the restore decision only:
  - skip `ShowWindow(SW_RESTORE)` when the window is visible, `SW_SHOWNORMAL`, not minimized, and not maximized
  - keep restore for minimized / maximized / hidden / placement-read-failed / other show-state cases
- topmost pulse and `SetForegroundWindow` ordering remain unchanged.
- rect restore / persistence was not introduced.
- Human-side manual validation on current `main` reported:
  - snap 配置維持成功
  - CASE 切替でも位置崩れなし
  - CASE close follow-up activate でも維持
  - 実機 OK

### 4. 現在の重要方針

- observability-first
- inventory-first
- minimal runtime surgery
- no large refactor
- no service explosion
- no abstraction-first
- user placement ownership を system が壊さない
- 「位置を保存して戻す」より「不要な restore を打たない」を優先する

### 5. 現在まだ frozen な領域

- create/remove ordering surgery
- explicit unbinding
- retry policy 変更
- `WorkbookOpen` orchestration surgery
- `_hostsByWindowKey` ownership move
- rect persistence
- `WindowState` policy 全体変更
- `Activate` / foreground ordering surgery

### 6. 積み残し課題

| Priority | Item | Current reading |
| --- | --- | --- |
| P1 | metadata consumer observability 拡張 | diagnostics-only で `WorkbookFullName` / `LastRenderSignature` の消費点をより比較しやすくし、後続 surgery の前提を固める |
| P1 | activation reuse observability | retained host reuse と activation 後再利用の分岐を、挙動不変の trace だけで見えるようにする |
| P1 | render-skip / show-existing coupling trace | `display-entry-decision` と render skip の関係をもう一段追えるようにし、 visible-pane early-complete との差を固定する |
| P1 | window mutation trace の追加解析 | placement mutation の前後でどの Win32 / Excel state が先に動くかを追加観測する |
| P1 | `rcNormalPosition` がいつ保存されるか未解明 | placement 崩れの根因切り分けに残る最大の unknown であり、まず観測を優先する |
| P2 | retry / fallback policy inventory 拡張 | ready retry `80ms` と pending retry `400ms` の依存差分を inventory-first で整理する |
| P2 | event lifetime / reuse coupling 解析 | retained host reuse と dispose-driven teardown の間にある event lifetime の実態を追加で固定する |
| P2 | `WorkbookOpen` downstream dependency 整理 | workbook-only と window-dependent の downstream 依存を崩さずに整理するための前段 inventory を増やす |
| P3 | foreground recovery 最適化候補 | runtime surgery に入るため、先行 observability と placement unknown の縮小後に扱う |
| P3 | visible-pane early-complete surgery 候補 | ready-show / metadata / visibility / foreground をまたぐため、単独 boundary として切るまで凍結する |

### 7. GO / STOP 原則

#### GO

- diagnostics-only
- local state observability
- minimal surgery
- 1 boundary / 1 commit
- behavior unchanged trace

#### STOP

- constructor fan-out 増加
- logger injection 波及
- service 追加
- ownership move
- timing normalisation
- create/remove ordering surgery
- `WorkbookOpen` / retry / foreground を同一 diff で混ぜる
