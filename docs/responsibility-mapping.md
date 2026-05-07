# Responsibility Mapping

## 位置づけ

この文書は、`docs/current-flow-source-of-truth.md` を基準座標として、案件情報System の現行責務を「変更理由ごとの ownership」という観点で観測するための docs です。目的は理想設計を書くことではなく、将来 CODEX に対して「この 1 責務だけを既存挙動不変で切り出せ」と依頼できる粒度を先に固定することです。

- 参照前提:
  - `AGENTS.md`
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/taskpane-architecture.md`
  - `docs/current-flow-source-of-truth.md`
  - `docs/case-workbook-lifecycle-current-state.md`
  - `docs/taskpane-refactor-current-state.md`
  - `docs/a2-window-visibility-current-state.md`
  - `docs/a4-c2-current-state.md`
  - `docs/a-priority-service-responsibility-inventory.md`
  - `docs/taskpane-manager-responsibility-inventory.md`
  - `docs/thisaddin-boundary-inventory.md`

この文書でいう `Current Owner` は理想 owner ではなく、現行 `main` で実際にその責務を持っている class / service を指します。1 つの責務が複数 owner にまたがる場合は、分裂した current owner をそのまま書きます。

1 つの責務に複数の sensitivity label が付くことがあります。`pure orchestration` と `pure business rule` は責務分類であり、owner class 全体の性質を断定するものではありません。同じ class 内に安全に寄せやすい責務と runtime-sensitive な責務が同居している場合があります。

## 0. 観測前提と固定ガード

- fail-closed:
  - context、binding、path、window、workbook が揃わないときは補正せず止まるのが current-state です。
  - これは「次善の推測で進める」よりも「誤った workbook / root / pane へ流さない」ことを優先しているためです。
- 再検索禁止:
  - `WorkbookContext` と `SYSTEM_ROOT` が source-of-truth の経路では、context-less fallback open や暗黙の workbook 推測を再導入しません。
  - 特に Kernel workbook 選択は「今開いているどれかを探し直す」よりも「要求元文脈に閉じる」ことが優先されます。
- close 後再参照禁止:
  - managed close / quiet close / hidden session cleanup は `Close` 後に対象 workbook を再参照しない current-state を維持します。
  - これは COM lifecycle と release timing を最小原則で固定するためです。
- foreground stabilization 必須:
  - CASE 表示完了は「window が visible」で終わらず、ready-show、suppression、protection、final foreground guarantee まで連鎖しています。
  - foreground 処理は見た目だけの問題ではなく、TaskPane refresh 成功条件と activation 再入抑止の一部です。
- ordering 固定:
  - `WorkbookOpen -> WorkbookActivate -> WindowActivate`
  - `release -> EnsureVisible -> SuppressUpcomingCasePaneActivationRefresh -> ShowWorkbookTaskPaneWhenReady`
  - `WriteToMasterList -> TASKPANE_MASTER_VERSION +1 -> Kernel save -> Base snapshot sync -> InvalidateCache`
  - `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up`
- 不明として残す事項:
  - `80ms` / `400ms` retry や `5秒` protection の正式な業務根拠
  - final foreground 安定化の正式 UX 完了条件
  - helper 非経由 close 全件の完全棚卸し

## 1. Responsibility Mapping Table

### 1.1 startup / shutdown

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| startup / shutdown | startup / shutdown 上位順序の保持 | `ThisAddIn` | event wiring、HOME 表示判定、startup refresh、hidden app-cache cleanup | ordering-sensitive | ordering 固定後でないと危険 | `InitializeStartupDiagnostics -> Compose -> Apply -> Hook -> startup 判定 -> startup refresh` の順序自体が current-state。 |
| startup / shutdown | startup context 事実収集と表示判定 | `KernelStartupContextInspector` + `KernelWorkbookStartupDisplayPolicy` | `ThisAddIn` の startup 呼び出し順 | pure orchestration, fail-closed-sensitive | 比較的安全に切り出せる | `WorkbookOpen` を window 安定境界に昇格させない前提を守る必要がある。 |
| startup / shutdown | application event wiring / unwiring | `ThisAddIn` + `ApplicationEventSubscriptionService` | VSTO lifecycle、event 順序前提 | ordering-sensitive | ordering 固定後でないと危険 | `WorkbookOpen -> WorkbookActivate -> WindowActivate` 前提を壊す変更は不可。 |

### 1.2 Kernel HOME

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| Kernel HOME | `KernelHomeForm` instance ownership | `ThisAddIn` | placeholder 表示、TaskPane hide、suppression 発行 | UI-only | 比較的安全に切り出せる | `thisaddin-boundary-inventory.md` では最小着手候補としてこの ownership だけが明示されている。 |
| Kernel HOME | HOME placeholder 表示シーケンス | `ThisAddIn` + `KernelWorkbookService` | form lifecycle、TaskPane hide、display preparation、WinForms show/activate | ordering-sensitive, window-sensitive | ordering 固定後でないと危険 | `unbound` HOME は placeholder-only。Kernel workbook の探索・open・自動 bind をしない。 |
| Kernel HOME | HOME binding / `SYSTEM_ROOT` 整合 | `KernelWorkbookBindingService` | display/close facade (`KernelWorkbookService`) | context-sensitive, fail-closed-sensitive | runtime-sensitive のため後回し | valid binding 不成立を補正する fallback open はしない。 |
| Kernel HOME | HOME close fail-closed handshake | `KernelWorkbookCloseService` | pending close、`FormClosing` cancel、`FormClosed` finalization、visibility release | fail-closed-sensitive, ordering-sensitive | ordering 固定後でないと危険 | backend close 成功前に Form を閉じない current-state を維持する。 |

### 1.3 CASE 作成

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| CASE 作成 | create plan / path resolve / name rule 解決 | `KernelCaseCreationService` | hidden session route 選択、save 前 window 正規化 | pure orchestration, context-sensitive | 比較的安全に切り出せる | `SYSTEM_ROOT` と出力先が揃わなければ進めない fail-closed 境界。 |
| CASE 作成 | hidden create session open / close mechanics | `CaseWorkbookOpenStrategy` | retained hidden app-cache、route 名ごとの差異、session cleanup | COM-sensitive, ordering-sensitive | COM lifecycle に密着している | hidden session は一般手段ではなく CASE 作成専用の例外。 |
| CASE 作成 | save 前の workbook window 正規化 | `KernelCaseCreationService` | hidden create route、interactive / batch 差分 | window-sensitive, ordering-sensitive | runtime-sensitive のため後回し | `visible + normal` への正規化は保存状態正規化であり、表示完了ではない。 |
| CASE 作成 | interactive route の shared app handoff | `KernelCasePresentationService` | wait UI、visibility recovery、ready-show、foreground promotion | ordering-sensitive, window-sensitive | ordering 固定後でないと危険 | hidden create session close 後にだけ表示責務が移る。 |

### 1.4 CASE 表示

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| CASE 表示 | wait UI lifecycle | `KernelCasePresentationService` + `CreatedCasePresentationWaitService` | suppression、cursor positioning、final foreground | UI-only, ordering-sensitive | 比較的安全に切り出せる | ただし close timing は ready-show / foreground 完了条件と近接する。 |
| CASE 表示 | ready-show 前の workbook visibility ensure | `WorkbookWindowVisibilityService` | `KernelCasePresentationService`、`WorkbookTaskPaneReadyShowAttemptWorker`、`ExcelWindowRecoveryService` | window-sensitive | runtime-sensitive のため後回し | CASE 表示経路と ready-show attempt の両方で共有される。 |
| CASE 表示 | initial cursor resolve と temporary read-only Kernel access | `KernelCasePresentationService` + `KernelWorkbookResolverService` | wait UI close、HOME activate、temporary COM release | context-sensitive, COM-sensitive | ordering 固定後でないと危険 | `Range.Select()` 後の local COM release を伴う。 |
| CASE 表示 | one-shot foreground promotion | `KernelCasePresentationService` | wait UI close、window visibility、deferred presentation completion | window-sensitive, ordering-sensitive | runtime-sensitive のため後回し | foreground は CASE 表示 UX の completion state に含まれる。 |

### 1.5 TaskPane lifecycle

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| TaskPane lifecycle | TaskPane runtime composition wiring | `AddInTaskPaneCompositionFactory` + `TaskPaneManagerRuntimeGraphFactory` + `TaskPaneHostFactory` / `TaskPaneHost` / `ThisAddIn` create-remove chain | shared host map owner、VSTO host wiring | pure orchestration | 比較的安全に切り出せる | manager attach surface は runtime-consumed collaborator に縮小済み。 |
| TaskPane lifecycle | refresh precondition 判定 | `TaskPaneRefreshPreconditionPolicy` + `TaskPaneRefreshOrchestrationService.RefreshPreconditionEvaluator` + `TaskPaneHostFlowService` entry gate consume | protection/suppression entry、host-flow entry hide-all/skip | pure orchestration | 比較的安全に切り出せる | `WorkbookOpen` 直後の window-dependent refresh skip と、`Unknown` role / empty `windowKey` の host-flow entry gate の正本。 |
| TaskPane lifecycle | ready-show attempt / early-complete | `WorkbookTaskPaneReadyShowAttemptWorker` | visibility ensure、host metadata、retry fallback | ordering-sensitive, window-sensitive | runtime-sensitive のため後回し | visible CASE pane がある場合は success 相当で終える current-state。 |
| TaskPane lifecycle | pending retry fallback | `PendingPaneRefreshRetryService` | active CASE context fallback、window resolve 再試行 | ordering-sensitive | runtime-sensitive のため後回し | workbook を見失っても active CASE context があれば継続する。 |
| TaskPane lifecycle | host flow / render / show | `TaskPaneManager` + `TaskPaneHostFlowService` + `TaskPaneDisplayCoordinator` + `TaskPaneHostLifecycleService` | shared host map、role 別 render、stale cleanup | ordering-sensitive, window-sensitive | runtime-sensitive のため後回し | facade、state owner、flow owner が分裂した current-state。 |
| TaskPane lifecycle | CASE post-action display 再入 | `TaskPaneActionDispatcher` + `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)` | CASE action dispatch、VSTO display entry | ordering-sensitive, window-sensitive | ordering 固定後でないと危険 | action dispatch が display 入口で閉じていない。 |
| TaskPane lifecycle | snapshot source selection / CASE cache 昇格 | `TaskPaneSnapshotBuilderService` + `TaskPaneSnapshotCacheService` | Base cache / Master rebuild read path | pure orchestration, context-sensitive | 比較的安全に切り出せる | 表示補助であり正本ではない。表示後追随を勝手にしない current-state が前提。 |

### 1.6 window activation / foreground stabilization

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| window activation / foreground stabilization | `WindowActivate` gating sequence | `WindowActivatePaneHandlingService` | protection predicate、external workbook 検知、suppression 判定、refresh 入口 | ordering-sensitive | ordering 固定後でないと危険 | 分岐順は `protection -> external workbook -> suppression -> refresh`。 |
| window activation / foreground stabilization | suppression / protection state | `KernelHomeCasePaneSuppressionCoordinator` | Kernel HOME suppression、CASE pane suppression、CASE foreground protection | ordering-sensitive, window-sensitive | runtime-sensitive のため後回し | 同一 coordinator に 3 種類の state が同居している。 |
| window activation / foreground stabilization | foreground recovery primitive | `ExcelWindowRecoveryService` | restore decision、Win32 API 呼び出し、placement guard | window-sensitive | COM lifecycle に密着している | `SW_RESTORE` の副作用回避を含む実行 owner。 |
| window activation / foreground stabilization | refresh 成功後の final foreground guarantee | `TaskPaneRefreshCoordinator` | ready-show success 境界、CASE protection 開始 | window-sensitive, ordering-sensitive | runtime-sensitive のため後回し | refresh が成功したときだけ foreground guarantee と protection 開始が走る。 |

### 1.7 Workbook close / COM release timing

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| Workbook close / COM release timing | dirty prompt / managed close orchestration | `CaseWorkbookLifecycleService` | session dirty、folder offer pending、CASE HOME 表示補正 | ordering-sensitive, fail-closed-sensitive | ordering 固定後でないと危険 | `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` を崩さない。 |
| Workbook close / COM release timing | no visible workbook 時の post-close quit | `PostCloseFollowUpScheduler` | Excel busy retry、visible workbook 判定、`Quit` | COM-sensitive, ordering-sensitive | COM lifecycle に密着している | CASE close 後の白 Excel 防止が目的。 |
| Workbook close / COM release timing | hidden session owned cleanup | `KernelUserDataReflectionService` | owned workbook close、owned application quit、quiet mode restore | COM-sensitive | COM lifecycle に密着している | `CloseWorkbookQuietly -> QuitApplicationQuietly -> FinalRelease` を owner 内で閉じる。 |
| Workbook close / COM release timing | temporary COM object local release | `KernelCasePresentationService` + `KernelUserDataReflectionService` + `AccountingSetKernelSyncService` など各 owner | business action 本線 | COM-sensitive | COM lifecycle に密着している | COM release timing は 1 箇所に集約されず、owner ごとの finally に分散している。 |

### 1.8 Accounting workbook open / sync

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| Accounting workbook open / sync | CASE からの accounting create / init / ready-show | `AccountingSetCreateService` | path resolve、wait UI、`File.Copy`、failure cleanup、ready-show handoff | context-sensitive, ordering-sensitive | ordering 固定後でないと危険 | `File.Copy` が最初の実副作用境界。 |
| Accounting workbook open / sync | Kernel からの accounting sync / quiet open-save-close | `AccountingSetKernelSyncService` | transfer plan、shared app quiet scope、owned workbook cleanup | COM-sensitive, fail-closed-sensitive | COM lifecycle に密着している | 不要な別 `Excel.Application` fallback を再導入しない current-state。 |
| Accounting workbook open / sync | application state snapshot / restore | `AccountingWorkbookService.BeginInitializationScope()` + `ExcelApplicationStateScope` | workbook open、cell write、本線 save | pure orchestration, COM-sensitive | runtime-sensitive のため後回し | `ScreenUpdating` と `EnableEvents` は必ず復元する。 |

### 1.9 reflection / navigation

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| reflection / navigation | reflection precondition / context validation | `KernelUserDataReflectionService` | hidden session owner、Base/Accounting 反映 plan | context-sensitive, fail-closed-sensitive | runtime-sensitive のため後回し | `WorkbookContext` / Kernel / `SYSTEM_ROOT` 不一致は補正しない。 |
| reflection / navigation | Base / Accounting 反映 plan 適用 | `KernelUserDataReflectionService` | quiet mode、hidden session cleanup、save 前 visibility restore | pure orchestration, COM-sensitive | runtime-sensitive のため後回し | plan 自体は分けられても owner は hidden session cleanup と近接している。 |
| reflection / navigation | CASE -> Kernel navigation / HOME suppression | `DocumentCommandService` + `ThisAddIn` + `KernelWorkbookService` | UI transition、sheet show、TaskPane refresh | ordering-sensitive, context-sensitive | ordering 固定後でないと危険 | bound context が無ければ進めず、再検索で補わない。 |

### 1.10 publication / template sync

| Flow | Responsibility | Current Owner | Mixed With | Runtime Sensitivity | Separation Difficulty | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| publication / template sync | template registration validation rules | `KernelTemplateSyncPreflightService` + `WordTemplateRegistrationValidationService` | preflight orchestration 呼び出し境界 | pure business rule, fail-closed-sensitive | 比較的安全に切り出せる | preflight failure では副作用を起こさない。 |
| publication / template sync | context-bound Kernel resolve / preflight orchestration | `KernelTemplateSyncService` | app state scope、sheet protection restore、result build | pure orchestration, context-sensitive | 比較的安全に切り出せる | `WorkbookContext` が唯一の入口。root 不一致を補正しない。 |
| publication / template sync | publication side-effect 実行順序 | `PublicationExecutor` | master list write、version bump、Kernel save、Base sync、invalidate | ordering-sensitive, fail-closed-sensitive | ordering 固定後でないと危険 | `Kernel save` が commit boundary。Base sync はその後だけ。 |
| publication / template sync | Base snapshot sync / managed close adjacency | `PublicationExecutor` + `KernelTemplateSyncService` | Base workbook open/save/close、managed close scope | COM-sensitive, ordering-sensitive | COM lifecycle に密着している | publication でありながら lifecycle close path に接続する。 |

## 2. Runtime Sensitivity Classification

責務分類は service 単位ではなく責務単位で付与します。特に `pure orchestration` は「runtime-sensitive ではない責務」を見分けるためのラベルであり、同じ class が他の row では runtime-sensitive になりえます。

| Classification | このシステムでの意味 | 代表責務 | 代表 owner |
| --- | --- | --- | --- |
| `COM-sensitive` | workbook / application close、quiet close、owned workbook release、Win32 window restore など、COM lifecycle と release timing に直結する責務 | hidden create session mechanics、hidden reflection cleanup、post-close quit、Base snapshot close adjacency | `CaseWorkbookOpenStrategy`、`KernelUserDataReflectionService`、`PostCloseFollowUpScheduler`、`ExcelWindowRecoveryService` |
| `ordering-sensitive` | 現行順序そのものが安全装置であり、前後を崩すだけで挙動が変わる責務 | startup/shutdown、HOME placeholder 表示、CASE display handoff、ready-show、publication side effects、managed close path | `ThisAddIn`、`KernelCasePresentationService`、`TaskPaneRefreshCoordinator`、`PublicationExecutor` |
| `window-sensitive` | visible window resolve、foreground、placement、pane target window など、window 安定化前提に依存する責務 | visibility ensure、foreground recovery、final foreground guarantee、one-shot promotion | `WorkbookWindowVisibilityService`、`ExcelWindowRecoveryService`、`KernelCasePresentationService` |
| `fail-closed-sensitive` | 欠損や不一致を補正せず止める current-state 自体が仕様の一部である責務 | HOME close handshake、Kernel workbook resolve、reflection precondition、publication preflight | `KernelWorkbookCloseService`、`KernelUserDataReflectionService`、`KernelTemplateSyncService` |
| `context-sensitive` | `WorkbookContext`、`SYSTEM_ROOT`、CASE context のどれを source-of-truth にするかが重要な責務 | Kernel resolve、CASE create plan、reflection、navigation、snapshot source selection | `KernelWorkbookBindingService`、`KernelCaseCreationService`、`KernelUserDataReflectionService`、`TaskPaneSnapshotBuilderService` |
| `pure orchestration` | 複数 service の呼び順や precondition を調停するが、自身は runtime primitive を持たない責務 | startup context fact collection、TaskPane runtime composition wiring、refresh precondition 判定、publication preflight orchestration | `KernelStartupContextInspector`、`AddInTaskPaneCompositionFactory`、`TaskPaneRefreshPreconditionPolicy`、`KernelTemplateSyncService` |
| `pure business rule` | workbook/window/COM の lifecycle から独立した判定・検証責務 | template registration validation rules | `KernelTemplateSyncPreflightService`、`WordTemplateRegistrationValidationService` |
| `UI-only` | wait UI や form instance の見せ方・寿命に閉じた責務 | `KernelHomeForm` instance ownership、CASE wait UI lifecycle | `ThisAddIn`、`CreatedCasePresentationWaitService` |

補足:

- `UI-only` は低危険度の同義ではありません。UI-only でも順序が近接していると `ordering-sensitive` を併発します。
- `pure orchestration` は「今すぐ class を分ける」意味ではありません。将来の安全単位を選ぶための観測タグです。
- `fail-closed-sensitive` は「柔軟に動くべき」ではなく、「推測で救済しない current-state」を維持すべき責務を指します。

## 3. Separation Safety Classification

`Separation Difficulty` 列は次の 4 区分で読むものとします。複数にまたがる場合は、より危険な側へ倒して扱います。

| Safety Class | 判定基準 | 代表責務 | 備考 |
| --- | --- | --- | --- |
| 比較的安全に切り出せる | runtime primitive よりも wiring、plan build、policy、message build に寄る | startup context 事実収集、TaskPane runtime composition wiring、refresh precondition 判定、snapshot source selection、template validation rules | ただし call site の順序は固定したまま、owner だけを切る前提。 |
| runtime-sensitive のため後回し | 順序だけでなく refresh 成功条件、visible pane 判定、visibility retain、suppression/protection state へ波及する | ready-show attempt、pending retry、host flow/render/show、suppression/protection state、final foreground guarantee | 先に inventory と frozen line を固定してからでないと危険。 |
| ordering 固定後でないと危険 | 現行順序が安全装置であり、責務自体は切れても reorder すると壊れる | startup/shutdown、HOME placeholder 表示、CASE display handoff、WindowActivate gating sequence、publication side-effect order、managed close path | 「抽出」より「順序不変」を優先する。 |
| COM lifecycle に密着している | close / quit / final release / owned workbook cleanup / Win32 restore の owner であり、再参照禁止や cleanup 完結性が本体 | hidden create session、post-close quit、hidden reflection cleanup、Base snapshot close adjacency、temporary COM release scope | 最後寄りで扱う。別責務と抱き合わせにしない。 |

この文書での原則:

- `pure orchestration` を先に扱う。
- `runtime-sensitive` は current-state の frozen line が十分に固定された後へ回す。
- `ordering-sensitive` は class 分割より順序固定を優先する。
- `COM-sensitive` は他の整理と抱き合わせず最後寄りに扱う。

## 4. Current Mixed Ownership Inventory

| Focus Area | Current mixed owners | 何が混在しているか | future task で凍結すべきもの |
| --- | --- | --- | --- |
| `TaskPaneManager` | `TaskPaneManager`、`TaskPaneManagerRuntimeGraphFactory`、`TaskPaneHostRegistry`、`TaskPaneHostFlowService`、`TaskPaneDisplayCoordinator`、`TaskPaneActionDispatcher` | facade entry surface、shared host state owner、manager attach surface の外にある registration orchestration、role 別 render、CASE post-action display 再入、visible pane bridge | `_hostsByWindowKey` owner、`CreateTaskPane/RemoveTaskPane` 境界、ready-show / protection / retry 本線 |
| `ThisAddIn` | `ThisAddIn`、`ApplicationEventSubscriptionService`、`TaskPaneRefreshOrchestrationService`、`KernelWorkbookService` | VSTO lifecycle、event handler、HOME form instance、TaskPane display entry、`CustomTaskPane` create/remove adapter、automation public surface、suppression / protection predicate bridge | startup/shutdown 順序、event 順序、`WorkbookOpen -> WorkbookActivate -> WindowActivate` 前提、`ScreenUpdating` restore |
| Workbook lifecycle 系 | `KernelWorkbookCloseService`、`CaseWorkbookLifecycleService`、`PostCloseFollowUpScheduler`、各 owner service の finally | HOME close fail-closed handshake、dirty prompt、managed close、post-close quit、temporary COM release、CASE HOME 表示補正 | close 後再参照禁止、`Quit` 成功後 restore しない、`ExcelApplicationStateScope` を managed close / quit に持ち込まない |
| CASE create / open | `KernelCaseCreationService`、`CaseWorkbookOpenStrategy`、`KernelCasePresentationService`、`WorkbookWindowVisibilityService`、`ExcelWindowRecoveryService` | create plan、hidden session route、save 前 window 正規化、wait UI、visibility recovery、ready-show handoff、initial cursor、final foreground | hidden create sessionを一般化しない、interactive handoff 前に close し切れない限り表示へ昇格しない |
| publication / template sync | `KernelTemplateSyncService`、`PublicationExecutor`、`KernelTemplateSyncPreflightService`、`MasterTemplateCatalogService` | preflight orchestration、application state scope、sheet protection restore、master list write、version bump、Kernel save、Base snapshot sync、invalidate | `WorkbookContext` / `SYSTEM_ROOT` 文脈、`Kernel save` commit boundary、base sync failure 時の warning semantics |
| foreground / window stabilization | `WindowActivatePaneHandlingService`、`KernelHomeCasePaneSuppressionCoordinator`、`WorkbookWindowVisibilityService`、`ExcelWindowRecoveryService`、`TaskPaneRefreshCoordinator`、`KernelCasePresentationService` | `WindowActivate` gating、suppression/protection state、visible ensure、Win32 restore、ready-show downstream、one-shot promotion、final foreground guarantee | `WorkbookOpen` を window 安定境界にしない、protection 3 入口の整合、restore semantics、foreground completion 条件 |

観測上のポイント:

- `TaskPaneManager` は「巨大だから分割対象」ではなく、「変更理由 owner が複数層に分裂している」ことが主問題です。
- `ThisAddIn` は委譲が進んでいても、VSTO 境界と UI 境界が残るため、runtime-sensitive な ownership が自然に集まり続けます。
- Workbook lifecycle は 1 つの close owner に集約されていません。むしろ「close 本線 owner」「post-close quit owner」「temporary COM owner」が分散していることが current-state です。
- CASE create / open は hidden session と visible presentation の owner が明確に分かれている一方、その handoff 順序が固定点になっています。
- publication / template sync は `PublicationExecutor` に side effects を集約済みですが、Base snapshot close adjacency まで含めると lifecycle 密着度が高いままです。
- foreground / window stabilization は owner が多いこと自体が問題ではなく、「1 つの completion condition が複数 owner にまたがる」ことが危険点です。

## 5. Recommended Refactor Order

壊れにくさ優先で future task の順序を並べると、現時点では次の順になります。

1. `pure orchestration` と wiring owner だけを切る。
  - 例: `TaskPaneManagerRuntimeGraphFactory` wiring / manager attach surface 整理、startup context fact collection、preflight orchestration、snapshot source selection。
   - 条件: runtime order と public surface は固定のまま、owner だけを整理する。
2. `pure business rule` と message/result build を切る。
   - 例: template registration validation rules、publication result/message build。
   - 条件: side-effect order と fail-closed 条件は触らない。
3. `UI-only` だが runtime core に食い込まない instance ownership を切る。
   - 例: `KernelHomeForm` instance ownership、wait UI host。
   - 条件: 表示順序、suppression、ready-show、foreground completion は変えない。
4. `context-sensitive` だが COM primitive を持たない plan / resolver を切る。
   - 例: CASE create plan、initial cursor resolver、Kernel navigation coordinator 候補。
   - 条件: `WorkbookContext` / `SYSTEM_ROOT` source-of-truth と再検索禁止を維持する。
5. `ordering-sensitive` の中でも、まず sequence だけを docs とテストで固定する。
   - 対象: startup/shutdown、CASE display handoff、publication side-effect order、managed close path。
   - 条件: reorder を伴う抽出を同時に行わない。
6. `window-sensitive` / foreground 系は dedicated phase に分ける。
   - 対象: visible window resolve ownership、suppression/protection state owner、foreground retry semantics。
   - 条件: ready-show、visible pane early-complete、protection 3 入口、restore semantics を同時変更しない。
7. `COM-sensitive` は最後寄りに扱う。
   - 対象: hidden create session、hidden reflection cleanup、post-close quit、Base snapshot close adjacency、temporary COM release catalog。
   - 条件: close 後再参照禁止、minimum COM 原則、owned workbook / owned application cleanup 完結性を崩さない。
8. `ThisAddIn` / VSTO create-remove boundary は独立タスクとして最後に扱う。
   - 対象: `CreateTaskPane(...)` / `RemoveTaskPane(...)`、event handler 境界、startup/shutdown lifecycle。
   - 例外: `KernelHomeForm` instance ownership だけは別安全単位として先行候補にできる。

この順序は「SOLID っぽさ」ではなく、「runtime-sensitive core に当たる前に、先に owner のにじみだけ減らす」ことを優先した current-state の推奨順です。
