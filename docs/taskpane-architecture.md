# TaskPane Architecture

## 位置づけ

この文書は、TaskPane リファクタリング後の現行設計を固定するための正本です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- TaskPane refresh policy: `docs/taskpane-refresh-policy.md`
- 優先度A現在地の補足: `docs/taskpane-refactor-current-state.md`
- metadata / cache / snapshot の補足: `docs/template-metadata-inventory.md`
- 読取経路の補足: `docs/template-metadata-read-path-inventory.md`

この文書で扱う対象は、CASE 向け TaskPane の表示、TaskPane 用 metadata の正本と派生 cache、文書ボタン実行時の解決責務です。

## 結論

TaskPane 設計の現行正本は、次の整理で固定します。

1. 文書ボタン定義の runtime 正本は Kernel `雛形一覧` と Kernel `TASKPANE_MASTER_VERSION` です。
2. Base 埋込 snapshot と CASE snapshot cache は、どちらも正本ではなく派生 cache です。
3. TaskPane の表示は `CASE cache -> Base cache -> Master rebuild` の順で解決します。
4. 表示中の CASE は、後から成功した雛形登録・更新に自動追随しません。
5. 文書名入力は CASE cache にのみ従い、実行時 template 解決だけが master fallback を持ちます。
6. UI 制御は専用サービス経由で行い、`WorkbookOpen` 直後に直接表示制御しません。

## 正本と派生情報

### 正本

- Kernel `雛形一覧`
  - A列: `key`
  - B列: `TemplateFileName`
  - C列: `caption`
  - D列: 文書ボタン色
  - E列: タブ名
  - F列: タブ色
- Kernel `TASKPANE_MASTER_VERSION`

### 派生 cache

- Base 埋込 snapshot
  - `TASKPANE_BASE_SNAPSHOT_COUNT`
  - `TASKPANE_BASE_SNAPSHOT_XX`
  - `TASKPANE_BASE_MASTER_VERSION`
  - Base 側 `TASKPANE_MASTER_VERSION`
- CASE snapshot cache
  - `TASKPANE_SNAPSHOT_CACHE_COUNT`
  - `TASKPANE_SNAPSHOT_CACHE_XX`
  - CASE 側 `TASKPANE_MASTER_VERSION`

### 扱いの原則

- snapshot / cache は表示補助です。
- snapshot / cache を保存・生成・実行判断の正本にしてはいけません。
- `TemplatePath` は保存正本を持たず、`DocumentTemplateResolver` が都度導出します。

## 主要責務

### `KernelTemplateSyncService`

- `SYSTEM_ROOT\雛形` と `CaseList_FieldInventory` を前提に雛形登録・更新を行う
- Kernel `雛形一覧` を更新する
- `TASKPANE_MASTER_VERSION` を成功時に無条件で `+1` する
- Base 用 snapshot と version を再生成して埋め込む

### `TaskPaneSnapshotBuilderService`

- CASE 表示時の snapshot 解決元を選ぶ
- 優先順は `CASE cache -> Base cache -> Master rebuild`
- 必要時に CASE cache を更新する

### `MasterWorkbookReadAccessService`

- `MasterTemplateCatalogService` と `TaskPaneSnapshotBuilderService` が共有する Master 読み取り境界である
- Master path 解決、read-only open、所有 workbook の close、hidden window 化を一元化する

### `TaskPaneSnapshotCacheService`

- 文書ボタン実行時の CASE cache lookup を担う
- Base snapshot の on-demand promote を担う
- cache 互換性不一致時の clear を担う

### `TaskPaneManager`

- TaskPane 側の facade / composition root である
- `TaskPaneHostRegistry`、`TaskPaneDisplayCoordinator`、`TaskPaneHostLifecycleService`、`TaskPaneHostFlowService`、`TaskPaneActionDispatcher` などを組み立てる
- role 別 render 切替、CASE pane action wiring、周辺 service への委譲入口を担う
- `RefreshPane(...)` 本線は `TaskPaneHostFlowService` に委譲し、自身は retry / ready-show / protection / window resolve を持たない
- lightweight helper / policy として `TaskPaneManagerDiagnosticHelper`、`TaskPaneHostReusePolicy`、`TaskPaneRenderStateEvaluator`、`TaskPaneShowExistingPolicy`、`TaskPaneShowWithRenderPolicy` を使う

やってはいけないこと:

- ready-show / retry / pending timer / protection を保持しない
- `WorkbookContext` 生成や workbook/window resolve を持たない
- Excel event の購読や `WorkbookOpen` / `WindowActivate` 入口判定を持たない

### `TaskPaneHostFlowService`

- refresh-time の host flow を担当する
- 対象は precondition による hide-all / skip、stale kernel host cleanup の実行順、host selection、CASE host reuse、render 要否判定、show 前調停、最終 show である
- `TaskPaneHostLifecycleService` には lifecycle primitive を要求し、`TaskPaneDisplayCoordinator` には表示調停だけを要求する

やってはいけないこと:

- host 集合の長期保持、register / dispose all、workbook 単位 cleanup の主責務を持たない
- ready-show / pending retry / protection / workbook window resolve を持たない
- `WorkbookContext` 生成や Excel event handling を持たない

### `TaskPaneHostLifecycleService`

- registry-backed な host lifecycle primitive を担当する
- `TaskPaneHostRegistry` を通した get-or-replace / register / remove / dispose / workbook 単位 cleanup を担う
- refresh-time stale cleanup は flow service から要求されたときだけ実行する

やってはいけないこと:

- render / show / show 前調停 / reuse 判定を持たない
- ready-show / retry / protection / workbook window resolve を持たない
- Excel event 入口や `WorkbookContext` 生成を持たない

### `TaskPaneActionDispatcher`

- CASE pane の UIイベント受付を担う
- `TaskPaneBusinessActionLauncher` を通して `doc` / `accounting` / `caselist` 実行へ接続する
- `TaskPanePostActionRefreshPolicy` に従って post-action refresh の skip / defer / 即時再描画を調停する

### `TaskPaneBusinessActionLauncher`

- `doc` 実行前の `DocumentNamePromptService.TryPrepare(...)` 順序を固定する
- prompt 準備後に `DocumentCommandService` を呼び出す

### `TaskPaneRefreshOrchestrationService`

- Excel event / explicit request / ready-show から入る refresh orchestration の入口である
- `ShowWorkbookTaskPaneWhenReady(...)` を ready-show 入口として持ち、ready-show attempt 本体は `WorkbookTaskPaneReadyShowAttemptWorker` へ委譲する
- `ScheduleTaskPaneReadyRetry(...)` により ready-show retry `80ms` の scheduling を担う
- `RefreshPreconditionEvaluator`、`RefreshDispatchShell`、`PendingPaneRefreshRetryService`、`WorkbookPaneWindowResolver` を使い、precondition、dispatch、pending retry fallback、window resolve 入口を調停する
- `TaskPaneRefreshCoordinator` へ dispatch し、host selection / render / show 自体は行わない
- retry / protection / ready-show の policy 正本は `docs/taskpane-refresh-policy.md` を参照する
- `WindowActivatePaneHandlingService` は `WindowActivate` をこの orchestration へ接続する sibling entry service であり、host flow / host lifecycle ではない

やってはいけないこと:

- `TaskPaneHostRegistry` / `TaskPaneHostLifecycleService` / `TaskPaneHost` へ直接触れない
- ready-show attempt 本体や visible pane early-complete 判定を戻さない
- role 別 render や show/hide を持たない
- `WorkbookContext` の最終採用判断より後の host UI 制御を持たない

### `WorkbookTaskPaneReadyShowAttemptWorker`

- ready-show attempt 本体の実行境界である
- `TaskPaneDisplayRetryCoordinator` と `WorkbookTaskPaneDisplayAttemptCoordinator` を内側 helper として使い、attempt 1 の即時実行と retry 継続可否を進める
- attempt 1 のときだけ `WorkbookWindowVisibilityService.EnsureVisible(...)` を呼び、Workbook Window 可視化を前処理する
- window 解決後に `HasVisibleCasePaneForWorkbookWindow(...)` を確認し、既存 visible CASE pane があれば success 相当で early-complete する
- early-complete が成立しない場合だけ `TryRefreshTaskPane(...)` へ refresh を handoff する
- ready-show 側の試行が尽きた後は、自身で pending retry を持たず orchestration 側の fallback へ戻す

やってはいけないこと:

- `400ms` pending retry state / timer を持たない
- protection 判定や Excel event 入口判定を持たない
- CASE 専用 visible pane early-complete を accounting へ広げない

### `WorkbookWindowVisibilityService`

- workbook window visible ensure の共通責務を担う
- 対象 workbook の visible window を優先解決し、必要時だけ `Window.Visible = true` を補助する
- `KernelCasePresentationService` の ready-show 前処理と `WorkbookTaskPaneReadyShowAttemptWorker` の attempt 1 前処理で共用する

やってはいけないこと:

- ready-show retry scheduling / pending retry / protection 判定 / event flow を持たない
- refresh dispatch や host UI 制御を持たない

### `PendingPaneRefreshRetryService`

- `400ms` pending retry の fallback 経路を担う
- workbook target と active target を分けて追跡し、対象 workbook を再取得できる間はその workbook を追う
- ready-show / deferred refresh が即時成功しなかった後に retry sequence を開始する
- 対象 workbook を見失っても active CASE context があれば active refresh fallback を継続する

やってはいけないこと:

- ready-show attempt 本体や visible pane early-complete 判定を持たない
- workbook window visible ensure や protection 判定を持たない

### `TaskPaneRefreshPreconditionPolicy`

- `ShouldSkipWorkbookOpenWindowDependentRefresh(...)` を `WorkbookOpen` 直後の window-dependent refresh skip 判定の正本として持つ
- `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` はこの policy を利用し、skip 条件を重複保持しない
- 判定は pure であり、ログ出力・状態変更・COMメンバーアクセス・UI操作を持たない

### `DocumentNamePromptService`

- CASE cache から `caption` を引けた場合だけ prompt 初期値へ使う
- master fallback は行わない

### `DocumentTemplateResolver`

- CASE cache 優先で `key -> DocumentName / TemplateFileName` を解決する
- CASE cache miss 時だけ master catalog に fallback する
- `WORD_TEMPLATE_DIR` または `SYSTEM_ROOT\雛形` から `TemplatePath` を導出する

## 呼び出し関係

- Excel event handling
  - `ThisAddIn` / `WorkbookLifecycleCoordinator` / `WindowActivatePaneHandlingService` が入口を持ち、TaskPane 更新要求は `TaskPaneRefreshOrchestrationService` に渡す
- Ready-show path
  - `KernelCasePresentationService` / `AccountingSetCreateService` の ready-show 要求は `TaskPaneRefreshOrchestrationService` に入り、`WorkbookTaskPaneReadyShowAttemptWorker` が attempt 実行を担当し、失敗時だけ `PendingPaneRefreshRetryService` 側 fallback へ戻す
- Workbook 文脈
  - `TaskPaneRefreshCoordinator` が `WorkbookSessionService` と `ResolveWorkbookPaneWindow(...)` を使って `WorkbookContext` を確定し、その後で `TaskPaneManager.RefreshPane(...)` に渡す
- Host flow
  - `TaskPaneManager.RefreshPane(...)` は `TaskPaneHostFlowService` に委譲し、同 service が `TaskPaneHostLifecycleService`、`TaskPaneDisplayCoordinator`、role 別 render を順に調停する
- Control event handling
  - `TaskPaneHostRegistry` / `TaskPaneManager` が host と control の wiring を持ち、CASE pane UIイベントは `TaskPaneActionDispatcher`、非 CASE action は `TaskPaneNonCaseActionHandler` へ渡す

## 禁止境界

- `TaskPaneManager` に ready-show / retry / protection / workbook window resolve を戻さない
- `TaskPaneHostFlowService` に pending retry / ready-show / `WorkbookContext` 生成を持たせない
- `TaskPaneHostLifecycleService` に render / show / reuse 判定を戻さない
- `TaskPaneRefreshOrchestrationService` から `TaskPaneHostRegistry` / `TaskPaneHostLifecycleService` を直接触らせない
- `TaskPaneRefreshOrchestrationService` に ready-show attempt 本体を戻さない
- `WorkbookTaskPaneReadyShowAttemptWorker` に pending retry state を持たせない
- `WorkbookWindowVisibilityService` に ready-show / retry / protection 判定を持たせない
- CASE 専用 visible pane early-complete を accounting に広げない
- `WorkbookOpen` を window 安定通知とみなす前提を、host flow / lifecycle 側へ持ち込まない

## フロー固定

### 1. 雛形登録・更新成功時

1. `KernelTemplateSyncService` が Kernel `雛形一覧` を更新する
2. `TASKPANE_MASTER_VERSION` を `+1` する
3. Base 用 snapshot を再生成する
4. Base に `TASKPANE_BASE_SNAPSHOT_*`、`TASKPANE_BASE_MASTER_VERSION`、`TASKPANE_MASTER_VERSION` を保存する
5. master catalog cache を無効化する

注意:

- version 更新は内容差分比較ではなく、成功時無条件 `+1` です
- この方針は変えません

### 2. 新規 CASE 作成時

1. Base を物理コピーして CASE を作成する
2. Base 埋込 snapshot / version を新規 CASE が引き継ぐ
3. `CaseTemplateSnapshotService` が初期 promote を行う
4. 新規 CASE は原則として最新 snapshot を持った状態で開始する

### 3. 既存 CASE 表示時

1. `TaskPaneSnapshotBuilderService` が CASE cache を確認する
2. CASE cache が使えなければ Base snapshot を確認する
3. どちらも使えない場合だけ Master から rebuild する
4. いったん生成した Pane / host / control は、CASE を閉じるまで維持する

注意:

- `WorkbookActivate` / `WindowActivate` のたびに version 比較して Pane を作り直す設計ではありません
- 開いている CASE が最新雛形へ自動追随しないことは現行仕様です

### 4. 文書ボタン押下時

1. `TaskPaneActionDispatcher` がボタン押下を受ける
2. `TaskPaneBusinessActionLauncher` が `DocumentNamePromptService` を通して CASE cache の `caption` だけを使って prompt 初期値を準備する
3. `TaskPaneBusinessActionLauncher` が `DocumentCommandService` 実行に進める
4. `DocumentExecutionEligibilityService` が `DocumentTemplateResolver` を通して実行前確認を行う
5. `DocumentTemplateResolver` は CASE cache 優先、miss 時のみ master fallback で `TemplateFileName` を解決する
6. `DocumentCreateService` が文書生成する

責務分離:

- prompt 初期値は表示中 CASE に整合する補助情報
- 実行時解決は保存・生成判断に必要な正本確認入口
- UIイベント dispatch と post-action refresh は `TaskPaneActionDispatcher` に分離済み

### 5. 案件一覧登録後

1. CASE 側 `TASKPANE_SNAPSHOT_CACHE_COUNT` を `0` に戻す
2. `TASKPANE_SNAPSHOT_CACHE_XX` を削除する
3. Base 側 `TASKPANE_BASE_*` には触れない

## UI 制御方針

- TaskPane は左ドック固定です
- 表示制御は専用サービス経由です
- `WorkbookOpen` 直後に直接 UI 表示制御を追加しません
- 遅延表示、一時抑止、Window 単位の再利用を前提にします
- `ScreenUpdating` を変更した場合は必ず復元します

### WorkbookOpen と window 安全境界

- `WorkbookOpen` は workbook が開いた通知であり、window 安定通知ではありません。
- `WorkbookOpen` 時点では `ActiveWorkbook` / `ActiveWindow` が未確定な場合があります。
- workbook-only 処理と window-dependent 処理を混ぜない方針を維持します。
- `WorkbookOpen` 直後に workbook は取得できても window が未解決な refresh は、`TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` で skip し、後続イベントへ委ねます。
- pane 対象 window の確定、window key 依存の host 再利用、window 前提の表示調停は `WorkbookActivate` 以降、必要なら `WindowActivate` 以降を安全境界として扱います。
- `ResolveWorkbookPaneWindow` が安全に成功する条件は、対象 workbook の visible window を取得できること、または active workbook が対象 workbook と一致し active window を取得できることです。
- 単体生成 CASE 再オープン時の白 Excel 調査では、`WorkbookOpen` 時点の `ActiveWorkbook` / `ActiveWindow` 未確定が確認されました。startup context 系を再導入する前に、このイベント境界の安定化を優先します。
- `TaskPaneManagerOrchestrationPolicyTests` は、この skip 境界を `TaskPaneRefreshPreconditionPolicy` に対して直接検証します。

## 触ってはいけない固定点

- `TASKPANE_MASTER_VERSION` の成功時無条件 `+1`
- Base snapshot 埋め込み
- CASE cache 優先の表示・lookup
- `DocumentNamePromptService` の cache-only policy
- `DocumentTemplateResolver` の CASE cache 優先 + master fallback
- `WorkbookActivate` / `WindowActivate` の host 再利用経路
- snapshot / cache を正本扱いする変更
- `WorkbookOpen` 直接依存の表示制御
- `WorkbookOpen` を window 安定通知とみなす変更

## 今後課題として残すもの

- `TaskPaneHostRegistry`
  - 独立クラスとして host 生成、差し替え、破棄、workbook 単位掃除を担います。
  - `ThisAddIn` を通した VSTO `CustomTaskPane` 生成と action event 配線に関わるため、分離リスクが高い領域です。
  - 次に触る場合は `TaskPaneHostRegistry` だけを対象にし、action dispatch や refresh 本線には触れない方針を維持します。
- `ThisAddIn` 境界
  - `ThisAddIn` は VSTO lifecycle、application event、custom task pane 生成、TaskPane 表示要求の入口です。
  - `TaskPaneManager` / `TaskPaneHostRegistry` との依存境界を急に変えると起動、終了、pane 表示に波及するため、現時点では現状メモと依存関係棚卸しを優先します。

## 不明として残す事項

- Pane 再利用判定の全条件
- retry 間隔や protection 秒数の正式な仕様根拠
- 実機観測を伴うちらつきや ready-show の最終挙動詳細

コードだけで確定できない事項は、この文書でも断定しません。
