# TaskPane Refresh Policy

## 位置づけ

この文書は、現行 `main` における TaskPane refresh 周辺の policy 正本です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- TaskPane 設計正本: `docs/taskpane-architecture.md`
- Workbook / Window 境界の補足: `docs/workbook-window-activation-notes.md`
- 調査メモ: `docs/taskpane-protection-ready-show-investigation.md`
- 現在地文書: `docs/taskpane-refactor-current-state.md`
- 実機観測 baseline: `docs/taskpane-protection-baseline.md`
- 実機観測手順: `docs/taskpane-protection-observation-checklist.md`

この文書では、次を明示的に分けて扱います。

- 確定 policy
- 現行実装上の事実
- 未確定事項
- 今後の削減・整理候補

今回扱う対象は、`retry`、`protection`、`ready-show`、`WorkbookOpen / WorkbookActivate / WindowActivate` 境界、visible pane early-complete、pending retry `400ms`、ready retry `80ms`、attempts count、protection `5秒` です。

## 対象サービス

- `KernelCasePresentationService`
- `TaskPaneRefreshOrchestrationService`
- `WorkbookTaskPaneReadyShowAttemptWorker`
- `WorkbookWindowVisibilityService`
- `PendingPaneRefreshRetryService`
- `TaskPaneRefreshPreconditionPolicy`
- `TaskPaneRefreshCoordinator`
- `WorkbookLifecycleCoordinator`
- `WindowActivatePaneHandlingService`
- `KernelHomeCasePaneSuppressionCoordinator`

## 現在地

- `TaskPaneRefreshOrchestrationService` は、現在の `main` では refresh 本線の順序調停に寄っており、`RefreshPreconditionEvaluator`、`RefreshDispatchShell`、`PendingPaneRefreshRetryService`、`WorkbookPaneWindowResolver` への helper split が main に反映済みです。
- ready-show attempt 本体は `WorkbookTaskPaneReadyShowAttemptWorker` に分離済みで、`TaskPaneRefreshOrchestrationService` へ戻さない前提で読むべきです。
- workbook window visible ensure は `WorkbookWindowVisibilityService` に分離済みで、ready-show / protection / event flow の判定から切り離されています。
- protection / visible pane 判定 / ready-show 要求に関わる case-pane 系 `ThisAddIn` 依存は `ICasePaneHostBridge` 経由へ整理済みです。
- `TaskPaneRefreshCoordinator` は `KernelFlickerTrace` の structured trace を維持し、`04150a7` で obsolete route に付随していた duplicate plain log を削除済みです。

## サービス境界

### `TaskPaneRefreshOrchestrationService`

- ready-show / explicit refresh / Excel event 由来 refresh の入口です。
- `ShowWorkbookTaskPaneWhenReady(...)` を ready-show 入口として持ち、attempt 実行本体は `WorkbookTaskPaneReadyShowAttemptWorker` に委譲します。
- `TaskPaneReadyShowRetryScheduler` が ready retry `80ms` の scheduling を担います。
- `ScheduleWorkbookTaskPaneRefresh(...)` と `PendingPaneRefreshRetryService` が ready-show 失敗後の fallback refresh を受け持ちます。
- `ResolveWorkbookPaneWindow(...)` は ready-show と refresh orchestration が共有する window resolve 入口です。

### `WorkbookTaskPaneReadyShowAttemptWorker`

- ready-show attempt 本体を実行します。
- attempt 1 のときだけ `WorkbookWindowVisibilityService.EnsureVisible(...)` を呼びます。
- window 解決後に `HasVisibleCasePaneForWorkbookWindow(...)` を使い、CASE 専用 visible pane early-complete を判定します。
- early-complete が成立しない場合だけ `TryRefreshTaskPane(...)` へ refresh を handoff します。
- 自身では pending retry state / timer を持たず、attempt 枯渇後は orchestration 側 fallback へ戻します。

### `WorkbookWindowVisibilityService`

- workbook window visible ensure の共通責務です。
- ready-show の前処理で使いますが、ready-show / retry / protection / event flow の判定は持ちません。
- 返すのは visible ensure の outcome だけで、refresh dispatch や host UI 制御は持ちません。

### `PendingPaneRefreshRetryService`

- `400ms` pending retry fallback を担います。
- workbook target tracking と active target tracking を分けて持ちます。
- 対象 workbook を見失っても active CASE context があれば active refresh fallback を継続します。
- ready-show 側 retry の失敗後に入る fallback 先であり、ready-show attempt 本体は持ちません。

## 禁止境界

- `TaskPaneRefreshOrchestrationService` に ready-show attempt 本体を戻さない
- `WorkbookTaskPaneReadyShowAttemptWorker` に pending retry state を持たせない
- `WorkbookWindowVisibilityService` に ready-show / retry / protection 判定を持たせない
- CASE 専用 visible pane early-complete を accounting に広げない

## WorkbookOpen 境界

### 確定 policy

- `WorkbookOpen` 直後は window-dependent refresh の完了境界として扱いません。
- `WorkbookOpen` は workbook-only 境界として扱い、window 安定は `WorkbookActivate` / `WindowActivate` 以降で扱います。
- `WorkbookOpen` で workbook は取得できても window が未確定なら、必要に応じて defer / retry に回します。
- `WorkbookOpen` を window 安定境界として扱う方向へ docs や実装を戻しません。

### 現行実装上の事実

- `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` が shared skip policy の正本です。
- skip 条件は `reason == WorkbookOpen` かつ `workbook != null` かつ `window == null` です。
- `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` はこの policy を利用する側であり、同じ skip 条件を個別に持ちません。
- window 解決は、対象 workbook から visible window を取得できるか、active workbook が対象 workbook と一致して active window を取得できる場合に成功します。

### 未確定事項

- すべての実行環境で `WorkbookActivate` と `WindowActivate` のどちらを最終安全境界とみなすべきかは、コードだけでは確定しません。

## Ready-Show Policy

### 確定 policy

- CASE 表示直後の pane 表示は ready-show 経由で安定化します。
- ready-show は refresh dispatch 本体とは責務を分け、即時 1 回で決め打ちしない遅延表示経路として扱います。
- CASE 表示直後の順序は、transient suppression release、Workbook Window 可視化、CASE pane activation suppression 設定、ready-show 要求の順を崩しません。
- visible pane early-complete が成立する場合は、再描画完了そのものではなくても成功相当で終了してよい現行 policy として扱います。

### 現行実装上の事実

- `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)` が ready-show 入口であり、`WorkbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(...)` へ委譲します。
- `WorkbookTaskPaneReadyShowAttemptWorker` は内部 helper として `TaskPaneDisplayRetryCoordinator` と `WorkbookTaskPaneDisplayAttemptCoordinator` を使い、attempt 1 を即時実行します。
- `WorkbookPaneWindowResolveAttempts = 2` が ready-show 側の window 解決 attempt 上限として使われます。
- ready retry は `TaskPaneReadyShowRetryScheduler` により `80ms` 間隔で行われます。
- `WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce(...)` では attempt 1 のときだけ `WorkbookWindowVisibilityService` による Workbook Window 可視化補助を行います。
- visible pane early-complete の成立条件は、window 解決成功後に `HasVisibleCasePaneForWorkbookWindow(...)` が真になることです。
- visible pane early-complete は既存 CASE pane の不要な refresh 回避に使われ、この分岐は CASE 専用です。
- early-complete が成立しない場合、worker は `TryRefreshTaskPane(...)` へ refresh を handoff します。
- ready-show 側の試行が尽きた場合だけ `ScheduleWorkbookTaskPaneRefresh(...)` / `PendingPaneRefreshRetryService` の fallback へ移ります。

### 未確定事項

- ready-show がこの順序を必要とする正式な設計根拠は、既存 docs とコードだけでは完全には確定しません。
- ready-show 完了の最終 UX 定義は、実機観測なしには断定しません。

## Pending Retry Policy

### 確定 policy

- pending retry `400ms` は fallback refresh のための補助経路です。
- attempts count は retry 継続可否の管理値であり、window 解決や refresh dispatch の意味自体を変更しません。
- active target と workbook target は別物として扱います。
- retry は window 解決や refresh dispatch を直接書き換える責務ではありません。

### 現行実装上の事実

- `PendingPaneRefreshIntervalMs = 400`
- `PendingPaneRefreshMaxAttempts = 3`
- `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` は workbook target を `PendingPaneRefreshRetryService` に登録し、timer 開始前に一度 `TryRefreshTaskPane(...)` を試し、成功時は timer を開始しません。
- `PendingPaneRefreshRetryService` は対象 workbook を追う経路と、active target を追う経路を分けています。
- ready-show 側 retry の試行が尽きた場合、fallback は `ScheduleWorkbookTaskPaneRefresh(...)` からこの service に入ります。
- 対象 workbook を見失った場合でも、active context が CASE なら `TryRefreshTaskPane(reason, null, null)` による active refresh fallback を継続します。

### 未確定事項

- `400ms` と `3 attempts` の正式な仕様根拠は未確認です。
- active CASE context fallback が本番運用でどの程度必要かは、実機観測なしには確定しません。

## Protection Policy

### 確定 policy

- protection 中に無理な pane refresh をしません。
- protection ignore は `WorkbookActivate`、`WindowActivate`、`TaskPaneRefresh` の 3 入口にまたがる policy です。
- protection は CASE foreground 回復中の不要な再描画や再入を抑止する目的で扱います。
- protection の意味や適用入口を、実機観測なしに narrower / broader へ変更しません。

### 現行実装上の事実

- CASE pane activation suppression は対象 workbook の `FullName` と、`WorkbookActivate` 用 1 回、`WindowActivate` 用 1 回のカウントを持ちます。
- CASE pane activation suppression は、両カウント消費後または期限切れで解除されます。
- CASE foreground protection は CASE refresh 成功後に開始されます。
- suppression と foreground protection はどちらも `SuppressionDuration = TimeSpan.FromSeconds(5)` を使います。
- `ShouldIgnoreWorkbookActivateDuringProtection(...)` は対象 workbook と active window が protected target と一致する場合に止めます。
- `ShouldIgnoreWindowActivateDuringProtection(...)` は対象 workbook と event window が protected target と一致する場合に止めます。
- `ShouldIgnoreTaskPaneRefreshDuringProtection(...)` は active window が protected target と一致する場合に止めます。
- したがって、`TaskPaneRefresh` だけは入力 workbook / window 一致ではなく、active window 基準で広めに止めているのが現行実装上の事実です。

### 未確定事項

- protection `5秒` が UX 要件か暫定値かは未確認です。
- `TaskPaneRefresh` を active window 基準で止める正式な設計意図は未確定です。

## 数値と意味

次は現行コードで確認できる数値ですが、この文書では正式仕様と断定しません。

| 項目 | 現行実装上の事実 | この段階で断定しないこと |
| --- | --- | --- |
| ready-show window resolve attempts | `2` | 2 回で十分かどうか |
| ready retry 間隔 | `80ms` | UX 要件か経験則か |
| pending retry 間隔 | `400ms` | 業務要件としての必須性 |
| pending retry attempts | `3` | 適正回数の根拠 |
| suppression / protection duration | `5秒` | 正式設計値か暫定値か |

## 実機未確認として残す事項

- CASE 表示直後に、Workbook Window 可視化の後で TaskPane が従来どおり自然に出るか
- `WorkbookActivate` / `WindowActivate` の連続発火時に protection が効きすぎて TaskPane が出なくならないか
- ready-show retry が `80ms` 側で終わる代表ケースと `400ms` 側へ落ちる代表ケース
- visible pane early-complete により余計な refresh を避けられているか
- active CASE context fallback が実機でどの頻度で必要になるか

実機観測は `docs/taskpane-protection-baseline.md` と `docs/taskpane-protection-observation-checklist.md` を基準に扱います。

## 今後の削減・整理候補

### docs 上で整理候補として残すもの

- `ScheduleActiveTaskPaneRefresh` が production route か dead route 候補かの整理
- `80ms` / `400ms` / `3 attempts` / `5秒` を「正式仕様値」として固定するか、「現行実装値」として扱い続けるかの整理
- active target / workbook target を分けている `PendingPaneRefreshRetryService` 状態管理の必要性
- active CASE context fallback の必要条件
- visible pane early-complete を成立させる host metadata 依存の明文化
- protection 3入口の判定差を、タイミングを変えずに説明できる共通 policy 名へ寄せること

### 実機確認が必要な削減候補

- protection `5秒` の短縮・延長・廃止
- ready retry `80ms` と pending retry `400ms` の値変更
- attempts count の削減
- `TaskPaneRefresh` protection を active window 基準より狭くする変更
- visible pane early-complete 条件の単純化
- active CASE context fallback の削減

### まだ削ってはいけない条件

- `WorkbookOpen` shared skip policy
- suppression release -> Workbook Window 可視化 -> CASE pane activation suppression 設定 -> ready-show 要求の順序
- `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の protection 3入口
- visible pane early-complete を成功相当として扱う分岐
- pending retry の active target / workbook target 区別

## この文書から言えること

- retry / protection / ready-show は、単なる値や helper の問題ではなく、Workbook / Window 境界、suppression、foreground 回復、host 可視状態と結び付いた policy 群です。
- 現行 `main` では、`WorkbookOpen` を window 安定境界にしないこと、ready-show を遅延表示経路として扱うこと、protection を 3 入口で効かせることが前提です。
- 一方で、数値根拠と最終 UX は未確定の部分が残るため、次フェーズで削減に入る場合も、まずはこの policy と実機観測結果の突き合わせを先に行います。
