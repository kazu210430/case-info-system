# TaskPane ready-show retry contract investigation

## 位置づけ

この文書は、R06 ready-show retry scheduler ownership を再開する前に、`ready-show max attempts` の source-of-truth を分離するための調査メモです。

今回の目的は実装方針の決定ではありません。現行 runtime truth、docs truth、contract truth を分け、R06 / R07 / R08 の retry ownership 境界を曖昧なまま動かさないための判断材料を固定します。

参照した正本:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-display-recovery-current-state.md`
- `docs/taskpane-refresh-orchestration-responsibility-inventory.md`
- `docs/taskpane-refresh-orchestration-target-boundary-map.md`
- `docs/taskpane-display-recovery-freeze-line.md`
- `docs/taskpane-refresh-policy.md`
- `docs/case-display-recovery-protocol-current-state.md`
- `docs/visibility-foreground-boundary-current-state.md`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookTaskPaneReadyShowAttemptWorker.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneDisplayRetryCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRetryTimerLifecycle.cs`
- `dev/CaseInfoSystem.ExcelAddIn/AddInCompositionRoot.cs`
- `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`

この文書は調査 docs です。コード変更、retry 順序変更、trace 名変更、callback 意味変更、completion 条件変更、display session 変更、R06 ownership 分離、timer lifecycle 変更、Deploy / runtime Addins 更新、Release package 作成は行いません。

## 結論サマリ

ready-show max attempts の contract truth は A「docs / freeze line の `2 attempts` を正式 contract truth とする」を採用します。

runtime truth は `3 attempts` です。production composition では pending retry 用の `PendingPaneRefreshMaxAttempts = 3` が `TaskPaneDisplayRetryCoordinator` へ渡されており、ready-show retry coordinator の max attempts としても効いています。

docs truth は `2 attempts` です。`docs/taskpane-display-recovery-freeze-line.md` は ready-show retry を `attempt 1 -> 80ms retry attempt 2 -> pending fallback` として固定しています。

contract truth は次です。

```text
attempt 1 -> 80ms attempt 2 -> pending fallback
```

`3 attempts` は pending retry constant の composition leakage と扱います。ready-show retry と pending retry は別責務であり、pending retry `400ms / 3 attempts` は ready-show exhaustion 後の fallback retry です。R06 ownership 分離を再開する前に、runtime / composition / tests を `2 attempts` に揃える必要があります。

## 1. retry ownership 境界

### ready-show retry

ready-show retry は created CASE 表示後の TaskPane 表示安定化 route です。

現行構造:

- `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)`
  - ready-show request を受理します。
  - created CASE display session を開始します。
  - `WorkbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(...)` へ委譲します。
- `WorkbookTaskPaneReadyShowAttemptWorker`
  - 1 attempt の owner です。
  - attempt 1 のみ `WorkbookWindowVisibilityService.EnsureVisible(...)` を実行します。
  - `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` を呼びます。
  - already-visible 判定または refresh delegate 呼び出しを行います。
  - pending retry state は持ちません。
- `TaskPaneDisplayRetryCoordinator`
  - attempt number の進行と max attempts 判定を持ちます。
  - `scheduleRetry` callback 経由で delayed attempt を発火します。
  - max attempts 超過時に fallback callback を呼びます。
- `TaskPaneReadyShowRetryScheduler`
  - ready-show delayed retry scheduler です。
  - `80ms` timer を `TaskPaneRetryTimerLifecycle` へ登録し、timer firing 時に retry action を呼びます。
- `TaskPaneRetryTimerLifecycle`
  - R16 timer lifecycle owner です。
  - timer の create / register / stop / unregister / dispose を持ちます。
  - retry sequence の意味や max attempts は持ちません。

### pending retry

pending retry は ready-show attempts exhausted 後の fallback refresh route です。

現行構造:

- `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)`
  - ready-show fallback handoff を受けます。
  - `WorkbookOpen` window-dependent skip を確認します。
  - target workbook を tracking します。
  - timer 開始前に immediate refresh を 1 回試します。
  - immediate success しない場合だけ pending retry を開始します。
- `PendingPaneRefreshRetryService`
  - pending retry state / callback owner です。
  - pending retry `400ms / 3 attempts` を持ちます。
  - target workbook retry と active CASE context fallback を持ちます。
- `TaskPaneRetryTimerLifecycle`
  - pending timer の lifecycle だけを持ちます。

## 2. `2 attempts` が導入された経緯

`WorkbookTaskPaneReadyShowAttemptWorker` の `ReadyShowMaxAttempts = 2` と `ReadyShowRetryDelayMs = 80` は、commit `da3bd98` (`diagnostics: correlate ready-show retry and fallback handoff`) で導入されています。

この commit の主目的は diagnostics / trace correlation です。追加された `ReadyShowMaxAttempts` は、主に次に使われています。

- `wait-ready-entry` / `wait-ready-attempt-start` / attempt refresh logs の `maxAttempts` field
- `ready-show-attempts-exhausted` 相当の trace 発火条件

重要な点として、この `ReadyShowMaxAttempts = 2` は `TaskPaneDisplayRetryCoordinator` の `_maxAttempts` へ渡されていません。したがって、現行 production runtime の retry 上限を直接制御していません。

## 3. `3 attempts` が composition に残っている理由

`PendingPaneRefreshMaxAttempts = 3` は `TaskPaneRefreshOrchestrationService` の初期追加 commit `e76b4fa` (`Add dev projects`) から存在しています。

同じ commit で、`AddInCompositionRoot` は次の形で `TaskPaneDisplayRetryCoordinator` を作成しています。

```csharp
var taskPaneDisplayRetryCoordinator = new TaskPaneDisplayRetryCoordinator(_pendingPaneRefreshMaxAttempts);
```

現在も production composition はこの構造を引き継ぎ、`ThisAddIn` から `TaskPaneRefreshOrchestrationService.PendingPaneRefreshMaxAttempts` を composition root へ渡しています。

このため、`3` は名前上は pending retry max attempts ですが、composition 経由で ready-show retry coordinator の `_maxAttempts` としても使われています。

現時点で確認できる範囲では、`3` が ready-show 専用の意図で残っている根拠は見つかっていません。より正確には、pending retry 用の定数が ready-show generic retry coordinator に流用されている状態です。

## 4. production runtime で実際に何回 retry されるか

現行 production runtime では、ready-show attempt は最大 3 回実行されます。

流れ:

1. `TaskPaneDisplayRetryCoordinator.ShowWhenReady(...)` が attempt 1 を即時実行します。
2. attempt 1 が false の場合、attempt 2 を `scheduleRetry` へ渡します。
3. 調査時点では `TaskPaneRefreshOrchestrationService.ScheduleTaskPaneReadyRetry(...)` が `80ms` timer を登録していました。R06 safe unit 後は `TaskPaneReadyShowRetryScheduler` が同じ `80ms` timer scheduling を担います。
4. attempt 2 が false の場合、attempt 3 を `scheduleRetry` へ渡します。
5. attempt 3 も false の場合、次の attempt number が max attempts を超え、fallback callback が呼ばれます。
6. fallback callback は `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` です。

したがって runtime sequence は、静的コード上は次です。

```text
attempt 1 -> 80ms retry attempt 2 -> 80ms retry attempt 3 -> pending fallback
```

これは freeze line の `attempt 1 -> 80ms retry attempt 2 -> pending fallback` とは一致しません。

## 5. `3` の意味

現時点の読み分け:

| 値 | runtime 上の効き先 | 名前 / docs 上の意味 | 実際の状態 |
| --- | --- | --- | --- |
| `PendingPaneRefreshMaxAttempts = 3` | pending retry max attempts | pending retry `400ms / 3 attempts` | 正しい pending retry 用値として使われる |
| `_pendingPaneRefreshMaxAttempts` | `TaskPaneDisplayRetryCoordinator` constructor | composition root の引数名も pending retry 寄り | ready-show max attempts としても使われる |
| `ReadyShowMaxAttempts = 2` | worker logs / exhausted trace condition | ready-show max attempts | production retry controller には未接続 |
| `WorkbookPaneWindowResolveAttempts = 2` | window resolve attempts と scheduler trace の `maxAttempts` | window resolve attempts | ready-show max attempts と混同されやすい |

したがって `3` は、pending retry 用でありながら ready-show retry 用にも効いています。coordinator 全体 retry 用の値として扱われている、というのが runtime truth です。

## 6. freeze line docs の性質

`docs/taskpane-display-recovery-freeze-line.md` は、自身の読み方として「今後の refactor で変えてはいけない runtime 契約」と定義しています。

その文書は ready-show retry について次を固定しています。

- ready-show max attempts は `2`
- ready-show retry は attempt 1 後の `80ms` attempt 2 だけ
- pending retry `400ms / 3 attempts` は ready-show exhaustion 後の fallback retry
- pending fallback は attempt 1 と attempt 2 が表示成立しない場合だけ入る

ただし、上記は現行 production runtime の静的事実とは一致していません。

したがって今回の調査では、freeze line docs は「runtime を記録しているつもりの contract docs」ですが、この一点では runtime 事実を正確に写していない可能性が高い、と扱います。

## 7. 不一致は危険か harmless か

この不一致は harmless discrepancy ではありません。

理由:

- pending fallback handoff の発火タイミングが 1 attempt 分ずれます。
- runtime では attempt 3 成功により `HandleWorkbookTaskPaneShown(...)` callback が呼ばれ得ます。
- completion details / trace に attempt number `3` が入り得ます。
- `ready-show-attempts-exhausted` は worker 側の `ReadyShowMaxAttempts = 2` を見て attempt 2 失敗時点で出得ますが、runtime はまだ attempt 3 を行います。
- 調査時点の `ScheduleTaskPaneReadyRetry(...)` trace は `maxAttempts=2` を出しますが、attempt 3 を schedule / firing し得ました。runtime alignment 後は attempt 3 は ready-show path に存在しません。
- `ready-show-fallback-handoff` は docs truth より 80ms 遅れます。

一方で、callback 意味、completion 条件、display session owner そのものは、この不一致だけでは別 owner に移っていません。危険なのは、観測契約と retry sequence contract が runtime と docs で割れている点です。

## 8. trace / callback / completion への影響

### trace

影響は大きいです。

想定される矛盾:

- `attempt=3, maxAttempts=2` の retry trace が出得る。
- `ready-show-attempts-exhausted` が attempt 2 failure で出ても、その後に attempt 3 が走り得る。
- `ready-show-fallback-handoff` は attempt 3 failure 後にだけ出る。
- trace contract 上の `attempt 1 -> 80ms retry attempt 2 -> pending fallback` と runtime event sequence が一致しない。

### callback

callback の意味は変わっていません。

`HandleWorkbookTaskPaneShown(...)` は、ready-show attempt が shown と判定された後に raw facts を normalized outcomes へ変換し、created CASE display session completion へ接続する再収束点です。

ただし runtime では attempt 3 success によっても callback が呼ばれ得るため、docs truth の `attempt 2` 上限とは異なる completion source が成立し得ます。

### completion

completion 条件そのものは変わっていません。

`case-display-completed` は引き続き `TaskPaneRefreshOrchestrationService.TryCompleteCreatedCaseDisplaySession(...)` 相当の条件でのみ emit されます。

ただし、ready-show attempt 3 success が completion chain に入れるため、completion 到達までの retry sequence と trace details は freeze line と一致しません。

## 9. A/B/C/D 判断

採用判断は A です。

A「docs の `2 attempts` が正しく、production composition `3` は修正対象」を ready-show retry contract truth として固定します。

理由:

- ready-show retry は created CASE 表示直後の短い安定化吸収であり、`attempt 1 -> 80ms attempt 2 -> pending fallback` が責務に合います。
- pending retry `400ms / 3 attempts` は ready-show exhaustion 後の fallback retry であり、ready-show retry そのものではありません。
- `PendingPaneRefreshMaxAttempts = 3` は pending retry 用の名前と意味を持つ値です。これが ready-show retry coordinator に渡される現行 runtime は、composition leakage と扱います。
- attempt 3 を ready-show contract に採用すると、pending fallback handoff が 80ms 遅れ、trace / completion details に docs truth と異なる attempt 3 が入り得ます。
- Phase 3 freeze line は `2 attempts` を既に変更禁止契約として固定しており、これを `3 attempts` に変更するより、runtime / composition / tests を `2 attempts` に揃えるほうが R06/R07/R08 の ownership 境界を明確にできます。

B は採用しません。現行 runtime は確かに `3 attempts` ですが、`3` が ready-show 専用 contract として設計された根拠は確認できません。

C は採用しません。`2` と `3` は別レイヤ値として整理できる状態ではなく、現行 runtime では `3` が ready-show max attempts として実際に効いています。

D は終了します。追加 inventory は有用ですが、R06 前提判断としては A を採用するだけの材料があります。

## 10. R06 再開可否

R06 は、A 採用後もまだ再開しません。

R06 ready-show retry scheduler ownership を安全に分離する前に、次を満たす必要があります。

- production runtime が `2 attempts` に揃っている。
- ready-show retry の tests が `attempt 1 -> 80ms attempt 2 -> pending fallback` を固定している。
- freeze line と runtime が一致している。
- pending retry `400ms / 3 attempts` は別責務として維持されている。
- trace 名 / trace 意味が不変である。
- callback 意味が不変である。
- completion 条件が不変である。
- display session boundary が不変である。
- R16 `TaskPaneRetryTimerLifecycle` と R06 ready-show retry scheduler ownership を混ぜていない。

## 次に必要な作業

1. implementation alignment safe unit として、ready-show retry coordinator の runtime max attempts を `2` に揃える。
2. `PendingPaneRefreshMaxAttempts = 3` を pending retry 専用値として扱い、ready-show retry max attempts へ流用しない。
3. tests で `attempt 1 -> 80ms attempt 2 -> pending fallback`、attempt 3 不在、pending retry `400ms / 3 attempts` 維持を固定する。
4. trace / callback / completion / display session boundary が変わっていないことを確認する。
5. 上記完了後に R06 ready-show retry scheduler ownership 分離を再開する。

## Runtime alignment 完了メモ

2026-05-09 の runtime alignment safe unit で、ready-show retry coordinator の production composition は `ReadyShowMaxAttempts = 2` を使う形に揃えました。

この alignment の範囲:

- ready-show retry max attempts は `2`。
- ready-show retry delay は `80ms`。
- production runtime の ready-show sequence は `attempt 1 -> 80ms attempt 2 -> pending fallback`。
- `PendingPaneRefreshMaxAttempts = 3` は pending retry `400ms / 3 attempts` 専用値として維持する。
- `PendingPaneRefreshMaxAttempts` を ready-show retry coordinator へ渡さない。
- attempt 3 は ready-show path に存在しないことを tests で固定する。

この alignment では、R06 ownership 分離、callback 意味、completion 条件、display session boundary、trace 名、trace 意味、R16 timer lifecycle ownership は変更しません。
