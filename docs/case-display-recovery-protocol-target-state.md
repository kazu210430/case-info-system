# CASE Display Recovery Protocol Target State

## 位置づけ

この文書は、`docs/case-display-recovery-protocol-current-state.md` を前提に、CASE display / recovery protocol の target-state を定義するための docs-only 設計記録です。

- 基準コード:
  - `2026-05-08` 時点で `main` と `origin/main` の一致を確認した `e41feb5d607f79077e112a1945e81ac0a76d95a4`
  - foreground guarantee ownership target-state 追記時点の `main` / `origin/main`: `3d6f2441f84dfefe46393508d4eae02ebe06b886`
  - visibility recovery ownership target-state 追記時点の `main` / `origin/main`: `79c4823537c881b81582d3456145f8fc5f09466f`
  - rebuild fallback ownership target-state 追記時点の `main` / `origin/main`: `ca23a651a2c811eb19f81ade2348277af19fa0c3`
  - refresh source ownership target-state 追記時点の `main` / `origin/main`: `b9f0ab8b1534b083160a4c709e1cf33c753975a3`
- 参照正本:
  - `AGENTS.md`
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/readyshow-recovery-observation-points-2026-05-08.md`
  - `docs/codex-review-first-safe-unit-plan.md`

この文書は、retry 回数、visibility 制御、hidden session、foreground 条件、rebuild fallback 条件を変えるものではありません。今回は owner と completion definition を明確にすることだけを目的とします。

## この文書で固定すること

1. CASE display / recovery protocol を構成する protocol unit の target-state 定義。
2. `CASE display completed` を誰が定義するべきか。
3. `workbook open completed`、`display handoff completed`、`pane visible`、`refresh completed`、`foreground guarantee completed`、`CASE display completed` の関係。
4. ready-show request / enqueue / attempt / early-complete の owner 境界。
5. foreground guarantee、visibility recovery、rebuild fallback、refresh source、WindowActivate の責任境界。
6. 最初に実装すべき安全単位候補。

この文書で固定しないこと:

- 実機不安定の根本原因断定
- retry 値や timing の変更
- recovery 実行条件の変更
- hidden session 修正
- CASE 作成本体修正
- build / test / DeployDebugAddIn

## Target-State Design Principles

- `WorkbookOpen` は window 安定境界ではない。
- `CASE display completed` は success-only の protocol terminal state とする。
- `pane visible`、`refresh completed`、`foreground guarantee completed` は別概念として維持する。
- `CASE display completed` を lower-level service が勝手に確定しない。
- decision owner と execution owner を分けてもよいが、protocol unit の completion owner は 1 箇所に置く。
- `refresh source` と `snapshot source` を同じ意味で扱わない。
- `rebuild fallback` は refresh / snapshot subprotocol に留め、display completion 条件へ昇格させない。
- current-state の fail-closed、window recovery、ready-show retry、host reuse、post-release protection、white Excel 回避を落とさない。

## Protocol Unit Definitions

| Protocol unit | 定義 | Target-state owner | CASE display completed との関係 |
| --- | --- | --- | --- |
| `workbook open completed` | shared/current app 上で CASE workbook の open / reopen が完了し、一時的な application state も restore 済みで、presentation 側へ引き渡せる状態。 | `CaseWorkbookOpenStrategy` | upstream 前提段階。completion 条件そのものには含めない。 |
| `display handoff completed` | CASE 表示要求が ready-show orchestration に受理され、created-case display session の owner が presentation 層から display protocol 層へ移った状態。 | `TaskPaneRefreshOrchestrationService` | upstream 前提段階。session start 条件。 |
| `pane visible` | 対象 workbook / window に対して visible な CASE pane が成立している状態。新規 show でも retained host の visible 維持でもよい。 | `TaskPaneHostFlowService` が visible state owner。`WorkbookTaskPaneReadyShowAttemptWorker` は already-visible を観測するだけ。 | 必要条件。 |
| `refresh completed` | refresh path が実行され、pane refresh unit が terminal になった状態。already-visible path では発生しない。 | `TaskPaneRefreshCoordinator` | 補助条件。pane visible に到達する 1 経路だが、必要条件ではない。 |
| `foreground guarantee completed` | 同一 display session に対して foreground obligation が残っていない状態。`Required` なら recovery 実行後、`NotRequired` なら skip 判定の terminal 化までを含む。 | decision / outcome / trace owner は `TaskPaneRefreshOrchestrationService`。execution primitive owner は `ExcelWindowRecoveryService`。 | 条件付き必要条件。`Required` の場合だけ必要。 |
| `CASE display completed` | created-case display session が success として閉じ、target pane visible が成立し、foreground obligation も terminal になった状態。 | `TaskPaneRefreshOrchestrationService`。将来専用 service を切る場合も、この orchestration 境界に置く。 | 最終 terminal state。 |

## CASE Display Completed Definition

target-state では、`CASE display completed` を次のように定義する。

1. success-only であること
2. 同一 created-case display session に対する completion であること
3. `pane visible` が成立していること
4. foreground obligation が terminal であること
5. worker / coordinator / host-flow の個別成功を、そのまま completion と見なさないこと

逆に、次は `CASE display completed` に含めない。

- wait UI が閉じたこと
- 初期カーソル移動が完了したこと
- `refresh completed`
- `snapshot source` が `CaseCache` / `BaseCache` / `MasterListRebuild` のどれだったか
- `rebuild fallback` を使ったかどうか
- protection start が呼ばれたこと

## 必要条件と補助条件

### 前提段階

- `workbook open completed`
- `display handoff completed`

これらは `CASE display completed` より前に通るべき upstream stage であり、最終 completion 条件そのものではない。

### 必要条件

- `pane visible`

`pane visible` は hard requirement とする。`CASE display completed` を `pane visible` なしに成立させない。

### 条件付き必要条件

- `foreground guarantee completed`

ただしこれは「必ず foreground recovery を実行する」という意味ではない。`TaskPaneRefreshOrchestrationService` が `NotRequired` と判断した場合も、その判断自体を terminal 化してから completion へ進む。

### 補助条件

- `refresh completed`

`refresh completed` は refresh path でだけ通る補助条件とする。already-visible path では `refresh completed` なしで completion できる。

## Owner Boundary Decisions

### 1. CASE display completed definition owner

`CASE display completed` の canonical owner は `TaskPaneRefreshOrchestrationService` とする。

理由:

- worker は 1 attempt の owner であり、fallback handoff や後続 foreground obligation を知らない。
- coordinator は refresh unit と foreground execution bridge の owner だが、already-visible path を知らない。
- host-flow service は `pane visible` state owner だが、created-case display session 全体の terminal owner ではない。
- orchestration service だけが request / enqueue / retry / attempt result / fallback / final completion を同一 session で束ねられる。

したがって target-state では、worker と coordinator は completion そのものを emit せず、normalized outcome を orchestration service へ返す構造に寄せる。

### 2. ready-show request / enqueue / attempt / early-complete

- `ready-show request`
  - owner は `KernelCasePresentationService`
  - created-case post-release の presentation owner として request を発行する
- `ready-show enqueue`
  - owner は `TaskPaneRefreshOrchestrationService`
  - queue 受理と display session 作成を担当する
- `ready-show attempt`
  - owner は `WorkbookTaskPaneReadyShowAttemptWorker`
  - 1 attempt の window resolve、already-visible 確認、refresh delegate 呼出しを担当する
- `early-complete`
  - already-visible の検出 owner は `WorkbookTaskPaneReadyShowAttemptWorker`
  - それを `CASE display completed` とみなす semantic owner は `TaskPaneRefreshOrchestrationService`

つまり target-state では、worker は `visible already satisfied` を返せても、`case-display-completed` を自分で確定しない。

### 3. foreground guarantee の decision / execution

- foreground guarantee protocol unit の decision owner は `TaskPaneRefreshOrchestrationService`
- foreground guarantee protocol unit の outcome / completion trace owner も `TaskPaneRefreshOrchestrationService`
- `TaskPaneRefreshCoordinator` は refresh raw result と foreground execution bridge を返す側に寄せる。
- 実際の app/window/foreground recovery primitive の execution owner は `ExcelWindowRecoveryService`

この関係は「decision / outcome / emit と execution primitive を分け、protocol unit completion owner は orchestration に固定する」と読む。

`ExcelWindowRecoveryService` は execution result を返すが、`CASE display completed` の可否や foreground obligation の skip / required 判定までは持たない。

### 4. visibility recovery の orchestration owner

target-state では visibility recovery を次の 2 層で固定する。

- pre-handoff presentation preparation
  - `KernelCasePresentationService`
  - hidden reopen 後の最初の見せ方準備だけを持つ
- post-handoff display / recovery orchestration
  - `TaskPaneRefreshOrchestrationService`
  - ready-show 以降に visibility primitive をいつ使うかを調停する

primitive owner 自体は維持する。

- lightweight workbook visible ensure:
  - `WorkbookWindowVisibilityService`
- full application/window/foreground recovery:
  - `ExcelWindowRecoveryService`

つまり target-state で整理したいのは primitive の統合ではなく、「誰が protocol 上でそれらを呼ぶ順序責任を持つか」である。

### 5. rebuild fallback の protocol 上の位置

`rebuild fallback` は CASE display protocol の top-level unit ではなく、refresh path の内部にある snapshot acquisition subprotocol として扱う。

- owner:
  - `TaskPaneSnapshotBuilderService`
- protocol position:
  - `TaskPaneRefreshCoordinator` 配下の refresh render path の内部
- 扱い:
  - `refresh completed` の内部要因にはなりうる
  - `CASE display completed` の直接条件にはしない

これにより `rebuild fallback` を ready-show owner や foreground owner へ誤って昇格させない。

### 6. refresh source owner

`refresh source` は raw string `reason` の別名にしてはいけない。target-state では、structured request の field として upstream で一度だけ確定する。

source setter は entry ごとに分ける。

- `KernelCasePresentationService`
  - created-case post-release ready-show request
- `WorkbookLifecycleCoordinator`
  - `WorkbookOpen` / `WorkbookActivate`
- `WindowActivatePaneHandlingService`
  - `WindowActivate`
- action 系 caller
  - post-action refresh

downstream の扱いは次のとおり。

- `TaskPaneRefreshOrchestrationService`
  - source を session に保持し、そのまま下流へ渡す
- `TaskPaneRefreshCoordinator`
  - source の consumer であり、source owner ではない
- `TaskPaneSnapshotBuilderService`
  - 独立した `snapshot source` owner のまま維持する

要するに、target-state の `refresh source owner` は coordinator ではなく request creation boundary である。

### 7. WindowActivate event capture と refresh dispatch の境界

- Excel event capture / observation owner:
  - `ThisAddIn`
  - `WorkbookEventCoordinator`
- `WindowActivate` 特有の protection / suppression / dispatch owner:
  - `WindowActivatePaneHandlingService`
- display protocol owner:
  - `TaskPaneRefreshOrchestrationService`

target-state でも event capture と dispatch は分けてよい。ただし capture 側に refresh decision や source 再解釈を持ち込まない。

### 8. display handoff completion owner

`display handoff completed` の owner は `TaskPaneRefreshOrchestrationService` とする。

target-state の役割分担:

- `CaseWorkbookOpenStrategy`
  - `workbook open completed`
- `KernelCasePresentationService`
  - initial preparation
  - `ready-show request`
- `TaskPaneRefreshOrchestrationService`
  - request accepted
  - `ready-show enqueue`
  - `display handoff completed`

この整理により、current-state の `display-handoff-open-completed` 二重観測を解消する。

## Target-State Flow

1. `CaseWorkbookOpenStrategy` が `workbook open completed` を成立させる。
2. `KernelCasePresentationService` が pre-handoff preparation を終え、created-case display request を出す。
3. `TaskPaneRefreshOrchestrationService` が request を enqueue し、display session を作成して `display handoff completed` を成立させる。
4. `WorkbookTaskPaneReadyShowAttemptWorker` が ready-show attempt を実行する。
5. already-visible なら worker は `pane visible already satisfied` を返す。
6. refresh が必要なら `TaskPaneRefreshCoordinator` が refresh path を実行する。
7. `TaskPaneHostFlowService` が show / reuse により `pane visible` state を成立させる。
8. `TaskPaneRefreshOrchestrationService` が foreground unit を `Completed` / `NotRequired` / `RequiredDegraded` などの outcome として閉じる。
9. `TaskPaneRefreshOrchestrationService` が同一 session の `pane visible` と foreground terminal を確認して、`CASE display completed` を 1 回だけ成立させる。

## First Safe Implementation Unit

最初の安全単位は、「`CASE display completed` の ownership を orchestration 境界へ寄せる」ことに限定する。

### Implementation status (2026-05-08 first safe unit)

- Implemented: created-case display session in `TaskPaneRefreshOrchestrationService`.
- Implemented: `case-display-completed` is emitted only by `TaskPaneRefreshOrchestrationService`.
- Implemented: worker / coordinator / host-flow now return normalized outcomes instead of owning final display completion.
- Implemented: `display-handoff-completed` is emitted at ready-show acceptance in `TaskPaneRefreshOrchestrationService`.
- Preserved: retry counts, ready-show timing, foreground and visibility recovery conditions, rebuild fallback, hidden session behavior, CASE creation behavior, and fail-closed conditions.

### Completion record (2026-05-08, merged main `e41feb5d607f79077e112a1945e81ac0a76d95a4`)

第1実装安全単位の目的は、CASE display / recovery protocol 全体を作り替えることではなく、`CASE display completed` definition と `case-display-completed` emit owner を 1 箇所に固定することでした。

current-state で見えていた ownership 分裂:

- already-visible path では `WorkbookTaskPaneReadyShowAttemptWorker` が completion 相当の判断を持っていた。
- refresh path では `TaskPaneRefreshCoordinator` が refresh / foreground 後の completion 相当の判断を持っていた。
- `TaskPaneHostFlowService` は `pane visible` を成立させるが、display session 全体の terminal owner ではなかった。
- foreground guarantee は `TaskPaneRefreshCoordinator` の decision / completion と `ExcelWindowRecoveryService` の execution に分かれていた。
- display handoff completion trace は presentation / open-strategy 側にもあり、ready-show acceptance との境界が観測上重なっていた。

target-state で固定した completion definition:

- `CASE display completed` は success-only の terminal state とする。
- 同一 created-case display session に対して成立する。
- `pane visible` が成立している。
- foreground obligation が terminal である。
- `refresh completed` は補助条件であり、already-visible path では必須ではない。
- `pane visible`、`refresh completed`、`foreground guarantee completed` のいずれか単独を `CASE display completed` の別名にしない。

completion owner を `TaskPaneRefreshOrchestrationService` に置いた理由:

- ready-show request acceptance、queue、retry / fallback handoff、attempt outcome、refresh outcome を同一 session で束ねられる。
- already-visible path と refresh path の両方を見られる境界である。
- Worker / Coordinator / HostFlowService の lower-level 成功を final completion と誤認しない境界である。
- hidden protocol や CASE 作成本体へ ownership を戻さず、display protocol の orchestration 層で閉じられる。

各 service に残した責務:

- `WorkbookTaskPaneReadyShowAttemptWorker`: 1 attempt の window resolve、already-visible 検知、refresh delegate 呼び出し、attempt outcome 返却。
- `TaskPaneRefreshCoordinator`: refresh unit、foreground execution bridge、refresh attempt raw result 返却。
- `TaskPaneHostFlowService`: host reuse / render / show と `pane visible` state の返却。
- `ExcelWindowRecoveryService`: application / workbook window / foreground recovery primitive の execution。
- `TaskPaneSnapshotBuilderService`: CASE cache / Base cache / MasterListRebuild を含む snapshot source と rebuild fallback。

created-case display session を導入した理由:

- CASE display completion を単発ログではなく、created-case post-release の表示要求に紐づく protocol session として扱うため。
- `pane visible`、`refresh completed`、`foreground guarantee terminal` を同じ session の材料として束ねるため。
- completion emit の重複を防ぎ、`case-display-completed` を 1 session につき 1 回だけ成立させるため。

already-visible path と refresh path を収束させた理由:

- already-visible は `refresh completed` を持たないが、CASE display completion としては `pane visible` と foreground terminal が満たされれば success にできる。
- refresh path は refresh / render / foreground を経由するが、最終的には同じ created-case display session の terminal 判定へ戻す必要がある。
- path ごとに completion owner を分けると、同じ CASE 表示完了が別 semantic になり、重複 emit や観測ずれを再発させる。

維持したもの:

- retry 回数
- ready-show timing
- foreground guarantee 条件
- visibility recovery 条件
- rebuild fallback 条件
- hidden session behavior
- CASE 作成本体
- fail closed

実機確認結果:

- 新規CASE作成は正常。
- `created-case-display-session-started -> display-handoff-completed -> case-display-completed` は 1 回だけ出る。
- already-visible path / refresh path は同じ completion definition へ収束する。
- 既存CASE reopen は正常。
- 白Excelなし。
- ぐるぐるなし。
- 雛形更新後の新規CASE作成も体感改善済み。

今回まだ触っていない ownership:

- foreground guarantee owner
- visibility recovery owner
- rebuild fallback owner
- refresh source owner
- WindowActivate ownership

次の安全単位候補:

- foreground guarantee を 1 protocol unit として、decision / execution / terminal trace の境界を固定する。
- visibility recovery を lightweight workbook visible ensure と full application/window/foreground recovery に分けて、caller orchestration owner を固定する。
- refresh source を `reason` の再掲ではなく structured source として request boundary で固定する。
- rebuild fallback を refresh / snapshot subprotocol の観測 unit として整理し、display completion 条件へ昇格させないことを明文化する。
- `WindowActivate` event capture と refresh dispatch の ownership を、`ThisAddIn` / `WorkbookEventCoordinator` / `WindowActivatePaneHandlingService` / orchestration service の境界で整理する。

### この安全単位に含めたこと

- `TaskPaneRefreshOrchestrationService` に created-case display session を導入する
- `WorkbookTaskPaneReadyShowAttemptWorker` から `case-display-completed` emit を外し、attempt outcome を返す
- `TaskPaneRefreshCoordinator` から `case-display-completed` emit を外し、refresh / foreground outcome を返す
- `CaseWorkbookOpenStrategy` 側の completion trace 名を `workbook open completed` 系へ寄せる
- `display handoff completed` を enqueue acceptance 側へ一本化する

### この安全単位に含めなかったこと

- retry 回数 / delay の変更
- visibility primitive の呼出し順変更
- foreground recovery 条件変更
- `WorkbookOpen -> WorkbookActivate -> WindowActivate` 境界変更
- rebuild fallback 条件変更
- hidden session / CASE 作成本体修正

### この安全単位を最初に選んだ理由

- current-state で最も分散している completion definition だけを切り出せる
- worker / coordinator / host-flow の責務そのものは変えず、final completion owner だけを寄せられる
- retry、visibility、foreground の危険領域へ直ちに踏み込まずに済む
- current-state 正本が指摘している最重要 owner 不明箇所にそのまま対応できる

## foreground guarantee ownership target-state (2026-05-08)

### target-state summary

この節は、`docs/case-display-recovery-protocol-current-state.md` の `foreground guarantee ownership current-state (2026-05-08 docs-only)` を受けて、次の実装安全単位へ進むための target-state を固定する。

- 第1実装安全単位では、foreground guarantee の decision / outcome / `final-foreground-guarantee-*` trace owner を `TaskPaneRefreshOrchestrationService` 側へ寄せる。
- build / test は `build.ps1` を入口に確認する。`DeployDebugAddIn` は実機確認が必要な場合の別入口として扱う。
- retry 条件、ready-show timing、visibility recovery 条件、foreground recovery 条件、rebuild fallback 条件、`WindowActivate` 挙動は変更しない。
- service 分割、helper 切り出し、ガード追加による上書きは行わない。
- `CASE display completed` の emit owner は引き続き `TaskPaneRefreshOrchestrationService` とする。
- `foreground guarantee` は CASE display completion の材料であり、CASE display completion そのものではない。

target-state では、foreground guarantee を次の 3 層に分けて扱う。

| 層 | owner | target-state の責務 |
| --- | --- | --- |
| decision / outcome / trace | `TaskPaneRefreshOrchestrationService` | refresh attempt の事実から foreground obligation を判定し、必要なら execution primitive を呼び、normalized outcome と `final-foreground-guarantee-*` trace を確定する。 |
| execution bridge | `TaskPaneRefreshCoordinator` | refresh path の raw result を返し、既存の `ExcelWindowRecoveryService` 呼び出しと post-guarantee protection を実行する。foreground outcome は確定しない。 |
| execution primitive | `ExcelWindowRecoveryService` | workbook window / application の recovery、activation、foreground promotion を実行し、実行結果だけを返す。 |
| display-session consumption | `TaskPaneRefreshOrchestrationService` | foreground outcome を同一 created-case display session の材料として消費し、`case-display-completed` を success-only で emit するか判断する。 |

### foreground guarantee completed definition

`foreground guarantee completed` は、「foreground recovery が成功した」という単一意味ではなく、同一 refresh attempt または created-case display session に対する foreground obligation が terminal になった状態を指す。

`foreground guarantee completed` は次のいずれかで成立する。

| outcome | completed | display-completable | 定義 |
| --- | --- | --- | --- |
| `NotRequired` | yes | yes | refresh path ではない、または foreground recovery を要求する条件が成立していない。 |
| `SkippedAlreadyVisible` | yes | yes | already-visible path で pane visible が成立しており、final foreground execution を要求しない。 |
| `SkippedNoKnownTarget` | yes | no | workbook / window / active fallback の target が protocol 上確定できず、推測で補完しない。 |
| `RequiredSucceeded` | yes | yes | `TaskPaneRefreshOrchestrationService` が foreground recovery required と判定し、`ExcelWindowRecoveryService` の execution が成功した。 |
| `RequiredDegraded` | yes | yes, but degraded | execution は走ったが、OS foreground promotion や recovery primitive が完全成功を返さない。ただし対象 workbook / window と pane visible は維持され、追加 retry / timing hack へ進まない。 |
| `RequiredFailed` | yes | no | foreground recovery required だが、target mismatch、例外、execution 不成立などで foreground obligation を display-completable と扱えない。 |
| `Unknown` | no | no | owner が outcome を正規化できていない。target-state では fail-closed とし、success completion に使わない。 |

この定義では、`completed` と `success` を分ける。

- `completed`
  - protocol 上、その foreground obligation に対してこれ以上同じ owner が処理を続けない terminal 状態。
- `display-completable`
  - `TaskPaneRefreshOrchestrationService` が `CASE display completed` の success-only 判定材料として消費できる状態。
- `degraded`
  - foreground execution が best-effort に留まったことを観測できる状態。`RequiredDegraded` は `case-display-completed` を即座に禁止する failure ではないが、成功と同じ意味に丸めない。
- `failed`
  - success-only の `case-display-completed` に使ってはいけない状態。新しい retry や fallback を足して覆わず、既存 flow の範囲で fail closed に扱う。

### owner boundary

#### foreground guarantee emit / decision owner

- `TaskPaneRefreshOrchestrationService`
  - foreground guarantee の decision owner。
  - refresh attempt / created-case display session 内の foreground guarantee outcome owner。
  - `foreground-recovery-decision` と `final-foreground-guarantee-started` / `final-foreground-guarantee-completed` 相当の trace owner。
  - `ExcelWindowRecoveryService` の execution result を normalized foreground outcome へ変換する owner。
- `TaskPaneRefreshCoordinator`
  - refresh unit の owner。
  - foreground recovery execution bridge と post-guarantee protection owner。
  - raw execution result を返し、foreground outcome / `case-display-completed` を確定しない。
- `ExcelWindowRecoveryService`
  - execution primitive owner。
  - `Required` / `NotRequired` / `Skipped` の protocol 判定は持たない。
  - `CASE display completed` の emit 可否を判断しない。

#### retry / recovery / fallback との境界

- ready retry `80ms` と pending retry `400ms` は `TaskPaneRefreshOrchestrationService` / `PendingPaneRefreshRetryService` の既存責務に残す。
- foreground guarantee outcome を理由に retry 回数や delay を変更しない。
- lightweight visibility ensure は `WorkbookWindowVisibilityService` の責務に残す。
- full window recovery / foreground promotion は `ExcelWindowRecoveryService` の責務に残す。
- rebuild fallback は `TaskPaneSnapshotBuilderService` の snapshot acquisition subprotocol に残し、foreground guarantee の owner や completion 条件へ昇格しない。

#### WindowActivate との境界

- `WindowActivate` は window が activate された観測点であり、foreground guarantee completed の代替ではない。
- event capture は `ThisAddIn` / `WorkbookEventCoordinator` に残す。
- `WindowActivate` 特有の suppression / protection / dispatch は `WindowActivatePaneHandlingService` に残す。
- `WindowActivate` 発火だけを `RequiredSucceeded` とみなさない。
- `WindowActivate` dispatch から refresh path に入り、`TaskPaneRefreshOrchestrationService` が foreground outcome を返した場合だけ foreground guarantee の protocol outcome として扱う。

#### ready-show / visibility recovery / rebuild fallback との境界

- `WorkbookTaskPaneReadyShowAttemptWorker`
  - ready-show attempt、window resolve、already-visible detection、refresh delegate 呼び出しを担当する。
  - already-visible path では `SkippedAlreadyVisible` 相当の foreground obligation を outcome として返せる。
  - `final-foreground-guarantee-*` trace や `case-display-completed` は emit しない。
- `WorkbookWindowVisibilityService`
  - workbook window visible ensure だけを返す。
  - foreground guarantee completed を返さない。
- `TaskPaneSnapshotBuilderService`
  - `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` の snapshot source decision owner。
  - rebuild fallback の発生有無を foreground outcome に混ぜない。

### allowed / forbidden responsibilities

#### allowed

- `TaskPaneRefreshOrchestrationService` が foreground requirement を判定し、`TaskPaneRefreshCoordinator` の execution bridge 経由で `ExcelWindowRecoveryService` を呼び、normalized `ForegroundGuaranteeOutcome` を返す。
- `ExcelWindowRecoveryService` が app / workbook window / Win32 foreground primitive を実行し、実行結果だけを返す。
- `TaskPaneRefreshOrchestrationService` が `pane visible` と foreground outcome を同一 created-case display session で突き合わせる。
- `WorkbookTaskPaneReadyShowAttemptWorker` が already-visible path を success material として返す。
- `TaskPaneHostFlowService` が host reuse / render / show と `pane visible` source を返す。
- `TaskPaneSnapshotBuilderService` が snapshot source と rebuild fallback を観測可能にする。

#### forbidden

- lower-level service が `case-display-completed` を emit する。
- `ExcelWindowRecoveryService` が `CASE display completed` の可否を判断する。
- `WindowActivate` 発火だけで foreground guarantee completed とみなす。
- rebuild fallback の発生有無を foreground guarantee success / failure の条件にする。
- foreground outcome の不足を新しい guard で覆う。
- active workbook fallback を target 不明時の推測 promotion として広げる。
- retry 回数、delay、foreground recovery 条件、visibility recovery 条件、rebuild fallback 条件をこの owner 整理で変更する。
- `Application.DoEvents()`、sleep、単なる timing delay で completed を作る。
- hidden session / retained hidden app-cache / COM lifetime を foreground owner 整理のために広げる。

### normalized outcome design

次の実装安全単位では、まず outcome を明示するだけに留め、挙動条件は変えない。

#### ForegroundGuaranteeOutcome

target-state の normalized outcome は少なくとも次を持つ。

| field | 意味 |
| --- | --- |
| `Status` | `NotRequired` / `SkippedAlreadyVisible` / `SkippedNoKnownTarget` / `RequiredSucceeded` / `RequiredDegraded` / `RequiredFailed` / `Unknown` |
| `WasRequired` | foreground recovery execution が protocol 上必要だったか。 |
| `WasExecutionAttempted` | `ExcelWindowRecoveryService` の foreground primitive を呼んだか。 |
| `IsTerminal` | foreground obligation が terminal か。`Unknown` は false。 |
| `IsDisplayCompletable` | `CASE display completed` の材料として使えるか。 |
| `TargetKind` | `ExplicitWorkbookWindow` / `ActiveWorkbookFallback` / `AlreadyVisible` / `NoKnownTarget` など。 |
| `RecoverySucceeded` | execution result。not-required / skipped では null を許容する。 |
| `Reason` | skip / degraded / failed の事実ベース理由。 |

#### Worker / Coordinator / HostFlowService

- `WorkbookTaskPaneReadyShowAttemptWorker`
  - `IsPaneVisible`
  - `PaneVisibleSource`
  - `RefreshAttempted`
  - `RefreshResult`
  - `ForegroundGuaranteeOutcome`
  - already-visible path では `SkippedAlreadyVisible` を返す。
- `TaskPaneRefreshCoordinator`
  - `IsPaneVisible`
  - `IsRefreshCompleted`
  - `RefreshSource`
  - `WindowTarget`
  - foreground decision / completion owner ではなく、raw foreground execution facts を上位へ返す。
- `TaskPaneHostFlowService`
  - `IsPaneVisible`
  - `PaneVisibleSource`
  - `HostReused`
  - `Rendered`
  - foreground outcome と CASE display completion は持たない。

#### TaskPaneRefreshOrchestrationService

`TaskPaneRefreshOrchestrationService` は次だけを判断する。

1. outcome が同一 created-case display session に属するか。
2. `pane visible` が成立しているか。
3. `ForegroundGuaranteeOutcome.IsTerminal` が true か。
4. `ForegroundGuaranteeOutcome.IsDisplayCompletable` が true か。
5. 既存 retry / fallback が outstanding ではないか。

この条件を満たした場合だけ `case-display-completed` を emit する。`RequiredFailed`、`SkippedNoKnownTarget`、`Unknown` は success-only completion に使わない。

### CASE display completed との関係

`CASE display completed` は引き続き `TaskPaneRefreshOrchestrationService` の created-case display session terminal state とする。

- `foreground guarantee completed` は `CASE display completed` の条件付き材料である。
- `foreground guarantee completed` は `pane visible` の代替ではない。
- `foreground guarantee completed` は `refresh completed` の別名ではない。
- `RequiredDegraded` は display-completable だが degraded として観測可能に残す。
- `RequiredFailed` は display-completable ではないため、success-only の `case-display-completed` に使わない。
- already-visible path は final foreground execution を要求しないが、`SkippedAlreadyVisible` として terminal 化し、pane visible と合わせて completion 材料にできる。
- rebuild fallback は refresh / snapshot 内部の source decision であり、CASE display completion の直接条件にしない。

この関係により、`CASE display completed` の owner を増やさず、foreground guarantee の結果だけを normalized input として扱う。

### constraints to preserve

- 白Excel対策を落とさない。
  - `PostCloseFollowUpScheduler` の no visible workbook quit と foreground recovery を混同しない。
  - foreground outcome 整理を白Excel対策ガードの追加で覆わない。
- COM解放を落とさない。
  - hidden create session / retained hidden app-cache / temporary workbook close の cleanup 境界を変えない。
  - foreground outcome のために workbook / window / application COM reference lifetime を広げない。
- Excel状態制御を落とさない。
  - `ScreenUpdating` / `DisplayAlerts` / `EnableEvents` は既存 scope で restore する。
  - `ExcelWindowRecoveryService` の state restore は recovery primitive として扱い、恒常状態変更にしない。
- fail closed を維持する。
  - context / workbook / window が不明な場合に推測で補完しない。
  - `Unknown` を success に丸めない。
- timing hack に逃げない。
  - `Application.DoEvents()`、sleep、単なる delay 追加は禁止する。
- ガード追加で覆わない。
  - foreground / visibility / rebuild fallback 条件を新しい guard で隠さない。
  - `WorkbookOpen` を window 安定境界へ戻さない。

### 次の実装安全単位候補

1. `ForegroundGuaranteeOutcome` の taxonomy を result 型に追加し、ログ / trace 上で `NotRequired`、`SkippedAlreadyVisible`、`RequiredSucceeded`、`RequiredDegraded`、`RequiredFailed`、`Unknown` を観測できるようにする。挙動条件は変えない。
2. `TaskPaneRefreshOrchestrationService` の foreground decision / execution result 変換を 1 箇所に寄せ、`ExcelWindowRecoveryService` の `recovered=false` を `RequiredDegraded` として返す。retry / recovery 条件は変えない。
3. `WorkbookTaskPaneReadyShowAttemptWorker` の already-visible path を `SkippedAlreadyVisible` として orchestration へ渡す。`case-display-completed` emit owner は増やさない。
4. `TaskPaneRefreshOrchestrationService` が `IsDisplayCompletable` を見て success-only completion を判断する。
5. `WindowActivate` は event capture / dispatch / refresh request の境界整理だけを行い、foreground completed の代替にしない。

## visibility recovery ownership target-state (2026-05-08 docs-only)

### target-state summary

この節は、`docs/case-display-recovery-protocol-current-state.md` の `visibility recovery ownership current-state (2026-05-08 docs-only)` を受けて、次の実装安全単位へ進むための target-state を固定する。

- 調査開始時の `main` / `origin/main`: `79c4823537c881b81582d3456145f8fc5f09466f`
- 参照した docs:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/case-display-recovery-protocol-target-state.md`
- 今回は docs-only であり、コード変更、service 分割、helper 切り出し、ready retry / pending retry / visibility recovery / foreground guarantee / rebuild fallback / `WindowActivate` 条件変更は行わない。
- docs-only のため build / test / `DeployDebugAddIn` は実行しない。

target-state では、visibility recovery を「workbook window を visible にしたか」だけではなく、created-case display session に対して `pane visible` へ到達できたかを正規化する protocol unit として扱う。ただし primitive owner は統合しない。

| 層 | owner | target-state の責務 |
| --- | --- | --- |
| protocol decision / normalized outcome / completion trace | `TaskPaneRefreshOrchestrationService` | ready-show / refresh / pending handoff の結果から `VisibilityRecoveryOutcome` を確定し、`CASE display completed` の材料として消費する。 |
| ready-show attempt facts | `WorkbookTaskPaneReadyShowAttemptWorker` | 1 attempt の window resolve、already-visible 判定、lightweight ensure の local result、refresh delegate result を返す。protocol outcome は確定しない。 |
| refresh raw facts | `TaskPaneRefreshCoordinator` | refresh path の window target、pre-context recovery local result、refresh completion、host-flow result を返す。visibility completion は確定しない。 |
| pane visible state | `TaskPaneHostFlowService` | host reuse / render / show による actual pane visible transition を返す。created-case display session の terminal owner ではない。 |
| lightweight primitive | `WorkbookWindowVisibilityService` | workbook window visible ensure の local outcome を返す。CASE pane visible や display completion は返さない。 |
| full recovery primitive | `ExcelWindowRecoveryService` | application / workbook window recovery、minimized restore、必要時の activation / foreground primitive を実行し、execution facts を返す。protocol outcome は確定しない。 |

### visibility recovery completed definition

`visibility recovery completed` は、同一 created-case display session の visibility obligation が terminal になり、canonical な `pane visible` が成立した状態を指す。workbook window の `Visible=true`、`WindowActivate` 発火、または `ExcelWindowRecoveryService` の bool 成功だけでは completed としない。

target-state の `pane visible` は、次のいずれかで成立する。

- already-visible path:
  - 対象 window key が解決できる。
  - host registry に対象 window key の host がある。
  - host の `WorkbookFullName` が対象 workbook と一致する。
  - host role が `WorkbookRole.Case` である。
  - host の `CustomTaskPane.Visible` が true である。
- refresh path:
  - `TaskPaneHostFlowService` が `ReusedShown` または `RefreshedShown` 相当の visible result を返す。
  - その result が対象 workbook / window / CASE role と矛盾しない。

この定義では、`workbook window visible` は `pane visible` の前提になりうるが十分条件ではない。`WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible` / `MadeVisible` は workbook window の local outcome であり、CASE pane visible とは同義にしない。

### completed / skipped / degraded / failed

`VisibilityRecoveryOutcome` は、少なくとも次の status を持つ。

| status | terminal | display-completable | 定義 |
| --- | --- | --- | --- |
| `Completed` | yes | yes | visibility recovery が必要または実行され、canonical `pane visible` が成立した。 |
| `Skipped` | yes | conditional | recovery primitive を実行しなかった。`AlreadyVisible` で canonical `pane visible` が true の場合だけ display-completable。`WorkbookOpenWindowUnstable`、`Suppressed`、`NoCreatedCaseSession`、`NoKnownTarget` などは success completion に使わない。 |
| `Degraded` | yes | yes, but degraded | recovery primitive が partial failure / unverifiable state を返したが、canonical `pane visible` は成立している。成功に丸めず trace に残す。 |
| `Failed` | yes | no | target workbook / window / context が不明、target mismatch、例外、または allowed recovery / show attempt 後も canonical `pane visible` が成立しない。 |
| `Unknown` | no | no | owner が outcome を正規化できていない。fail-closed とし、success completion に使わない。 |

`Degraded` と `Failed` の境界は `pane visible` で切る。`VisibleAfterSet=false/null` や full recovery の `recovered=false` があっても、その後の host reuse / render / show で canonical `pane visible` が成立していれば `Degraded` として扱える。`pane visible` が成立しない場合は `Failed` であり、追加 guard、sleep、DoEvents、rebuild fallback で覆わない。

### owner boundary

#### visibility 判定 owner

- canonical `pane visible` の事実 owner は `TaskPaneDisplayCoordinator` / `TaskPaneHostFlowService` に置く。
- `WorkbookTaskPaneReadyShowAttemptWorker` は already-visible path で判定を呼び、結果を observation として返す。
- `TaskPaneRefreshOrchestrationService` はその observation を同一 created-case display session の `VisibilityRecoveryOutcome` へ正規化する。
- `WorkbookWindowVisibilityService` と `ExcelWindowRecoveryService` は workbook / application / window の local recovery owner であり、CASE pane visible 判定 owner ではない。

#### recovery trigger owner

- pre-handoff presentation preparation は `KernelCasePresentationService` に残す。
  - hidden reopen 後の initial visibility ensure、initial full recovery without showing、ready-show request 前 visibility ensure を扱う。
  - ここでは created-case display session の normalized `VisibilityRecoveryOutcome` を確定しない。
- post-handoff display / recovery orchestration は `TaskPaneRefreshOrchestrationService` に寄せる。
  - ready-show attempt、refresh handoff、pending retry handoff の結果を同一 session で束ねる。
  - normalized outcome と `visibility-recovery-*` protocol trace の owner になる。

#### retry / pending retry との境界

- ready retry `80ms` は `TaskPaneRefreshOrchestrationService.ScheduleTaskPaneReadyRetry(...)` の既存責務に残す。
- pending retry `400ms` は `PendingPaneRefreshRetryService` の既存責務に残す。
- `VisibilityRecoveryOutcome.Failed` を理由に retry 回数、delay、ready 条件、pending 条件を変更しない。
- pending retry が active CASE context fallback に入る場合、outcome には `TargetKind=ActiveWorkbookFallback` などを残す。explicit workbook target と同一視できない場合は fail-closed とする。

#### foreground guarantee との境界

- visibility recovery は `pane visible` へ到達できたかを扱う。
- foreground guarantee は pane visible / refresh 後に foreground obligation が terminal かを扱う。
- `ExcelWindowRecoveryService` は両方で使われうる primitive だが、visibility outcome と `ForegroundGuaranteeOutcome` は別 result とする。
- `Window.Visible = true`、`workbook.Activate()`、`WindowActivate` 発火は foreground guarantee completed の代替ではない。
- `CASE display completed` は visibility outcome と foreground outcome の両方を `TaskPaneRefreshOrchestrationService` が消費して判断する。

#### rebuild fallback との境界

- rebuild fallback の owner は引き続き `TaskPaneSnapshotBuilderService`。
- visibility recovery 失敗は即 rebuild fallback ではない。
- rebuild fallback は refresh path が context 解決、refresh precondition、render path へ進んだ後の snapshot acquisition subprotocol でだけ発生する。
- visibility failure により window resolve / context 解決 / refresh precondition が fail-closed で止まる場合、rebuild fallback までは到達しない。

#### WindowActivate との境界

- `WindowActivate` は visible window が activate された観測点、または refresh request source の入口であり、visibility recovery completed の代替ではない。
- event capture は `ThisAddIn` / `WorkbookEventCoordinator` に残す。
- protection / suppression / dispatch は `WindowActivatePaneHandlingService` に残す。
- refresh / visibility / foreground outcome は `TaskPaneRefreshOrchestrationService` が判断する。
- protection / suppression により dispatch が return する場合、event 発火だけで visibility recovery terminal とは扱わない。

### allowed / forbidden responsibilities

#### allowed

- `TaskPaneRefreshOrchestrationService` が lower-level result を同一 created-case display session に束ね、`VisibilityRecoveryOutcome` を確定する。
- `WorkbookTaskPaneReadyShowAttemptWorker` が already-visible 判定、attempt 1 の lightweight ensure、refresh delegate 呼び出しを行い、facts を返す。
- `TaskPaneRefreshCoordinator` が refresh pre-context recovery と refresh raw result を返す。
- `TaskPaneHostFlowService` が `ReusedShown` / `RefreshedShown` として pane visible source を返す。
- `WorkbookWindowVisibilityService` が workbook window visible ensure の local outcome を返す。
- `ExcelWindowRecoveryService` が app / window recovery primitive を実行し、execution facts を返す。
- `TaskPaneSnapshotBuilderService` が snapshot source と rebuild fallback を観測可能にする。

#### forbidden

- lower-level service が `case-display-completed` を emit する。
- `WorkbookWindowVisibilityService.EnsureVisible(...)` の local outcome をそのまま `visibility recovery completed` とみなす。
- `ExcelWindowRecoveryService` が `CASE display completed` や protocol-level visibility outcome を判断する。
- `WindowActivate` 発火だけで `Completed` / `RequiredSucceeded` / `CASE display completed` とみなす。
- visibility recovery failure を rebuild fallback へ直接接続する。
- `Degraded` を retry / fallback / timing hack で覆う。
- ready retry、pending retry、visibility recovery、foreground guarantee、rebuild fallback、`WindowActivate` 条件を owner 整理の名目で変更する。
- `Application.DoEvents()`、sleep、単なる delay 追加で completed を作る。
- hidden session / retained hidden app-cache / COM lifetime を visibility outcome のために広げる。
- `WorkbookOpen` を window 安定境界へ戻す。

### normalized outcome design

#### VisibilityRecoveryOutcome

target-state の normalized outcome は少なくとも次を持つ。

| field | 意味 |
| --- | --- |
| `Status` | `Completed` / `Skipped` / `Degraded` / `Failed` / `Unknown` |
| `Reason` | `AlreadyVisible`、`MadeVisibleThenShown`、`RefreshShown`、`WorkbookOpenWindowUnstable`、`NoKnownTarget`、`TargetMismatch`、`Exception` などの事実ベース理由。 |
| `IsTerminal` | visibility obligation が terminal か。`Unknown` は false。 |
| `IsPaneVisible` | canonical `pane visible` が成立しているか。 |
| `IsDisplayCompletable` | `CASE display completed` の success-only 材料として使えるか。 |
| `TargetKind` | `ExplicitWorkbookWindow` / `ActiveWorkbookFallback` / `AlreadyVisible` / `NoKnownTarget` など。 |
| `PaneVisibleSource` | `AlreadyVisibleHost` / `ReusedShown` / `RefreshedShown` / `None` など。 |
| `WorkbookWindowEnsureStatus` | `WorkbookWindowVisibilityService` の local status。null 可。 |
| `FullRecoveryAttempted` | `ExcelWindowRecoveryService` の full recovery primitive を呼んだか。 |
| `FullRecoverySucceeded` | full recovery execution result。not attempted では null を許容する。 |
| `DegradedReason` | partial failure / unverifiable state を success と区別して残す理由。 |

#### Worker / Coordinator / HostFlowService

- `WorkbookTaskPaneReadyShowAttemptWorker`
  - `IsPaneVisible`
  - `PaneVisibleSource`
  - `WindowTarget`
  - `WorkbookWindowEnsureStatus`
  - `RefreshAttempted`
  - `RefreshResult`
  - `FailureReason`
  - protocol-level `VisibilityRecoveryOutcome` と `case-display-completed` は持たない。
- `TaskPaneRefreshCoordinator`
  - `IsRefreshCompleted`
  - `WindowTarget`
  - `RefreshPrecondition`
  - `PreContextRecoveryFacts`
  - `HostFlowResult`
  - `RawFailureReason`
  - normalized visibility outcome / completion trace は確定しない。
- `TaskPaneHostFlowService`
  - `IsPaneVisible`
  - `PaneVisibleSource`
  - `HostReused`
  - `Rendered`
  - `WorkbookFullName`
  - `WindowKey`
  - visibility recovery status、foreground outcome、CASE display completion は持たない。

#### TaskPaneRefreshOrchestrationService

`TaskPaneRefreshOrchestrationService` は次だけを判断する。

1. lower-level result が同一 created-case display session に属するか。
2. canonical `pane visible` が成立しているか。
3. recovery primitive の local failure を `Degraded` として残すか、`Failed` として fail-closed にするか。
4. `VisibilityRecoveryOutcome.IsTerminal` が true か。
5. `VisibilityRecoveryOutcome.IsDisplayCompletable` が true か。
6. foreground obligation が terminal か。
7. 既存 retry / fallback が outstanding ではないか。

この条件を満たした場合だけ、foreground outcome と合わせて `case-display-completed` を success-only で emit する。`Failed`、`Unknown`、および `Skipped` でも `IsPaneVisible=false` のものは completion に使わない。

### trace emit / completion emit の責務分離

- primitive trace
  - `WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService` は local mutation / execution trace を維持してよい。
- pane visible trace
  - `TaskPaneHostFlowService` は `taskpane-reused-shown` / `taskpane-refreshed-shown` 相当の actual show trace を維持する。
- protocol trace
  - `TaskPaneRefreshOrchestrationService` が `visibility-recovery-decision` と `visibility-recovery-completed` / `visibility-recovery-skipped` / `visibility-recovery-degraded` / `visibility-recovery-failed` 相当の normalized trace owner になる。
- completion emit
  - `case-display-completed` は引き続き `TaskPaneRefreshOrchestrationService` だけが emit する。

### rebuild fallback への接続条件

visibility recovery と rebuild fallback の接続は、次で固定する。

- visibility recovery 失敗は即 rebuild fallback ではない。
- rebuild fallback は `TaskPaneSnapshotBuilderService` が refresh / render / snapshot acquisition 内で `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` を選ぶ場合だけ成立する。
- `Completed` / `Reason=AlreadyVisible` の `Skipped` / `Degraded` は rebuild fallback を要求しない。
- `Failed` は rebuild fallback を要求しない。既存 ready retry / pending retry / fail-closed の範囲で扱う。
- refresh path に到達し、snapshot acquisition が `MasterListRebuild` を選んだ結果 pane visible が成立した場合、visibility outcome は `Completed` または `Degraded` になりうる。ただし fallback 使用有無は visibility status ではなく snapshot source として残す。
- fallback 判断 owner は `TaskPaneSnapshotBuilderService` であり、`TaskPaneRefreshOrchestrationService` は fallback 使用有無を `CASE display completed` の直接条件にしない。

### CASE display completed との関係

`CASE display completed` は引き続き `TaskPaneRefreshOrchestrationService` の created-case display session terminal state とする。

- visibility recovery outcome は `CASE display completed` の材料であり、completion そのものではない。
- `Completed`、`Reason=AlreadyVisible` かつ `IsPaneVisible=true` の `Skipped`、`Degraded` は display-completable になりうる。
- `Failed`、`Unknown`、`IsPaneVisible=false` の `Skipped` は success-only completion に使わない。
- `CASE display completed` は `VisibilityRecoveryOutcome.IsPaneVisible` と `ForegroundGuaranteeOutcome.IsTerminal / IsDisplayCompletable` を両方確認する。
- workbook window visibility ensure の local outcome を直接 completion 条件にしない。
- rebuild fallback の発生有無を直接の completion 条件にしない。

### constraints to preserve

- 白Excel対策を落とさない。
  - `PostCloseFollowUpScheduler` の no visible workbook quit と visibility recovery を混同しない。
  - 白 Excel を覆うだけの guard を追加しない。
- TaskPane が出ない regression を防ぐ。
  - ready-show、already-visible early-complete、pending retry、host reuse / render / show の現行条件を変更しない。
- COM解放を落とさない。
  - hidden create session、retained hidden app-cache、一時 workbook close の cleanup 境界を変えない。
  - visibility outcome のために workbook / window / application COM reference lifetime を広げない。
- Excel状態制御を落とさない。
  - `ScreenUpdating` / `DisplayAlerts` / `EnableEvents` の既存 restore scope を変えない。
  - `ExcelWindowRecoveryService.EnsureScreenUpdatingEnabled(...)` を恒常状態変更として扱わない。
- fail closed を維持する。
  - workbook / window / context が不明な場合に推測で補完しない。
  - active workbook fallback を target 不明時の広域 promotion として拡大しない。
- timing hack に逃げない。
  - `Application.DoEvents()`、sleep、単なる delay 追加は禁止する。
  - ready retry `80ms`、pending retry `400ms`、attempt count は今回変更しない。
- ガード追加で覆わない。
  - visibility / foreground / rebuild fallback 条件を新しい guard で隠さない。
  - `WorkbookOpen` を window 安定境界へ戻さない。

### 次の実装安全単位候補

1. `VisibilityRecoveryOutcome` の taxonomy を orchestration 層に追加し、lower-level result から `Completed` / `Skipped` / `Degraded` / `Failed` / `Unknown` を観測できるようにする。挙動条件は変えない。
2. `WorkbookTaskPaneReadyShowAttemptWorker` の already-visible / lightweight ensure / refresh delegate result を structured facts として返し、`case-display-completed` emit owner は増やさない。
3. `TaskPaneRefreshCoordinator` と `TaskPaneHostFlowService` の pane visible facts を orchestration が消費できる形に整理する。host reuse / render / show 条件は変えない。
4. `TaskPaneRefreshOrchestrationService` が `VisibilityRecoveryOutcome` と `ForegroundGuaranteeOutcome` を同一 created-case display session で突き合わせる。success-only completion の意味は変えない。
5. `visibility-recovery-*` normalized trace を orchestration へ寄せる。primitive trace と pane shown trace は既存 owner に残す。
6. rebuild fallback は `TaskPaneSnapshotBuilderService` の snapshot source decision として残し、visibility failure から直接起動しないことを test / trace 上でも確認する。

## rebuild fallback ownership target-state (2026-05-08 docs-only)

### target-state summary

この節は、`docs/case-display-recovery-protocol-current-state.md` の rebuild fallback current-state を受けて、次の実装安全単位へ進むための target-state を固定する。

- 参照した docs:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/taskpane-refresh-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/case-display-recovery-protocol-target-state.md`
- 今回は docs-only であり、コード変更、service 分割、helper 切り出し、ready retry / pending retry / visibility recovery / foreground guarantee / rebuild fallback / `WindowActivate` 条件変更は行わない。
- docs-only のため build / test / `DeployDebugAddIn` は実行しない。

target-state では、rebuild fallback を TaskPane refresh render path 内の `snapshot acquisition subprotocol` として扱う。`MasterListRebuild` を選んだ事実が rebuild fallback required の中心であり、visibility recovery failure、foreground guarantee failure、ready retry exhausted、pending retry exhausted をそのまま rebuild fallback required とは呼ばない。

`BaseCacheFallback` は cache source decision の fallback ではあるが、この節でいう rebuild fallback required ではない。`BaseCacheFallback` が snapshot を供給できる場合は、`MasterListRebuild` を要求しないため rebuild fallback は `Skipped` として扱う。

### rebuild fallback required definition

`rebuild fallback required` は、次をすべて満たす場合にだけ成立する。

1. refresh path が fail-closed せず、対象 workbook / window / context が refresh render path へ到達している。
2. host reuse / already-visible path ではなく、render が必要である。
3. `TaskPaneSnapshotBuilderService` の snapshot source decision が `CaseCache` / `BaseCache` / `BaseCacheFallback` で表示用 snapshot を供給できない。
4. `TaskPaneSnapshotBuilderService` が `MasterListRebuild` を選んだ、または選ぶべき facts に到達した。

required の判定材料は、snapshot acquisition の raw facts に限定する。

| fact / outcome | required 扱い | 理由 |
| --- | --- | --- |
| `CaseCache` usable | no | CASE cache で snapshot acquisition が完了する。 |
| `BaseCache` usable | no | Base cache を CASE cache へ昇格して snapshot acquisition が完了する。 |
| `BaseCacheFallback` selected | no | latest master version が読めない場合でも Base snapshot が使えるため、`MasterListRebuild` は要求しない。 |
| `CaseCacheIncompatible` / `CaseCacheStale` | maybe | 単独では required ではない。Base cache / Base fallback でも満たせない場合だけ `MasterListRebuild` required の reason になる。 |
| `BaseCacheIncompatible` / `BaseCacheStale` | maybe | CASE cache も使えず、Base cache も使えない場合に `MasterListRebuild` required の reason になる。 |
| `CacheUnavailable` | yes, if render path reached | render path で使える cache source がないため `MasterListRebuild` required。 |
| `MasterListRebuild` selected | yes | 実装上の rebuild fallback required 相当の canonical fact。 |

visibility recovery failure との接続は直接ではない。

- visibility recovery failure により window resolve / context resolve / refresh precondition が fail-closed で止まる場合、snapshot acquisition へ到達しないため rebuild fallback は required にならない。
- visibility recovery が `Completed` / `Degraded` で render path へ到達し、その中で `MasterListRebuild` が選ばれた場合だけ rebuild fallback required になる。
- rebuild fallback は workbook window visible 化、Excel foreground promotion、白 Excel 回復の代替ではない。

ready retry / pending retry exhausted との接続も直接ではない。

- ready-show attempt が尽きた場合の handoff 先は `ScheduleWorkbookTaskPaneRefresh(...)` / pending retry であり、即 `MasterListRebuild` ではない。
- pending retry が対象 workbook または active CASE context fallback から refresh path に入り、render が必要になった場合だけ snapshot acquisition へ到達する。
- ready retry exhausted / pending retry exhausted だけを reason にして rebuild fallback 条件を広げない。

### rebuild fallback completed definition

`rebuild fallback completed` は、required だった `MasterListRebuild` path が terminal になり、refresh render path が次の host show evaluation へ進める snapshot result を得た状態を指す。これは `pane visible`、`foreground guarantee completed`、`CASE display completed` と同義ではない。

`RebuildFallbackOutcome` は少なくとも次の status を持つ。

| status | terminal | refresh can continue | 定義 |
| --- | --- | --- | --- |
| `Skipped` | yes | conditional | rebuild fallback が required ではなかった。already-visible、host reuse、`CaseCache` / `BaseCache` / `BaseCacheFallback` 採用、または upstream fail-closed で snapshot acquisition に到達しない場合。 |
| `Completed` | yes | yes | `MasterListRebuild` が正常に snapshot text を構築し、render path が通常 snapshot として扱える。 |
| `Degraded` | yes | yes, but degraded | `MasterListRebuild` は required で、表示可能な fallback/error snapshot などを返せたが、CASE cache 更新失敗、Master 読み取りの一部不確実性、または diagnostic error を伴う。成功に丸めず trace に残す。 |
| `Failed` | yes | no | `MasterListRebuild` が required だったが、render path が扱える snapshot result を得られない、または target / root / master path / owned read access が fail-closed になった。 |
| `Unknown` | no | no | owner が raw facts を正規化できていない。success completion に使わず fail-closed とする。 |

fallback 後に再評価するものは、通常 refresh path と同じである。

1. `TaskPaneHostFlowService` が render result を使って host show / reuse 結果を評価する。
2. `TaskPaneRefreshOrchestrationService` が pane visible facts を visibility recovery outcome に正規化する。
3. `TaskPaneRefreshOrchestrationService` が foreground guarantee outcome を評価する。
4. `TaskPaneRefreshOrchestrationService` が同一 created-case display session の `case-display-completed` 可否を判断する。

`Completed` / `Degraded` は refresh render path が続行できることを示すだけで、CASE display completion を直接成立させない。`Failed` は snapshot acquisition subprotocol の failure であり、追加 guard、sleep、DoEvents、visibility recovery の再解釈で覆わない。

### owner boundary

#### fallback decision owner

- `TaskPaneSnapshotBuilderService` が snapshot source decision と rebuild fallback required / skipped の raw decision owner である。
- decision の入力は CASE cache facts、Base cache facts、latest master version availability、format compatibility、stale 判定、cache availability に限定する。
- `TaskPaneRefreshOrchestrationService`、`WorkbookTaskPaneReadyShowAttemptWorker`、`TaskPaneRefreshCoordinator` は rebuild fallback required 条件を再実装しない。

#### fallback execution owner

- `TaskPaneSnapshotBuilderService` が `MasterListRebuild` execution owner である。
- Master workbook の read-only open / close / owned cleanup 境界は `MasterWorkbookReadAccessService` に残す。
- `readAccess.CloseIfOwned()` の `finally` 境界、COM release、hidden session / retained hidden app-cache の cleanup 境界を rebuild fallback target-state で広げない。

#### fallback result normalization owner

- `TaskPaneSnapshotBuilderService` は `SnapshotSource`、fallback reason list、`MasterListRebuildUsed`、`UpdatedCaseSnapshotCache`、snapshot text availability、raw failure facts を返す。
- `TaskPaneRefreshCoordinator` / `TaskPaneHostFlowService` はその raw facts を refresh / render result に含めて上位へ渡す。
- `TaskPaneRefreshOrchestrationService` が created-case display session 上の `RebuildFallbackOutcome` へ正規化する。
- 正規化は観測と completion 消費のためであり、fallback 条件、retry 条件、host show 条件を変えるためのものではない。

#### trace emit owner

- primitive / diagnostic trace は `TaskPaneSnapshotBuilderService` と `MasterWorkbookReadAccessService` に残してよい。
- normalized protocol trace は `TaskPaneRefreshOrchestrationService` が `rebuild-fallback-required` / `rebuild-fallback-skipped` / `rebuild-fallback-completed` / `rebuild-fallback-degraded` / `rebuild-fallback-failed` 相当として emit する。
- trace owner を増やして `case-display-completed` の重複 emit を再導入しない。

#### CASE display completed owner との境界

- `TaskPaneSnapshotBuilderService` は `case-display-completed` を emit しない。
- `TaskPaneHostFlowService` は pane visible facts を返すが、created-case display session の terminal owner ではない。
- `TaskPaneRefreshOrchestrationService` だけが visibility outcome、foreground outcome、rebuild fallback outcome を同一 session の facts として消費し、success-only の `case-display-completed` を判断する。
- rebuild fallback の使用有無は completion の直接条件ではない。

### allowed / forbidden responsibilities

#### allowed

- `TaskPaneSnapshotBuilderService` が snapshot source decision と `MasterListRebuild` execution facts を返す。
- `MasterWorkbookReadAccessService` が Master read-only access と owned cleanup を閉じる。
- `TaskPaneHostFlowService` が render result 後の `ReusedShown` / `RefreshedShown` / show failure facts を返す。
- `TaskPaneRefreshCoordinator` が refresh raw result に snapshot acquisition facts を含める。
- `TaskPaneRefreshOrchestrationService` が lower-level facts を `RebuildFallbackOutcome` に正規化し、normalized trace を emit する。
- `TaskPaneRefreshOrchestrationService` が fallback outcome を display completion の diagnostic fact として記録する。

#### forbidden

- visibility recovery failure を rebuild fallback required に直結する。
- foreground guarantee failure を rebuild fallback required に直結する。
- ready retry exhausted / pending retry exhausted だけを reason に `MasterListRebuild` を要求する。
- host reuse / already-visible path を無理に render path へ流す。
- `BaseCacheFallback` を `MasterListRebuild` required と混同する。
- lower-level service が `case-display-completed` を emit する。
- `TaskPaneSnapshotBuilderService` に visibility recovery、foreground guarantee、CASE display completion を判断させる。
- fallback の成功率を上げる目的で `Application.DoEvents()`、sleep、単なる delay、timing hack を追加する。
- fallback 条件、ready retry 条件、pending retry 条件、visibility recovery 条件、foreground guarantee 条件、`WindowActivate` 挙動を owner 整理の名目で変更する。
- context-less fallback open、暗黙の workbook 推測、Master path 推測を追加する。
- COM lifetime、`ScreenUpdating` / `DisplayAlerts` / `EnableEvents` restore scope、hidden session cleanup 境界を広げる。

### normalized outcome design

#### RebuildFallbackOutcome

target-state の normalized outcome は少なくとも次を持つ。

| field | 意味 |
| --- | --- |
| `Status` | `Skipped` / `Completed` / `Degraded` / `Failed` / `Unknown` |
| `IsRequired` | `MasterListRebuild` が必要だったか。 |
| `IsTerminal` | rebuild fallback subprotocol が terminal か。`Unknown` は false。 |
| `CanContinueRefresh` | render / host show evaluation へ進める snapshot result があるか。 |
| `SnapshotSource` | `CaseCache` / `BaseCache` / `BaseCacheFallback` / `MasterListRebuild` / `None` |
| `FallbackReasons` | `CaseCacheIncompatible`、`CaseCacheStale`、`BaseCacheIncompatible`、`BaseCacheStale`、`LatestMasterVersionUnavailable`、`CacheUnavailable` などの diagnostic reason list。 |
| `MasterListRebuildAttempted` | MasterList rebuild を実行したか。 |
| `MasterListRebuildSucceeded` | MasterList rebuild が通常 snapshot を返せたか。 |
| `SnapshotTextAvailable` | render path が扱える snapshot text があるか。 |
| `UpdatedCaseSnapshotCache` | CASE cache を更新したか。display completion 条件ではない。 |
| `FailureReason` | `NoSystemRoot`、`MasterPathUnavailable`、`MasterReadFailed`、`SnapshotBuildException`、`NoSnapshotText` など。 |
| `DegradedReason` | cache update failure、error snapshot fallback、partial/unverifiable source など。 |

#### Worker / Coordinator / HostFlowService / SnapshotBuilder

- `WorkbookTaskPaneReadyShowAttemptWorker`
  - `RefreshAttempted`
  - `RefreshResult`
  - `IsPaneVisible`
  - `PaneVisibleSource`
  - `FailureReason`
  - rebuild fallback required / completed は判断しない。
- `TaskPaneRefreshCoordinator`
  - `IsRefreshCompleted`
  - `WindowTarget`
  - `RefreshPrecondition`
  - `ContextAccepted`
  - `HostFlowResult`
  - `SnapshotAcquisitionFacts`
  - normalized `RebuildFallbackOutcome` と completion trace は確定しない。
- `TaskPaneHostFlowService`
  - `HostReused`
  - `Rendered`
  - `IsPaneVisible`
  - `PaneVisibleSource`
  - `RenderResult`
  - `SnapshotAcquisitionFacts`
  - visibility recovery status、foreground outcome、CASE display completion は持たない。
- `TaskPaneSnapshotBuilderService`
  - `SnapshotSource`
  - `FallbackReasons`
  - `MasterListRebuildUsed`
  - `UpdatedCaseSnapshotCache`
  - `SnapshotTextAvailable`
  - `BuildFailureReason`
  - host show、foreground、CASE display completion は持たない。

#### TaskPaneRefreshOrchestrationService

`TaskPaneRefreshOrchestrationService` は次だけを判断する。

1. lower-level facts が同一 created-case display session に属するか。
2. `RebuildFallbackOutcome` が `Skipped` / `Completed` / `Degraded` / `Failed` / `Unknown` のどれか。
3. rebuild fallback failure が refresh continuation を止める raw failure か。
4. `pane visible` が成立しているか。
5. visibility outcome が display-completable か。
6. foreground outcome が terminal / display-completable か。
7. success-only の `case-display-completed` を emit してよいか。

`RebuildFallbackOutcome.Completed` / `Degraded` だけでは completion しない。`Failed` / `Unknown` は success completion に使わない。`Skipped` は fallback が不要だったという診断であり、pane visible / visibility / foreground が別途成立していれば completion を妨げない。

### visibility recovery / foreground guarantee との接続

#### visibility recovery

- visibility recovery failure は rebuild fallback required ではない。
- visibility recovery が fail-closed で refresh path に到達しない場合、rebuild fallback は `Skipped` または `Unknown` として扱い、`MasterListRebuild` を起動しない。
- rebuild fallback が `Completed` / `Degraded` になった後も、`pane visible` は `TaskPaneHostFlowService` の facts で再評価する。
- rebuild fallback は workbook window visible ensure の代替ではない。

#### foreground guarantee

- foreground guarantee は pane visible / refresh 後の foreground obligation を閉じる protocol unit であり、snapshot acquisition の fallback ではない。
- rebuild fallback outcome を foreground guarantee success / failure の条件にしない。
- `MasterListRebuild` が成功しても foreground guarantee required が消えるわけではない。
- foreground guarantee が `RequiredFailed` / `RequiredDegraded` になっても、それだけで rebuild fallback を要求しない。

### CASE display completed との関係

`CASE display completed` は引き続き `TaskPaneRefreshOrchestrationService` の created-case display session terminal state とする。

- rebuild fallback outcome は diagnostic / refresh continuation fact であり、completion の直接条件ではない。
- `Completed` / `Degraded` は、通常 refresh path と同じく host show、visibility outcome、foreground outcome の評価へ戻る。
- `Skipped` は fallback 不要を表すだけで、completion 成否は pane visible / visibility / foreground で判断する。
- `Failed` / `Unknown` は fail-closed とし、snapshot acquisition failure を visibility recovery や foreground guarantee で覆って success completion にしない。
- `case-display-completed` は lower-level service から emit しない。

### constraints to preserve

- 白Excel対策を落とさない。
  - rebuild fallback を白 Excel 対策 guard として扱わない。
  - `PostCloseFollowUpScheduler` の no visible workbook quit と snapshot acquisition fallback を混同しない。
- TaskPane が出ない regression を防ぐ。
  - ready-show、already-visible early-complete、pending retry、host reuse / render / show の現行条件を変更しない。
  - fallback 条件を広げすぎない。
- COM解放を落とさない。
  - `MasterWorkbookReadAccessService` の owned read access cleanup、hidden create session、retained hidden app-cache、一時 workbook close の cleanup 境界を変えない。
- Excel状態制御を落とさない。
  - `ScreenUpdating` / `DisplayAlerts` / `EnableEvents` の既存 restore scope を変えない。
  - Master read access を表示制御へ昇格しない。
- fail closed を維持する。
  - workbook / context / `SYSTEM_ROOT` / Master path が不明な場合に推測で補完しない。
- timing hack に逃げない。
  - `Application.DoEvents()`、sleep、単なる delay 追加は禁止する。
  - ready retry `80ms`、pending retry `400ms`、attempt count は今回変更しない。
- ガード追加で覆わない。
  - visibility / foreground / rebuild fallback 条件を新しい guard で隠さない。
  - `WorkbookOpen` を window 安定境界へ戻さない。

### 次の実装安全単位候補

1. `TaskPaneSnapshotBuilderService` の snapshot source decision を raw facts として返す。`SnapshotSource`、fallback reason list、`MasterListRebuildUsed`、`UpdatedCaseSnapshotCache`、failure / degraded reason を含める。挙動条件は変えない。
2. `CasePaneSnapshotRenderService` / `TaskPaneHostFlowService` / `TaskPaneRefreshCoordinator` が snapshot acquisition facts を上位へ運ぶ。host reuse / render / show 条件は変えない。
3. `TaskPaneRefreshOrchestrationService` に `RebuildFallbackOutcome` normalization と normalized trace emit を追加する。`case-display-completed` emit owner は増やさない。
4. `RebuildFallbackOutcome` を visibility outcome / foreground outcome と同一 created-case display session で突き合わせる。ただし completion 条件に rebuild fallback 使用有無を直接足さない。
5. tests / trace では、visibility recovery failure、foreground guarantee failure、ready retry exhausted、pending retry exhausted がそれ単独では rebuild fallback required にならないことを確認する。
6. `BaseCacheFallback` と `MasterListRebuild` を protocol outcome 上で区別し、fallback 条件を広げないことを確認する。

## refresh source ownership target-state (2026-05-08 docs-only)

### target-state summary

この節は、`docs/case-display-recovery-protocol-current-state.md` の refresh source ownership current-state を受けて、次の実装安全単位へ進むための target-state を固定する。

参照した正本:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/case-display-recovery-protocol-current-state.md`
- `docs/case-display-recovery-protocol-target-state.md`

今回は docs-only であり、コード変更、service 分割、helper 切り出し、source 採用順序変更、cache 条件変更、rebuild fallback 条件変更、visibility / foreground / `WindowActivate` 条件変更は行わない。docs-only のため build / test / `DeployDebugAddIn` は実行しない。

target-state では、`refresh source` を raw string `reason` の別名にしない。TaskPane refresh には少なくとも次の 4 種の source-like field があるため、同じ `source` という名前で混ぜない。

| 種別 | 定義 | owner | snapshot source selection との関係 |
| --- | --- | --- | --- |
| trigger reason | refresh が要求された契機や診断文字列。`WorkbookActivate`、`WindowActivate`、`CreatedCaseReadyShow`、`RibbonCasePaneRefresh` など。 | refresh request を作る entry service / caller。 | 採用 source ではない。render path に到達するまで snapshot source は決まらない。 |
| display request source | `TaskPaneDisplayRequest.Source` が表す structured な表示要求元。`WindowActivate`、`PostActionRefresh.<actionKind>` など。 | request creator と `TaskPaneRefreshOrchestrationService`。 | display entry / show / hide / reject の前段情報であり、cache 採用順序を決めない。 |
| snapshot source | CASE pane render が必要になったとき、表示用 snapshot text をどこから採用したか。 | `TaskPaneSnapshotBuilderService`。 | `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` の順でだけ決まる。 |
| log component source | trace / log を emit した component 名や logger 文脈。 | emit した service。 | protocol decision ではない。`refreshSource=(reason)` のような既存 log 名は normalized source として読まない。 |

この節で正本化する `refresh source ownership` の中心は、上表のうち `snapshot source` と、それを上位 protocol outcome へ正規化する境界である。trigger reason や display request source は保持してよいが、`CaseCache` / `BaseCache` / `BaseCacheFallback` / `MasterListRebuild` の採用判断とは分ける。

### source 種別定義

| source | target-state の定義 | selection reason | degraded / fallback / rebuild required との関係 |
| --- | --- | --- | --- |
| `CaseCache` | CASE workbook の `TASKPANE_SNAPSHOT_CACHE_*` が存在し、format compatible で、version 条件を満たすため、CASE cache を表示用 snapshot として採用した状態。 | `CaseCacheUsable`。CASE cache が表示中 Pane と整合する最優先 source として使える。 | degraded ではない。cache fallback ではない。rebuild fallback required ではない。 |
| `BaseCache` | CASE cache が使えず、Base 埋込 `TASKPANE_BASE_SNAPSHOT_*` が存在し、format compatible で、version 条件を満たすため、Base cache を CASE cache へ昇格して採用した状態。 | `BaseCachePromoted`。Base 埋込 snapshot が表示用 source として使える。 | degraded ではない。cache fallback ではない。rebuild fallback required ではない。 |
| `BaseCacheFallback` | Base snapshot が存在し format compatible だが latest master version を読めないため、`LatestMasterVersionUnavailable` を diagnostic reason として持ったまま Base snapshot を採用した状態。 | `LatestMasterVersionUnavailable`。latest version 不明でも Base snapshot 自体は表示用 source として使える。 | cache fallback である。`SelectionQuality=Fallback` または `DegradedSelected` として観測してよいが、`MasterListRebuild` required ではない。 |
| `MasterListRebuild` | CASE cache / Base cache / Base fallback のいずれも表示用 snapshot を供給できず、Master path を解決して read-only open し、`雛形一覧` から snapshot text を再構築する状態。 | `CacheUnavailable`、`CaseCacheIncompatible`、`CaseCacheStale`、`BaseCacheIncompatible`、`BaseCacheStale` など、cache source で完了できない reasons。 | rebuild fallback required の canonical source。rebuild outcome は `Completed` / `Degraded` / `Failed` に正規化する。 |
| workbook snapshot | CASE / Base workbook の DocProperty chunk、master version、format compatibility、latest master version availability などの raw facts。 | `CaseWorkbookSnapshotFacts` / `BaseWorkbookSnapshotFacts`。採用前の材料であり、採用 source そのものではない。 | degraded / fallback / rebuild required を直接意味しない。`TaskPaneSnapshotBuilderService` がこの raw facts を評価して上記 source に変換する。 |
| `None` | already-visible、host reuse、display entry reject、precondition skip、context reject、non-CASE render、または source selection failure で snapshot source を持てない状態。 | `SnapshotAcquisitionNotReached`、`ContextRejected`、`WorkbookMissing`、`SelectionFailed` など。 | `None` だけで rebuild required とは扱わない。selection failure で renderable snapshot がない場合は terminal failure として fail closed する。 |

`workbook snapshot` は source 採用前の raw material であり、normalized `SelectedSource` にはしない。`CaseCache` / `BaseCache` / `BaseCacheFallback` は workbook snapshot facts から採用された source であり、`MasterListRebuild` は workbook snapshot facts では満たせなかった場合の rebuild source である。

### owner boundary

#### source 採用判断 owner

- `TaskPaneSnapshotBuilderService` が `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` の採用順序と採用可否を判断する。
- `TaskPaneRefreshOrchestrationService`、`WorkbookTaskPaneReadyShowAttemptWorker`、`TaskPaneRefreshCoordinator`、`TaskPaneHostFlowService` はこの順序を再実装しない。
- host reuse / already-visible / precondition skip / context reject は snapshot source selection の前段であり、source 採用失敗として扱わない。

#### raw facts owner

- CASE / Base workbook snapshot facts は `TaskPaneSnapshotBuilderService` が読み取り、`TaskPaneSnapshotCacheService` など既存 cache helper の責務境界を変えない。
- Master path 解決、read-only open、owned workbook close、COM cleanup は `MasterWorkbookReadAccessService` に残す。
- `TaskPaneBuildResult`、`TaskPaneHostFlowResult`、`TaskPaneRefreshAttemptResult` は raw facts を上位へ運んでよいが、採用順序を再判定しない。

#### source normalization owner

- `TaskPaneRefreshOrchestrationService` が lower-level raw facts を `RefreshSourceSelectionOutcome` として正規化する。
- normalization は selected source、selection reason、degraded / fallback / rebuild required flags、terminal status を観測可能にするためのものであり、cache 条件や rebuild 条件を変えるためのものではない。
- `TaskPaneRefreshCoordinator` の既存 `refreshSource=(reason)` は target-state では `TriggerReason` または diagnostic reason として扱い、normalized source とは呼ばない。

#### trace emit owner

- raw diagnostic trace は `TaskPaneSnapshotBuilderService`、`MasterWorkbookReadAccessService`、`TaskPaneRefreshCoordinator` に残してよい。
- normalized protocol trace は `TaskPaneRefreshOrchestrationService` が emit する。
- normalized trace は `refresh-source-selected` / `refresh-source-skipped` / `refresh-source-degraded` / `refresh-source-failed` 相当として扱い、`case-display-completed` の emit owner を増やさない。
- log component source は emit 元の説明であり、snapshot source や trigger reason の代替名にしない。

#### rebuild fallback owner との境界

- `MasterListRebuild` を選ぶ raw decision は `TaskPaneSnapshotBuilderService` の snapshot source decision に含まれる。
- rebuild fallback required / completed / degraded / failed の normalized outcome は、既存の rebuild fallback target-state どおり `TaskPaneRefreshOrchestrationService` が `RebuildFallbackOutcome` として正規化する。
- `RefreshSourceSelectionOutcome` は `RebuildFallbackOutcome` を参照または内包してよいが、rebuild fallback 条件を再判定しない。

### allowed / forbidden responsibilities

#### allowed

- entry service / caller が trigger reason と display request source を設定する。
- `TaskPaneRefreshOrchestrationService` が trigger reason / display request source を session に保持し、downstream へ渡す。
- `TaskPaneSnapshotBuilderService` が workbook snapshot facts を評価し、snapshot source と selection reason を返す。
- `TaskPaneSnapshotBuilderService` が `BaseCacheFallback` と `MasterListRebuild` を別 source として返す。
- `TaskPaneHostFlowService` / `TaskPaneRefreshCoordinator` が snapshot source facts を result に含めて上位へ伝播する。
- `TaskPaneRefreshOrchestrationService` が `RefreshSourceSelectionOutcome` と normalized trace を作る。
- `case-display-completed` details に selected source / rebuild fallback outcome を diagnostic fact として含める。

#### forbidden

- trigger reason を `CaseCache` / `BaseCache` / `MasterListRebuild` の採用 source と混同する。
- display request source を cache 採用順序や rebuild fallback required 判定に使う。
- log component source を protocol source として扱う。
- `TaskPaneRefreshCoordinator` が `refreshSource=(reason)` を根拠に snapshot source を決める。
- `BaseCacheFallback` を `MasterListRebuild` required と混同する。
- `CaseCache` / `BaseCache` / `BaseCacheFallback` 採用時に rebuild fallback required を立てる。
- source normalization のために host reuse / already-visible / ready retry / pending retry / render / show 条件を変える。
- source selection failure を visibility recovery、foreground guarantee、追加 guard、sleep、`Application.DoEvents()` で覆って success に丸める。
- source selection のために context-less fallback open、暗黙の workbook 推測、Master path 推測を追加する。
- COM lifetime、`ScreenUpdating` / `DisplayAlerts` / `EnableEvents` restore scope、hidden session cleanup 境界を広げる。

### normalized outcome design

#### RefreshSourceSelectionOutcome

target-state の normalized outcome は少なくとも次を持つ。

| field | 意味 |
| --- | --- |
| `Status` | `NotReached` / `Selected` / `DegradedSelected` / `Failed` / `Unknown` |
| `SelectedSource` | `None` / `CaseCache` / `BaseCache` / `BaseCacheFallback` / `MasterListRebuild` |
| `SelectionReason` | `CaseCacheUsable`、`BaseCachePromoted`、`LatestMasterVersionUnavailable`、`CacheUnavailable`、`SnapshotAcquisitionNotReached`、`SelectionFailed` など。 |
| `TriggerReason` | upstream が渡した raw reason。source 採用判断には使わない diagnostic field。 |
| `DisplayRequestSource` | structured request source。存在する場合のみ保持する。 |
| `WorkbookSnapshotFactsAvailable` | CASE / Base workbook snapshot facts を評価できたか。 |
| `FallbackReasons` | cache が使えなかった diagnostic reason list。 |
| `IsCacheFallback` | `BaseCacheFallback` のように cache source 内の fallback を使ったか。 |
| `IsRebuildRequired` | `MasterListRebuild` が required だったか。 |
| `RebuildFallbackOutcome` | `MasterListRebuild` required 時の normalized rebuild fallback outcome。required でない場合は `Skipped` 相当。 |
| `IsTerminal` | source selection protocol が terminal か。`Unknown` は false。 |
| `CanContinueRefresh` | render / host show evaluation へ進める snapshot result があるか。 |
| `FailureReason` | `WorkbookMissing`、`ContextRejected`、`MasterPathUnavailable`、`SnapshotBuildException`、`NoSnapshotText` など。 |
| `DegradedReason` | latest version unavailable、cache update failure、error snapshot fallback、partial / unverifiable source など。 |

`Status` の意味は次で固定する。

| status | terminal | refresh can continue | 定義 |
| --- | --- | --- | --- |
| `NotReached` | yes | conditional | already-visible、host reuse、precondition skip、context reject などで snapshot acquisition に到達しなかった。失敗とは限らない。 |
| `Selected` | yes | yes | `CaseCache`、`BaseCache`、または正常な `MasterListRebuild` で renderable snapshot を得た。 |
| `DegradedSelected` | yes | yes, but degraded | `BaseCacheFallback`、または `MasterListRebuild` degraded など、表示継続は可能だが diagnostic reason を持つ。 |
| `Failed` | yes | no | snapshot source が必要だったが renderable snapshot を得られなかった。fail closed で扱う。 |
| `Unknown` | no | no | owner が facts を正規化できていない。success completion に使わない。 |

source selection failure の terminal outcome は `Status=Failed` とし、`SelectedSource=None`、`IsTerminal=true`、`CanContinueRefresh=false` とする。`MasterListRebuild` attempted かつ error snapshot text が renderable な場合は `Failed` ではなく `DegradedSelected` とし、degraded fact を trace に残す。

### rebuild fallback との接続条件

refresh source と rebuild fallback の接続は、`SelectedSource=MasterListRebuild` または raw facts 上 `MasterListRebuild` を選ぶべき状態に到達した場合だけ成立する。

- `CaseCache`
  - CASE cache が renderable snapshot を供給できるため、rebuild fallback required ではない。
- `BaseCache`
  - Base snapshot を CASE cache へ昇格して renderable snapshot を供給できるため、rebuild fallback required ではない。
- `BaseCacheFallback`
  - latest master version が読めない degraded / cache fallback であっても、Base snapshot を供給できるため、rebuild fallback required ではない。
- `MasterListRebuild`
  - cache source が snapshot を供給できず、Master list から再構築するため、rebuild fallback required である。
- `None`
  - snapshot acquisition に到達していない、または source selection failure の状態である。到達していない場合は required ではない。selection failure の場合は facts に応じて failed outcome とするが、trigger reason だけを根拠に `MasterListRebuild` を要求しない。

`MasterListRebuild` required になるためには、refresh path が fail-closed せず、CASE render が必要になり、`TaskPaneSnapshotBuilderService` が CASE / Base / Base fallback のいずれでも renderable snapshot を得られないことが必要である。visibility recovery failure、foreground guarantee failure、ready retry exhausted、pending retry exhausted、`WindowActivate` 発火有無は、それ単独では rebuild fallback required ではない。

### CASE display completed との関係

`CASE display completed` は引き続き `TaskPaneRefreshOrchestrationService` の created-case display session terminal state とする。

- selected source は completion の直接条件ではない。
- `CaseCache` / `BaseCache` / `BaseCacheFallback` / `MasterListRebuild` のどれを使っても、completion は pane visible、visibility outcome、foreground outcome で判断する。
- `BaseCacheFallback` や `MasterListRebuild` degraded は diagnostic fact であり、pane visible と foreground terminal が display-completable なら completion の材料を妨げない。
- `RefreshSourceSelectionOutcome.Failed` / `Unknown` は fail-closed とし、success-only の `case-display-completed` には使わない。
- already-visible / host reuse では source selection が `NotReached` でも、pane visible と foreground terminal が成立していれば display completion できる。
- `case-display-completed` は lower-level service から emit しない。

### constraints to preserve

- 既存の source 採用順序を変えない。
  - `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` を維持する。
- cache / snapshot / rebuild 条件を変えない。
  - format compatibility、master version 比較、latest master version unreadable 時の `BaseCacheFallback`、Master read-only rebuild 条件を維持する。
- 白Excel対策を落とさない。
  - source selection を Excel window visibility recovery や post-close quit の代替にしない。
- TaskPane 不表示 regression を防ぐ。
  - already-visible、host reuse、display entry、ready retry、pending retry、render / show 条件を source 正規化の名目で変えない。
- COM解放を落とさない。
  - `MasterWorkbookReadAccessService` の owned read access cleanup と `readAccess.CloseIfOwned()` の finally 境界を維持する。
- fail closed を維持する。
  - workbook / context / `SYSTEM_ROOT` / Master path が不明な場合に推測で補完しない。
- ガード追加で覆わない。
  - source failure を新しい guard、sleep、DoEvents、visibility / foreground 条件の拡大で隠さない。

### 次の実装安全単位候補

1. `TaskPaneSnapshotBuilderService` が `SnapshotSourceSelectionFacts` を返す。`SelectedSource`、selection reason、fallback reason list、workbook snapshot facts availability、failure / degraded reason を含め、採用順序と条件は変えない。
2. `CasePaneSnapshotRenderService` / `TaskPaneHostFlowService` / `TaskPaneRefreshCoordinator` が source selection facts を上位へ伝播する。host reuse / render / show 条件は変えない。
3. `TaskPaneRefreshOrchestrationService` が `RefreshSourceSelectionOutcome` を正規化し、created-case display session 上で normalized trace を emit する。`case-display-completed` emit owner は増やさない。
4. `TaskPaneRefreshCoordinator` の `refreshSource=(reason)` 相当 log を、互換を保ちながら `triggerReason` / `displayRequestSource` / `snapshotSource` の区別へ移す。
5. tests / trace では、`BaseCacheFallback` が rebuild fallback required ではないこと、`CaseCache` / `BaseCache` 採用時に fallback required を立てないこと、source selection failure が fail closed になることを確認する。
6. already-visible / host reuse / precondition skip / context reject では `SnapshotSource=None` または `NotReached` になり、`MasterListRebuild` required にならないことを確認する。

## 一言まとめ

target-state では、`CASE display completed` を `pane visible` の別名にも `refresh completed` の別名にもせず、created-case display session を閉じる orchestration-level terminal state として定義する。

その owner は `TaskPaneRefreshOrchestrationService` に置く。worker は attempt、coordinator は refresh unit / foreground execution bridge、host-flow は visible state、snapshot builder は rebuild fallback をそれぞれ持ち、final completion を奪わない構造が target-state である。
