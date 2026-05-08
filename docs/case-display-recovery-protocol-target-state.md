# CASE Display Recovery Protocol Target State

## 位置づけ

この文書は、`docs/case-display-recovery-protocol-current-state.md` を前提に、CASE display / recovery protocol の target-state を定義するための docs-only 設計記録です。

- 基準コード:
  - `2026-05-08` 時点で `main` と `origin/main` の一致を確認した `e41feb5d607f79077e112a1945e81ac0a76d95a4`
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
| `foreground guarantee completed` | 同一 display session に対して foreground obligation が残っていない状態。`Required` なら recovery 実行後、`NotRequired` なら skip 判定の terminal 化までを含む。 | completion owner は `TaskPaneRefreshCoordinator`。execution owner は `ExcelWindowRecoveryService`。 | 条件付き必要条件。`Required` の場合だけ必要。 |
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

ただしこれは「必ず foreground recovery を実行する」という意味ではない。`TaskPaneRefreshCoordinator` が `NotRequired` と判断した場合も、その判断自体を terminal 化してから completion へ進む。

### 補助条件

- `refresh completed`

`refresh completed` は refresh path でだけ通る補助条件とする。already-visible path では `refresh completed` なしで completion できる。

## Owner Boundary Decisions

### 1. CASE display completed definition owner

`CASE display completed` の canonical owner は `TaskPaneRefreshOrchestrationService` とする。

理由:

- worker は 1 attempt の owner であり、fallback handoff や後続 foreground obligation を知らない。
- coordinator は refresh / foreground unit の owner だが、already-visible path を知らない。
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

- foreground guarantee protocol unit の decision owner は `TaskPaneRefreshCoordinator`
- foreground guarantee protocol unit の completion owner も `TaskPaneRefreshCoordinator`
- 実際の app/window/foreground recovery primitive の execution owner は `ExcelWindowRecoveryService`

この関係は「decision と execution を分けるが、protocol unit completion owner は coordinator に固定する」と読む。

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
8. `TaskPaneRefreshCoordinator` が foreground unit を `Completed` または `NotRequired` で閉じる。
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
- `TaskPaneRefreshCoordinator`: refresh unit、foreground guarantee decision / terminal outcome、refresh attempt outcome 返却。
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

## 一言まとめ

target-state では、`CASE display completed` を `pane visible` の別名にも `refresh completed` の別名にもせず、created-case display session を閉じる orchestration-level terminal state として定義する。

その owner は `TaskPaneRefreshOrchestrationService` に置く。worker は attempt、coordinator は refresh / foreground unit、host-flow は visible state、snapshot builder は rebuild fallback をそれぞれ持ち、final completion を奪わない構造が target-state である。
