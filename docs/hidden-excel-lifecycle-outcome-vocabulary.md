# Hidden Excel Lifecycle / Outcome Vocabulary

## 位置づけ

この文書は、hidden Excel / isolated app / retained hidden app-cache / white Excel lifecycle redesign の最初の実装安全単位として、lifecycle terms、normalized outcome、trace vocabulary、owner vocabulary を docs-only で固定する正本です。

- 開始時の `main` / `origin/main`: `13a6c1ace4ecc290141cc5c28e91309726586eb8`
- 参照した正本:
  - `AGENTS.md`
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/case-display-recovery-protocol-target-state.md`
  - `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`
  - `docs/hidden-excel-isolated-app-white-excel-lifecycle-target-state.md`

この文書は vocabulary を固定するだけです。コード変更、enum 実装、helper / service 抽出、runtime 条件変更、visibility 制御変更、cleanup 条件変更、reopen 条件変更、build / test / `DeployDebugAddIn` 実行は行いません。

## Vocabulary Principles

- raw facts と normalized outcome を混ぜません。
- primitive owner と outcome owner を混ぜません。
- trigger owner と recovery owner を混ぜません。
- visibility owner と foreground owner を混ぜません。
- foreground owner と hidden cleanup owner を混ぜません。
- WindowActivate は display / refresh trigger であり、hidden cleanup / retained cleanup / white Excel prevention / foreground guarantee / CASE display completion の owner ではありません。
- shared/current app は原則 user-owned です。no visible workbook quit の close / quit protocol 以外で業務処理側が終了しません。
- isolated app は生成 owner が close / quit / COM release まで閉じます。
- retained hidden app-cache は `CaseWorkbookOpenStrategy` の例外境界です。session close と cached `Application` cleanup を混同しません。
- `WorkbookClose` 後は、close 前に採取済みの immutable facts だけを使い、close 済み workbook COM object を再参照しません。

## Lifecycle Terms

| term | definition | owner | lifecycle scope | cleanup scope | 触ってよい owner | 触ってはいけない owner |
| --- | --- | --- | --- | --- | --- | --- |
| `shared app` | Add-in が接続している利用者操作中の Excel `Application`。`current app` と同義に扱う場合があるが、user-owned である点を強調する呼称。 | user / Excel host。例外的な no visible workbook quit は `PostCloseFollowUpScheduler`。 | 利用者が操作する Excel instance の lifetime。 | 原則 cleanup 対象外。CASE close 後に visible workbook が無い場合だけ white Excel prevention として quit 判定対象。 | `PostCloseFollowUpScheduler` は no visible workbook quit の既存条件内でのみ触れる。各 shared-current app 利用 service は snapshot / restore 範囲で state を触る。 | hidden cleanup owner、foreground owner、WindowActivate dispatch owner、retained cache owner は shared app を推測で quit しない。 |
| `current app` | 現在の処理が利用している Excel `Application`。shared/current app 文脈では user-owned、isolated app 文脈では service-owned になり得るため ownership facts とセットで扱う。 | 文脈依存。shared/current app は user / Excel host、isolated app は生成 service。 | 呼び出し元が保持している `Application` reference の有効範囲。 | ownership が明確な場合のみ cleanup scope を持つ。 | owner facts を持つ service。 | ownership 不明な caller、WindowActivate、foreground recovery。 |
| `isolated app` | 処理 owner が専用に生成した hidden Excel `Application`。 | 生成 service。CASE create mechanics は `CaseWorkbookOpenStrategy`、reflection は `KernelUserDataReflectionService`。 | `Create -> Open -> Work -> Save/Close -> Quit -> COM release` の owner-owned session。 | owned workbook close、owned isolated app quit、COM final release。 | isolated app を生成した service と、その service が委譲した mechanics owner。 | user-owned shared/current app owner、WindowActivate、foreground guarantee、visibility recovery。 |
| `retained instance` | session 終了後も cache owner が保持する hidden `Application` instance。 | `CaseWorkbookOpenStrategy` hidden app-cache。 | retained hidden app-cache の cache entry lifetime。 | idle return、timeout、poison、feature flag disabled、shutdown cleanup。 | `CaseWorkbookOpenStrategy` cache owner。 | session owner 以外の caller、WindowActivate、foreground recovery、process 名だけを見た cleanup。 |
| `retained hidden app-cache` | CASE 新規作成専用 hidden create route の内部最適化として、hidden `Application` を再利用する cache。 | `CaseWorkbookOpenStrategy`。 | `CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE` の例外境界内。 | session close は workbook close / release まで、cached app cleanup は idle timeout / poison / shutdown だけ。 | `CaseWorkbookOpenStrategy`。 | `KernelCaseCreationService` が cached app cleanup を直接奪わない。WindowActivate / post-close quit / foreground recovery は cache cleanup しない。 |
| `hidden Excel` | hidden `Application` または hidden workbook window を使う作業状態の総称。一般的な表示制御手段ではなく、owner と cleanup が閉じた managed hidden session だけ許容する。 | session owner または mechanics owner。 | hidden create、hidden reflection、hidden-for-display、read-only hidden open など用途限定。 | owner-owned workbook / app / COM object に限定。 | managed hidden session owner、read access owner、display handoff owner。 | unrelated service、WindowActivate、foreground guarantee、visibility recovery。 |
| `orphaned instance` | owner / cache entry / session facts が失われ、user-owned Excel と区別できない可能性がある `EXCEL.EXE`。 | current-state では top-level owner 未定義。owner facts が残る場合は元 owner。 | 運用監視上の未定義領域。 | target-state docs だけでは broad cleanup しない。 | owner facts が揃う cache owner / isolated owner。 | process 名や PID だけを根拠にした cleanup、WindowActivate、foreground recovery。 |
| `white Excel` | visible workbook が無いのに Excel shell だけが残る close / quit protocol 上の状態。 | `PostCloseFollowUpScheduler`。 | CASE managed close / post-close follow-up。 | no visible workbook quit。 | `PostCloseFollowUpScheduler` が既存 visible workbook 判定内で触る。 | WindowActivate、visibility recovery、foreground guarantee、hidden cleanup owner。 |
| `managed close` | service が明示的に close intent と save / prompt / follow-up を調停する workbook close。 | CASE は `CaseWorkbookLifecycleService` / `ManagedCloseState` / `PostCloseFollowUpScheduler`。Kernel は Kernel close lifecycle owner。 | close 前 facts 採取から workbook close、post-close follow-up 予約まで。 | close 対象 workbook と close 後 follow-up。 | close lifecycle owner。 | close 後に workbook を再参照する caller、display / foreground owner。 |
| `unmanaged close` | user 操作または Excel event から入る close で、managed close scope ではないもの。 | event capture は `ThisAddIn` / `WorkbookLifecycleCoordinator`。role ごとの lifecycle service が扱える範囲だけ扱う。 | event-driven close observation。 | owner が確定した範囲のみ。 | role lifecycle service。 | hidden cleanup / white Excel prevention を推測で広げる caller。 |
| `reopen` | close 済み workbook の延長ではなく、新しい workbook / window facts を取得し直す open。interactive created CASE では hidden create close 後に shared/current app へ reopen する。 | `KernelCasePresentationService` / `CaseWorkbookOpenStrategy`。 | new workbook COM object、window、role、context、snapshot の再取得。 | reopen 前の古い COM object は cleanup 対象外。 | reopen owner と downstream display owner。 | close 済み workbook を再利用する caller。 |
| `visibility restore` | hidden / minimized workbook window を visible / normal へ戻す、または display handoff 後に対象 window を見える状態へ戻す処理。 | 保存状態正規化は session owner。lightweight ensure は `WorkbookWindowVisibilityService`。full recovery primitive は `ExcelWindowRecoveryService`。 | save normalization、presentation preparation、ready-show / recovery primitive。 | visibility 操作対象の owner-owned workbook window。 | session owner、`WorkbookWindowVisibilityService`、`ExcelWindowRecoveryService`。 | retained cleanup owner の代替、foreground guarantee の代替、WindowActivate 単独根拠。 |
| `foreground guarantee` | display / refresh protocol 内で対象 Excel / workbook window を前面に戻す obligation を terminal outcome にする unit。 | decision / outcome / trace は `TaskPaneRefreshOrchestrationService`、execution primitive は `ExcelWindowRecoveryService`。 | created-case display session / refresh attempt の foreground obligation。 | cleanup ではない。 | `TaskPaneRefreshOrchestrationService` と `ExcelWindowRecoveryService`。 | hidden cleanup owner、retained cleanup owner、post-close quit owner、WindowActivate。 |
| `hidden cleanup` | hidden session が所有する workbook close、isolated app quit、COM release、cache return / poison 判定。 | hidden session owner。CASE create mechanics は `CaseWorkbookOpenStrategy`、reflection は `KernelUserDataReflectionService`。 | hidden create session、managed hidden reflection session、retained cache session close。 | owner-owned workbook / app / COM object。 | hidden session owner / cache owner。 | foreground guarantee、visibility recovery、WindowActivate、user prompt owner。 |
| `retained cleanup` | retained hidden app-cache に保持された cached `Application` を idle timeout / poison / shutdown 等で破棄する cleanup。 | `CaseWorkbookOpenStrategy` cache owner。 | retained cache entry lifetime。 | cached `Application` quit / COM final release。 | `CaseWorkbookOpenStrategy` cache owner。 | session owner 以外の service、post-close white Excel owner、WindowActivate、foreground recovery。 |
| `fail closed` | owner facts、workbook / window / context、cleanup eligibility が不明な場合に成功扱いせず、推測補完や条件拡大を行わないこと。 | 各 protocol owner。 | 全 lifecycle / outcome 判定。 | 不明な対象には cleanup や visibility / foreground mutation を広げない。 | protocol owner が failure / skipped / unknown として扱う。 | guard 追加、sleep、DoEvents、visibility 条件拡大で覆う実装。 |

## Owner Vocabulary

| owner term | definition | 正しい使い方 | 混ぜてはいけないもの |
| --- | --- | --- | --- |
| `decision owner` | protocol unit を実行するか、skip / required / failed / unknown のどれにするかを判断する owner。 | facts を見て decision を作る。例: foreground guarantee decision は `TaskPaneRefreshOrchestrationService`。 | primitive owner、trigger owner。 |
| `primitive owner` | COM 操作、window 操作、close / quit / release など実際の mutation を行う owner。 | 既存条件内で primitive を実行し、execution facts を返す。例: full recovery primitive は `ExcelWindowRecoveryService`。 | outcome owner。primitive 成功を protocol completion と同一視しない。 |
| `outcome owner` | raw facts と primitive result を protocol vocabulary に正規化する owner。 | normalized outcome を 1 箇所で決める。 | raw trace owner、lower-level helper。 |
| `trace owner` | raw event trace または normalized trace を emit する owner。 | raw trace と normalized outcome trace を分ける。 | action owner。trace を出しただけで cleanup / recovery owner にならない。 |
| `recovery owner` | visibility recovery / foreground recovery など回復 protocol を調停する owner。 | recovery が必要かを判断し、primitive owner へ委譲する。 | trigger owner、cleanup owner。 |
| `cleanup owner` | owner-owned resource を close / quit / release / cache cleanup する owner。 | ownership facts が揃う範囲だけ cleanup する。 | visibility owner、foreground owner、UX owner。 |
| `UX owner` | user prompt、folder offer、manual guidance など利用者向け判断や表示を担う owner。 | close prompt は `CaseClosePromptService`、folder offer は `CaseFolderOpenService`。 | hidden cleanup completion、post-close quit eligibility。 |

固定する不等式:

- primitive owner ≠ outcome owner。
- trigger owner ≠ recovery owner。
- visibility owner ≠ foreground owner。
- foreground owner ≠ hidden cleanup owner。
- trace owner ≠ action owner。
- cleanup owner ≠ UX owner。

## Normalized Outcome Vocabulary

この節の outcome 名は docs 上の vocabulary です。まだ enum 実装しません。各 outcome は raw facts ではなく、owner が正規化した protocol result として扱います。

### Visibility Outcome

owner:

- decision / outcome owner: 呼び出し元 protocol owner。ready-show / display recovery では `TaskPaneRefreshOrchestrationService`。
- primitive owner: `WorkbookWindowVisibilityService` または `ExcelWindowRecoveryService`。

| outcome | meaning | trigger | guarantee | recovery | degraded / fail 条件 |
| --- | --- | --- | --- | --- | --- |
| `VisibilityRestoreNotRequired` | 対象 window は既に必要な visibility 条件を満たす、またはこの protocol では visibility restore が不要。 | display / refresh / save normalization の事前判定。 | visibility mutation は行わない。success terminal として扱えるのは owner が対象 facts を持つ場合だけ。 | なし。 | facts 不足なら `NotRequired` にしない。 |
| `VisibilityRestoreCompleted` | owner が必要な visibility restore を実行し、対象 window が visible / normal 等の必要条件を満たした。 | save normalization、presentation preparation、ready-show recovery。 | visibility restore は完了。ただし foreground guarantee completed ではない。 | 後続 foreground obligation があれば別 outcome へ渡す。 | foreground failure をこの outcome に混ぜない。 |
| `VisibilityRestoreSkipped` | 既存条件により visibility restore を実行しなかった。失敗とは限らない。 | role mismatch、対象外 route、owner-owned window ではない等。 | mutation なし。 | 必要なら fail closed または downstream retry。 | skip 理由が不明な場合は `Unknown`。 |
| `VisibilityRestoreDegraded` | restore は一部成功したが、期待状態を完全には保証できない。 | primitive execution partial success。 | display completion の材料にする場合は owner が degraded を明示的に許容した時だけ。 | retry / foreground recovery など別 owner へ事実を渡す。 | window resolve 不安定、activation 不成立など。 |
| `VisibilityRestoreFailed` | visibility restore が必要だったが失敗した。 | primitive exception、対象 window 解決失敗、COM 操作失敗。 | success terminal にしない。 | fail closed。必要なら既存 retry path へ委ねる。 | 失敗を foreground success や WindowActivate で覆わない。 |
| `VisibilityRestoreUnknown` | raw facts が不足し outcome を正規化できない。 | instrumentation gap。 | success completion に使わない。 | B 以降の trace 正本化候補。 | unknown を skipped / completed に丸めない。 |

### Hidden Cleanup Outcome

owner:

- hidden session owner / cleanup owner。
- CASE create mechanics は `CaseWorkbookOpenStrategy`。
- reflection は `KernelUserDataReflectionService`。

| outcome | meaning | trigger | guarantee | recovery | degraded / fail 条件 |
| --- | --- | --- | --- | --- | --- |
| `HiddenExcelCleanupNotRequired` | hidden session が開始されていない、または owner-owned hidden resource が無い。 | hidden route not entered、already open shared workbook reuse。 | cleanup mutation なし。 | なし。 | hidden resource の有無が不明なら使わない。 |
| `HiddenExcelCleanupCompleted` | owned workbook close、owned isolated app quit、COM release、または retained session return が owner scope 内で完了。 | session close / abort / finally。 | hidden cleanup scope は閉じた。foreground / visibility / white Excel success ではない。 | なし。 | retained cached app が残る場合は retained outcome と併記する。 |
| `HiddenExcelCleanupSkipped` | owner が既存条件により cleanup を実行しない。 | shared/current app reuse、owner 不一致、cleanup eligibility なし。 | user-owned resource へ触らない。 | 必要なら owner facts を trace。 | owner 不明を skipped に丸める場合は理由を必ず持つ。 |
| `HiddenExcelCleanupDegraded` | workbook close はできたが app quit / release / cache return に不確実性が残る等。 | partial cleanup、release exception handled、cache poison。 | success completion と同一視しない。 | poison / retry / shutdown cleanup へ接続。 | orphan 可能性、COM release failure。 |
| `HiddenExcelCleanupFailed` | cleanup が必要だったが owner scope 内で失敗した。 | close / quit / release exception。 | success にしない。 | fail closed。cache owner は poison 等を検討。 | foreground recovery、WindowActivate、guard で覆わない。 |
| `HiddenExcelCleanupUnknown` | cleanup facts が不足し正規化できない。 | trace gap。 | success completion に使わない。 | B フェーズで owner trace を正本化する候補。 | raw log だけで completed と呼ばない。 |

### Retained Cleanup Outcome

owner:

- `CaseWorkbookOpenStrategy` hidden app-cache。

| outcome | meaning | trigger | guarantee | recovery | degraded / fail 条件 |
| --- | --- | --- | --- | --- | --- |
| `RetainedInstanceKept` | healthy cached `Application` を保持する。 | session close 後、cache owner が keep / idle return 可能と判断。 | cached app は cleanup されない。 | idle timeout / poison / shutdown が後続 cleanup trigger。 | owner facts 不明なら keep にしない。 |
| `RetainedInstanceReturnedToIdle` | session close 後、owned workbook を閉じたうえで cached app を idle cache に戻した。 | app-cache route session close。 | session cleanup と cache keep の境界が分かれる。 | timeout / shutdown まで保持。 | workbook が残る場合は degraded / failed。 |
| `RetainedInstanceCleanupNotRequired` | retained instance が存在しない、または cleanup trigger がない。 | app-cache disabled、no cache entry。 | cleanup mutation なし。 | なし。 | cache state 不明なら unknown。 |
| `RetainedInstanceCleanupCompleted` | cache owner が cached app を quit / COM final release した。 | idle timeout、poison、feature flag disabled、shutdown。 | retained cleanup scope は閉じた。 | なし。 | user-owned shared app に触っていないことが前提。 |
| `RetainedInstanceCleanupSkipped` | cleanup trigger があったが既存条件で実行しない、または別 route に委ねた。 | in-use、owner 不一致、shutdown ordering 等。 | success ではない。 | next cache lifecycle event で再評価。 | skip 理由なしは unknown。 |
| `RetainedInstanceCleanupDegraded` | cleanup は一部実行されたが quit / release / state 確認に不確実性が残る。 | quit exception handled、release exception、cache poison。 | completed と同一視しない。 | owner trace と poison facts を残す。 | orphan risk。 |
| `RetainedInstanceCleanupFailed` | retained cleanup が必要だったが失敗した。 | quit / release failure。 | success にしない。 | fail closed。user guidance は未定義事項。 | process kill へ直行しない。 |
| `RetainedInstancePoisoned` | cached app を再利用不可として扱う。 | cleanup failure、unhealthy detection。 | reuse しない。cleanup completed ではない。 | cache owner の cleanup trigger へ接続。 | poison と cleanup completed を混同しない。 |
| `RetainedInstanceOwnershipUnknown` | cache entry / owner facts が不足して retained instance と断定できない。 | diagnostic gap。 | cleanup 不可。 | 別安全単位で owner facts を確認。 | broad process cleanup 禁止。 |

#### Retained Cleanup Protocol Mapping

retained cleanup vocabulary は `CaseWorkbookOpenStrategy` の retained hidden app-cache にだけ使います。session close の hidden cleanup facts と、cached `Application` disposal facts を分けて読むため、次の mapping を正本とします。

| protocol event | raw facts | normalized retained outcome | outcome owner | 使用禁止の読み替え |
| --- | --- | --- | --- | --- |
| app-cache session close succeeds and cache remains healthy | route、workbook close attempted / completed、cache returned to idle、cache poisoned false。 | `RetainedInstanceReturnedToIdle` または keep decision。 | `CaseWorkbookOpenStrategy`。 | `IsolatedAppReleased`、`RetainedInstanceCleanupCompleted`、foreground / visibility success。 |
| app-cache session close cannot safely return to idle | route、workbook close failure、return-to-idle failure、health check failure、cache poisoned true。 | `RetainedInstancePoisoned`、必要に応じて hidden cleanup degraded / failed。 | `CaseWorkbookOpenStrategy`。 | poison を cleanup completed と呼ぶこと。 |
| cached app is disposed by timeout / poison / feature flag disabled / shutdown | cleanup reason、owner facts、app quit attempted、app quit completed、release facts。 | `RetainedInstanceCleanupCompleted` / `RetainedInstanceCleanupDegraded` / `RetainedInstanceCleanupFailed`。 | `CaseWorkbookOpenStrategy`。 | hidden workbook close completed を cached app cleanup completed と呼ぶこと。 |
| cache entry absent or trigger absent | cache entry missing、cleanup trigger absent。 | `RetainedInstanceCleanupNotRequired`。 | `CaseWorkbookOpenStrategy`。 | trace gap を not required に丸めること。 |
| cache-owned facts are false or insufficient | owner mismatch、cache entry / process facts insufficient。 | `RetainedInstanceCleanupSkipped` / `RetainedInstanceOwnershipUnknown`。 | `CaseWorkbookOpenStrategy`。 | PID / process name だけで cleanup すること。 |

retained cleanup の raw facts は diagnostic であり、それ自体は success / failure ではありません。normalized outcome は cache owner が決めます。`WindowActivate`、white Excel prevention、foreground guarantee、visibility restore は retained cleanup outcome owner ではありません。

### Workbook Close Outcome

owner:

- close lifecycle owner。CASE は `CaseWorkbookLifecycleService` / `ManagedCloseState` / `PostCloseFollowUpScheduler`。

| outcome | meaning | trigger | guarantee | recovery | degraded / fail 条件 |
| --- | --- | --- | --- | --- | --- |
| `WorkbookCloseNotRequired` | 対象 close が不要、またはこの owner の close scope 外。 | role mismatch、already closed facts、owner scope 外。 | close しない。 | なし。 | close 必要性不明なら unknown / failed。 |
| `WorkbookCloseCompleted` | managed close が完了し、close 後に対象 workbook を再参照しない状態に入った。 | managed close path。 | close 前 immutable facts だけが後続へ渡る。 | post-close follow-up は別 outcome。 | close 後 COM 参照が必要なら設計不備。 |
| `WorkbookCloseSkipped` | close owner が既存条件により close しなかった。 | prompt cancel、visible workbook reuse、owner mismatch。 | workbook は開いたまま。 | user prompt / next event。 | skipped を failed と混同しない。 |
| `WorkbookCloseDegraded` | close は部分的に進んだが follow-up / state consistency に不確実性が残る。 | close exception handled、post-close scheduling uncertainty。 | success-only terminal ではない。 | fail closed / trace。 | close 後再参照で補完しない。 |
| `WorkbookCloseFailed` | close が必要だったが失敗した。 | `Workbook.Close` exception、save failure、prompt flow failure。 | close completed としない。 | prompt / retry の既存 owner へ戻す。 | hidden cleanup / foreground recovery で覆わない。 |
| `WorkbookReferenceInvalidAfterClose` | close 済み workbook reference が無効であり、再参照禁止であることを示す diagnostic outcome。 | close completion 後。 | old COM object を使わない。 | reopen は新しい facts を取得。 | close 後 workbook member access。 |

### Workbook Reopen Outcome

owner:

- `KernelCasePresentationService` / `CaseWorkbookOpenStrategy`。

| outcome | meaning | trigger | guarantee | recovery | degraded / fail 条件 |
| --- | --- | --- | --- | --- | --- |
| `WorkbookReopenNotRequired` | batch route など reopen が不要。 | `CreateCaseBatch`、display handoff なし route。 | shared/current app open は行わない。 | folder display 等の別 route。 | interactive route で誤用しない。 |
| `WorkbookReopenRequired` | close 後に shared/current app で新しい workbook facts を取得する必要がある。 | interactive created CASE display。 | decision outcome。reopen 完了ではない。 | `OpenHiddenForCaseDisplay(...)` へ進む。 | required を completed と混同しない。 |
| `WorkbookReopenCompleted` | new workbook COM object / window / context を取得し、display handoff へ渡せる状態。 | `OpenHiddenForCaseDisplay(...)` completion。 | old closed workbook は使わない。 | downstream display / visibility / foreground。 | window-dependent completion ではない。 |
| `WorkbookReopenSkipped` | reopen を既存条件で行わなかった。 | route mismatch、prompt cancel、owner scope 外。 | display success ではない。 | caller outcome へ戻す。 | reason なし skipped は unknown。 |
| `WorkbookReopenFailed` | reopen が必要だったが失敗した。 | open exception、path missing、context failure。 | display handoff へ進めない。 | fail closed。 | stale workbook / window facts で補完しない。 |
| `WorkbookReopenUnknown` | reopen facts が不足。 | trace gap。 | success に使わない。 | B 以降の trace 正本化対象外または別安全単位。 | required / skipped へ丸めない。 |

### Isolated App Outcome

owner:

- isolated app 生成 service。

| outcome | meaning | trigger | guarantee | recovery | degraded / fail 条件 |
| --- | --- | --- | --- | --- | --- |
| `IsolatedAppReleaseNotRequired` | isolated app が生成されていない、または retained app として保持判断された。 | shared/current app route、app-cache keep。 | quit / release なし。 | retained outcome があればそちらへ渡す。 | isolated/shared ownership 不明なら使わない。 |
| `IsolatedAppReleased` | owner-owned isolated app の quit / COM final release が完了。 | session close / abort / finally。 | isolated app lifetime は閉じた。 | なし。 | workbook close 未完了なら completed にしない。 |
| `IsolatedAppReleaseSkipped` | existing condition により release しない。 | retained healthy cache、owner mismatch、already released。 | success とは限らない。 | cache lifecycle または owner trace。 | skip reason なしは unknown。 |
| `IsolatedAppReleaseDegraded` | quit / release の一部に不確実性が残る。 | quit failure handled、release exception。 | released と同一視しない。 | poison / trace / shutdown cleanup。 | orphan risk。 |
| `IsolatedAppReleaseFailed` | release が必要だったが失敗した。 | quit / release exception。 | success にしない。 | fail closed。 | foreground recovery / WindowActivate で覆わない。 |
| `IsolatedAppOwnershipUnknown` | 対象 app が isolated owner-owned と断定できない。 | missing owner facts。 | cleanup しない。 | owner facts 正本化。 | user-owned shared app を誤って quit しない。 |

### White Excel Prevention Outcome

owner:

- close / quit protocol owner。現行 owner は `PostCloseFollowUpScheduler`。

| outcome | meaning | trigger | guarantee | recovery | degraded / fail 条件 |
| --- | --- | --- | --- | --- | --- |
| `WhiteExcelPreventionNotRequired` | queued key の target が current open workbook として残っている、visible workbook が残っている、または white Excel prevention scope 外。 | post-close still-open check / visible workbook check。 | shared app を quit しない。 | なし。 | still-open facts と visible workbook facts のどちらも判定できない場合は not required にしない。 |
| `WhiteExcelPreventionQueued` | close 後 follow-up として no visible workbook quit 判定が予約された。 | managed close completion。 | quit completed ではない。 | scheduler が retry / check へ進む。 | queued を completed と混同しない。 |
| `WhiteExcelPreventionCompleted` | visible workbook が無いことを確認し、shared/current app quit が完了した。 | post-close follow-up。 | white Excel prevention scope は閉じた。 | なし。 | quit 成功後の終了中 app は restore しない。 |
| `WhiteExcelPreventionSkipped` | 既存条件により quit しなかった。 | visible workbook exists、ownership unknown、busy / retry exhausted 等。 | shared app は残る。 | retry / user guidance は既存または未定義 owner。 | skipped reason を trace。 |
| `WhiteExcelPreventionDegraded` | quit 判定または実行が部分的で、Excel 残存可能性がある。 | busy retry、partial failure、trace gap。 | completed と同一視しない。 | fail closed / diagnostic。 | foreground recovery で success に丸めない。 |
| `WhiteExcelPreventionFailed` | no visible workbook quit が必要だったが失敗した。 | `Application.Quit` exception 等。 | success にしない。 | user-facing guidance は未定義事項。 | visibility restore / WindowActivate で補完しない。 |
| `WhiteExcelPreventionUnknown` | white Excel prevention outcome を正規化できない。 | insufficient facts。 | success に使わない。 | trace 正本化候補。 | raw log だけで completed と呼ばない。 |

E-2 current-state 補足:

- `WhiteExcelPreventionNotRequired` は current emitted outcome として 2 つの主な意味を持ちます。`outcomeReason=targetWorkbookStillOpen` は queued key と current open workbook key が一致したため quit しないこと、`outcomeReason=visibleWorkbookExists` は queued key の target は closed と読めるが shared/current app に visible workbook が残るため quit しないことを表します。
- `targetWorkbookStillOpen` の判定は、close 前に queue へ積まれた key と、dequeue 時点の current `Application.Workbooks` 列挙から得た fresh open workbook key の比較です。close 済み workbook COM object の再参照ではありません。
- immediate reopen 近接時に fresh reopened workbook が queued key と一致した場合、current-state では follow-up cancel ではなく still-open skip として `WhiteExcelPreventionNotRequired` を記録します。queued key を fresh reopen facts で置き換えません。
- `WhiteExcelPreventionCompleted` は still-open false かつ visible workbook false の後に `Application.Quit()` が完了した場合だけです。reopen / display / foreground の成功を completed に読み替えません。
- `WhiteExcelPreventionSkipped` は vocabulary 上の候補として残しますが、現行 `PostCloseFollowUpScheduler` の primary emitted outcome は `Queued` / `NotRequired` / `Completed` / `Failed` です。visible workbook がある場合の現行 emitted outcome は `Skipped` ではなく `NotRequired` です。
- `WhiteExcelPreventionFailed` は no visible workbook quit の試行後に失敗した場合です。foreground recovery、visibility restore、WindowActivate dispatch、hidden cleanup の outcome で補完しません。

## Trace Vocabulary

| trace term | definition | trace owner | action owner との関係 |
| --- | --- | --- | --- |
| `raw facts` | event received、workbook path、role、window visibility、owner facts、exception など未正規化の観測事実。 | facts を採取した owner。WindowActivate raw event は `ThisAddIn`。 | raw facts は success / failure ではない。 |
| `normalized outcome` | raw facts と primitive result を protocol vocabulary に正規化した結果。 | protocol outcome owner。 | lower-level primitive owner が勝手に上位 outcome を emit しない。 |
| `lifecycle transition` | created、opened、returned-to-idle、closed、released、queued、completed 等の state transition。 | lifecycle owner。 | display trace や foreground trace と混同しない。 |
| `cleanup decision` | cleanup required / not required / skipped / degraded / failed を決める trace。 | cleanup owner。 | foreground / visibility owner は cleanup decision を持たない。 |
| `visibility decision` | visibility restore required / not required / failed 等を決める trace。 | display / recovery protocol owner。 | hidden cleanup completed の代替にしない。 |
| `reopen decision` | reopen required / not required / failed 等を決める trace。 | reopen owner。 | close 済み workbook facts を再利用しない。 |
| `foreground decision` | foreground guarantee required / not required / degraded / failed 等を決める trace。 | `TaskPaneRefreshOrchestrationService`。 | hidden cleanup / white Excel prevention へ持ち込まない。 |
| `dispatch decision` | trigger を downstream protocol に渡すか ignored / deferred / failed とする trace。 | trigger-specific dispatch owner。WindowActivate は `WindowActivatePaneHandlingService`。 | dispatch success は display success ではない。 |

trace 固定ルール:

- raw facts と normalized outcome を同じ field / log message の意味にしません。
- trace owner と action owner を混ぜません。trace を emit した owner が cleanup / recovery primitive owner になるわけではありません。
- WindowActivate event capture は raw event trace だけを持てます。WindowActivate を hidden cleanup、retained cleanup、white Excel prevention、foreground guarantee、CASE display completion の normalized trace owner に戻しません。
- `case-display-completed` は `TaskPaneRefreshOrchestrationService` だけが emit します。
- cleanup trace と display / foreground trace を同じ success terminal に丸めません。
- `Unknown` は success completion に使いません。

## Implementation Prohibitions Before Runtime Work

この vocabulary がコードへ反映される前に、次の方針を禁止事項として固定します。

- `Application.DoEvents()` を追加しない。
- sleep / timing hack を追加しない。
- visibility 条件を拡大しない。
- `WindowActivate` を cleanup owner 化しない。
- guard 追加で cleanup failure、visibility failure、foreground failure を覆う設計にしない。
- `WorkbookClose` 後に close 済み workbook を再参照しない。
- hidden cleanup と foreground recovery を混在させない。
- user-owned shared app cleanup を行わない。
- enum 実装、helper / service 抽出、runtime 条件変更をこの docs-only フェーズに混ぜない。
- build / test / `DeployDebugAddIn` を docs-only フェーズで実行しない。

## B Phase Handoff

次の B フェーズは「hidden Excel cleanup owner の trace 正本化」です。今回の vocabulary を前提に、既存 runtime 条件を変えずに trace の owner と outcome の owner を揃える範囲だけを扱います。

B フェーズで扱う範囲:

- hidden session cleanup の owner facts を trace 上で区別する。
- `HiddenExcelCleanupCompleted`、`Skipped`、`Degraded`、`Failed`、`Unknown` を success に丸めない。
- retained app-cache の session close と cached app cleanup を trace 上で分ける。
- isolated app release と hidden workbook close を trace 上で分ける。
- trace emit owner と cleanup primitive owner の関係を明示する。

B フェーズで扱わない範囲:

- hidden workbook open / close 条件変更。
- isolated app quit 条件変更。
- shared/current app quit 条件変更。
- visibility restore / foreground guarantee 条件変更。
- `WindowActivate` dispatch 変更。
- orphaned `EXCEL.EXE` broad cleanup。
- post-close white Excel prevention 条件変更。
- service 分割、helper 抽出、enum の広範囲導入。

不明なまま残す事項:

- orphaned instance detection / user-facing UX の top-level owner。
- post-close quit failure 時の利用者向け guidance。
- retained hidden app-cache を将来も維持するかどうか。
- read-only / temporary workbook close 全般を hidden lifecycle 正本に含めるかどうか。

## 一言まとめ

この文書で固定するのは、runtime の動作ではなく言葉の境界です。hidden cleanup、retained cleanup、visibility restore、foreground guarantee、WorkbookClose / reopen、white Excel prevention、WindowActivate dispatch は別 unit であり、ある unit の成功を別 unit の成功へ読み替えません。
