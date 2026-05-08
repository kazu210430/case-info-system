# CODEX Review First Safe Unit Plan

## 1. 位置づけ

この文書は、CODEX レビューで指摘された責務混在箇所を、現行 `main` と docs 正本に揃えて再整理し、第1安全単位を 1 つ選ぶための docs-only 記録です。

- 正本として使う docs:
  - `AGENTS.md`
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
- 補助的に参照した current-state docs:
  - `docs/current-flow-source-of-truth.md`
  - `docs/responsibility-mapping.md`
  - `docs/taskpane-refactor-current-state.md`
  - `docs/taskpane-manager-responsibility-inventory.md`
- 判断原則:
  - `docs/flows.md` を source-of-truth とし、現コードをそこへ寄せる観点で整理する
  - hidden session、COM cleanup、Excel window 制御、fail-closed 条件をまたぐ大規模再設計はしない
  - 「最小差分」ではなく「1責務境界を完結して移せる安全単位」を優先する

## 2. レビュー指摘とフロー正本の対応表

| CODEXレビュー指摘 | `docs/flows.md` 上の対応フロー | 正本上の責務境界 | 現コードで主に関係する service |
| --- | --- | --- | --- |
| CASE作成と hidden session 制御の分離 | `新規 CASE 作成` | create plan / path resolve は `KernelCaseCreationService`、hidden workbook open は `CaseWorkbookOpenStrategy`、interactive 表示 handoff は `KernelCasePresentationService` | `KernelCaseCreationService`, `CaseWorkbookOpenStrategy`, `KernelCasePresentationService`, `KernelWorkbookCloseService` |
| Kernel lifecycle から管理シート保護/定義検証を分離 | `雛形登録・更新フロー` | context-bound Kernel resolve / preflight orchestration は `KernelTemplateSyncService`、template registration validation rules は `KernelTemplateSyncPreflightService` と `WordTemplateRegistrationValidationService`、publication side effects は `PublicationExecutor` | `KernelTemplateSyncService`, `KernelTemplateSyncPreflightService`, `WordTemplateRegistrationValidationService`, `PublicationExecutor` |
| close と Excel 状態制御プロトコルの一本化 | `CASE ライフサイクル`, `Kernel HOME close / managed close / post-close quit`, `Kernel ユーザー情報反映` | HOME close fail-closed、CASE managed close、post-close quit、hidden reflection cleanup は別 owner で固定される | `CaseWorkbookLifecycleService`, `KernelWorkbookCloseService`, `PostCloseFollowUpScheduler`, `KernelUserDataReflectionService` |
| TaskPane runtime compose の分割 | `TaskPane 更新` | event-side refresh orchestration、host flow / render / show、runtime composition wiring、VSTO create/remove chain は別境界 | `AddInTaskPaneCompositionFactory`, `TaskPaneManagerRuntimeGraphFactory`, `TaskPaneManager`, `TaskPaneHostRegistry`, `TaskPaneHostFactory`, `ThisAddIn` |

## 3. 各指摘で現コードが持ちすぎている責務

### 3.1 CASE作成と hidden session 制御の分離

- 正本上の境界:
  - `KernelCaseCreationService` は create plan / path resolve の owner
  - `CaseWorkbookOpenStrategy` は hidden workbook open / close mechanics と retained hidden app-cache owner
  - `KernelCasePresentationService` は hidden session close 後の表示 owner
- 現コードが持ちすぎている責務:
  - `KernelCaseCreationService.CreateSavedCase(...)` が create plan に加えて transient suppression、hidden create route 分岐、mode 別初期化、save 前 window 正規化、hidden session close 前後の DisplayAlerts 操作まで持つ
  - `KernelCaseCreationService.CreateSavedCaseWithoutShowing(...)` が create owner でありながら hidden session 実行 detail に深く食い込んでいる
  - `CaseWorkbookOpenStrategy` は hidden session mechanics に加えて retained hidden app-cache の寿命と cleanup も持つ
- 整理メモ:
  - `docs/flows.md` では owner は分かれているが、実装上は create owner が hidden-session-sensitive な detail をまだ多く抱える
  - hidden session は例外経路であり、最初の安全単位として触るには window / COM / handoff の連鎖が重い

### 3.2 Kernel lifecycle から管理シート保護/定義検証を分離

- 正本上の境界:
  - `KernelTemplateSyncService` は `WorkbookContext` 起点の Kernel resolve と publication 入口
  - `KernelTemplateSyncPreflightService` / `WordTemplateRegistrationValidationService` は validation 側
  - `PublicationExecutor` は side-effect order owner
- 現コードが持ちすぎている責務:
  - `KernelTemplateSyncService.Execute(...)` が `ExcelApplicationStateScope`、管理シート取得、sheet protection save/unprotect/restore、`CaseList_FieldInventory` 定義読取、preflight request 組立、preflight failure result build、publication 呼出しを 1 本で持つ
  - `TemporarySheetProtectionRestoreScope` が `KernelTemplateSyncService` の private 実装に閉じており、管理シート access 境界が service 外から読めない
  - `LoadDefinedTemplateTags(...)` が preflight 用 fact collection であるにもかかわらず `Execute(...)` に直結している
- 整理メモ:
  - `docs/flows.md` 上は preflight と publication side effects の境界が既に明確
  - hidden session や foreground を跨がず、`WorkbookContext` / `SYSTEM_ROOT` / preflight failure no-side-effect を守りながら owner を揃えやすい

### 3.3 close と Excel 状態制御プロトコルの一本化

- 正本上の境界:
  - HOME close fail-closed は `KernelWorkbookCloseService`
  - dirty prompt / managed close ordering は `CaseWorkbookLifecycleService`
  - post-close quit は `PostCloseFollowUpScheduler`
  - hidden reflection cleanup は `KernelUserDataReflectionService`
- 現コードが持ちすぎている責務:
  - `CaseWorkbookLifecycleService` が dirty prompt / folder offer / managed close dispatch に加えて `DisplayAlerts` snapshot / restore を抱える
  - `KernelWorkbookCloseService` が HOME close handshake と `Quit` 時の `DisplayAlerts` restore protocol を抱える
  - `KernelUserDataReflectionService` が quiet mode、owned workbook visibility restore、owned workbook close、owned application quit、final release まで抱える
- 整理メモ:
  - 問題は「1 service が巨大」というより、close protocol が複数 owner に分散しつつ各 owner が business rule と Excel state control を同居させていること
  - COM lifecycle と再参照禁止契約に直結するため、第1安全単位には向かない

### 3.4 TaskPane runtime compose の分割

- 正本上の境界:
  - runtime composition wiring は `AddInTaskPaneCompositionFactory` / `TaskPaneManagerRuntimeGraphFactory`
  - host flow / display / lifecycle は `TaskPaneHostFlowService` / `TaskPaneDisplayCoordinator` / `TaskPaneHostLifecycleService`
  - VSTO create/remove は `TaskPaneHostRegistry` / `TaskPaneHostFactory` / `TaskPaneHost` / `ThisAddIn`
- 現コードが持ちすぎている責務:
  - `AddInTaskPaneCompositionFactory` が wide delegate surface を受けて compose している
  - `TaskPaneManagerRuntimeGraphFactory` が dispatcher subtree、host factory、host registry、display coordinator を束ね、compose owner と VSTO boundary 直前の wiring が近接している
  - `TaskPaneHostFactory` が CASE / Kernel / Accounting の control create と `ActionInvoked` binding を一括所有している
- 整理メモ:
  - compose owner の整理自体は比較的安全寄りだが、残っている論点は VSTO create/remove と event lifetime に近い
  - `docs/taskpane-refactor-current-state.md` でも、ここから先は runtime-sensitive boundary として慎重に扱う前提になっている

## 4. 安全単位候補一覧

| 候補 | 対応レビュー指摘 | 単位の内容 | 安全性評価 | 今回の優先度 |
| --- | --- | --- | --- | --- |
| A | Kernel lifecycle から管理シート保護/定義検証を分離 | `KernelTemplateSyncService.Execute(...)` から、管理シート access / protection / defined-tag 読取 / preflight 実行を 1 owner に寄せる | 高い。`flows.md` 上の preflight 境界と一致し、hidden session / foreground / close protocol に触れない | 第1候補 |
| B | TaskPane runtime compose の分割 | `AddInTaskPaneCompositionFactory` と `TaskPaneManagerRuntimeGraphFactory` の wiring owner をさらに整理する | 中。pure orchestration 寄りだが、残件は VSTO create/remove と event lifetime に近い | 第2候補 |
| C | CASE作成と hidden session 制御の分離 | create plan owner と hidden session execution owner の境界をさらに明文化する | 低め。window 正規化、retained hidden app-cache、interactive handoff をまたぐ | 後回し |
| D | close と Excel 状態制御プロトコルの一本化 | `DisplayAlerts` / quiet close / quit / final release の protocol を統一する | 最も低い。COM lifecycle と fail-closed を跨ぐ | 最後寄り |

## 5. 第1安全単位の選定

第1安全単位は、候補 A の「`KernelTemplateSyncService` から管理シート保護/定義検証の owner を分離する」を採用する。

選定理由:

- `docs/flows.md` で `preflight` と `publication side effects` がすでに別段で定義されている
- `KernelTemplateSyncPreflightService` と `WordTemplateRegistrationValidationService` が既に存在し、完全なゼロ再実装を避けやすい
- hidden session、CASE 表示、TaskPane ready-show、close/quit の本線に触れない
- `WorkbookContext` 必須、root mismatch fail-closed、preflight failure no-side-effect という current-state の安全装置を守りやすい
- TaskPane compose は安全そうに見えても、残件が `ThisAddIn` / `TaskPaneHostRegistry` / `TaskPaneHostFactory` の VSTO boundary に近い
- close protocol 一本化は、複数 owner の COM-sensitive な protocol をまたぐため初手としては重い

## 6. 第1安全単位で守るべき既存挙動・実機対策・fail-closed 条件

- `WorkbookContext` を唯一の入口とし、`ResolveKernelWorkbook(context)` 失敗時は中止する
- `SYSTEM_ROOT` 不一致や暗黙の Kernel workbook 推測を追加しない
- `ExcelApplicationStateScope` による `ScreenUpdating=false` / `EnableEvents=false` の外側制御は維持し、必ず復元する
- 管理シート protection の save / unprotect / restore 契約を崩さない
- `CaseList_FieldInventory` を定義検証の source-of-truth として維持する
- preflight failure では `PublicationExecutor.PublishValidatedTemplates(...)` を呼ばず、副作用を起こさない
- publication side-effect order を変えない
  - `WriteToMasterList -> TASKPANE_MASTER_VERSION +1 -> Kernel save -> Base snapshot sync -> InvalidateCache`
- `kernel save` failure では Base sync / invalidate へ進めない current-state を維持する
- Base sync failure では invalidate 実行と success + warning semantics を維持する
- Base snapshot sync の close path は、今回の安全単位に含めない
- `build` / `test` / `DeployDebugAddIn` は今回の docs-only 作業では実施しない

## 7. 次に CODEX へ投げる実装プロンプト案

以下を、そのまま次回の実装依頼の叩き台として使う。

```text
目的:
`docs/flows.md` の `雛形登録・更新フロー` を正本として、`KernelTemplateSyncService.Execute(context)` から管理シート保護/定義検証の owner を 1 安全単位で分離する。

必須参照:
- AGENTS.md
- docs/architecture.md
- docs/flows.md
- docs/ui-policy.md
- docs/current-flow-source-of-truth.md の `publication / template sync`
- docs/codex-review-first-safe-unit-plan.md

対象フロー要約:
- `KernelTemplateSyncService.Execute(context)` は `WorkbookContext` から対象 Kernel workbook を確定する
- preflight failure では副作用を起こさない
- success 時だけ `PublicationExecutor` が
  `WriteToMasterList -> TASKPANE_MASTER_VERSION +1 -> Kernel save -> Base snapshot sync -> InvalidateCache`
  を実行する

今回の実装単位:
- `KernelTemplateSyncService` から、
  - 管理シート取得
  - 管理シート protection save/unprotect/restore
  - `CaseList_FieldInventory` 定義読取
  - `KernelTemplateSyncPreflightService.Run(...)` 呼出し
  を 1 owner へ寄せる
- `PublicationExecutor` の順序、Base sync、invalidate、close path には踏み込まない
- hidden session / TaskPane / close protocol には触らない

変更対象候補:
- dev/CaseInfoSystem.ExcelAddIn/App/KernelTemplateSyncService.cs
- 必要なら preflight orchestration 用の新規 App service 1 ファイル
- 必要なら関連テスト:
  - dev/CaseInfoSystem.Tests/KernelTemplateSyncServiceTests.cs

守ること:
- `WorkbookContext` 必須
- root mismatch を補正しない
- preflight failure no-side-effect
- `ScreenUpdating` / `EnableEvents` の restore
- 管理シート protection restore
- publication side-effect order 不変
- Base sync failure の success + warning semantics 不変

やらないこと:
- PublicationExecutor の並び変更
- Base snapshot sync の close protocol 変更
- hidden session / CASE作成 / TaskPane / HOME close への波及
- 大規模な service 再設計

完了条件:
- 安全単位が `KernelTemplateSyncService` 周辺だけで閉じる
- 挙動差分は owner の整理に限定される
- docs と矛盾しない
```

## 8. 今回の結論

- 第1安全単位は `KernelTemplateSyncService` 周辺の preflight owner 整理とする
- これは「管理シート保護/定義検証」を `flows.md` 上の preflight 境界へ寄せる作業であり、publication side effects や close protocol を同時に触らない
- CASE 作成 hidden session、本格的な close protocol 一本化、TaskPane VSTO boundary 整理は、いずれも今回より危険度が高く、第1単位には採らない

## 9. 実装結果記録（2026-05-08）

- 本節は実装後の結果記録であり、`1` から `8` までは実装前の計画記録として保持する
- 実装コミット:
  - `38e4f3fedc6c6da05aca1ed00cb47256379182f3`
- 実装結果:
  - `KernelTemplateSyncPreparationService` を追加し、管理シート解決、sheet protection の一時解除/復元、`SYSTEM_ROOT` 解決、`CaseList_FieldInventory` 定義タグ読取、`KernelTemplateSyncPreflightService` 呼出しを新 owner へ移した
  - `KernelTemplateSyncService` は `WorkbookContext` 起点の kernel resolve、`ExcelApplicationStateScope`、preflight 成否分岐、`PublicationExecutor` 呼出し、結果整形に寄せた
  - `PublicationExecutor` の side-effect order と Base sync/close path には変更を入れていない
- 既存挙動として維持したこと:
  - preflight failure no-side-effect
  - `ScreenUpdating` / `EnableEvents` の restore
  - sheet protection restore と restore failure の握りつぶし
  - `SYSTEM_ROOT` の property/path fallback
  - `WriteToMasterList -> TASKPANE_MASTER_VERSION +1 -> Kernel save -> Base snapshot sync -> InvalidateCache` の順序
  - Base sync failure 時の success + warning semantics
- 実装結果として追加した確認:
  - `KernelTemplateSyncServiceTests` に、preflight failure 時の保護復元と、preparation 例外時の保護復元を固定するテストを追加した

## 10. 実機 NG 調査記録（2026-05-08）

- 本節は、第1安全単位の実装後に実機で観測した NG と、その原因調査記録である
- 観測した症状:
  - 雛形更新後に新規 CASE を作成したところ、表示準備中にぐるぐるループした
  - いったん終了後、対象 CASE を開いたところ白 Excel になった
  - 白 Excel はウインドウ再表示で復元した
  - 他の既存 CASE は、ボタンパネル更新で表示できた
- 調査で分かったこと:
  - `KernelTemplateSyncService` から `KernelTemplateSyncPreparationService` へ移した処理について、`ExcelApplicationStateScope`、sheet protection restore、preflight、publication の前後関係は変わっていない
  - publication 本線の順序 `WriteToMasterList -> TASKPANE_MASTER_VERSION +1 -> Kernel save -> Base snapshot sync -> InvalidateCache` は変わっていない
  - Base 本体 `案件情報System_Base.xlsx` は、調査時点で `TASKPANE_BASE_MASTER_VERSION=68`、`TASKPANE_MASTER_VERSION=68`、`TASKPANE_BASE_SNAPSHOT_COUNT=14` を持っていた
  - 2026-05-08 15:50 以後に新規作成した `20260508_テスト\\案件情報_テスト.xlsx` は、`TASKPANE_BASE_MASTER_VERSION=68`、`TASKPANE_MASTER_VERSION=68`、`TASKPANE_SNAPSHOT_CACHE_COUNT=14` で正常だった
  - 一方で、問題として開かれた `20260508_白フラッシュ出ちゃった。\\案件情報_白フラッシュ出ちゃった。.xlsx` は、`TASKPANE_BASE_MASTER_VERSION=65` を保持していた
  - 同 CASE の reopen ログでは、`caseMasterVersion=65, latestMasterVersion=68`、`embeddedMasterVersion=65, latestMasterVersion=68` から `MasterListRebuild` と foreground recovery に入っていた
  - 問題 CASE のフォルダ作成時刻は 2026-05-08 00:26:46 で、15:50 の雛形更新より前だった
- `KernelTemplateSyncPreparationService` 分離が直接原因ではなさそうな理由:
  - refactor 前後で publication 本線の順序変更がない
  - `ExcelApplicationStateScope` / protection restore / preflight / publication の関係が同一である
  - Base 本体は update 後に version 68 / snapshot count 14 へ到達している
  - update 後に作成した新規 CASE は version 68 / cache 正常で、今回の refactor だけで新規 CASE を壊した形跡がない
  - 実機 NG の対象は、雛形更新後に作った CASE ではなく、更新前状態を引きずった stale CASE reopen と読む方が事実に合う
- 本命の残課題:
  - stale CASE reopen 時の `TaskPaneSnapshotBuilderService` による `MasterListRebuild`
  - rebuild 後の foreground recovery
  - ready-show / window recovery と白 Excel 観測の関係
- 次の安全単位候補として扱う対象:
  - `TaskPaneSnapshotBuilderService`
  - ready-show / window recovery 側の orchestration
- revert 判断:
  - `KernelTemplateSyncPreparationService` 分離 commit の revert は第一推奨ではない
  - 理由は、今回の NG が「refactor 後に作った新規 CASE の破損」ではなく、「stale CASE reopen で既存の rebuild / recovery 経路が表面化した事象」である可能性が高いため
