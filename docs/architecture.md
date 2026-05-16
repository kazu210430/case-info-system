# システム構成

## 概要

案件情報System は、Excel ブックと VSTO Add-in を中心に構成されています。主要な構成要素は `Kernel`、`Base`、`CASE`、会計書類セット、Excel Add-in、Word Add-in、Excel Launcher です。

## 主要構成要素

- `Kernel`
  - 起点となるブックは `案件情報System_Kernel.xlsx` です。
  - Excel Add-in はファイル名や DocProperty をもとに Kernel を判定します。
  - Kernel では HOME 相当の画面、設定反映、CASE 作成、案件一覧遷移などが扱われます。
- `Base`
  - `案件情報System_Base.xlsx` を CASE 作成時のコピー元として扱います。
  - ファイル名または `ROLE=BASE` の DocProperty が判定材料です。
- `CASE`
  - 個別案件のブックです。
  - `ROLE=CASE`、`SYSTEM_ROOT`、対応拡張子、既知パス情報などをもとに CASE として扱われます。
- 会計書類セット
  - CASE から派生して作成される別 Workbook です。
  - `CASEINFO_WORKBOOK_KIND=ACCOUNTING_SET` や `SOURCE_CASE_PATH` などの DocProperty を持ちます。

## Add-in の役割分担

### Excel Add-in

Excel Add-in は、実行時の中核制御を担当します。

- WorkbookRole 判定
- Excel イベント購読
- Kernel HOME 表示制御
- CASE 作成
- CASE 表示制御
- TaskPane 構築と更新
- 文書作成コマンド実行
- 会計書類セット作成
- Workbook 保存前・クローズ前制御
- Excel ウィンドウ復旧と前面化

### Word Add-in

Word Add-in は、Word 側の補助機能を担当します。

- Word 起動時の初期化
- スタイルペイン表示制御
- ContentControl の Title / Tag 一括置換

Word Add-in の存在は確認できますが、各帳票の詳細差し込み仕様はこの文書の対象外です。

## WorkbookRole の考え方

Excel Add-in は Workbook を役割ごとに分類して処理を切り替えます。

- `Kernel`
  - `案件情報System_Kernel.xlsx` または `案件情報System_Kernel.xlsm` を優先判定します。
- `Base`
  - `案件情報System_Base.xlsx` または `案件情報System_Base.xlsm`、または `ROLE=BASE` を持つブックとして扱います。
- `CASE`
  - Kernel / Base / 会計書類セット以外の対象ブックです。
  - `ROLE=CASE`、`SYSTEM_ROOT`、対応拡張子などが判定に使われます。
- 会計書類セット
  - `CASEINFO_WORKBOOK_KIND=ACCOUNTING_SET` や会計用シート構成、`SOURCE_CASE_PATH` などで判定されます。

## 実行時の主要入口

Excel Add-in は起動時にサービスを組み立て、次の Excel イベントを購読します。

- `WorkbookOpen`
- `WorkbookActivate`
- `WorkbookBeforeSave`
- `WorkbookBeforeClose`
- `WindowActivate`
- `SheetActivate`
- `SheetSelectionChange`
- `SheetChange`
- `AfterCalculate`

これらを入口に、Workbook ライフサイクル、TaskPane 更新、表示制御が連動します。

## サービス構成の大枠

Excel Add-in の組み立ては `AddInCompositionRoot` で行われます。責務は大きく次の単位に分かれます。

- Kernel 系
  - Kernel 解決、設定、CASE 作成、CASE 表示、Kernel HOME 関連。
- CASE / Lifecycle 系
  - `CaseWorkbookLifecycleService` は orchestration 寄りで、CASE / Base 初回初期化、dirty session 状態管理、before-close / managed close / post-close follow-up の調停、created case folder offer pending 状態管理、CASE HOME 表示補正を担います。
  - close prompt は `CaseClosePromptService`、保存先フォルダ解決・存在確認・Explorer 起動は `CaseFolderOpenService`、Kernel の name rule 読み取りは `KernelNameRuleReader` が担当します。
  - `ManagedCloseState` は managed close の入れ子状態を、`PostCloseFollowUpScheduler` は close 後 follow-up / retry / no visible workbook 時の Excel 終了判定を担当します。
- Document 系
  - テンプレート解決、出力名解決、実行可否判定、Word 生成、保存、待機 UI。
  - `DocumentExecutionEligibilityService` は登録済みテンプレートを前提に、VSTO 実行に必要な基本適格性を確認します。
  - allowlist / review の旧 runtime policy 系は撤去済みです。
  - `DocumentExecutionModeService` は mode の読取と運用スイッチ管理を担います。現行コードで確認できる主用途は Word warm-up 制御であり、gating 本体ではありません。
- Accounting 系
  - 会計書類セット作成、会計ブック制御、補助フォーム、保存別名処理。
  - 会計フォーム / import prompt の close lifecycle は `docs/accounting-close-lifecycle-current-state.md` を正本とし、実機安定化済みの現行順序を helper 化・共通化・再整理の対象にしません。
- TaskPane 系
  - スナップショット構築、描画、リフレッシュ調停、Window 単位の表示管理、CASE pane UIイベント dispatch。
  - 現在は `TaskPaneManager` を facade に、`TaskPaneHostFlowService` が refresh-time host flow、`TaskPaneHostLifecycleService` が registry-backed host lifecycle、`TaskPaneDisplayCoordinator` が show/hide 調停、`TaskPaneRefreshOrchestrationService` / `TaskPaneRefreshCoordinator` が event-side refresh orchestration を担う構造です。
  - その周辺で `TaskPaneRefreshPreconditionPolicy`、`CasePaneSnapshotRenderService`、`CasePaneCacheRefreshNotificationService`、`TaskPaneActionDispatcher` などへ主責務が分離されています。
- Infrastructure 系
  - Excel / Word Interop、パス互換、フォルダ表示、ウィンドウ復旧、ログなど。

## 大型調停クラス改善の現在地

直近の大型クラス改善は、責務統合ではなく、判断・分類・facts / trace 組み立ての collaborator 化として固定します。owner / lifecycle / callback / completion / close / COM release は安易に移動していません。

- `TaskPaneRefreshOrchestrationService`
  - precondition / fail-closed decision、observation 後 decision、retry continuation decision は `TaskPaneRefreshPreconditionDecisionService`、`TaskPaneRefreshObservationDecisionService`、`TaskPaneRefreshRetryContinuationDecisionService` へ分離済みです。
  - retry timer、ready-show callback、pending retry lifecycle、completion / emit、created-case session、TaskPane 表示命令、foreground 実行の owner は orchestration 側に残します。
- `CaseWorkbookOpenStrategy`
  - route decision、cleanup outcome 分類、presentation handoff facts、hidden app lifecycle support facts / reason / trace は `CaseWorkbookOpenRouteDecisionService`、`CaseWorkbookOpenCleanupOutcomeService`、`CaseWorkbookPresentationHandoffService`、`CaseWorkbookHiddenAppLifecycleSupportService` へ分離済みです。
  - Excel application 作成、workbook open / close、hidden session owner、isolated app lifecycle、shared app handoff、retained app-cache owner、cleanup 実行、app quit、COM release、CASE 表示 recovery owner は strategy 側に残します。
- `ThisAddIn`
  - 現行 `main` の `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs` は 1510 行です。大型化していた 2187 行規模の状態から、startup / execution boundary と runtime diagnostics を collaborator へ分離した到達点として扱います。
  - VSTO Startup / Shutdown 入口、composition root 呼び出し、Application event wiring / unwiring、event handler 入口、CustomTaskPane create / remove、shutdown cleanup 呼び出しは `ThisAddIn` 側に残します。
  - startup guard、empty startup quit、suppression count、ScreenUpdating bridge、DisplayAlerts bridge、runtime diagnostics は `AddInStartupBoundaryCoordinator`、`AddInExecutionBoundaryCoordinator`、`AddInRuntimeExecutionDiagnosticsService` に寄せています。
- `KernelWorkbookResolverService`
  - `ResolveOrOpen(...)` / `ResolveOrOpenReadOnly(...)` のような業務都合で open を内包する限定境界では、temporary open の close ownership を `KernelWorkbookAccessResult.CloseIfOwned()` に寄せます。
  - `CloseIfOwned()` は resolver が一時 open した workbook だけを一度だけ閉じ、既に open 済みの Kernel workbook は閉じません。

大型クラス削減の目的は、巨大クラスを小さくするだけではなく、変更理由と安全境界を読みやすくすることです。一方で、今回の成果として行数削減と責務削減は実際に進んでいます。

次に触る候補は、TaskPane completion / emit 境界の整理、`ThisAddIn` に残る event handler glue のさらなる整理、`CaseWorkbookOpenStrategy` の owner / lifecycle を動かさず切れる追加削減に限ります。実機安定化済みの hidden Excel / white Excel / Book1 / close lifecycle / TaskPane lifecycle は不用意に変更しません。

## KernelWorkbookService 3分割の到達点

- `KernelWorkbookService`
  - 現在は facade として残し、既存の呼び出し面と composition root の受け口を維持します。
  - binding / display / close の具体責務は内包せず、各境界へ委譲します。
- `KernelWorkbookBindingService`
  - Kernel workbook 解決、HOME binding、設定読取・保存の境界です。
  - HOME binding は Kernel workbook と正規化済み `SYSTEM_ROOT` が一致する場合だけ成立し、不一致や binding 不成立は補正せず fail-closed で扱います。
- `KernelWorkbookDisplayService`
  - HOME 表示準備、Excel / workbook window visibility、`ReleaseHomeDisplay`、prepared display state の解放を担います。
  - display / visibility の責務を持ちますが、HOME close finalization や managed close backend は持ちません。
- `KernelWorkbookCloseService`
  - HOME close backend 調停、CASE 作成中 skip-restore 分岐、pending close、`FormClosed` 後 finalization、managed close bridge を担います。
  - close / finalization の責務を持ちますが、root 解決や visibility policy 自体の owner ではありません。
- 到達点の読み方
  - HOME binding / root 整合は binding 境界で閉じます。
  - display / visibility と close / finalization の混在は解消済みで、`KernelWorkbookService` は facade としてだけ残します。

## Kernel HOME close / managed close / post-close quit の安定化境界

この節で固定するのは、close / quit のうち `Kernel HOME close`、`Kernel managed close`、`CASE managed close`、`post-close quit` の到達点だけです。全 workbook close 経路の一般ルールではありません。`KernelUserDataReflectionService` の未 open Base / Accounting 反映は、service-owned な `managed hidden reflection session` の例外として別節で定義します。`MasterWorkbookReadAccessService`、`CaseWorkbookOpenStrategy` などの読み取り専用 / 一時 workbook close は別 owner の境界として扱います。会計フォーム / import prompt の「Excelを閉じる」直 close 経路は `docs/accounting-close-lifecycle-current-state.md` で安定化済み契約として固定します。

- Form / Service の責務分離
  - `KernelHomeForm` は close の意思表示と `FormClosing` cancel による close 可否制御を担います。
  - `KernelWorkbookService` は facade として close 要求を受け、HOME session close の backend 調停と `FormClosed` 後 finalization は `KernelWorkbookCloseService`、Kernel managed close は `KernelWorkbookLifecycleService`、CASE managed close は `CaseWorkbookLifecycleService`、post-close quit は `PostCloseFollowUpScheduler` が担います。
- HOME close は fail-closed
  - backend close 成功後にのみ HOME session / binding / visibility を解放します。
  - close 失敗時は Form を閉じず、binding / visibility を維持します。
  - HOME session の finalization は `FormClosed` 後にだけ行います。
- 今回安定化対象の managed close / quit 経路
  - `WorkbookCloseInteropHelper` を経由し、`Workbook.Close` の named argument は使いません。
  - optional 引数は `Type.Missing` / `false` を明示して渡します。
  - close 後に対象 workbook を再参照しません。
  - `Save` / `DisplayAlerts` / `Close` / `Quit` 以外の COM 操作を増やさず、`ExcelApplicationStateScope` は使いません。
  - `DisplayAlerts` は個別の `try/finally` で扱い、`Quit` 成功後は終了中 `Application` を restore しません。restore は `Quit` 失敗時だけに限定します。
- 既知の残課題
  - helper 非経由 close 全件の一般棚卸しはこの節の確定範囲外です。ただし、会計フォーム / import prompt の「Excelを閉じる」直 `workbook.Close()` は残課題ではなく、`docs/accounting-close-lifecycle-current-state.md` の固定済み例外です。
  - `WorkbookPromptSuppressionHelper` による `Workbook.Saved` 操作も今回の確定範囲外です。

## hidden session / 裏Excel 設計原則

この節は、`docs/flows.md` の CASE 作成 / CASE 表示 / 会計書類セット / CASE ライフサイクルと、`docs/ui-policy.md` の UI 制御原則を前提に、priority A フェーズ完了時点の hidden session / 裏Excel の正本を固定するための節です。hidden session は一般的な実装テクニックとして推奨せず、owner と cleanup が閉じた例外だけを許容します。

hidden Excel / isolated app / retained hidden app-cache / white Excel lifecycle の protocol 単位の current-state は、`docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md` を参照します。この節は設計原則、同文書は現行 owner / cleanup / visibility / close-reopen 接続点の正本として扱います。owner / protocol の target-state は `docs/hidden-excel-isolated-app-white-excel-lifecycle-target-state.md` を参照します。lifecycle / outcome / trace / owner vocabulary は `docs/hidden-excel-lifecycle-outcome-vocabulary.md` を参照します。white Excel prevention / recovery の current-state と target boundary は `docs/white-excel-prevention-boundary-current-state.md` を参照します。

### 1. 原則

- 既定は `shared/current application`
  - 利用者が操作中の Excel `Application` は caller-owned であり、業務処理側は終了責務を持ちません。
  - shared/current app を使う経路では `Application` 状態変更を snapshot / restore 前提で局所化し、`Application.Quit()` は行いません。
- 裏Excelは一般許可しない
  - hidden workbook / hidden `Application` を使うのは、用途限定の `managed hidden session` だけに制限します。
  - hidden open は可視状態の戦略であり、ownership の例外ではありません。
- retained hidden app-cache は例外扱い
  - retained instance を持つのは `CaseWorkbookOpenStrategy` の retained hidden app-cache だけです。
  - これは one-shot isolated lifecycle の一般形ではなく、CASE新規作成専用 route の内部最適化としてだけ扱います。

### 2. 許容される managed hidden session

- `KernelUserDataReflectionService`
  - 未 open の Base / Accounting 反映だけに `managed hidden reflection session` を許容します。
- CASE新規作成
  - `KernelCaseCreationService.CreateSavedCaseWithoutShowing(...)` から入る CASE新規作成専用 `managed hidden create session` だけを許容します。

### 3. 原則避けるもの

- 補助処理のためだけの別 `Excel.Application` fallback
  - `AccountingSetKernelSyncService` の未 open fallback は撤去済みで、今後も一般化しません。
- owner 不明な hidden app
  - workbook close、`Application.Quit`、COM release、orphan cleanup が service 内で閉じない hidden app は許容しません。
- route 名と実装のズレ
  - `experimental-isolated-inner-save` のように route 名が専用 hidden `Application` を示す場合、その実装も同じ意味に揃えます。

### 4. shared app / isolated app / retained hidden app-cache の境界

- `shared/current app`
  - 現在の Add-in が共有している `Application` を使う経路です。
  - 既に open 済みの workbook を再利用した場合、再利用側はその workbook や `Application` を勝手に close / quit しません。
- `isolated app`
  - 専用 `Application` を生成する経路です。
  - `Create -> Open -> Work -> Save/Close -> Quit -> COM release` を生成側サービスが最後まで担当します。
- `retained hidden app-cache`
  - workbook close までは各 hidden session が担当し、cached `Application` 自体の破棄は cache 側の idle return / timeout / poison / shutdown だけで行います。

### 5. AccountingSetKernelSyncService の到達点

- 不要な別 `Excel.Application` 生成 fallback は撤去済みです。
- 未 open workbook を扱う場合も `kernelWorkbook.Application` を shared/current app として使い、`AccountingWorkbookService.OpenInCurrentApplication(...)` で開きます。
- 既に open 済みの会計 workbook は再利用し、save はしても close しません。
- 自分で open した workbook だけを hidden window のまま反映し、save 後に quiet close します。
- `DisplayAlerts` / `ScreenUpdating` / `EnableEvents` の restore は `ExcelApplicationStateScope` の局所スコープに閉じます。

### 6. KernelUserDataReflectionService の到達点

- 未 open 対象だけ `managed hidden reflection session` を開始します。
- 既に open 済みの Base / Accounting workbook は再利用し、save はしても close しません。
- owner は `KernelUserDataReflectionService` です。未 open 対象に限り hidden な isolated `Application` を生成し、open した対象 workbook だけを扱います。
- save 前には owned workbook window visibility を restore し、hidden window state を保存ファイルへ残さない契約を持ちます。
- cleanup は `CloseWorkbookQuietly`、`Application.Quit`、`ComObjectReleaseService.FinalRelease` まで含めて service 内で完結します。
- shared Excel 側の quiet mode (`DisplayAlerts` / `EnableEvents` / `ScreenUpdating`) は restore まで含めて shared app 内で閉じます。

### 7. CASE新規作成専用 managed hidden create session の到達点

- この例外は `KernelCaseCreationService.CreateSavedCaseWithoutShowing(...)` から `CaseWorkbookOpenStrategy.OpenHiddenWorkbook(...)` を呼ぶ CASE新規作成に限定します。
- 現コードでは `ShouldUseHiddenCreateSession()` が `true` 固定のため、`NewCaseDefault` / `CreateCaseSingle` / `CreateCaseBatch` の全モードが hidden create 分岐を通ります。
- session owner は `KernelCaseCreationService` です。hidden workbook open / close mechanics は `CaseWorkbookOpenStrategy` が担当し、retained hidden app-cache を使う場合だけ cached `Application` 自体の owner は `CaseWorkbookOpenStrategy` に残ります。
- CASE workbook open の route / hidden route decision は `CaseWorkbookOpenRouteDecisionService` が値として組み立てます。これは判断責務の分離であり、Excel app lifecycle、hidden session owner、workbook close owner、cleanup owner、COM release owner は `CaseWorkbookOpenStrategy` に残します。
- interactive route (`NewCaseDefault` / `CreateCaseSingle`) は hidden session 内の save 前に workbook window を `Visible=true` へ戻さず、必要なら `WindowState=xlNormal` だけを整えて save / close します。その後に shared app の `OpenHiddenForCaseDisplay(...)` と `KernelCasePresentationService` / `WorkbookWindowVisibilityService` が表示責務を引き継ぎます。
- 実機確認済みの禁止契約として、interactive route の hidden create session 中に保存前正規化を理由に `Visible=true` を実行してはなりません。これは白フラッシュ再発と終了時 Excel / Book1 発生の再露出を招いたためであり、final visible / normal presentation は `KernelCasePresentationService` 側の責務として遅延します。
- batch route (`CreateCaseBatch`) も save 前に workbook window を `visible + normal` へ正規化してから save / close しますが、表示経路へは昇格させず CASE workbook の reopen も行いません。

| route 名 | 実装上の `Application` | retained app | owner / cleanup | 事実ベースのメモ |
| --- | --- | --- | --- | --- |
| `legacy-isolated` | `CaseWorkbookOpenStrategy` が専用 hidden `Application` を生成 | なし | `KernelCaseCreationService` が session を開始し、`HiddenCaseWorkbookSession.Close/Abort` が workbook close と `Application.Quit`、COM final release を完結 | current/shared app は使いません。 |
| `experimental-isolated-inner-save` | 専用 hidden `Application` を生成 | なし | cleanup は `legacy-isolated` と同じ | `CASEINFO_EXPERIMENT_DEDICATED_HIDDEN_INNER_SAVE` 指定時の opt-in route です。current/shared app は使わず、close 時 inner save を含みます。 |
| `app-cache` | `CaseWorkbookOpenStrategy` cache が専用 hidden `Application` を生成または再利用 | あり | session は workbook close まで担当し、cache owner は idle return / poison / timeout / shutdown で cached `Application` を破棄 | retained hidden app-cache の唯一の例外です。shared/current app は使いません。 |
| `app-cache-bypass-inuse` | cache が使用中の時だけ専用 hidden `Application` を別途生成 | なし | cleanup は `legacy-isolated` と同じ | 内部 fallback route であり retained app を引き継ぎません。 |

- `KernelHomeForm` からの Kernel HOME close は、interactive な CASE 表示が始まった後に `KernelWorkbookCloseService` / `KernelHomeSessionDisplayPolicy` が `skipDisplayRestoreForCaseCreation` を判定し、CASE 作成フロー中は Kernel を前景へ戻さない契約で閉じます。
- `ThisAddIn_Shutdown` は retained hidden app-cache の最終 cleanup を `CaseWorkbookOpenStrategy.ShutdownHiddenApplicationCache()` 経由で実行します。これは retained hidden app-cache cleanup を指す shutdown API です。
- 旧環境変数 `CASEINFO_EXPERIMENT_SHARED_HIDDEN_EXCEL` は互換 alias としてのみ内部対応し、契約上の正本は `CASEINFO_EXPERIMENT_DEDICATED_HIDDEN_INNER_SAVE` です。

### 8. 残課題

- retained hidden app-cache の実運用上の必要性確認
- 旧 alias `CASEINFO_EXPERIMENT_SHARED_HIDDEN_EXCEL` の将来撤去判断
- retained hidden app-cache に起因する orphaned `EXCEL.EXE` の運用監視
- `MasterWorkbookReadAccessService` は shared/current app の read-only open 境界として別扱いを維持すること
- `TaskPaneManager` の runtime composition 整理は次フェーズ候補として扱うこと

## Startup事実収集の分離（KernelStartupContextInspector）

本システムにおいて、startup 時の Kernel HOME 表示判定は、

- 事実収集（Context生成）
- 判定（Policy）

が分離されています。

### 構造

`KernelWorkbookStateService`
↓
`KernelStartupContextInspector`（事実収集のみ）
↓
`KernelStartupContext`（DTO）
↓
`KernelWorkbookStartupDisplayPolicy`（判定）

### 設計原則

- `WorkbookOpen` は window 安定境界ではありません。
- window 依存処理は `WorkbookActivate` / `WindowActivate` 以降で扱います。
- `WorkbookOpen` 直後の window-dependent refresh skip 判定は `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` を正本とします。
- この policy は pure 判定のみを持ち、ログ出力・状態変更・COMメンバーアクセス・UI操作を持ちません。
- Inspector は UI制御・window制御・判定ロジックを持ちません。
- `ActiveWorkbook` の読み取りタイミングは旧実装から変更しません。
- 振る舞い不変を最優先とします。

### 備考

- Window列挙（`workbook.Windows` / `window.Visible`）による可視判定は現状維持しています。
- startup 時の `HasOpenKernelWorkbook` は HOME 表示可否のための事実収集として扱い、表示後に任意の open Kernel workbook を選んで binding する用途には使いません。
- これは将来分離可能な技術的負債として扱います。

## 雛形管理の設計方針

本システムでは、雛形の品質担保は登録時に行います。

- 実行時ではなく登録時に不正な雛形を `雛形一覧` から排除します。
- 実装上の検証は `CaseList_FieldInventory` を基準にした最小限の妥当性確認です。
- 雛形の修正責任は利用者側にあります。
- 文書実行時の安全性は runtime allowlist gating ではなく、登録前 validation によって担保します。
- 実行時は登録済み `templateSpec` を前提に処理し、文書作成本線は `DocumentExecutionEligibilityService` の基本適格性で進みます。
- allowlist / review の旧 runtime policy サービスは撤去済みです。

これにより次を狙います。

- TaskPane 表示の安定化
- 文書作成時エラーの削減
- 問題発生時の切り分け容易化

## Document 実行ポリシーの現状

- `allowlist`
  - runtime gating には使っていません。
  - config ファイル、csproj 同梱設定、専用 tools、旧 runtime policy サービスは撤去済みです。
- `review`
  - runtime safety には使っていません。
  - config ファイル、csproj 同梱設定、専用 tools、旧 runtime policy サービスは撤去済みです。
- `mode`
  - runtime gating 目的ではありません。
  - 現行コードで確認できる主用途は Word warm-up 制御などの運用スイッチです。
  - allowlist / review とは分けて扱い、現時点では撤去対象に含めません。

## Document 系サービスの補足

- `DocumentExecutionModeService`
  - mode 読み取りと運用スイッチ管理を担当します。
  - Word warm-up 制御に関与します。
  - gating 本体ではありません。

## タグ定義運用

実装上、雛形登録時の Tag 検証で直接参照される定義元は Kernel の管理シート `CaseList_FieldInventory` です。Base の `ホーム` シート A列は、システムから直接参照されません。

ただし、運用ルールとしては次を採用します。

- Base `ホーム` シート A列をタグ定義の正本とします。
- `CaseList_FieldInventory` は Base `ホーム` シート A列と一致させて管理します。
- Base `ホーム` シート A列を変更した場合は `CaseList_FieldInventory` を更新します。
- VSTO 側の同期入口はリボンの `Base定義更新` とし、Base `ホーム` シート A列から `CaseList_FieldInventory.ProposedFieldKey` へ反映する補助コマンドとして扱います。
- 同期コマンドは `SourceCell` / `ProposedNamedRange` / `Label` / `DataType` / `NormalizeRule` などの既存メタ情報を正当化なく変更せず、行対応は `SourceCell=B{Base HOME row}` を基準に fail-closed で扱います。
- 同期後も Word 雛形の CC Tag は利用者が同じキーへ修正し、雛形登録・更新で再検証する運用を維持します。

## TaskPane と HOME の位置づけ

- CASE 向け UI は主に Excel の Custom Task Pane として表示されます。
- TaskPane のタイトルは `案件情報System` で、左ドックに配置されます。
- Kernel HOME は TaskPane ではなく、WinForms の独立フォームとして表示されます。
- Kernel HOME は valid binding を持たない `unbound` 状態でも表示され得ます。
- `unbound` HOME は placeholder-only の UI セッションとして扱い、Kernel workbook / Kernel window の自動選択・自動 bind・自動復元は行いません。
- sheet 遷移、案件作成、設定変更などの bound 前提処理は、valid binding がある場合だけ実行します。

## TaskPane snapshot と version 管理

CASE の文書ボタンパネルは、Master 一覧を都度直接読むのではなく、DocProperty に保持した snapshot と version を使って構成します。主な責務分担は次のとおりです。

- `KernelTemplateSyncService`
  - `shMasterList` / `雛形一覧` を更新し、`TASKPANE_MASTER_VERSION` を進めます。
  - Base に TaskPane 用 snapshot と master version を埋め込みます。
- `TaskPaneSnapshotBuilderService`
  - CASE 表示時に `CASE cache -> Base cache -> MasterList rebuild` の順で snapshot を解決します。
  - MasterList から再構築した snapshot を CASE cache に保存します。
- `MasterWorkbookReadAccessService`
  - `TaskPaneSnapshotBuilderService` と `MasterTemplateCatalogService` が共有する Master 読み取り境界です。
  - Master path 解決、read-only open、所有 workbook close、window 非表示化を一元化します。
- `TaskPaneSnapshotCacheService`
  - 文書実行時に表示中 Pane と整合する CASE cache を優先して参照します。
  - 必要に応じて Base 埋込 snapshot を CASE cache へ昇格します。

| プロパティ | 保存先 | 用途 | 更新タイミング |
| --- | --- | --- | --- |
| `TASKPANE_MASTER_VERSION` | Kernel, Base, CASE | Master 一覧に対応する現在 version。型は `long`、形式は `yyyyMMddNNN` です。CASE 側では CASE cache がどの master を前提にしたかの記録にも使います。 | 雛形登録・更新成功時に Kernel で同日連番を進めます。Base 反映時に Base にも保存。CASE では Base からの昇格時と MasterList rebuild 時に更新されます。 |
| `TASKPANE_BASE_MASTER_VERSION` | Base, 新規 CASE | Base に埋め込まれた snapshot がどの master version 由来かを示します。 | 雛形登録・更新成功後、Base snapshot 更新時に書き込みます。CASE では Base 埋込状態を引き継ぎます。 |
| `TASKPANE_SNAPSHOT_CACHE_COUNT` | CASE | CASE cache の chunk 数です。`0` は cache 無効を表します。 | CASE cache 保存時に更新。案件一覧登録後は削除せず `0` に戻します。 |
| `TASKPANE_SNAPSHOT_CACHE_XX` | CASE | 表示中 Pane と整合する CASE snapshot 本体です。`XX` は 2 桁連番です。 | Base から CASE cache へ昇格する時、または MasterList rebuild で再構築した時に保存します。案件一覧登録後は chunk を削除します。 |
| `TASKPANE_BASE_SNAPSHOT_COUNT` | Base, 新規 CASE | Base に埋め込んだ snapshot の chunk 数です。 | 雛形登録・更新成功後、Base snapshot 更新時に保存します。新規 CASE はこの埋込値を引き継ぎます。 |
| `TASKPANE_BASE_SNAPSHOT_XX` | Base, 新規 CASE | Base に埋め込んだ TaskPane snapshot 本体です。`XX` は 2 桁連番です。 | 雛形登録・更新成功後、Base snapshot 更新時に保存します。既存 CASE の案件一覧登録後整理では触りません。 |

### 補足

- Base に snapshot / version を埋め込む目的は、新規 CASE 作成直後に不要な MasterList rebuild を避けることです。
- TaskPane / Master / Snapshot の version は `yyyyMMddNNN` 形式の `long` として扱います。同日複数更新時は `NNN` を 001 から進め、999 を超える場合は fail-closed とします。旧整数方式の値は次回更新時に当日 `yyyyMMdd001` へ移行します。候補 version が既存 version 以下になる場合は、時計ずれ対策として単調増加を優先します。
- `DocumentTemplateResolver` は `TaskPaneSnapshotCacheService` 経由で CASE cache を先に参照し、cache にない場合だけ master catalog にフォールバックします。
- `WorkbookActivate` / `WindowActivate` のたびに version 比較して Pane を再生成する構成ではありません。
- 正本 / 派生 cache / snapshot / Base / CASE の境界整理は `docs/template-metadata-inventory.md` を参照してください。

## SYSTEM_ROOT 文脈境界と Kernel workbook 選択

- `DocumentTemplateResolver`、`AccountingTemplateResolver`、`MasterTemplateCatalogService` などの template 解決系は、CASE workbook や対象 workbook から `SYSTEM_ROOT` を解決して文脈境界を切ります。
- `MasterTemplateCatalogService` の master catalog cache も、現在は resolved master path 単位で保持・invalidate されます。
- 雛形登録・更新フローの入口である `KernelCommandService -> KernelTemplateSyncService` は、Kernel pane 由来の `WorkbookContext` を保持したまま進みます。
- `KernelTemplateSyncService` は `_kernelWorkbookService.ResolveKernelWorkbook(context)` により、まず `context.Workbook` の Kernel 文脈を優先し、それが使えない場合だけ `WorkbookContext.SystemRoot` に対応する open Kernel workbook を解決します。
- これにより、master catalog cache の境界だけでなく、雛形登録・更新の操作対象 Kernel workbook も `SYSTEM_ROOT` 単位で確定します。
- 複数 Kernel workbook や hidden workbook が同時に存在する場合でも、雛形登録・更新、snapshot 反映、cache invalidate が別 root に流れる余地を減らします。
- `GetOpenKernelWorkbook()` のような文脈なしで 1 冊の Kernel workbook を返す API は廃止し、Kernel workbook の選択は `ResolveKernelWorkbook(context)` / `ResolveKernelWorkbook(systemRoot)` に集約します。
- `WorkbookContext` を Kernel 操作の唯一の source-of-truth とし、root 不一致は補正せず fail-closed とします。
- 許容される open は、明示的な `WorkbookContext` / `SYSTEM_ROOT` 文脈からの open と user action 起点の open に限ります。
- context-less fallback open や暗黙の workbook 推測は、この境界では禁止します。
- なお `KernelWorkbookResolverService.ResolveOrOpen(...)` 系は、業務都合で open 内包責務を残した限定境界として扱います。
- `KernelWorkbookResolverService` が未 open Kernel workbook を一時 open した場合の close ownership は `KernelWorkbookAccessResult.CloseIfOwned()` に寄せます。呼び出し側は result 型を通じて閉じ、既に open 済みの Kernel workbook は勝手に close しません。
- Kernel workbook の選択は引き続き `SYSTEM_ROOT` 文脈に閉じ、`GetOpenKernelWorkbook()` のような文脈なし再検索へ戻しません。

## 不明点

- Kernel ブックや Base ブックのシート内部仕様は、この文書では詳細化していません。
- CASE 判定に使われるすべての DocProperty の運用意図までは、コードだけでは確定しません。
- 会計書類セット判定に使うシート構成の業務上の意味は、この文書では扱いません。
