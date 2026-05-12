# 主要フロー

## 対象

この文書では、コードから確認できる主要フローのみを扱います。帳票ごとの差し込み詳細や業務ルールは対象外です。

- TaskPane refresh の policy 正本: `docs/taskpane-refresh-policy.md`
- TaskPane 表示回復の current-state 正本: `docs/taskpane-display-recovery-current-state.md`

## 新規 CASE 作成

新規 CASE 作成は `KernelCaseCreationService` を起点として処理されます。コード上では少なくとも次のモードが存在します。

- `NewCaseDefault`
- `CreateCaseSingle`
- `CreateCaseBatch`

### 基本の流れ

1. `KernelCaseCreationService` が Kernel から `SYSTEM_ROOT`、`NAME_RULE_A`、`NAME_RULE_B`、Base の場所を解決します。
2. `KernelCaseCreationService` が作成先フォルダを決定します。
3. `KernelCaseCreationService` が CASE フォルダ名と CASE ブック名を決定します。
4. `KernelCaseCreationService` が Base ブックを物理コピーして CASE ブックを作成します。
5. `CaseWorkbookInitializer` が CASE ブックに対して初期化処理を実行します。
6. モードに応じて `KernelCasePresentationService` が CASE 表示またはフォルダ表示へ進めます。

### モード差分

- `NewCaseDefault`
  - `KernelCaseCreationCommandService` が Kernel の `DEFAULT_ROOT` を優先使用します。
  - `KernelCaseCreationCommandService` が未設定時にフォルダ選択を行い、その結果を Kernel に保存します。
- `CreateCaseSingle`
  - `KernelCaseCreationCommandService` がフォルダ選択を行って 1 件作成します。
- `CreateCaseBatch`
  - `KernelCaseCreationCommandService` がフォルダ選択を行う複数作成向けの分岐です。
  - 作成後は `KernelCasePresentationService` が CASE ブックを直接表示せず、フォルダ表示へ進める実装があります。

### CASE新規作成専用 managed hidden create session

1. `KernelCaseCreationService.CreateSavedCase(...)` は現コードで `ShouldUseHiddenCreateSession() == true` のため、全モードで `CreateSavedCaseWithoutShowing(...)` を通します。
2. `CaseWorkbookOpenStrategy.OpenHiddenWorkbook(...)` が hidden create route を選びます。優先順は `app-cache`、未使用時は `legacy-isolated`、環境変数 `CASEINFO_EXPERIMENT_DEDICATED_HIDDEN_INNER_SAVE` 指定時だけ `experimental-isolated-inner-save` です。
3. `NewCaseDefault` / `CreateCaseSingle` は hidden create session で `InitializeForVisibleCreate(...)` を実行し、`NormalizeInteractiveWorkbookWindowStateBeforeSave(...)` では save 前に workbook window を `Visible=true` へ戻さず、必要なら `WindowState=xlNormal` だけを整えて save / hidden session close を完了します。
4. interactive 表示開始後は `KernelHomeForm.CloseKernelAfterCaseCreation()` と `KernelWorkbookCloseService` が Kernel HOME close を完了します。CASE 作成フロー中は `KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(...)` により Kernel を前景へ戻しません。
5. interactive route の hidden session close 後は、`KernelCasePresentationService` が shared app の `OpenHiddenForCaseDisplay(...)` で CASE を reopen し、表示責務を shared/current app 側へ渡します。
6. `CreateCaseBatch` は hidden create session で `InitializeForHiddenCreate(...)` を使い、`NormalizeBatchWorkbookWindowStateBeforeSave(...)` で save 前に workbook window を `visible + normal` へ正規化したあとに save / close して完了します。CASE workbook の reopen は行わず、フォルダ表示と HOME 継続へ分岐します。

補足:

- hidden create session の owner は `KernelCaseCreationService` です。hidden workbook open / close mechanics は `CaseWorkbookOpenStrategy` が担い、retained hidden app-cache を使う場合だけ cached `Application` 自体の owner は `CaseWorkbookOpenStrategy` に残ります。
- interactive route では、保存前正規化を理由に hidden create session 中の workbook window を visible 化しません。実機確認で `Visible=true` の前倒しは白フラッシュと終了時 Excel / Book1 発生を再露出させたため、作成側は hidden のまま初期化・保存・handoff までを完了し、visible / normal の最終表示は `KernelCasePresentationService` が担います。
- `experimental-isolated-inner-save` は route 名どおり、current/shared app ではなく dedicated hidden `Application` を生成し、close 時の inner save を含む経路です。
- 互換のため旧環境変数 `CASEINFO_EXPERIMENT_SHARED_HIDDEN_EXCEL` でも同 route に到達しますが、契約上の正本は `CASEINFO_EXPERIMENT_DEDICATED_HIDDEN_INNER_SAVE` です。
- `app-cache` は one-shot isolated session ではなく、`CaseWorkbookOpenStrategy` が所有する retained hidden app-cache の例外です。
- hidden Excel / isolated app / retained hidden app-cache / white Excel lifecycle の current-state は `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`、target-state は `docs/hidden-excel-isolated-app-white-excel-lifecycle-target-state.md`、lifecycle / outcome / trace / owner vocabulary は `docs/hidden-excel-lifecycle-outcome-vocabulary.md`、white Excel prevention / recovery の current-state と target boundary は `docs/white-excel-prevention-boundary-current-state.md` を参照します。この節は新規 CASE 作成フローの順序、同文書は instance / visibility / cleanup owner の正本です。

### 不明点

- `CaseWorkbookInitializer` が初期化時に書き込む全項目の一覧は、この文書では確定しません。

## CASE 表示

CASE 表示は `KernelCasePresentationService` を起点として処理されます。

### 確認できる処理

1. `KernelCasePresentationService` が作成済み CASE のパスを既知パスとして登録します。
2. `KernelCasePresentationService` が一時的な TaskPane 表示抑止を設定します。
3. `KernelCasePresentationService` が必要に応じて非表示オープンを経由して表示準備を行います。
4. `ExcelWindowRecoveryService` が Excel ウィンドウ復旧を試行します。
5. `KernelCasePresentationService` が CASE の Workbook Window を可視化します。
6. `TaskPaneRefreshOrchestrationService` が TaskPane の準備完了表示を予約します。
7. `KernelCasePresentationService` が初期カーソル位置を CASE HOME 上の定義済み位置へ移動します。

### 注意

- CASE 表示には待機 UI が使われます。
- 画面ちらつき抑止や一時的な pane 抑止が入るため、通常の WorkbookOpen だけではなく表示専用の補助処理があります。
- ready-show / retry / protection の詳しい policy は `docs/taskpane-refresh-policy.md` を正本とします。

## 文書作成ボタン

文書作成ボタンは `DocumentCommandService` を起点として処理されます。TaskPane のアクション種別には `doc`、`accounting`、`caselist` があります。

### `doc` の流れ

1. `TaskPaneActionDispatcher` が CASE pane の選択ボタンから `actionKind` と文書キーを受け取ります。
2. `TaskPaneBusinessActionLauncher` が `doc` 実行前に `DocumentNamePromptService.TryPrepare` を呼び、文書名入力ダイアログの初期値を準備します。
3. `DocumentNamePromptService` は `DocumentTemplateLookupService.TryResolveFromCaseCache` を通して CASE cache だけを参照し、`caption` を prompt 初期値に使います。
4. CASE cache に対象 key が無い場合、文書名入力側では master catalog へフォールバックせず、空欄のまま prompt を開きます。
5. prompt で確定した値は `DocumentNameOverrideScope` により一時 DocProperty として保持されます。
6. `TaskPaneBusinessActionLauncher` が `DocumentCommandService` へ文書キーを渡します。
7. `DocumentExecutionModeService` が `DocumentExecutionMode.txt` を読み込みます。
8. `DocumentExecutionEligibilityService` が登録済みテンプレートを前提に `DocumentTemplateResolver` で `templateSpec` を解決し、テンプレート種別、マクロ有無、出力先、CASE コンテキストを確認します。
9. `DocumentTemplateResolver` は `DocumentTemplateLookupService.TryResolveWithMasterFallback` を使い、まず CASE cache を参照し、解決できない場合だけ CASE workbook から解決した `SYSTEM_ROOT` 文脈の `MasterTemplateCatalogService` master catalog にフォールバックします。
10. `DocumentCommandService` は runtime の allowlist / review block を行わず、そのまま `DocumentCreateService` に進みます。
11. `DocumentCreateService` が `templateSpec.DocumentName` と一時 override を使って文書名を解決し、`DocumentOutputService` が出力先を解決します。
12. `MergeDataBuilder` が CASE データから差し込み用データを構築します。
13. `DocumentPresentationWaitService` が待機 UI を表示します。
14. `WordInteropService` が Word アプリケーションを取得または再利用します。
15. `WordInteropService` がテンプレートから文書を生成し、`DocumentMergeService` が差し込み処理を行います。
16. `DocumentMergeService` が ContentControl の除去処理を行います。
17. `DocumentSaveService` が保存し、`WordInteropService` が Word 文書を表示します。

補足:

- `DocumentNamePromptService` が使う snapshot / CASE cache は表示状態に合わせた補助情報であり、文書生成の正本ではありません。
- 保存・生成・実行判断は、`DocumentExecutionEligibilityService` と `DocumentTemplateResolver` が正本側の確認を行う前提です。

### 現在の安全モデル

- 文書実行時の主防御は runtime allowlist gating ではなく、雛形登録前 validation です。
- `KernelTemplateSyncService` と `WordTemplateRegistrationValidationService` が、不正な雛形や不正な定義を登録前に排除します。
- 実行時は、登録済み `templateSpec` を前提に `DocumentExecutionEligibilityService` が基本適格性を確認します。
- `DocumentExecutionEligibilityService` は fail-closed を維持し、template 解決、template path、出力先、CASE context のいずれかが欠ける場合は実行へ進めません。
- allowlist / review の旧 runtime policy サービスは撤去済みで、文書作成本線の runtime 実行可否には関与しません。

### 実行モードと制御ファイル

- 文書実行モードを読む `DocumentExecutionMode.txt` の存在はコードで確認できます。
- `allowlist` / `review` の runtime policy は撤去済みです。
- allowlist / review の config ファイル、csproj 同梱設定、専用 tools、旧 runtime policy サービスは撤去済みです。
- pilot は runtime 本線で未使用だったため撤去済みです。
- `mode` は runtime gating 目的ではありません。現行コードで確認できる主用途は Word warm-up 制御などの運用スイッチであり、allowlist / review とは分けて扱い、現時点では撤去対象に含めません。

### テンプレート配置

- `DocumentTemplateResolver` は `WORD_TEMPLATE_DIR` が設定されている場合はそちらを優先し、未設定時は `SYSTEM_ROOT\雛形` をテンプレート配置先として解決します。
- `DocumentTemplateResolver` は `.docx`、`.dotx`、`.dotm` を対応テンプレートとして扱います。
- `DocumentExecutionEligibilityService` は VSTO 実行可否判定時に、`.doc` / `.docm` を実行対象にせず、`.dotm` をマクロ有効テンプレートとして制限対象に扱います。
- したがって `.dotm` 対応は resolver 側の lookup 補助であり、そのまま実行許可を意味しません。

### 不明点

- 文書ごとの差し込み項目と命名規則の最終業務ルールは、コードだけでは確定しません。
- `DocumentExecutionMode.txt` などの制御ファイルの詳細な運用手順は、この文書では確定しません。

## 雛形登録・更新フロー

雛形登録・更新は `KernelCommandService` から `KernelTemplateSyncService` を呼び出して実行されます。利用者が配置した Word 雛形を検証し、適正なもののみを `雛形一覧` に登録する処理です。

### フロー

1. `KernelCommandService` が Kernel pane 由来の `WorkbookContext` を `KernelTemplateSyncService` へ渡し、`KernelTemplateSyncService` がその文脈の Kernel workbook または `WorkbookContext.SystemRoot` に対応する open Kernel workbook を解決します。
2. `KernelTemplateSyncService` が Kernel の管理シート `CaseList_FieldInventory` を読み取り、定義済み Tag 一覧を構築します。
3. `WordTemplateRegistrationValidationService` が雛形フォルダ直下の候補ファイルを走査します。
4. 各ファイルに対して登録前チェックを実施します。
5. OK 雛形のみを `shMasterList` / `雛形一覧` の一覧へ書き戻します。
6. NG 雛形は登録しません。
7. 登録除外理由と警告を結果メッセージに表示します。
8. `TASKPANE_MASTER_VERSION` を更新します。
9. Kernel 保存後に Base へ TaskPane 用 snapshot を更新します。
10. `MasterTemplateCatalogService` の当該 `SYSTEM_ROOT` 文脈に対応する master catalog cache を無効化します。

この登録前 validation が、現行実装における文書作成フローの主防御です。runtime 側の allowlist / review 判定は、登録済みテンプレートの実行可否を直接制御していません。

### publication side effects の固定点

- publication side effects は `PublicationExecutor` に集約されます。
- 順序は `WriteToMasterList -> TASKPANE_MASTER_VERSION 更新 -> Kernel save -> Base snapshot sync -> InvalidateCache` で固定します。
- preflight failure では副作用を発生させません。
- kernel save failure では Base sync / invalidate へ進めません。
- base sync failure では invalidate は実行し、success + warning の扱いを維持します。
- `SYSTEM_ROOT` 文脈、invalidate API、cache key 解決方式は変えません。

### Kernel workbook 選択仕様

- `KernelCommandService.Execute(context, actionId)` は `reflect-template` 分岐で `ExecuteReflectTemplate(context)` を呼び、`WorkbookContext` を `KernelTemplateSyncService.Execute(context)` へ引き渡します。
- `KernelTemplateSyncService.Execute(context)` は `WorkbookContext` を必須入力として扱い、対象 Kernel workbook を `_kernelWorkbookService.ResolveKernelWorkbook(context)` で確定します。
- `KernelOpenWorkbookLocator.ResolveKernelWorkbook(context)` は、まず `context.Workbook` が Kernel ならその workbook を使い、それ以外は `WorkbookContext.SystemRoot` に対応する Kernel workbook path を解決して open workbook を特定します。
- この経路では、複数 Kernel workbook や hidden workbook が同時に存在しても、雛形登録・更新、snapshot 反映、cache invalidate は要求元の `SYSTEM_ROOT` 文脈に対応する Kernel workbook へ閉じます。
- `GetOpenKernelWorkbook()` のような文脈なし API は使わず、雛形登録・更新フローの Kernel workbook 選択は `ResolveKernelWorkbook(context)` に限定します。
- 本フェーズの到達点として、Kernel 操作は `WorkbookContext` を唯一の入口とし、root 不一致は補正せず fail-closed とします。
- 許容される open は、明示的な `WorkbookContext` / `SYSTEM_ROOT` 文脈から行う open と user action 起点の open です。
- 禁止される open は、context-less fallback open と暗黙の workbook 推測です。
- `KernelWorkbookResolverService.ResolveOrOpen(...)` 系は、業務都合により open 内包責務を残した将来課題として扱います。

### 登録前チェック

- ファイル名先頭の key No. が 2 桁かを確認します。
- key No. が `01` から `99` の範囲内かを確認します。
- key 重複を確認します。
- 拡張子を確認します。候補走査対象は `.docx` / `.dotx` / `.docm` / `.dotm` ですが、`.docm` / `.dotm` は登録不可です。
- Word ファイルとして読み取れるかを確認します。
- テキスト / リッチテキスト ContentControl の Tag を検証します。

### Tag 検証

- `CaseList_FieldInventory` に定義された Tag のみ許可します。
- `Date` は特例として許可します。
- 未定義 Tag がある場合は登録不可です。

### 警告

- Tag 未設定のテキスト項目は警告になります。
- 警告のみの場合は登録を許可します。

### 非対象

- 非テキスト ContentControl は無視します。

### 出力

- 登録成功件数
- 登録除外件数
- 警告件数
- 各ファイルの除外理由
- 各ファイルの警告内容

## Base HOME フィールドキー同期

Base HOME フィールドキー同期は、リボンの `Base定義更新` から `KernelCommandService` と `BaseHomeFieldInventorySyncService` を呼び出して実行されます。Base `ホーム` A列を変更した後、Kernel `CaseList_FieldInventory.ProposedFieldKey` を同じキーへ合わせるための補助入口です。

### フロー

1. リボン入口が active workbook または Kernel HOME binding から `WorkbookContext` を組み立て、`KernelCommandService` が同期サービスへ渡します。
2. `BaseHomeFieldInventorySyncService` が対象 Kernel workbook と `SYSTEM_ROOT` の一致を確認します。
3. `SYSTEM_ROOT` から Base workbook を解決し、Base `ホーム` A列のキーを読み取ります。
4. Kernel `CaseList_FieldInventory` の既存行を読み取り、`SourceCell=B{Base HOME row}` を安定した行対応として使います。
5. 空欄、重複、制御文字、対応行欠落、同期後の `ProposedFieldKey` 重複、重要キー変更を検出した場合は書き込み前に中断します。
6. 問題がない場合だけ、対応する既存行の `ProposedFieldKey` を更新します。
7. `SourceCell`、`ProposedNamedRange`、`Label`、`DataType`、`NormalizeRule`、その他既存列は変更しません。
8. 同期結果を表示し、Word 雛形 CC Tag の更新と雛形登録・更新の実行を案内します。

### 非対象

- Word 雛形 CC Tag の自動書き換え
- `CaseList_FieldInventory` の行削除、全再作成、列構成変更
- 雛形登録検証ルールの緩和
- 文書作成フローの差し込み仕様変更

## Kernel ユーザー情報反映

Kernel ユーザー情報反映は `KernelUserDataRegistrationExecutionService` または `KernelCommandService` から `KernelUserDataReflectionService` を呼び出して実行されます。Kernel `shUserData` の値を Base HOME と会計書類セットへ反映する補助フローです。

### フロー

1. `KernelUserDataReflectionService` が `WorkbookContext` から Kernel workbook と `SYSTEM_ROOT` を確定します。
2. shared Excel 側では quiet mode を適用し、Kernel `shUserData` の snapshot を読み取ります。
3. Base / Accounting workbook が既に open なら、その workbook を再利用して反映し、save はしても close はしません。
4. Base / Accounting workbook が未 open なら、`KernelUserDataReflectionService` が service-owned な `managed hidden reflection session` を開始します。
5. session では hidden な isolated `Application` を生成し、対象 workbook を open して window を hidden のまま反映します。
6. 反映後は save 前に owned workbook window visibility を restore し、自分で open した対象 workbook だけを save / close し、生成した isolated `Application` だけを `Quit` します。

### managed hidden reflection session の境界

- owner は `KernelUserDataReflectionService` です。
- 対象 workbook が既に open なら hidden session は開始せず、その workbook を再利用します。
- hidden session は shared Excel を `Quit` しません。
- shared Excel 側の `DisplayAlerts` / `EnableEvents` / `ScreenUpdating` は quiet mode の restore まで含めて呼び出し側で閉じます。
- save 前の visibility restore は保存状態正規化のための owner-side cleanup であり、shared/current app の表示経路へ昇格させる意味ではありません。
- cleanup は `CloseWorkbookQuietly`、`Application.Quit`、`ComObjectReleaseService.FinalRelease` まで含めて hidden session 内で完結します。
- session の目的は、未 open の Base / Accounting workbook へ反映するための owner 付き hidden 作業を、処理後に orphaned `EXCEL.EXE` を残さず閉じることです。
- hidden reflection session と他の hidden lifecycle の接続は `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`、owner / protocol の target-state は `docs/hidden-excel-isolated-app-white-excel-lifecycle-target-state.md` を参照します。

## 会計書類セット

会計書類セットは `AccountingSetCommandService` を起点として処理されます。CASE では `AccountingSetCreateService` が作成処理を実行します。

### CASE から作成する流れ

1. `AccountingSetCreateService` が CASE コンテキストを取得します。
2. `AccountingTemplateResolver` がテンプレートファイルを `SYSTEM_ROOT\雛形` から解決します。
3. `DocumentOutputService` が出力先フォルダを解決し、`AccountingSetNamingService` が出力ファイル名を決定します。
4. `AccountingSetPresentationWaitService` が待機 UI を表示します。
5. `AccountingSetCreateService` がテンプレート Excel をコピーします。
6. `AccountingWorkbookService` が作成した会計ブックを現在の Excel アプリケーションで開きます。
7. `AccountingWorkbookService` が会計ブックを可視化します。
8. `AccountingSetCreateService` が次の DocProperty を設定またはコピーします。

- `CASEINFO_WORKBOOK_KIND=ACCOUNTING_SET`
- `SOURCE_CASE_PATH`
- `SYSTEM_ROOT`
- `NAME_RULE_A`
- `NAME_RULE_B`

9. `AccountingWorkbookService` が顧客名や関連情報を対象シートへ反映します。
10. `AccountingWorkbookService` が入力開始シートまたはセルへ誘導します。
11. `TaskPaneRefreshOrchestrationService` が TaskPane 表示を準備します。

### 補足

- `AccountingSetCreateService.Execute()` は `AccountingTemplateResolver`・`DocumentOutputService`・`AccountingSetNamingService` で template path / output folder / output path を決めてから `File.Copy(...)` に進む。
- `File.Copy(...)` が最初の実副作用ポイントで、その後に `SuppressPath(...)`・workbook open・failure cleanup delete が接続されるため、副作用境界として扱う。
- Kernel 側から会計関連の同期フローに入る分岐もあります。
- 会計補助フォームや支払履歴取込などの関連機能は存在しますが、詳細仕様はこの文書では扱いません。
- `AccountingWorkbookService.BeginInitializationScope()` は、初期化中だけ `Application.ScreenUpdating` と `Application.EnableEvents` の現在値を退避し、両方を `false` に設定する。
- 同 scope は `AccountingSetCreateService` では `using` で使われ、DocProperty 設定、初期セル反映、代理人反映の範囲だけを囲う。
- scope 終了時は `ApplicationStateScope.Dispose()` により、退避した `ScreenUpdating` と `EnableEvents` を元値へ戻す。
- `AccountingSetCreateService.Execute()` では、会計ブック open と visible 化の後に初期化ブロックへ入り、`BeginInitializationScope()` の内側で `CASEINFO_WORKBOOK_KIND` などの DocProperty 設定、顧客名の初期セル書込、`ReflectLawyers()` による代理人反映をまとめて実行する。
- 同初期化ブロックの外側には、workbook window visible、`ActivateInvoiceEntry()`、`ShowWorkbookTaskPaneWhenReady()` による ready-show handoff、overflow 時の `MessageBox`、失敗時 cleanup が残る。
- 会計系の save-as は `AccountingSaveAsService` から `AccountingWorkbookService.SaveAsMacroEnabled()` を呼び出す。
- `SaveAsMacroEnabled()` は save-as 専用境界として `Application.DisplayAlerts`、`Application.EnableEvents`、`Application.ScreenUpdating` の現在値を退避し、`SaveAs` 実行中だけ 3 つとも `false` に設定する。
- `SaveAsMacroEnabled()` は `try/finally` で `SaveAs` 後に 3 つの Application 状態を元値へ戻す。
- `AccountingWorkbookService` の cell write / range 操作には、`WriteCell`、`WriteSameValueToSheets`、range copy、named range write / clear、print area / alignment / sort などが含まれる。
- 現行テストで呼び出し観測が確認できる会計 workbook 書込は、`AccountingSetCreateService` と `AccountingSetKernelSyncService` が使う `WriteCell` と `WriteSameValueToSheets` が中心である。
- named range / copy / clear / format / print area / sort 系の操作は、現行テストから直接の観測点を確認できない。

### Kernel からの会計書類セット同期（`AccountingSetKernelSyncService`）

1. `AccountingSetKernelSyncService` は Kernel workbook から会計書類セット template path を解決し、対象 workbook が既に open かを確認します。
2. 対象 workbook が既に open なら、それを再利用して反映 / save し、close はしません。
3. 対象 workbook が未 open なら、別 `Excel.Application` は生成せず、`kernelWorkbook.Application` を shared/current app として使い、`AccountingWorkbookService.OpenInCurrentApplication(...)` で open します。
4. 自分で open した workbook は hidden window のまま反映 / save し、最後に owned workbook だけを quiet close します。
5. `DisplayAlerts` / `ScreenUpdating` / `EnableEvents` の restore は `ExcelApplicationStateScope` の局所スコープに閉じます。

補足:

- この経路は managed hidden session の例外には含めません。shared/current app 前提の補助処理として固定します。
- 不要な別 `Excel.Application` fallback は撤去済みで、今後の docs 上も再許容しません。

### 不明点

- 会計書類セットで各シートや各セルに反映する値の業務上の意味は、コードだけでは確定しません。

## CASE ライフサイクル

CASE / Base の lifecycle は `WorkbookLifecycleCoordinator` を入口にし、主調停は `CaseWorkbookLifecycleService` が担います。

### 初回初期化

1. `WorkbookOpen` / `WorkbookActivate` で `CaseWorkbookLifecycleService.HandleWorkbookOpenedOrActivated(...)` が呼ばれます。
2. `CaseWorkbookLifecycleInitializationPolicy` が対象外 / 既初期化 / Base / CASE を判定します。
3. CASE の初回初期化では `WorkbookRoleResolver.RegisterKnownCaseWorkbook(...)` と `KernelNameRuleReader.TryReadForCaseWorkbook(...)` による `NAME_RULE_A` / `NAME_RULE_B` 同期が行われます。
4. Base は初期化済みマークだけを更新します。

### dirty 状態

1. `SheetChange` では `CaseWorkbookSheetChangePolicy` が対象外 / managed close 中 / transient pane suppression 中を除外します。
2. 対象 workbook だけ session dirty として記録します。

### クローズ

1. `WorkbookBeforeClose` では `CaseWorkbookBeforeClosePolicy` が `Ignore` / `SuppressPromptForManagedClose` / `PromptForDirtySession` / `SchedulePostCloseFollowUp` を判定します。
2. dirty session では `CaseClosePromptService` が `保存しますか？` の Yes / No / Cancel を表示します。
3. `KernelCaseCreationCommandService` から pending が付与されていた workbook では、保存先フォルダが解決できる場合だけ folder offer を出し、`CaseFolderOpenService` が必要に応じて Explorer を起動します。
4. dirty close は `CaseWorkbookLifecycleService` が managed close を dispatcher 経由で予約し、`ManagedCloseState` のスコープ内で save 有無を処理し、今回安定化対象の managed close 経路では `WorkbookCloseInteropHelper.CloseWithoutSave(workbook)` を使って `false, Type.Missing, Type.Missing` の optional 引数を明示した close へ進めます。
5. managed close の内部、または clean close の before-close 処理では `PostCloseFollowUpScheduler` が予約されます。
6. `PostCloseFollowUpScheduler` は close 後に対象 workbook が残っていないことを確認し、Excel busy なら retry し、visible workbook が 1 つも無い場合だけ Excel 終了を試みます。`Quit` 成功後は終了中 `Application` を restore せず、`DisplayAlerts` の restore は失敗時だけに限定します。
7. close 継続時の workbook state / accounting state / TaskPane pane の片付けは `WorkbookLifecycleCoordinator` 側が後続で行います。

dirty path の大まかな順序は `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` です。

## Kernel HOME close / managed close / post-close quit

この節で固定するのは、close / quit のうち `Kernel HOME close`、`Kernel managed close`、`CASE managed close`、`post-close quit` だけです。全 workbook close 経路の一般ルールではありません。他の helper 非経由 close は別途棚卸し対象として扱います。

### Kernel HOME close

1. `KernelHomeForm` が close 意思を受け、`FormClosing` で service-mediated close を要求します。
2. `KernelWorkbookService.RequestCloseHomeSessionFromForm(...)` が backend close を調停します。
3. backend close が pending / rejected の間は `FormClosing` を cancel し、Form を閉じません。
4. backend close 成功後だけ `FormClosed` で `FinalizePendingHomeSessionCloseAfterFormClosed()` が走り、HOME session / binding / visibility を解放します。
5. close 失敗時は fail-closed で終了し、Form / binding / visibility を維持します。

### Kernel / CASE managed close

- 今回安定化対象の managed close 経路では `WorkbookCloseInteropHelper` を使います。
- `Workbook.Close(SaveChanges: false)` のような named argument は使いません。
- save ありの Kernel managed close では `Type.Missing, Type.Missing, Type.Missing`、save なしの Kernel / CASE managed close では `false, Type.Missing, Type.Missing` を明示して渡します。
- close 後に対象 workbook を再参照しません。

### managed close / post-close quit の最小 COM 原則

- 今回安定化対象の managed close / quit 経路では、`Save` / `DisplayAlerts` / `Close` / `Quit` 以外の COM 操作を増やしません。
- この経路では `ExcelApplicationStateScope` を使いません。
- `DisplayAlerts` は個別の `try/finally` または同等の局所 restore で扱います。
- `Quit` 成功後は終了中 `Application` を restore しません。
- `Quit` 失敗時だけ `DisplayAlerts` を restore します。

### CASE post-close quit

1. `PostCloseFollowUpScheduler` が visible workbook の有無を確認します。
2. visible workbook が残っていなければ `Quit` を試みます。
3. `Quit` 成功後は終了中 `Application` を restore しません。
4. `Quit` 失敗時だけ `DisplayAlerts` を restore します。
5. 設計目標は CASE close 後に白 Excel を残さないことです。

white Excel prevention / recovery の current-state、`targetWorkbookStillOpen` と `visibleWorkbookExists` の意味、G-1 で触るべき安全単位は `docs/white-excel-prevention-boundary-current-state.md` を参照します。

### Shadow copy / 実機反映

- Excel が起動中だと古い shadow copy DLL が使われ続けることがあります。
- 実機確認前は Excel を完全終了します。
- 実行 DLL の確認は `Runtime execution observed` ログの `assemblySha256` を使います。

### 既知の残課題

- helper 非経由 close が `MasterWorkbookReadAccessService`、`CaseWorkbookOpenStrategy` などに残っています。
- `KernelUserDataReflectionService` の未 open Base / Accounting 反映は、上記の `managed hidden reflection session` として明文化した例外です。
- `WorkbookPromptSuppressionHelper` の `Workbook.Saved` 操作は今回対象外です。
- これらは別途棚卸し対象であり、今回 docs の確定範囲外です。

### 補助サービス

- `CaseClosePromptService`
  - dirty prompt のタイトル解決と `保存しますか？` ダイアログ、created case folder offer prompt を担当します。
- `CaseFolderOpenService`
  - 保存先フォルダ解決、存在確認、Explorer 起動を担当します。
- `KernelNameRuleReader`
  - open 中 Kernel workbook または package `docProps/custom.xml` から name rule を読み取ります。
- `ManagedCloseState`
  - managed close の入れ子状態を workbook key 単位で管理します。
- `PostCloseFollowUpScheduler`
  - close 後 follow-up、Excel busy retry、no visible workbook 時の Excel 終了判定を担当します。

## TaskPane 更新

TaskPane 更新は `WorkbookLifecycleCoordinator`、`WindowActivatePaneHandlingService`、`TaskPaneRefreshOrchestrationService` を起点として処理されます。

- retry / protection / ready-show の詳細 policy は `docs/taskpane-refresh-policy.md` を参照します。

### Kernel HOME unbound セッション

- startup 時の自動表示や明示的な HOME 表示では、Kernel HOME が valid binding を持たない `unbound` 状態で開く場合があります。
- `unbound` HOME は placeholder-only とし、Kernel が既に open でも自動 bind せず、Kernel workbook / Kernel window の選択や復元を行いません。
- Kernel が open でない場合も、`unbound` HOME 表示のために Kernel workbook を探したり開いたりしません。
- `unbound` HOME を閉じるときも、Kernel workbook は managed close 対象や window 復元対象に含めません。
- startup 文脈で使う open Kernel workbook の有無は表示可否判定の事実であり、HOME 表示後に 1 冊の Kernel workbook を選ぶための入力には使いません。
- `unbound` HOME は fail-closed セッションであり、binding 不成立を補正するための fallback open は行いません。

### 更新の入口

- `TaskPaneRefreshOrchestrationService` が起動時の再描画要求を扱います。
- `WorkbookLifecycleCoordinator` が `WorkbookOpen` を入口にします。
- `WorkbookLifecycleCoordinator` が `WorkbookActivate` を入口にします。
- `WindowActivatePaneHandlingService` が `WindowActivate` を入口にします。
- `TaskPaneRefreshOrchestrationService` が明示的な再描画要求を扱います。
- `TaskPaneRefreshOrchestrationService` が準備完了後の遅延表示を扱います。

### 構築内容

`TaskPaneRefreshOrchestrationService` が更新を調停し、`TaskPaneRefreshCoordinator` と `TaskPaneManager` が TaskPane の表示内容をスナップショットとして組み立てます。

- `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` が `WorkbookOpen` 直後の window-dependent refresh skip 境界を定義します。
- `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` はこの policy を利用する側であり、skip 条件を個別に重複保持しません。
- `TaskPaneRefreshOrchestrationService` は `RefreshPreconditionEvaluator`、`RefreshDispatchShell`、`PendingPaneRefreshRetryState`、`WorkbookPaneWindowResolver` に helper split 済みで、現在は順序調停寄りに整理されています。
- `TaskPaneRefreshCoordinator` は `KernelFlickerTrace` の structured trace を維持し、`04150a7` で obsolete route に付随していた duplicate plain log を削除済みです。

- 特別ボタン
  - `案件一覧登録`
  - `会計書類セット`
- タブ
  - `全て` を含むタブ構成
- 文書ボタン
  - Master 一覧やキャッシュから再構成されるボタン群

### 取得元

- CASE ブックの DocProperty キャッシュ
- Base に埋め込まれたキャッシュ
- Master ブックの一覧シート

### 補足

- TaskPane は左ドックです。
- Window 単位で管理され、再利用と再描画の判定があります。
- 一時抑止、遅延再試行、WindowActivate 専用処理が実装されています。

### WorkbookOpen と window 確定の境界

- `WorkbookOpen` は workbook が開いた通知です。
- `WorkbookOpen` 時点では `ActiveWorkbook` と `ActiveWindow` が未確定な場合があります。
- `WorkbookOpen` 時点で workbook 自体は取得できても、対象 workbook の visible window や active window がまだ解決できないケースがあります。
- `WorkbookActivate` は、対象 workbook が active workbook として前面系の文脈に乗った後続イベントです。
- `WindowActivate` は、対象 window が実際に activate された後続イベントです。

確認できた順序:

1. `WorkbookOpen`
2. `WorkbookActivate`
3. `WindowActivate`

扱いの原則:

- workbook-only 処理は `WorkbookOpen` で扱ってよいです。
- window-dependent 処理は `WorkbookActivate` 以降、必要なら `WindowActivate` 以降を安全境界として扱います。
- `WorkbookOpen` 直後の `ActiveWorkbook` / `ActiveWindow` を前提に、window 解決・表示・前面化・pane 対象決定を確定させない方針を維持します。
- `WorkbookOpen` 直後に workbook は取得できても window が未解決な refresh は、`TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` により skip し、後続の `WorkbookActivate` / `WindowActivate` 側へ委ねます。

補足:

- `ResolveWorkbookPaneWindow` が安全に成功する条件は、対象 workbook の visible window が取得できること、または active workbook が対象 workbook と一致し active window が取得できることです。
- 単体生成 CASE の再オープン調査では、`WorkbookOpen` 時点で `ActiveWorkbook` / `ActiveWindow` が空のため window 解決に失敗し、その後 `WorkbookActivate` で回復するログが確認されました。
- `TaskPaneManagerOrchestrationPolicyTests` は、この skip 境界を `TaskPaneRefreshPreconditionPolicy` に対して直接検証します。
- startup context 系の再分解を再開する前に、このイベント境界の安定化を優先する必要があります。

### CASE 文書ボタンパネル更新仕様

#### 目的

CASE の文書ボタンパネル更新仕様は、次を同時に満たすためのものです。

- 新規 CASE は最新の文書ボタン定義で開始する
- 既に開いている CASE の Pane は勝手に変えない
- 不要な TaskPane 再構築を避ける
- 表示中 Pane と文書実行時の解決元を一致させる

#### 雛形登録・更新時の流れ

雛形登録・更新成功時は、次の順で TaskPane 更新元を進めます。

1. `KernelTemplateSyncService` が `shMasterList` / `雛形一覧` を更新します。
2. `KernelTemplateSyncService` が `TASKPANE_MASTER_VERSION` を `yyyyMMddNNN` 形式の `long` として更新します。
3. この version 更新では内容差分の有無を見ません。雛形登録・更新は利用者の明示操作なので、成功時に無条件で同日連番を進めてよい仕様です。旧整数方式の値、過去日、不正値は次回更新時に当日 `yyyyMMdd001` へ移行し、候補 version が既存 version 以下になる場合は時計ずれ対策として単調増加を優先します。
4. `KernelTemplateSyncService` が TaskPane 用 snapshot を組み立て、Base に `TASKPANE_BASE_SNAPSHOT_*` と `TASKPANE_BASE_MASTER_VERSION` を埋め込みます。
5. Base にも `TASKPANE_MASTER_VERSION` を保存し、新規 CASE が version ごと引き継げる状態にします。
6. `MasterTemplateCatalogService.InvalidateCache(openKernelWorkbook)` を実行して、選択された Kernel workbook から解決した `SYSTEM_ROOT` 文脈の master catalog cache を無効化します。

補足:

- 現在の実装では、この `openKernelWorkbook` は `ResolveKernelWorkbook(context)` によって要求元の `SYSTEM_ROOT` 文脈へ閉じた workbook として選ばれます。
- したがって cache invalidate の境界は root 単位に改善済みですが、その upstream にある Kernel workbook 選択境界は将来課題として残ります。
- `MasterTemplateCatalogService` と `TaskPaneSnapshotBuilderService` は、どちらも `MasterWorkbookReadAccessService` を共有して Master path 解決と read-only open を揃えています。

#### 新規 CASE 作成時の流れ

新規 CASE 作成では、TaskPane 更新仕様として次を前提にします。

1. `KernelCaseCreationService` が Base を物理コピーして CASE を作成します。
2. Base に埋め込まれていた `TASKPANE_BASE_SNAPSHOT_*` と `TASKPANE_BASE_MASTER_VERSION` は、新規 CASE にもそのまま入ります。
3. CASE 側では `TaskPaneSnapshotCacheService` などの処理により、必要時に Base 埋込 snapshot / version を CASE cache へ昇格できます。
4. このため、新規 CASE は原則として最新 snapshot を持った状態で始まり、初回表示時に不要な `shMasterList` 再構築を避けます。

#### 既存 CASE を開く時の流れ

既存 CASE の TaskPane 更新元は、`TaskPaneSnapshotBuilderService` で次の順に解決されます。

1. `TASKPANE_SNAPSHOT_CACHE_*` が有効で、かつ CASE の `TASKPANE_MASTER_VERSION` が最新 master version 以上なら CASE cache を使います。
2. CASE cache が空、または古い場合は `TASKPANE_BASE_SNAPSHOT_*` と `TASKPANE_BASE_MASTER_VERSION` を確認します。
3. Base 側が有効なら、その snapshot を CASE cache へ昇格して使います。
4. CASE cache / Base cache のどちらも使えない場合だけ `shMasterList` から再構築します。
5. ただし、いったん Pane / host / control が生成された後は、その CASE を閉じるまで表示中の Pane を維持します。

補足:

- Base 埋込 snapshot と CASE cache はいずれも派生 cache であり、global 正本ではありません。
- TaskPane snapshot は表示用断面であり、保存・生成・実行判断の正本にしてはいけません。

#### WorkbookActivate / WindowActivate の扱い

- `WorkbookActivate` と `WindowActivate` は、既存 host の再表示・再利用を優先する仕様です。
- `TaskPaneHostReusePolicy` は、同じ CASE workbook に対する `WorkbookActivate` / `WindowActivate` を host 再利用対象として扱います。
- この経路では毎回 version 比較して Pane を再生成する仕様ではありません。
- したがって、開いている CASE が、後から行われた雛形登録・更新に追随しないことは現行仕様です。
- この仕様は、表示中の CASE の UI を利用者の明示操作なしに変えないために維持します。

#### 表示中 Pane と文書実行時の cache 利用

- `DocumentNamePromptService` は文書名入力 UI 用の補助情報だけを扱い、CASE cache から `caption` を引けた場合にだけ prompt 初期値へ反映します。
- `DocumentNamePromptService` は実行可否判定や実体テンプレートファイル解決の正本ではありません。
- `DocumentNamePromptService` は CASE cache miss 時に master fallback しません。文書名入力 UI は、表示中 Pane と整合する CASE cache 表示状態に従います。
- `DocumentTemplateResolver` は、まず `TaskPaneSnapshotCacheService` を使って CASE cache から文書キーに対応する定義を解決します。
- CASE cache に解決対象がない場合だけ、対象 CASE workbook から解決した `SYSTEM_ROOT` 文脈の `MasterTemplateCatalogService` master catalog にフォールバックします。
- master fallback は `DocumentTemplateResolver` 側の実行時解決責務として扱います。
- そのため、開いている CASE では表示中 Pane と整合する CASE cache を使い続けてよく、master version だけを見ると stale に見える場合でも直ちに問題扱いしません。
- 文書名入力 UI と文書実行は責務を分離し、前者は現在の CASE 表示状態、後者は実行可能なテンプレート解決を担います。
- 文書ボタン実行も、表示中 Pane と一致する cache を優先してよい仕様です。
- 最新雛形を使いたい場合は、CASE を開き直して新しい snapshot 解決経路に入り直す運用とします。

#### 案件一覧登録後の cache 整理

- 案件一覧登録後は、CASE 側の `TASKPANE_SNAPSHOT_CACHE_COUNT` を `0` に更新して CASE cache を無効化します。
- 同時に `TaskPaneSnapshotCacheService.ClearCaseSnapshotCacheChunks()` により `TASKPANE_SNAPSHOT_CACHE_01` などの chunk を削除します。
- `TASKPANE_SNAPSHOT_CACHE_COUNT` 自体は削除せず、`0` として維持します。
- `TASKPANE_BASE_SNAPSHOT_*` と `TASKPANE_BASE_SNAPSHOT_COUNT` / `TASKPANE_BASE_MASTER_VERSION` には触れません。

#### 触ってはいけない注意点

- `WorkbookActivate` / `WindowActivate` の host 再利用経路を安易に問題扱いしないこと。
- 開いている CASE の Pane / host / control を close まで維持する仕様を壊さないこと。
- 雛形登録・更新成功時の `TASKPANE_MASTER_VERSION` 無条件更新を差分チェック方式に変えないこと。
- `DocumentTemplateResolver` の CASE cache 優先を安易に変更しないこと。
- Base snapshot 埋め込みを削らないこと。
- `TASKPANE_SNAPSHOT_CACHE_COUNT` を削除対象に含めないこと。

## 雛形更新直後の新規 CASE 作成〜初回表示〜再オープンの観測メモ（2026-05-08）

この節は、第1安全単位 merge 後に 1 度だけ観測した表示不安定について、観測事実と未確定事項を分けて残す補足です。原因断定や恒常不具合の認定はまだ行いません。

### 観測事実

- 雛形更新後、そのまま Kernel から新規 CASE を作成した。
- その新規 CASE は初回表示の準備中にぐるぐる状態になった。
- いったん Excel を終了した。
- その当該 CASE を開いたところ白 Excel になった。
- 白 Excel はウインドウ再表示で復元した。
- その後、同じ操作で再現を試したが、最初から問題なく表示された。
- したがって、現時点では恒常再現する不具合か、一過性の表示タイミング問題か、根本原因が潜んでいるかは未確定である。
- 以前の「古いCASEを開いたらぐるぐる」という要約は不正確であり、今回の一次観測は「雛形更新直後に新規 CASE を作成した直後の初回表示」から始まっている。

### この観測に関係しうる既存フロー

- 新規 CASE 作成直後は `CaseWorkbookInitializer` と `CaseTemplateSnapshotService` が Base 埋込 snapshot / master version を CASE へ引き継ぐ。
- 初回表示では `KernelCasePresentationService` が hidden create session 後の表示 handoff を行い、`WorkbookWindowVisibilityService` が workbook window の visible 化、`ExcelWindowRecoveryService` が initial recovery、`ShowWorkbookTaskPaneWhenReady(...)` が ready-show 予約を担当する。
- 再オープン後の TaskPane snapshot 解決では `TaskPaneSnapshotBuilderService` が `CASE cache -> Base cache -> MasterListRebuild` の順で解決し、refresh 完了後の foreground recovery は `TaskPaneRefreshCoordinator` が判断する。

### まだ断定できないこと

- 今回のぐるぐるが、新規 CASE 作成直後の初回表示タイミングだけで起きた一過性の現象か。
- `ShowWorkbookTaskPaneWhenReady(...)` 周辺の ready-show handoff が関与したか。
- `ExcelWindowRecoveryService` / `TaskPaneRefreshCoordinator` の foreground recovery / window recovery が白 Excel の復元前後で関与したか。
- `TaskPaneSnapshotBuilderService` の `MasterListRebuild` が今回の再オープン時に実際に走っていたか。
- version mismatch が起点だったか。
- Excel window visibility の変化が主因だったか。
- 第1安全単位 `KernelTemplateSyncPreparationService` 分離が直接原因かどうか。

### 追加で見たいログと確認項目

- 新規 CASE 作成直後から初回表示完了までの `NewCaseVisibilityObservation` と `KernelFlickerTrace`。
- `KernelCasePresentationService` の `initial-recovery-completed`、`post-release-suppression-prepared`、`ready-show-requested`、`ShowCreatedCase workbook window made visible before ready-show`。
- 再オープン時の `TaskPaneSnapshotBuilderService` による `caseMasterVersion`、`embeddedMasterVersion`、`latestMasterVersion`、`Task pane snapshot source=CaseCache|BaseCache|MasterListRebuild`。
- `TaskPaneRefreshCoordinator` の `foreground-recovery-decision`、`final-foreground-guarantee-start`、`final-foreground-guarantee-end`。
- `ExcelWindowRecoveryService` の `Excel window recovery evaluated` と `Excel window recovery mutation trace`。
- `WorkbookOpen -> WorkbookActivate -> WindowActivate` の順序と、その時点での対象 workbook / window 解決可否。
- `Application.Visible`、`ScreenUpdating`、workbook window `Visible`、`WindowState` の復元有無。

### 次の安全単位候補と判断

- 現時点では、白 Excel 対策ガードや stale CASE reopen 前提の分岐を先に足すより、追加観測を挟む判断を優先する。
- 次の安全単位候補は、`TaskPaneSnapshotBuilderService` 周辺の version mismatch / `MasterListRebuild` 観測整理と、`KernelCasePresentationService` / ready-show / `TaskPaneRefreshCoordinator` / `ExcelWindowRecoveryService` 周辺の表示回復経路の観測整理である。
- 第1安全単位の直接原因とは断定しない。
- stale CASE reopen が原因だとも断定しない。

## 不明点

- この文書の不明点は、該当する各節の `### 不明点` に記載します。
