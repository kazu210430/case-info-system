# 主要フロー

## 対象

この文書では、コードから確認できる主要フローのみを扱います。帳票ごとの差し込み詳細や業務ルールは対象外です。

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
- `DocumentExecutionEligibilityService` は VSTO 実行可否判定時に、マクロ有効テンプレートを制限対象として扱います。

### 不明点

- 文書ごとの差し込み項目と命名規則の最終業務ルールは、コードだけでは確定しません。
- `DocumentExecutionMode.txt` などの制御ファイルの詳細な運用手順は、この文書では確定しません。

## 雛形登録・更新フロー

雛形登録・更新は `KernelCommandService` から `KernelTemplateSyncService` を呼び出して実行されます。利用者が配置した Word 雛形を検証し、適正なもののみを `雛形一覧` に登録する処理です。

### フロー

1. `KernelTemplateSyncService` が `GetOpenKernelWorkbook()` により Kernel ブックを取得し、`SYSTEM_ROOT\雛形` を登録対象フォルダとして解決します。
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

### 現状の Kernel workbook 選択仕様

- `KernelCommandService.Execute(context, actionId)` は `reflect-template` 分岐で `ExecuteReflectTemplate()` を呼びますが、`context` 自体は `KernelTemplateSyncService.Execute()` へ渡しません。
- `KernelTemplateSyncService.Execute()` は CASE workbook context を受け取らず、`_kernelWorkbookService.GetOpenKernelWorkbook()` の戻り値をそのまま雛形登録・更新対象の Kernel workbook として扱います。
- `KernelOpenWorkbookLocator.GetOpenKernelWorkbook()` は `_application.Workbooks` を先頭から列挙し、Kernel と判定された最初の workbook を返します。
- この経路では active workbook、visible workbook、`WorkbookContext.SystemRoot`、表示中の CASE workbook は判定材料に使われません。
- そのため、`MasterTemplateCatalogService` の cache 境界が resolved master path 単位に改善された後も、「どの root の Kernel workbook に対して雛形登録・更新を行うか」は upstream で探索順依存のまま残ります。
- 通常の単一 Kernel workbook 運用では問題化しにくいですが、複数 Kernel workbook や hidden workbook が同時にある場合は、利用者の意図と異なる Kernel workbook を操作対象にする余地があります。
- これは今回の cache 修正で混入した問題ではなく、既存の Kernel workbook 選択仕様の設計課題です。
- 将来は command / UI / CASE 文脈から `SYSTEM_ROOT` を明示的に渡し、その文脈で Kernel workbook を確定する改善を検討します。
- `GetOpenKernelWorkbook()` は便利関数として残す場合でも、複数 root を跨ぐ経路では使用範囲を限定する前提で扱います。

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

- Kernel 側から会計関連の同期フローに入る分岐もあります。
- 会計補助フォームや支払履歴取込などの関連機能は存在しますが、詳細仕様はこの文書では扱いません。

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
4. dirty close は `CaseWorkbookLifecycleService` が managed close を dispatcher 経由で予約し、`ManagedCloseState` のスコープ内で save 有無を処理して `workbook.Close(SaveChanges: false)` へ進めます。
5. managed close の内部、または clean close の before-close 処理では `PostCloseFollowUpScheduler` が予約されます。
6. `PostCloseFollowUpScheduler` は close 後に対象 workbook が残っていないことを確認し、Excel busy なら retry し、visible workbook が 1 つも無い場合だけ Excel 終了を試みます。
7. close 継続時の workbook state / accounting state / TaskPane pane の片付けは `WorkbookLifecycleCoordinator` 側が後続で行います。

dirty path の大まかな順序は `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` です。

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
2. `KernelTemplateSyncService` が `TASKPANE_MASTER_VERSION` を `+1` します。
3. この version 更新では内容差分の有無を見ません。雛形登録・更新は利用者の明示操作なので、成功時に無条件で `+1` してよい仕様です。
4. `KernelTemplateSyncService` が TaskPane 用 snapshot を組み立て、Base に `TASKPANE_BASE_SNAPSHOT_*` と `TASKPANE_BASE_MASTER_VERSION` を埋め込みます。
5. Base にも `TASKPANE_MASTER_VERSION` を保存し、新規 CASE が version ごと引き継げる状態にします。
6. `MasterTemplateCatalogService.InvalidateCache(openKernelWorkbook)` を実行して、選択された Kernel workbook から解決した `SYSTEM_ROOT` 文脈の master catalog cache を無効化します。

補足:

- 現在の実装では、この `openKernelWorkbook` 自体が `GetOpenKernelWorkbook()` の探索順に依存して選ばれます。
- したがって cache invalidate の境界は root 単位に改善済みですが、その upstream にある Kernel workbook 選択境界は将来課題として残ります。

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
- 雛形登録・更新成功時の `TASKPANE_MASTER_VERSION` 無条件 `+1` を差分チェック方式に変えないこと。
- `DocumentTemplateResolver` の CASE cache 優先を安易に変更しないこと。
- Base snapshot 埋め込みを削らないこと。
- `TASKPANE_SNAPSHOT_CACHE_COUNT` を削除対象に含めないこと。

## 不明点

- この文書の不明点は、該当する各節の `### 不明点` に記載します。
