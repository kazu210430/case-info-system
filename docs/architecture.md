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
- TaskPane 系
  - スナップショット構築、描画、リフレッシュ調停、Window 単位の表示管理、CASE pane UIイベント dispatch。
  - 現在は `TaskPaneManager` を facade に、`TaskPaneHostFlowService` が refresh-time host flow、`TaskPaneHostLifecycleService` が registry-backed host lifecycle、`TaskPaneDisplayCoordinator` が show/hide 調停、`TaskPaneRefreshOrchestrationService` / `TaskPaneRefreshCoordinator` が event-side refresh orchestration を担う構造です。
  - その周辺で `TaskPaneRefreshPreconditionPolicy`、`CasePaneSnapshotRenderService`、`CasePaneCacheRefreshNotificationService`、`TaskPaneActionDispatcher` などへ主責務が分離されています。
- Infrastructure 系
  - Excel / Word Interop、パス互換、フォルダ表示、ウィンドウ復旧、ログなど。

## Excel Application 状態管理（second wave 完了）

この節は、`docs/flows.md` の CASE 作成 / CASE 表示 / 会計書類セット / CASE ライフサイクルと、`docs/ui-policy.md` の UI 制御原則を前提に、Excel `Application` 状態管理の正本を固定するための節です。`Application` 状態管理は個別機能の都合ではなく、shared app / isolated app / retained instance をまたぐ横断関心として扱います。

### 1. 設計原則（確定事項）

- `shared app`
  - 利用者が操作中の Excel `Application` は caller-owned であり、業務処理側は終了責務を持ちません。
  - `Application` 状態は必ず snapshot / restore 前提で扱います。
  - `DisplayAlerts` は API 境界（`Save` / `Close` / `Quit`）に限定します。
- `isolated app`
  - 専用に生成する hidden Excel `Application` は `Create -> Open -> Work -> Save/Close -> Quit` の lifecycle に閉じます。
  - `Quit` は restore ではなく cleanup です。
- `retained instance`
  - 例外は `CaseWorkbookOpenStrategy` の hidden application cache のみです。
  - これは one-shot isolated lifecycle の一般形ではなく、再利用のために idle へ戻す retained instance として扱います。

### 2. 境界の定義

- `shared app` と `isolated app` の違い
  - `shared app` は現在の Add-in が共有している `Application` を使う経路です。状態変更は restore まで含めて shared app 内で閉じ、`Application.Quit()` は行いません。
  - `isolated app` は専用 `Application` を生成する経路です。workbook だけでなく `Application` 自体の ownership も作成側サービスが持ち、`finally` で cleanup まで完結させます。
- hidden open の扱い
  - hidden open は可視状態の戦略であり、ownership の例外ではありません。
  - shared app での hidden open は workbook を hidden で開いても shared app の一部です。
  - isolated app での hidden open は one-shot isolated lifecycle の内部処理です。
- ownership（誰が close するか）
  - 自分で open した workbook は、その open を行ったサービスが close します。
  - shared app で既に open 済みの workbook を再利用した場合、再利用側は `Application` の close / quit を行いません。
  - isolated app では、生成したサービスが workbook `Close` と `Application.Quit` の両方を担当します。
  - retained instance では workbook close までは各 session が担当し、cached `Application` 自体の破棄は cache 側の健康判定 / timeout / poison 時にだけ行います。

### 3. 完了条件

- `Application` 状態管理が、CASE 作成・hidden open・保存・close・quit をまたぐ横断関心として一貫して説明できること。
- isolated instance が「例外的な逃げ道」ではなく、one-shot な hidden 作業を閉じるためのルールとして説明できること。
- retained instance が一般ルールではなく、`CaseWorkbookOpenStrategy` cache のみの例外だと説明できること。

### 4. あえてやっていないこと

次は second wave の対象外とし、third wave で扱います。

- `StatusBar`
- `Visible` / `WindowState`
- UI 正本化

これらは `Application` 状態管理そのものではなく、進捗表示・window 表示復旧・UI policy 正本化の論点として分離します。

### 5. 残課題

- retained instance の契約は、必要なら `CaseWorkbookOpenStrategy` の内部 route 名ではなく設計用語として別途明文化します。
- route 名 `experimental-shared` などは ownership モデルと一致しないため、必要なら命名改善を行います。

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
| `TASKPANE_MASTER_VERSION` | Kernel, Base, CASE | Master 一覧に対応する現在 version。CASE 側では CASE cache がどの master を前提にしたかの記録にも使います。 | 雛形登録・更新成功時に Kernel で `+1`。Base 反映時に Base にも保存。CASE では Base からの昇格時と MasterList rebuild 時に更新されます。 |
| `TASKPANE_BASE_MASTER_VERSION` | Base, 新規 CASE | Base に埋め込まれた snapshot がどの master version 由来かを示します。 | 雛形登録・更新成功後、Base snapshot 更新時に書き込みます。CASE では Base 埋込状態を引き継ぎます。 |
| `TASKPANE_SNAPSHOT_CACHE_COUNT` | CASE | CASE cache の chunk 数です。`0` は cache 無効を表します。 | CASE cache 保存時に更新。案件一覧登録後は削除せず `0` に戻します。 |
| `TASKPANE_SNAPSHOT_CACHE_XX` | CASE | 表示中 Pane と整合する CASE snapshot 本体です。`XX` は 2 桁連番です。 | Base から CASE cache へ昇格する時、または MasterList rebuild で再構築した時に保存します。案件一覧登録後は chunk を削除します。 |
| `TASKPANE_BASE_SNAPSHOT_COUNT` | Base, 新規 CASE | Base に埋め込んだ snapshot の chunk 数です。 | 雛形登録・更新成功後、Base snapshot 更新時に保存します。新規 CASE はこの埋込値を引き継ぎます。 |
| `TASKPANE_BASE_SNAPSHOT_XX` | Base, 新規 CASE | Base に埋め込んだ TaskPane snapshot 本体です。`XX` は 2 桁連番です。 | 雛形登録・更新成功後、Base snapshot 更新時に保存します。既存 CASE の案件一覧登録後整理では触りません。 |

### 補足

- Base に snapshot / version を埋め込む目的は、新規 CASE 作成直後に不要な MasterList rebuild を避けることです。
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
- なお `KernelWorkbookResolverService.ResolveOrOpen(...)` 系は、業務都合で open 内包責務を残した暫定境界として扱います。

## 不明点

- Kernel ブックや Base ブックのシート内部仕様は、この文書では詳細化していません。
- CASE 判定に使われるすべての DocProperty の運用意図までは、コードだけでは確定しません。
- 会計書類セット判定に使うシート構成の業務上の意味は、この文書では扱いません。
