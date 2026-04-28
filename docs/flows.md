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

1. `DocumentCommandService` が TaskPane の選択ボタンから文書キーを受け取ります。
2. `DocumentExecutionModeService` が `DocumentExecutionMode.txt` を読み込みます。
3. `DocumentExecutionEligibilityService` が登録済みテンプレートを前提に `DocumentTemplateResolver` で `templateSpec` を解決し、テンプレート種別、マクロ有無、出力先、CASE コンテキストを確認します。
4. `DocumentCommandService` は runtime の allowlist / review block を行わず、そのまま `DocumentCreateService` に進みます。`DocumentExecutionPolicyService` は現在 runtime 本線から外れており、この流れでは使われません。
5. `DocumentCreateService` が文書名を解決し、`DocumentOutputService` が出力先を解決します。
6. `MergeDataBuilder` が CASE データから差し込み用データを構築します。
7. `DocumentPresentationWaitService` が待機 UI を表示します。
8. `WordInteropService` が Word アプリケーションを取得または再利用します。
9. `WordInteropService` がテンプレートから文書を生成し、`DocumentMergeService` が差し込み処理を行います。
10. `DocumentMergeService` が ContentControl の除去処理を行います。
11. `DocumentSaveService` が保存し、`WordInteropService` が Word 文書を表示します。

### 現在の安全モデル

- 文書実行時の主防御は runtime allowlist gating ではなく、雛形登録前 validation です。
- `KernelTemplateSyncService` と `WordTemplateRegistrationValidationService` が、不正な雛形や不正な定義を登録前に排除します。
- 実行時は、登録済み `templateSpec` を前提に `DocumentExecutionEligibilityService` が基本適格性を確認します。
- `DocumentExecutionPolicyService` のソースは残っていますが、現在は runtime 本線から外れており、文書作成本線の runtime 実行可否には関与しません。

### 実行モードと制御ファイル

- 文書実行モードを読む `DocumentExecutionMode.txt` の存在はコードで確認できます。
- `DocumentExecutionAllowlist.txt`、`DocumentExecutionAllowlist.review.txt` の存在も確認できます。
- ただし、現行コードでは `DocumentExecutionAllowlist.txt`、`DocumentExecutionAllowlist.review.txt` は文書作成本線の runtime gating 本体としては使われていません。
- `allowlist` は runtime gating には使われておらず、過去の運用・検証用の残存要素です。今後の段階的撤去候補として扱います。
- `review` は runtime safety には寄与しておらず、PASS / HOLD / FAIL の記録媒体としての残存要素です。今後の段階的撤去候補として扱います。
- allowlist / review のファイルと csproj 同梱設定は残っています。専用 tools の撤去は完了しており、完全撤去は次フェーズです。
- `DocumentExecutionPolicyService` 自体もソース上は残っていますが、現在は runtime 本線から外れており、次フェーズの削除候補です。
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

1. `KernelTemplateSyncService` が Kernel ブックを取得し、`SYSTEM_ROOT\雛形` を登録対象フォルダとして解決します。
2. `KernelTemplateSyncService` が Kernel の管理シート `CaseList_FieldInventory` を読み取り、定義済み Tag 一覧を構築します。
3. `WordTemplateRegistrationValidationService` が雛形フォルダ直下の候補ファイルを走査します。
4. 各ファイルに対して登録前チェックを実施します。
5. OK 雛形のみを `shMasterList` / `雛形一覧` の一覧へ書き戻します。
6. NG 雛形は登録しません。
7. 登録除外理由と警告を結果メッセージに表示します。
8. `TASKPANE_MASTER_VERSION` を更新します。
9. Kernel 保存後に Base へ TaskPane 用 snapshot を更新します。
10. `MasterTemplateCatalogService.InvalidateCache()` を実行してキャッシュを無効化します。

この登録前 validation が、現行実装における文書作成フローの主防御です。runtime 側の allowlist / review 判定は、登録済みテンプレートの実行可否を直接制御していません。

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

## CASE クローズ

CASE クローズは `WorkbookLifecycleCoordinator` を起点として処理されます。CASE / Base 側では `CaseWorkbookLifecycleService` が処理します。

### 確認できる処理

1. `CaseWorkbookLifecycleService` が Workbook ライフサイクル調停に入ります。
2. `CaseWorkbookLifecycleService` が対象 Workbook が CASE / Base かどうかを確認します。
3. `CaseWorkbookLifecycleService` が managed close 中かどうかを確認します。
4. `CaseWorkbookLifecycleService` が dirty 状態であれば保存確認を表示します。
5. `CaseWorkbookLifecycleService` がクローズ後の後続処理を必要に応じて予約します。
6. `CaseWorkbookLifecycleService` が Workbook 状態や TaskPane 状態を片付けます。

### 確認できるダイアログ

- dirty 状態の保存確認として `保存しますか？` の Yes / No / Cancel が使われます。
- 新規作成直後の CASE には、保存後に保存先フォルダを開くか確認する後続導線があります。

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

## 不明点

- この文書の不明点は、該当する各節の `### 不明点` に記載します。
