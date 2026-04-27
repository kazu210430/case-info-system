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
2. `DocumentExecutionModeService` が実行モードを確認します。
3. `DocumentExecutionEligibilityService` が実行可否判定を行います。
4. `DocumentExecutionPolicyService` が Allowlist / Pilot 用の判定処理を行います。
5. `DocumentTemplateResolver` がテンプレートを解決します。
6. `DocumentExecutionEligibilityService` がテンプレート情報を含めて実行可否を判定します。
7. `DocumentCreateService` が文書名を解決し、`DocumentOutputService` が出力先を解決します。
8. `MergeDataBuilder` が CASE データから差し込み用データを構築します。
9. `DocumentPresentationWaitService` が待機 UI を表示します。
10. `WordInteropService` が Word アプリケーションを取得または再利用します。
11. `WordInteropService` がテンプレートから文書を生成し、`DocumentMergeService` が差し込み処理を行います。
12. `DocumentMergeService` が ContentControl の除去処理を行います。
13. `DocumentSaveService` が保存し、`WordInteropService` が Word 文書を表示します。

### 実行モードと制御ファイル

- 文書実行モードを読む `DocumentExecutionMode.txt` の存在はコードで確認できます。
- `DocumentExecutionPilot.txt`、`DocumentExecutionAllowlist.txt`、`DocumentExecutionAllowlist.review.txt` の存在も確認できます。
- ただし、各ファイルの中身や運用ルールはこの文書では扱いません。

### テンプレート配置

- `DocumentTemplateResolver` は `WORD_TEMPLATE_DIR` が設定されている場合はそちらを優先し、未設定時は `SYSTEM_ROOT\雛形` をテンプレート配置先として解決します。
- `DocumentTemplateResolver` は `.docx`、`.dotx`、`.dotm` を対応テンプレートとして扱います。
- `DocumentExecutionEligibilityService` は VSTO 実行可否判定時に、マクロ有効テンプレートを制限対象として扱います。

### 不明点

- 文書ごとの差し込み項目と命名規則の最終業務ルールは、コードだけでは確定しません。
- `DocumentExecutionMode.txt` などの制御ファイルの運用手順は、この文書では確定しません。

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
