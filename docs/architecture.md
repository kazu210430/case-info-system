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
  - CASE 初期化、dirty 状態管理、保存前後、クローズ前後の制御。
- Document 系
  - テンプレート解決、出力名解決、実行可否判定、Word 生成、保存、待機 UI。
  - `DocumentExecutionEligibilityService` は登録済みテンプレートを前提に、VSTO 実行に必要な基本適格性を確認します。
- `DocumentExecutionPolicyService` は allowlist / review 関連ファイルを読む互換レイヤーですが、現行実装では permissive に動作しており、runtime の実行可否そのものは制御しません。
  - `DocumentExecutionModeService` は mode の読取と運用スイッチ管理を担います。現行コードで確認できる主用途は Word warm-up 制御であり、gating 本体ではありません。
- Accounting 系
  - 会計書類セット作成、会計ブック制御、補助フォーム、保存別名処理。
- TaskPane 系
  - スナップショット構築、描画、リフレッシュ調停、Window 単位の表示管理。
- Infrastructure 系
  - Excel / Word Interop、パス互換、フォルダ表示、ウィンドウ復旧、ログなど。

## 雛形管理の設計方針

本システムでは、雛形の品質担保は登録時に行います。

- 実行時ではなく登録時に不正な雛形を `雛形一覧` から排除します。
- 実装上の検証は `CaseList_FieldInventory` を基準にした最小限の妥当性確認です。
- 雛形の修正責任は利用者側にあります。
- 文書実行時の安全性は runtime allowlist gating ではなく、登録前 validation によって担保します。
- 実行時は登録済み `templateSpec` を前提に処理し、`DocumentExecutionPolicyService` は現状 permissive な互換レイヤーとして残っています。

これにより次を狙います。

- TaskPane 表示の安定化
- 文書作成時エラーの削減
- 問題発生時の切り分け容易化

## Document 実行ポリシーの現状

- `allowlist`
  - 現状 runtime gating には使われていません。
  - 過去の運用・検証用の残存要素です。
  - 今後の段階的撤去候補として扱います。
- `review`
  - 現状 runtime safety には寄与していません。
  - PASS / HOLD / FAIL の記録媒体としての残存要素です。
  - 今後の段階的撤去候補として扱います。
- `mode`
  - runtime gating 目的ではありません。
  - 現行コードで確認できる主用途は Word warm-up 制御などの運用スイッチです。
  - allowlist / review とは分けて扱い、現時点では撤去対象に含めません。

## Document 系サービスの補足

- `DocumentExecutionPolicyService`
  - 現状は permissive な互換レイヤーです。
  - allowlist / review の過去互換要素を含みます。
  - 将来の段階的撤去対象候補です。
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

## 不明点

- Kernel ブックや Base ブックのシート内部仕様は、この文書では詳細化していません。
- CASE 判定に使われるすべての DocProperty の運用意図までは、コードだけでは確定しません。
- 会計書類セット判定に使うシート構成の業務上の意味は、この文書では扱いません。
