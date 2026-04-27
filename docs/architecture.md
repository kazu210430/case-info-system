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
- Accounting 系
  - 会計書類セット作成、会計ブック制御、補助フォーム、保存別名処理。
- TaskPane 系
  - スナップショット構築、描画、リフレッシュ調停、Window 単位の表示管理。
- Infrastructure 系
  - Excel / Word Interop、パス互換、フォルダ表示、ウィンドウ復旧、ログなど。

## TaskPane と HOME の位置づけ

- CASE 向け UI は主に Excel の Custom Task Pane として表示されます。
- TaskPane のタイトルは `案件情報System` で、左ドックに配置されます。
- Kernel HOME は TaskPane ではなく、WinForms の独立フォームとして表示されます。

## 不明点

- Kernel ブックや Base ブックのシート内部仕様は、この文書では詳細化していません。
- CASE 判定に使われるすべての DocProperty の運用意図までは、コードだけでは確定しません。
- 会計書類セット判定に使うシート構成の業務上の意味は、この文書では扱いません。
