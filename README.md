# 案件情報System

案件情報System は、Excel を起点に CASE ブック作成、文書作成、会計書類セット作成、案件一覧連携を行う Windows 向け業務ツールです。実装の中心は Excel VSTO Add-in で、補助として Word VSTO Add-in と Excel Launcher を持ちます。

この README は概要のみを扱います。詳細仕様は `docs/` 配下を参照してください。

## このシステムの前提

- Excel を UI として利用する業務システムである
- WorkbookRole によりブックの役割を判定する
- Add-in が全体制御の中心である
- ファイルコピー（Base → CASE）を基本とする

## 全体像

- `Kernel`
  - `案件情報System_Kernel.xlsx` を起点に、設定・画面遷移・CASE 作成を扱います。
- `Base`
  - `案件情報System_Base.xlsx` を CASE 作成時のコピー元として扱います。
- `CASE`
  - 個別案件の Excel ブックです。TaskPane から文書作成や会計書類セット作成を行います。
- `Excel Add-in`
  - WorkbookRole 判定、TaskPane 表示、CASE 作成、文書作成、会計書類セット作成、画面制御を担当します。
- `Word Add-in`
  - Word 側の表示補助と ContentControl の一括置換機能を担当します。
- `Excel Launcher`
  - 実行ファイルと同じ場所にある `案件情報System_Kernel.xlsx` を開きます。

## 読み順

1. `docs/architecture.md`
2. `docs/flows.md`
3. `docs/deployment.md`
4. `docs/ui-policy.md`
5. `docs/git-operations.md`
6. `AGENTS.md`

## ドキュメント一覧

- [docs/architecture.md](docs/architecture.md)
  - システム構成、WorkbookRole、Add-in の役割、サービス構成の大枠。
- [docs/flows.md](docs/flows.md)
  - 新規 CASE 作成、CASE 表示、文書作成、会計書類セット、CASE クローズ、TaskPane 更新。
- [docs/deployment.md](docs/deployment.md)
  - 実行時フォルダ構成、Addins 配置、雛形フォルダ、実行ファイルと Kernel の関係。
- [docs/ui-policy.md](docs/ui-policy.md)
  - 待機 UI、画面表示制御、前面化制御、Excel ウィンドウ復旧。
- [docs/git-operations.md](docs/git-operations.md)
  - 基準点固定、ブランチ運用、マージ時の注意点。
- [AGENTS.md](AGENTS.md)
  - ドキュメント更新ルール。

## 補足

- 各帳票の詳細仕様、差し込み内容、業務ルールはこの一式の対象外です。
- `.txt` 制御ファイルは存在説明に留め、内容や運用ルールは扱いません。
- コードとドキュメントに差異がある場合は、現行 `main` の実コードを優先してください。
