# 実行時配置

## 概要

案件情報System には、実行時ルートとソースルートがあります。この文書では、主に実行時に必要な配置を整理します。

## ルートの分け方

- 実行時ルート
  - `C:\Users\kazu2\Documents\案件情報System`
- ソースルート
  - `C:\Users\kazu2\Documents\案件情報System\開発物`

ドキュメント上で単に「ルート」と書くと混同しやすいため、必ず区別して記述します。

## 実行時フォルダ構成

実行時ルートには、少なくとも次の構成要素があります。

- `案件情報System.exe`
- `案件情報System_Kernel.xlsx`
- `案件情報System_Base.xlsx`
- `Addins\`
- `雛形\`

このほか、運用用ファイルや補助ディレクトリが存在する場合があります。

## 実行ファイルと Kernel の関係

- `CaseInfoSystem.ExcelLauncher` は、実行ファイルと同じ場所にある `案件情報System_Kernel.xlsx` を開きます。
- したがって、実行ファイルと Kernel ブックは同一ディレクトリに置かれる前提があります。

## Addins 配置

実行時の Add-in 配置先として、少なくとも次のディレクトリが使われます。

- `Addins\CaseInfoSystem.ExcelAddIn`
- `Addins\CaseInfoSystem.WordAddIn`

ソース上のプロジェクト設定でも、既定の出力先としてこの配置が参照されています。

## VSTO署名運用

- `*_TemporaryKey.pfx` は開発用の一時証明書として扱います。
- 一時証明書はリポジトリ管理しません。各開発者のローカル環境でのみ保持します。
- Release 配布物の署名には、リポジトリ外で管理される配布用証明書が必要です。
- Release build は `ReleaseCertificateKeyFile` または `ManifestKeyFile` で外部の `.pfx` を明示しないと失敗します。
- `*_TemporaryKey.pfx` を Release 署名へ流用した場合も build は失敗します。
- Debug build の成功と、Release 配布署名の成功は別扱いです。`docs/ci-tests.md` の compile-only 確認は Release 配布証明書の代替ではありません。

## Excel Add-in 配下

`Addins\CaseInfoSystem.ExcelAddIn` 配下には、少なくとも次の種類のファイルが存在します。

- `CaseInfoSystem.ExcelAddIn.dll`
- `CaseInfoSystem.ExcelAddIn.vsto`
- `.config` ファイル
- 文書実行制御用 `.txt` ファイル
- VSTO / Office 関連 DLL

### 文書実行制御用ファイル

存在が確認できるファイル:

- `DocumentExecutionMode.txt`

allowlist / review の config ファイルと旧 runtime policy サービスは撤去済みです。この文書では mode ファイルの存在説明のみを扱い、中身や運用ルールは扱いません。

## Word Add-in 配下

`Addins\CaseInfoSystem.WordAddIn` 配下には、少なくとも次の種類のファイルが存在します。

- `CaseInfoSystem.WordAddIn.dll`
- `CaseInfoSystem.WordAddIn.vsto`
- `.config` ファイル
- VSTO / Office 関連 DLL

## 雛形フォルダ

既定の雛形配置先は `SYSTEM_ROOT\雛形` です。実行時ルート直下に `雛形` フォルダが存在する構成を確認できます。

このフォルダには少なくとも次の種類のファイルがあります。

- Word 文書テンプレート
- 会計書類セット用 Excel テンプレート
- サブフォルダ配下の補助テンプレート類

各帳票の詳細仕様はこの文書では扱いません。

## ソースルート側の構成

`開発物` 配下には、少なくとも次の主要ディレクトリがあります。

- `dev`
  - Add-in 本体、Launcher、Tests、Deploy を含みます。
- `docs`
  - ドキュメント置き場です。
- `scripts`
  - 補助スクリプト置き場です。

## 配置と DocProperty の関係

- `SYSTEM_ROOT` はテンプレート解決や関連ブック解決の基準として使われます。
- CASE 作成時には `SYSTEM_ROOT`、`NAME_RULE_A`、`NAME_RULE_B` が CASE ブックへ設定されます。
- 会計書類セット作成時には、上記に加えて `CASEINFO_WORKBOOK_KIND` と `SOURCE_CASE_PATH` が設定されます。

## 不明点

- 配布パッケージの正式な作成手順は、この文書では扱いません。
- 実行時に必須となる Office / VSTO のインストール条件は、プロジェクト参照から存在は確認できますが、利用端末側の完全な要件としては未整理です。
