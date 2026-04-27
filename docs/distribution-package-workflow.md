# 配布パッケージ運用仕様

## 1. この文書の目的
この文書は、案件情報System を当面 ZIP で配布するための運用仕様を固定するものです。
開発・配布する側と利用者側の役割分担、および配布物の基準点を明確にすることを目的とします。

## 2. 基本方針
- 開発用フォルダと配布用フォルダは物理的に分離する
- 開発用と配布用は `案件情報System` 配下にまとめる
- ルートを分散させない
- `案件情報System` 配下の開発用 / 配布用を、配布運用の基準パスとして扱う
- 配布用フォルダは毎回新規生成する
- 開発用フォルダは Kernel / Base / 雛形 / PDF / logs の正本にしない
- ビルド成果物（`案件情報System.exe` / `Addins`）は Release 出力物から取得する
- 実行時資産（Kernel / Base / 雛形 / PDF）は実行時ルート直下の正本から取得する
- `logs` はコピー元を持たず、配布時に空フォルダを生成する
- 開発用フォルダ直下のファイルはコピー元にしない
- 開発用 Kernel / Base は正規化しない
- docprops 正規化は配布用にコピー済みの Kernel / Base にだけ行う
- `利用開始ガイド.pdf` はコピー時リネームで配布する
- 利用者には PowerShell 実行や docprops 操作を求めない
- 当面の配布方式は ZIP とし、将来の MSI 化は別検討とする

## 3. 想定フォルダ
- 実行時ルート: `C:\Users\kazu2\Documents\案件情報System`
- 開発用: `C:\Users\kazu2\Documents\案件情報System\開発用`
- 配布用: `C:\Users\kazu2\Documents\案件情報System\配布用`
- 最終配布物: `C:\Users\kazu2\Documents\案件情報System\案件情報System.zip`

この構成は、誤操作防止と構造の一貫性確保のために固定します。

## 4. 配布用フォルダに含めるもの
配布用フォルダには、次を含める前提で固定します。

- `案件情報System.exe`
- `案件情報System_Kernel.xlsx`
- `案件情報System_Base.xlsx`
- `利用開始ガイド.pdf`
- `Addins` フォルダ
- `雛形` フォルダおよびその中身一式
- `logs` フォルダ

PDF の正本は実行時ルート直下の `案件情報System_利用開始ガイド.pdf` とし、配布時は `利用開始ガイド.pdf` にリネームします。
`logs` は既存内容をコピーせず、配布時に空フォルダとして生成します。
`利用開始ガイド.pdf` は現状では仮案ですが、当面は配布物に含めます。
将来的には作り直す前提で扱います。

## 5. 役割分担
### 開発・配布する側
配布できる状態に整える人です。
Release、配布用生成、正規化、ZIP 化までを担当します。

### 利用者側
受け取って使う人です。
ZIP 展開と起動だけを担当します。

## 6. 開発・配布する側の標準フロー
1. 開発用フォルダで開発・実機確認する
2. Release ビルドする
3. `案件情報System.exe` と `Addins` を Release 出力物から選別する
4. 配布用フォルダを新規作成する
5. 実行時ルート直下の正本から Kernel / Base / 雛形 を配布用フォルダへコピーする
6. 実行時ルート直下の `案件情報System_利用開始ガイド.pdf` を `利用開始ガイド.pdf` としてコピーする
7. `logs` フォルダを空で生成する
8. 配布用フォルダにコピー済みの Kernel / Base の docprops を明示パス指定で正規化する
9. 雛形コピー時は `~$*` を除外する
10. 配布用フォルダを ZIP 化する
11. 利用者向けには ZIP 展開後のフォルダ名が `案件情報System` になる形で渡す
12. ZIP を利用者へ渡す

## 7. コピー元の前提
- ビルド成果物（`案件情報System.exe` / `Addins`）は Release 出力物をコピー元とします
- 実行時資産（Kernel / Base / 雛形 / PDF）は実行時ルート直下の正本をコピー元とします
- `logs` はコピーせず、配布時に空フォルダを新規作成します
- PDF の正本ファイル名は `案件情報System_利用開始ガイド.pdf` とし、配布時のファイル名は `利用開始ガイド.pdf` とします
- 雛形コピー時は Office 一時ファイル `~$*` を除外します
- 開発用フォルダ直下に存在する実行ファイル、Workbook、`Addins`、その他の運用ファイルは配布用コピー元にしません
- この前提により、Debug 混入や古いファイル混入を防ぎます

## 8. docprops 正規化の位置づけ
- 正規化対象は配布用フォルダ内の Kernel / Base のみです
- 開発用の Kernel / Base は対象外です
- 実行時ルート直下の正本 Kernel / Base に対しては実行しません
- 既存の `Normalize-DistributionWorkbookDocProps.ps1` は責務変更せずに使います
- 引数なし実行は避け、配布用にコピー済みファイルを明示パス指定する前提とします

## 9. 開発・配布する側の禁止事項
- 開発用 Kernel / Base を正規化しない
- 引数なしで Normalize スクリプトを実行しない
- 配布用フォルダを手で中途半端に編集しない
- 古い配布用フォルダを使い回さない
- Release ビルド前の古い `Addins` を配らない
- 開発用フォルダ直下のファイルをそのまま配布しない
- `logs` の中身を配布用へコピーしない
- 雛形コピー時に `~$*` を含めない

## 10. 利用者へ渡すもの
- `案件情報System.zip`
- `利用開始ガイド.pdf`

利用者側には、ZIP 展開後のフォルダ名が `案件情報System` になる形で渡します。
`配布用` は開発・配布する側の作業名であり、利用者側には見せません。

## 11. 今の段階での理想形
- 開発・配布する側: ワンクリックで `案件情報System.zip` を作る
- 利用者側: ZIP を展開して `案件情報System.exe` を起動する

この運用が固まった後に、将来の MSI 化を検討します。

## 12. 将来の自動化対象
将来的には次の処理をワンクリック化することを想定します。

- 配布用フォルダ作成
- Release 出力物からの `案件情報System.exe` / `Addins` コピー
- 実行時ルート正本からの Kernel / Base / 雛形 / PDF コピー
- `利用開始ガイド.pdf` へのリネームコピー
- `logs` 空フォルダ生成
- `~$*` 除外
- docprops 正規化
- ZIP 作成
- ログ出力

ただし、この文書ではまだスクリプト作成や build 連携の実装には入りません。

## 13. 配布モードと署名方針

### A. 配布モード

案件情報System の配布は、当面次の 2 モードに分けて扱います。

#### Internal（所内配布・モニター配布）

- 所内の特定少数への検証・モニター配布を目的とします
- 現時点の配布スクリプトは Internal 配布を主対象とします
- ただし VSTO Release package 生成には署名が必要です
- 署名なし Release package は原則採用しません
- 暫定署名または所内専用コード署名証明書を使う方針とします
- `TemporaryKey.pfx` は Release 署名に使いません

#### Public（社外配布・不特定多数配布）

- 社外ユーザーへの正式配布を目的とします
- 正式な外部コード署名証明書を前提とします
- 証明書、配布方式、更新方式、必要に応じて MSI / Bootstrapper などを別フェーズで設計します
- Internal 配布とは混同しません

### B. 証明書管理方針

- `.pfx` は Git 管理しません
- `.pfx` は repo 外に保管します
- 証明書パスや thumbprint は build / package 実行時に外部から渡します
- パスワードは docs、script、csproj、commit に含めません
- 署名用証明書は Excel Add-in / Word Add-in の Release package 生成時に使います

### C. 現時点の blocker

- 配布スクリプトは `feature/build-distribution-package-script` に実装済みです
- ただし `dev\Deploy\Package\CaseInfoSystem.WordAddIn` は未生成です
- 現行仕様では Word Add-in の Release package 生成に署名が必要です
- Internal 用署名証明書の運用が決まるまで、配布スクリプトの本番実行と `main` merge は保留します

### D. build / package の扱い

- `.\build.ps1 -Mode Compile` は compile-only であり、Release package 生成の代替ではありません
- Release package 生成では署名前提の MSBuild 経路を使います
- `SignManifests=false` による Release package は原則採用しません
