# 配布担当者ガイド

## 1. この文書の役割
- この文書は、配布する側の作業手順をまとめたものです
- 配布パッケージの設計方針は `distribution-package-workflow.md` を参照します
- 利用者向け手順は `distribution-package-user-guide.md` を参照します

## 2. 配布作業の全体像
- 配布担当者が行う作業は、基本的に `CreateDistributionPackage.bat` の実行です
- バッチ実行後に、生成された ZIP を確認して利用者へ渡します
- 手作業で配布用フォルダを組み立てたり、個別スクリプトを順番に実行したりする運用は標準にしません

## 3. 標準手順
- `開発用` フォルダを開きます
- `CreateDistributionPackage.bat` を実行します
- 完了後に `案件情報System.zip` が生成されたことを確認します
- 必要に応じて `配布用` フォルダの中身も確認します
- 利用者には `案件情報System.zip` を渡します
- `dev/Deploy/Package` は Release Add-in package の中間生成物であり、利用者へ渡す配布正本ではありません

## 4. 自動で行われる処理
- Release Add-in package 生成
- 配布用フォルダ生成
- 必要ファイルコピー
- `初回セットアップ.bat` 同梱
- `CaseInfoSystem.Internal.cer` 同梱
- docprops 正規化
- `logs` 空生成
- ZIP生成

## 5. 自動処理の中身
- Release Add-in package を生成します
- 既存の `配布用` フォルダがあれば作り直します
- Release 出力物から `案件情報System.exe` と `Addins` をコピーします
- 実行時ルートの正本から `Kernel`、`Base`、`雛形`、PDF をコピーします
- `distribution-assets\初回セットアップ.bat` を配布用へコピーします
- Release VSTO manifest から `CaseInfoSystem.Internal.cer` を書き出して同梱します
- 配布用にコピー済みの `Kernel` / `Base` に対して docprops を正規化します
- `logs` フォルダを空で生成します
- 雛形コピー時に `~$*` を除外します
- ZIP 展開後のルート名が `案件情報System` になる形で ZIP を生成します

## 6. 配布前チェック
- 作業場所が `C:\Users\kazu2\Documents\案件情報System\開発用` であること
- 実行時ルートが `C:\Users\kazu2\Documents\案件情報System` であること
- 実行時ルート直下の `案件情報System_Kernel.xlsx` と `案件情報System_Base.xlsx` が配布したい版であること
- 実行時ルート直下の `雛形` と `案件情報System_利用開始ガイド.pdf` が配布したい版であること
- Release 用証明書ファイルが既定パスに配置されていること
- 配布対象の版やタイミングが確定していること

## 7. 配布後チェック
- `案件情報System.zip` が生成されていること
- `配布用` フォルダが再生成されていること
- `配布用` フォルダ直下に次があること
- `案件情報System.exe`
- `案件情報System_Kernel.xlsx`
- `案件情報System_Base.xlsx`
- `利用開始ガイド.pdf`
- `初回セットアップ.bat`
- `CaseInfoSystem.Internal.cer`
- `Addins`
- `雛形`
- `logs`

## 8. モニター配布前 実機確認チェックリスト

所内モニター配布前は、build / test 成功、Debug Add-in 実機反映、Release 配布物生成を別扱いにして、少なくとも次を実機で確認します。

### TaskPane 表示回復

- created CASE 表示後の ready-show が `attempt 1 -> 80ms attempt 2 -> pending fallback` の順で扱われること。
- pending retry は `400ms / 3 attempts` の fallback として扱い、pending retry success を completion と読まないこと。
- pending / callback / WindowActivate / foreground / normalized outcome は completion ではないこと。
- completion trace は `case-display-completed` の one-time emit だけであること。
- TaskPane freeze line trace の payload field set / order / names / values を変えずに観測できること。

### CASE 表示

- 新規 CASE 初回表示で CASE workbook と TaskPane が表示されること。
- 既存 CASE 開き直しで TaskPane が表示されること。
- CASE 間切替で TaskPane が対象 CASE と混線しないこと。
- TaskPane 欠落、二重表示、白 Excel 化がないこと。

### Kernel HOME

- 初回起動で Kernel HOME が従来どおり表示されること。
- unbound 表示時は fail-closed として扱われ、Kernel workbook / Kernel window の自動選択や自動復元をしないこと。
- CASE 作成後に Kernel が不要に前面へ戻らないこと。

### 新規 CASE 作成

- hidden create が完了し、shared app handoff 後に CASE 表示へ進むこと。
- CASE workbook が可視化されること。
- 初期カーソル移動まで到達すること。
- TaskPane 表示完了まで進み、表示待ちのまま止まらないこと。

### 文書作成 / CC 流し込み

- 文書ボタンから文書作成が開始できること。
- 登録済みテンプレートが解決されること。
- Word 文書が表示されること。
- ContentControl 差し込みが行われること。
- 出力先へ保存されること。

### Release package / Add-in 反映

- `CreateDistributionPackage.bat` で配布物を生成すること。
- ZIP 内容に必要なファイルが含まれること。
- `Addins` が同梱されること。
- Release 署名済み package として生成されること。
- Debug 実機反映と Release 配布物生成を混同しないこと。

### 初回起動 / 既存環境上書き

- 利用者は既存フォルダへ上書きせず、新しいフォルダへ展開する前提で案内すること。
- 初回導入は `初回セットアップ.bat` から開始すること。

### ログ / trace / rollback

- `Runtime execution observed` ログで実行 DLL を確認すること。
- `assemblySha256` を使い、想定した Add-in が実行されていることを確認すること。
- TaskPane freeze line trace を確認し、completion と非 completion の読み替えが起きていないことを確認すること。
- 正式 rollback 手順は、この文書では未確定です。必要な場合は配布前に戻し先、戻し方法、利用者への案内を別途確定します。

## 9. 禁止事項
- `CreateDistributionPackage.bat` を使わずに手作業で配布物を組み立てない
- `scripts` を直接変更しない
- `csproj` を直接変更しない
- 開発用フォルダ直下のファイルをそのまま利用者へ渡さない
- 古い `配布用` フォルダを流用しない
- 配布担当者向け手順として PowerShell 個別実行を標準化しない
- 利用者へ `.vsto` の直接実行を案内しない
- 開発用 `Kernel` / `Base` を正規化しない
- 生成後の `配布用` フォルダを手で部分編集して整えない

## 10. トラブル切り分け

### `CreateDistributionPackage.bat` がすぐ止まる場合
- `build.ps1` があるか確認します
- `scripts\Build-DistributionPackage.ps1` があるか確認します
- Release 用証明書ファイルが既定パスにあるか確認します

### Release Add-in package 生成で止まる場合
- Release 用証明書の配置と指定内容を確認します
- Excel Add-in / Word Add-in の Release package が作れる状態か確認します
- 署名まわりの問題か、ビルド成果物の問題かを切り分けます

### ZIP 生成で止まる場合
- 実行時ルート直下の `Kernel` / `Base` / `雛形` / PDF が存在するか確認します
- `配布用` フォルダや `案件情報System.zip` を別プロセスが掴んでいないか確認します
- `配布用` フォルダを手で編集していないか確認します

### `CaseInfoSystem.Internal.cer` が入らない場合
- Release Excel Add-in と Release Word Add-in の `.vsto` が生成されているか確認します
- 両方の VSTO manifest が同じ証明書で署名されているか確認します

### 利用者からセットアップできないと言われた場合
- 利用者が `初回セットアップ.bat` を実行したか確認します
- 利用者が `.vsto` を直接開いていないか確認します
- 利用者が ZIP を展開せずに実行していないか確認します
- 利用者側には `distribution-package-user-guide.md` の手順を案内します

## 11. 利用者への渡し方
- 渡すものは `案件情報System.zip` を基本とします
- 利用者には ZIP 展開後に `初回セットアップ.bat` を先に実行するよう案内します
- 利用者には PowerShell 実行や `.vsto` の直接操作を案内しません
