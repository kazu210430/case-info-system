# 配布パッケージ設計

## 1. この文書の役割
- この文書は、配布パッケージの設計方針を固定するための文書です
- 実際の配布手順は `distribution-operator-guide.md` に委譲します
- 利用者向けの使い始め方は `distribution-package-user-guide.md` に委譲します
- 手順よりも、配布物の構成、責務分担、固定ルールを整理することを目的とします

## 2. 3文書の役割分担
- `distribution-package-workflow.md`
  - 配布方式、固定パス、配布物構成、責務分担を定義します
- `distribution-operator-guide.md`
  - 配布する側が何を確認し、何を実行するかを説明します
- `distribution-package-user-guide.md`
  - 利用者が ZIP 展開後に何を実行するかを説明します

## 3. 基本方針
- 開発用フォルダと配布用フォルダは物理的に分離します
- 配布用フォルダは毎回新規生成します
- 配布作業の標準入口は `CreateDistributionPackage.bat` とします
- 利用者には PowerShell 実行を求めません
- 利用者には `.vsto` の直接実行を求めません
- 利用者の初回導線は `初回セットアップ.bat` を前提にします
- 配布用 `Kernel` / `Base` の docprops 正規化は自動処理に含めます
- 開発用 `Kernel` / `Base` は正規化対象にしません
- `logs` はコピー元を持たず、配布時に空フォルダを生成します
- 当面の配布方式は ZIP とします
- 配布正本は `CreateDistributionPackage.bat` で生成される `案件情報System.zip` です
- `dev/Deploy/Package` は Release Add-in package の中間生成物であり、配布正本として扱いません
- 旧 Release package 由来の生成物はリポジトリに保持しません

## 4. 想定フォルダ
- 実行時ルート: `C:\Users\kazu2\Documents\案件情報System`
- 開発用: `C:\Users\kazu2\Documents\案件情報System\開発用`
- 配布用: `C:\Users\kazu2\Documents\案件情報System\配布用`
- 最終配布物: `C:\Users\kazu2\Documents\案件情報System\案件情報System.zip`

## 5. コピー元の設計
- `案件情報System.exe`
  - Release 出力物をコピー元にします
- `Addins`
  - Release Add-in package をコピー元にします
- `案件情報System_Kernel.xlsx`
  - 実行時ルート直下の正本をコピー元にします
- `案件情報System_Base.xlsx`
  - 実行時ルート直下の正本をコピー元にします
- `雛形`
  - 実行時ルート直下の正本をコピー元にします
- `利用開始ガイド.pdf`
  - 実行時ルート直下の `案件情報System_利用開始ガイド.pdf` をコピー時にリネームします
- `logs`
  - コピーせず、配布時に空フォルダを生成します
- `初回セットアップ.bat`
  - `distribution-assets\初回セットアップ.bat` を同梱します
- `CaseInfoSystem.Internal.cer`
  - Release VSTO manifest から書き出した証明書を同梱します

## 6. 配布物に含めるもの
- `案件情報System.exe`
- `案件情報System_Kernel.xlsx`
- `案件情報System_Base.xlsx`
- `利用開始ガイド.pdf`
- `初回セットアップ.bat`
- `CaseInfoSystem.Internal.cer`
- `Addins` フォルダ
- `雛形` フォルダ
- `logs` フォルダ

## 7. 自動化の責務
- `CreateDistributionPackage.bat` は配布作業の開始点です
- このバッチは Release Add-in package 生成と配布 ZIP 生成を順に実行します
- 配布用フォルダの再生成、必要ファイルのコピー、証明書の書き出し、docprops 正規化、ZIP 生成は自動処理として扱います
- 手作業は配布前チェックと、生成物の確認に限定します

## 8. 配布用フォルダ設計
- 配布用フォルダ名は `配布用` に固定します
- ZIP 展開後のルートフォルダ名は `案件情報System` に固定します
- 雛形コピー時は `~$*` を除外します
- `CaseInfoSystem.Internal.cer` は配布用フォルダ直下に置きます
- `初回セットアップ.bat` は配布用フォルダ直下に置きます

## 9. docprops 正規化の位置づけ
- 正規化対象は配布用にコピー済みの `Kernel` / `Base` のみです
- 開発用の `Kernel` / `Base` は対象外です
- 実行時ルート直下の正本 `Kernel` / `Base` も対象外です
- 正規化は配布処理の一部として自動実行します

## 10. 役割分担
- 配布する側
  - `CreateDistributionPackage.bat` を標準入口として配布物を生成します
  - 生成された ZIP と配布用フォルダを確認して利用者へ渡します
- 利用者
  - ZIP を展開します
  - `初回セットアップ.bat` を実行します
  - その後に `案件情報System.exe` を起動します

## 11. operator-guide へ委譲する事項
- 配布前チェックの実施項目
- `CreateDistributionPackage.bat` の実行手順
- 配布後の確認観点
- 配布作業での禁止事項
- トラブル切り分け

## 12. user-guide へ委譲する事項
- ZIP 展開手順
- `初回セットアップ.bat` の実行順
- Excel Add-in / Word Add-in インストーラー画面への対応
- セットアップ後の起動方法
- 利用者側の禁止事項

## 13. 証明書と署名の位置づけ
- Release Add-in package は署名済み成果物を前提にします
- 配布用の `CaseInfoSystem.Internal.cer` は、利用者の初回セットアップで使用します
- 利用者へ証明書ファイル単体の手動操作を求めません
- `.pfx` は利用者向け配布物に含めません

## 14. 禁止事項
- 開発用フォルダ直下の運用ファイルをそのまま配布しません
- 古い `配布用` フォルダを使い回しません
- 古い `dev/Deploy/Package` 生成物を配布正本として使い回しません
- 利用者向け手順として PowerShell を案内しません
- 利用者向け手順として `.vsto` の直接実行を案内しません
- 開発用 `Kernel` / `Base` を正規化しません

## 15. 補足
- 将来の MSI 化は別フェーズで検討します
- 現時点では ZIP 配布を前提に文書と運用をそろえます
