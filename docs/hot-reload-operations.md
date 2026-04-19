# Hot Reload運用メモ

このリポジトリでは、ビルド結果が `Addins/` 配下の実行用アドインへ同期される前提で運用しています。
1人で実機確認しながら育てる進め方自体は問題ありません。大事なのは、「即反映」とセットで「即復旧」できることです。

## おすすめ運用ルール

1. 日常の育成は `Debug` を基本にする。
2. 実機反映は `Debug` だけにする。
3. `Release` は配布確認や引き渡し確認に寄せる。
4. Rebuild 前に、実行用アドインと運用 workbook を必ず世代バックアップする。
5. 修正記録は短くてよいので、まず症状を書く。
6. いつでも戻せる既知の安定コミットかタグを最低1つ残す。
7. workbook 構造を変える日は、テンプレートも一緒に退避する。
8. ビルド後は「同期できたか」だけでなく、「同期元と同期先が一致したか」まで確認する。

## 追加したスクリプト

`scripts/Invoke-HotReloadGuard.ps1`

このスクリプトには 2 つのモードがあります。

1. `Backup`
2. `Verify`

### Backup

バックアップは次のような対象を `build/hot-reload-backups/<timestamp>/` へ世代退避します。

`ExcelAddIn` を選んだ場合:

- `Addins/CaseInfoSystem.ExcelAddIn`
- ルート直下で `*Kernel.xlsx` に一致する workbook
- ルート直下で `*Base.xlsx` に一致する workbook
- `dev`、`build`、`scripts`、`docs`、`tools` などの基盤フォルダを除いた、ルート直下の runtime 用フォルダ

`WordAddIn` を選んだ場合:

- `Addins/CaseInfoSystem.WordAddIn`

### Verify

`Verify` は `dev/Deploy/...` と `Addins/...` を比較し、次を確認します。

- 必須ファイルが揃っていること
- 同期先に不足ファイルがないこと
- 同期先に余分なファイルが残っていないこと
- 同名ファイルの SHA-256 ハッシュが一致すること

自動の `Backup` / `Verify` と実機反映は `Debug` のみです。`Release` は配布パッケージ更新用として使います。

## 手で実行する例

両方をバックアップ:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-HotReloadGuard.ps1 -Mode Backup -Project All -Configuration Debug
```

両方を検証:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-HotReloadGuard.ps1 -Mode Verify -Project All -Configuration Debug
```

Excel だけ検証:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-HotReloadGuard.ps1 -Mode Verify -Project ExcelAddIn -Configuration Debug
```

## Visual Studio への組み込み案

おすすめは、VSTO プロジェクトの `Debug` ビルド時だけ `Backup` と `Verify` が走る形です。

Excel Add-in の Pre-build event:

```powershell
powershell -ExecutionPolicy Bypass -File "$(SolutionDir)scripts\Invoke-HotReloadGuard.ps1" -Mode Backup -Project ExcelAddIn -Configuration $(ConfigurationName)
```

Excel Add-in の Post-build event:

```powershell
powershell -ExecutionPolicy Bypass -File "$(SolutionDir)scripts\Invoke-HotReloadGuard.ps1" -Mode Verify -Project ExcelAddIn -Configuration $(ConfigurationName)
```

Word Add-in の Pre-build event:

```powershell
powershell -ExecutionPolicy Bypass -File "$(SolutionDir)scripts\Invoke-HotReloadGuard.ps1" -Mode Backup -Project WordAddIn -Configuration $(ConfigurationName)
```

Word Add-in の Post-build event:

```powershell
powershell -ExecutionPolicy Bypass -File "$(SolutionDir)scripts\Invoke-HotReloadGuard.ps1" -Mode Verify -Project WordAddIn -Configuration $(ConfigurationName)
```

## 最低限これだけは入れる

まずは次の 3 つで十分です。

1. 日常運用は `Debug`
2. 実機反映は `Debug` だけ
3. Rebuild 前に `Backup`
4. Rebuild 後に `Verify`
