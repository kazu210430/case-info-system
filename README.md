# 案件情報System

案件情報System は、案件情報の作成・更新を補助するための Windows 向け業務ツール群です。リポジトリ上では Excel 用 VSTO アドイン、Word 用 VSTO アドイン、補助ランチャー、純粋ロジック中心のテスト プロジェクトを管理しています。

## 主要構成

- `dev/CaseInfoSystem.slnx`: 開発用ソリューション
- `dev/CaseInfoSystem.ExcelAddIn`: Excel VSTO アドイン
- `dev/CaseInfoSystem.WordAddIn`: Word VSTO アドイン
- `dev/CaseInfoSystem.ExcelLauncher`: Kernel workbook 起動用ランチャー
- `dev/CaseInfoSystem.Tests`: pure tests
- `build/packages`: `CaseInfoSystem.VstoBuildAssets.<version>.nupkg` の配置先
- `dev/Deploy`: 配布用 / デバッグ用パッケージ出力先
- `Addins`: ローカル実行用に同期されるアドイン配置先

## セットアップ

### 前提

- Windows 環境
- `dotnet` CLI
- .NET Framework 4.8 をビルドできる環境
- Microsoft Excel / Word
- VSTO 実行環境
- `build/packages` に配置した `CaseInfoSystem.VstoBuildAssets.<version>.nupkg`

`NuGet.config` は `build/packages` と `nuget.org` を参照します。`build/packages` に VSTO ビルド用パッケージが無い場合、`Microsoft.Office` などの参照が解決できず、アドイン プロジェクトのビルドが止まります。

各 VSTO プロジェクトは署名用ファイルとして以下を参照する設定です。

- `dev/CaseInfoSystem.ExcelAddIn/CaseInfoSystem.ExcelAddIn_TemporaryKey.pfx`
- `dev/CaseInfoSystem.WordAddIn/CaseInfoSystem.WordAddIn_TemporaryKey.pfx`

これらの `.pfx` はリポジトリにありますが、`dotnet build` で署名を有効にしたまま VSTO プロジェクトをビルドすると、環境によっては `ResolveKeySource` が証明書ストア参照で失敗します。開発中のコンパイル確認では `SignManifests=false` を付けた未署名ビルドを使ってください。

### 初回復元

```powershell
dotnet restore .\dev\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.csproj
dotnet restore .\dev\CaseInfoSystem.WordAddIn\CaseInfoSystem.WordAddIn.csproj
dotnet restore .\dev\CaseInfoSystem.Tests\CaseInfoSystem.Tests.csproj
```

## ビルド

通常の開発向けビルドは次を使用します。

```powershell
dotnet build .\dev\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.csproj -c Release -p:SignManifests=false -p:ManifestCertificateThumbprint=
dotnet build .\dev\CaseInfoSystem.WordAddIn\CaseInfoSystem.WordAddIn.csproj -c Release -p:SignManifests=false -p:ManifestCertificateThumbprint=
dotnet build .\dev\CaseInfoSystem.ExcelLauncher\CaseInfoSystem.ExcelLauncher.csproj -c Release
```

補足:

- `dev/CaseInfoSystem.slnx` に対する `dotnet restore` / `dotnet build` は、この環境では終了コード 1 で失敗することがあり、通常手順としては使っていません。
- VSTO プロジェクトは `SignManifests=false` を付けると `dotnet build` で DLL までのコンパイル確認ができます。
- 署名付き manifest / `.vsto` を含む配布ビルドは、証明書ストアと VSTO パッケージング前提を満たす別経路です。
- CI では `dotnet build <VSTO csproj> -p:SignManifests=false -p:ManifestCertificateThumbprint=` と `dotnet test` を組み合わせ、署名依存を外してコンパイル検証を行います。
- 配布物生成や実機同期の確認は CI の対象外です。

## テスト

純粋ロジック中心のテストは次を使用します。

```powershell
dotnet test .\dev\CaseInfoSystem.Tests\CaseInfoSystem.Tests.csproj -c Release --no-restore
```

補足:

- テスト対象は `dev/CaseInfoSystem.Tests` です。
- CI はソリューション ビルドに加えてこのテスト プロジェクトを実行します。
- Office / VSTO の実機挙動は CI では確認しません。
- CI の範囲メモは [docs/ci-tests.md](docs/ci-tests.md) を参照してください。

## 実機運用メモ

`dev` 側の修正を実機で確認しながら育てる運用を前提にする場合は、実機反映を `Debug` のみに寄せ、事前バックアップと事後検証を入れておくのがおすすめです。

- Hot reload 運用ルールと Visual Studio 組み込み例: [docs/hot-reload-operations.md](docs/hot-reload-operations.md)
- 追加したガードスクリプト: `scripts/Invoke-HotReloadGuard.ps1`

## 配布

### Excel アドイン

Excel 側には配布用ターゲットが定義されています。

- Release 配布ターゲット: `DeployReleaseAddIn`
- Debug 配布ターゲット: `DeployDebugAddIn`
- Release 配布先: `dev/Deploy/Package/CaseInfoSystem.ExcelAddIn`
- Debug 配布先: `dev/Deploy/DebugPackage/CaseInfoSystem.ExcelAddIn`
- 実行用同期先: `Addins/CaseInfoSystem.ExcelAddIn`

配布前の最低確認:

- `CaseInfoSystem.ExcelAddIn.dll`
- `CaseInfoSystem.ExcelAddIn.dll.manifest`
- `CaseInfoSystem.ExcelAddIn.vsto`
- `DocumentExecutionMode.txt`
- `DocumentExecutionPilot.txt`
- `DocumentExecutionAllowlist.txt`
- `DocumentExecutionAllowlist.review.txt`

### Word アドイン

Word 側にも配布用ターゲットが定義されています。

- Release 配布ターゲット: `DeployReleaseAddIn`
- Debug 配布ターゲット: `DeployDebugAddIn`
- Release 配布先: `dev/Deploy/Package/CaseInfoSystem.WordAddIn`
- Debug 配布先: `dev/Deploy/DebugPackage/CaseInfoSystem.WordAddIn`
- 実行用同期先: `Addins/CaseInfoSystem.WordAddIn`

## 詰まりやすい前提

- VSTO 参照は通常の NuGet パッケージだけでは完結せず、`build/packages` のローカル パッケージ前提があります。
- 実機確認は pure tests だけでは足りません。Excel / Word 上でのアドイン読込確認が別途必要です。
- 配布ターゲットは VSTO パッケージングが有効な環境を前提にしています。`dotnet` ベースの通常ビルドだけで配布物生成まで完了するかは未確認です。

## TODO / 未確定

- 署名用 `.pfx` の正式な配置手順
- 必須の Visual Studio ワークロード名や開発環境セットアップ手順の確定

