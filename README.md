# 案件情報System

案件情報System は、案件情報の作成・更新を補助する Windows 向け業務ツール群です。  
このリポジトリでは、Excel VSTO Add-in を中心に、Word VSTO Add-in、Kernel 起動用ランチャー、テストを管理しています。

## リポジトリ構成

- `dev/CaseInfoSystem.slnx`: 開発用ソリューション
- `dev/CaseInfoSystem.ExcelAddIn`: Excel VSTO Add-in 本体
- `dev/CaseInfoSystem.WordAddIn`: Word VSTO Add-in
- `dev/CaseInfoSystem.ExcelLauncher`: `案件情報System_Kernel.xlsx` 起動用ランチャー
- `dev/CaseInfoSystem.Tests`: 純粋ロジック中心のテスト
- `dev/Deploy`: 配布物出力先
- `Addins`: 実行用 Add-in 同期先
- `雛形`: Word / Excel 雛形格納先

## システム構成

### Excel Add-in

Excel Add-in は workbook を次の役割で扱います。

- `Kernel`
  - システム全体の入口です。
  - HOME 画面、各種設定、ユーザー情報、雛形登録、案件一覧、CASE 作成を担当します。
- `CASE`
  - 個別案件の workbook です。
  - task pane から文書作成、Accounting 作成、案件一覧登録を実行します。
- `Accounting`
  - CASE から派生する会計系 workbook です。
  - 専用シートと文書プロパティで識別されます。

### Word Add-in

Word Add-in は次の機能を提供します。

- 文書ごとの Style 作業ウィンドウ表示切替
- コンテンツコントロールの Title / Tag 一括置換

### Excel Launcher

`CaseInfoSystem.ExcelLauncher` は、配布先フォルダにある `案件情報System_Kernel.xlsx` を既定の Excel で起動します。

## 利用フロー

### Kernel HOME

Kernel workbook 到達時、条件を満たす場合は HOME 画面を表示します。  
HOME では以下を実行します。

- 顧客名入力
- 新規 CASE 作成
- 既存シートへの遷移
  - ユーザー情報
  - 雛形登録
  - 案件一覧
- 既定保存先の設定
- 命名ルールの設定

### 新規CASE作成フロー

新規 CASE 作成は 3 モードあります。

- `NewCaseDefault`
  - Kernel の `DEFAULT_ROOT` を使って新規 CASE を作成します。
  - `DEFAULT_ROOT` が未設定なら最初にフォルダ選択を行い、Kernel に保存します。
- `CreateCaseSingle`
  - 保存先フォルダを選択して CASE を 1 件作成します。
- `CreateCaseBatch`
  - 保存先フォルダを選択して CASE を連続作成する前提のモードです。

共通の処理は次のとおりです。

1. Kernel から `SYSTEM_ROOT` と Base workbook を解決する
2. 顧客名と Kernel 側の設定を使って保存先パスを決める
3. Base workbook をコピーして CASE workbook を作る
4. CASE workbook を hidden session で初期化して保存する

作成後の挙動:

- `NewCaseDefault` / `CreateCaseSingle`
  - 待機 UI を表示した上で、作成済み CASE を開きます
  - CASE 表示前にフォルダ表示を先行させる経路があります
- `CreateCaseBatch`
  - CASE 自体は自動表示せず、フォルダ表示を優先します
  - HOME を次の入力に戻します

## CASE からの操作

CASE の task pane では、アクション種別を `doc / accounting / caselist` に分岐しています。

- `doc`
  - 文書作成フローへ進みます
- `accounting`
  - Accounting workbook 作成 / 同期フローへ進みます
- `caselist`
  - 案件一覧登録を実行します

### 文書作成フロー

CASE の文書作成は VSTO 側で次の順に処理します。

1. 文書キーからテンプレート情報を解決する
   - task pane snapshot cache を優先
   - 解決できない場合は MasterList を参照
2. 実行可否を判定する
   - 対応 Word テンプレートか
   - テンプレートが存在するか
   - 出力先フォルダが解決できるか
   - CASE 文脈と差し込みデータが解決できるか
3. 文書実行モードと allowlist を確認する
4. Word 文書を生成する
   - テンプレートから文書作成
   - 差し込み
   - content control の整理
   - 保存
   - Word 画面表示

テンプレート解決の基準:

- `WORD_TEMPLATE_DIR` が設定されていればそれを使用
- 未設定なら `SYSTEM_ROOT\雛形` を使用

文書実行制御の基準ファイル:

- `DocumentExecutionMode.txt`
- `DocumentExecutionPilot.txt`
- `DocumentExecutionAllowlist.txt`
- `DocumentExecutionAllowlist.review.txt`

## 表示制御

### HOME 表示

HOME 表示は、Kernel workbook の起動直後や遷移直後に無条件で出すのではなく、startup 状態と suppression 状態を見て制御します。

- Kernel startup context があること
- Kernel workbook context が解決できること
- startup workbook が Kernel であるか、または visible な non-kernel workbook が無いこと

### Excel 本体 / workbook の可視状態

HOME 表示時は現在の workbook 状態に応じて、次を切り替えます。

- Kernel window だけを最小化する
- Kernel window を不可視化する
- Excel メインウィンドウ自体を隠す

これにより、HOME と workbook 表示の競合を避けます。

### CASE の非表示 Open と表示復帰

CASE 作成後の interactive mode では、作成済み CASE を hidden-for-display 経路で開きます。  
表示時は次を順に行います。

- workbook window の回復
- 必要に応じた window 可視化
- CASE pane の ready-show 要求
- 初期カーソル移動
- workbook window の前面化

### suppression と activate 保護

WorkbookActivate / WindowActivate / pane refresh の競合を避けるため、以下を個別に制御しています。

- Kernel HOME 表示 suppression
- CASE pane activation suppression
- CASE workbook activate protection

外部 workbook を検知した場合は、suppression 中を除いて HOME を閉じます。

## 待機UI

待機 UI は 2 系統あります。

- CASE 表示待機
  - 新規 CASE 作成後、interactive open 中に表示します
  - 必要に応じて owner form を一時的に最小化します
- 文書表示待機
  - Word 文書作成中に表示します
  - Word 表示完了後に閉じます

これとは別に、HOME には作成開始時の最小化と前面化リトライ処理があります。

## ログとデバッグ

### 基本ログ

Excel Add-in の主ログは次に出力されます。

- 主ログ: `logs\CaseInfoSystem.ExcelAddIn_trace.log`
- fallback: `%TEMP%\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn_trace.log`

主ログは、既存の `案件情報System` ルートを優先して解決します。

### 起動診断

Excel 起動時には process launch context を記録します。  
詳細 startup diagnostics は環境変数 `CASEINFOSYSTEM_EXCEL_STARTUP_TRACE` で有効化できます。

### 表示系トレース

HOME / pane / window まわりの表示制御は `[KernelFlickerTrace]` プレフィックスで追跡できます。

## 開発とビルド

### 前提

- Windows
- `dotnet` CLI
- .NET Framework 4.8 をビルドできる環境
- Microsoft Excel / Word
- VSTO 実行環境
- `build/packages` に配置した `CaseInfoSystem.VstoBuildAssets.<version>.nupkg`

`NuGet.config` は `build/packages` と `nuget.org` を参照します。

### 初回復元

```powershell
dotnet restore .\dev\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.csproj
dotnet restore .\dev\CaseInfoSystem.WordAddIn\CaseInfoSystem.WordAddIn.csproj
dotnet restore .\dev\CaseInfoSystem.ExcelLauncher\CaseInfoSystem.ExcelLauncher.csproj
dotnet restore .\dev\CaseInfoSystem.Tests\CaseInfoSystem.Tests.csproj
```

### コンパイル確認

`dotnet build` は compile-only 用です。  
Core MSBuild では VSTO packaging が無効なため、compile-only ビルドでは `AllowCoreBuildWithoutVstoPackaging=true` が必要です。

```powershell
dotnet build .\dev\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.csproj -c Release -p:SignManifests=false -p:ManifestCertificateThumbprint= -p:AllowCoreBuildWithoutVstoPackaging=true
dotnet build .\dev\CaseInfoSystem.WordAddIn\CaseInfoSystem.WordAddIn.csproj -c Release -p:SignManifests=false -p:ManifestCertificateThumbprint= -p:AllowCoreBuildWithoutVstoPackaging=true
dotnet build .\dev\CaseInfoSystem.ExcelLauncher\CaseInfoSystem.ExcelLauncher.csproj -c Release
```

### テスト

```powershell
dotnet test .\dev\CaseInfoSystem.Tests\CaseInfoSystem.Tests.csproj -c Release --no-restore
```

テストは pure logic 中心です。Office 実機挙動は別途確認が必要です。

## 配布

### Excel Add-in

- Release 配布ターゲット: `DeployReleaseAddIn`
- Debug 配布ターゲット: `DeployDebugAddIn`
- Release 配布先: `dev/Deploy/Package/CaseInfoSystem.ExcelAddIn`
- Debug 配布先: `dev/Deploy/DebugPackage/CaseInfoSystem.ExcelAddIn`
- 実行用同期先: `Addins/CaseInfoSystem.ExcelAddIn`

Excel 配布物には、アドイン本体に加えて次の文書実行制御ファイルが含まれます。

- `DocumentExecutionMode.txt`
- `DocumentExecutionPilot.txt`
- `DocumentExecutionAllowlist.txt`
- `DocumentExecutionAllowlist.review.txt`

### Word Add-in

- Release 配布ターゲット: `DeployReleaseAddIn`
- Debug 配布ターゲット: `DeployDebugAddIn`
- Release 配布先: `dev/Deploy/Package/CaseInfoSystem.WordAddIn`
- Debug 配布先: `dev/Deploy/DebugPackage/CaseInfoSystem.WordAddIn`
- 実行用同期先: `Addins/CaseInfoSystem.WordAddIn`

### Debug 配布の実行例

VSTO packaging と runtime 反映を伴う Debug 配布は `MSBuild.exe` ベースで行います。  
付属スクリプトでは次を実行します。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-DeployDebugAddIns.ps1 -Project ExcelAddIn
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-DeployDebugAddIns.ps1 -Project WordAddIn
```

## 補足

- `dotnet build` だけでは `Addins/` への runtime 反映は行いません
- 実機確認が必要なときは `DeployDebugAddIn` 系を使ってください
- Word / Excel の実機挙動確認は CI の対象外です
