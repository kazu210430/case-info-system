# Current-State Fact Review

## 位置づけ

この文書は、`docs/architecture.md`、`docs/flows.md`、`docs/ui-policy.md` などをレビュー時の判断材料として使う前に、current-state 表現を実装から再確認するための軽量手順です。設計方針を置き換えるものではなく、古い事実表現だけを根拠に改善を止める誤運用を防ぐための補助メモです。

## 扱う表現

次のような表現は、固定仕様ではなく確認時点のスナップショットとして扱います。

- 行数、class size、巨大クラス評価
- `現行 main`、`現在地`、`現状`、`current-state`、`current meaning`
- 現在のサービス構成、adapter / bridge 分離状況
- TaskPane / ThisAddIn / service 境界の「今どこに残っているか」という記述
- 基準 commit hash や「main / origin/main 一致時点」の記録

## レビュールール

- docs の数値や current-state 記述は固定仕様として扱わず、レビュー時に実装から再確認します。
- 改善判断では、docs の古い数値だけを根拠に STOP / GO 判定しません。
- current-state を根拠にする場合は、対象ファイルの現行行数、主要責務、差分範囲を確認します。
- design policy と current-state fact を分けて読みます。`WorkbookOpen` に直接依存した表示制御を追加しない、UI 制御は専用サービス経由で行う、などの方針は policy として扱います。
- docs と実装が矛盾する場合は、実装事実、参照 docs、判断に使うかどうかを分けて記録します。
- historical hash は履歴として残してよいですが、最新 main の事実として読む場合は再確認します。

## 軽量確認コマンド

PowerShell 例です。行数は `Measure-Object -Line` ではなく、物理行数として `ReadAllLines(...).Length` を使います。

```powershell
git status --short --branch
git rev-parse main
git rev-parse origin/main
([System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath .\dev\CaseInfoSystem.ExcelAddIn\ThisAddIn.cs))).Length
rg -n "ThisAddIn|TaskPane|service|adapter|class size|line count|行数|巨大|current-state|current state|現状|現在地|現行|現在は|到達点" docs
```

必要に応じて、実装側の現在の owner / adapter 境界を確認します。

```powershell
rg -n "class\s+ThisAddIn|class\s+.*Adapter|class\s+TaskPane|class\s+.*Service|CreateTaskPane|RemoveTaskPane|HookApplicationEvents|UnhookApplicationEvents" .\dev\CaseInfoSystem.ExcelAddIn
```

## 優先して見るファイル

- `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/VstoEventAdapter.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneEntryAdapter.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/HomeTransitionAdapter.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/ShutdownCleanupAdapter.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneHostRegistry.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneHostFactory.cs`
- `dev/CaseInfoSystem.ExcelAddIn/UI/TaskPaneHost.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`

## 更新時の書き方

- 行数を書く場合は、確認日、基準 hash、確認コマンドを併記し、固定仕様ではないことを明記します。
- 「ThisAddIn は N 行だから問題」とは書きません。行数はレビューの入口であり、問題判断は主要責務と変更範囲を確認してから行います。
- adapter / service 分離後の現状を書く場合は、owner と delegate surface を分けて書きます。
- current-state docs を根拠にする時は、古い hash や古い行数が混ざっていないか先に確認します。

## 今回の確認スナップショット

- 確認日: `2026-05-17`
- `main` / `origin/main`: `616b4af88d15571f0e57a95f447693f0426faeb7`
- `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`: 645 行
- 確認コマンド: `([System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath .\dev\CaseInfoSystem.ExcelAddIn\ThisAddIn.cs))).Length`
