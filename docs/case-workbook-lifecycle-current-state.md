# Case Workbook Lifecycle Current State

## 位置づけ

この文書は、`CaseWorkbookLifecycleService` 分割と、その後続整理である `CaseClosePromptService` / `CaseFolderOpenService` 分離後の現在地を、現行 `main` のコードベースに合わせて固定するための補助文書です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- TaskPane 側の現在地: `docs/taskpane-refactor-current-state.md`
- 残課題: `docs/technical-debt.md`

この文書は TaskPaneManager 側の追加整理を扱いません。lifecycle 側の安定到達点だけを明文化します。

## 現在の構成

- `AddInCompositionRoot`
  - `ManagedCloseState`、`CaseFolderOpenService`、`CaseClosePromptService`、`KernelNameRuleReader`、`PostCloseFollowUpScheduler` を生成し、`CaseWorkbookLifecycleService` へ注入します。
- `CaseWorkbookLifecycleService`
  - orchestration 寄りの中心サービスです。
  - CASE / Base 初回初期化、dirty session 状態管理、before-close / managed close / post-close follow-up の調停、created case folder offer pending 状態管理、CASE HOME 表示補正を担います。
- `KernelNameRuleReader`
  - CASE workbook から `SYSTEM_ROOT` を解決し、open 中 Kernel workbook または package `docProps/custom.xml` から `NAME_RULE_A` / `NAME_RULE_B` を読み取ります。
- `ManagedCloseState`
  - workbook key 単位で managed close の入れ子状態を管理します。
- `CaseClosePromptService`
  - dirty close prompt のタイトル解決と `保存しますか？` ダイアログ、created case folder offer prompt を担当します。
- `CaseFolderOpenService`
  - 保存先フォルダ解決、存在確認、Explorer 起動を担当します。
- `PostCloseFollowUpScheduler`
  - close 後 follow-up、Excel busy retry、visible workbook が無い場合の Excel 終了判定を担当します。

## 大まかな順序

1. `WorkbookOpen` / `WorkbookActivate`
   - `CaseWorkbookLifecycleService` が初回初期化要否を判定し、CASE だけ `RegisterKnownCaseWorkbook` と name rule 同期を行います。
2. `SheetChange`
   - 対象 workbook かつ managed close 中でも transient suppression 中でもない場合だけ session dirty を記録します。
3. `WorkbookBeforeClose`
   - `CaseWorkbookBeforeClosePolicy` が対象外 / managed close 中 / dirty session / clean close を判定します。
4. dirty prompt
   - dirty session の場合は `CaseClosePromptService` が `保存しますか？` を表示します。
5. folder offer
   - `KernelCaseCreationCommandService` から pending が付与されていた workbook では、保存先フォルダが解決できる場合だけ created case folder offer prompt を出し、必要時に `CaseFolderOpenService` が Explorer を開きます。
6. managed close
   - dirty path では `CaseWorkbookLifecycleService` が dispatcher 経由で managed close を予約し、`ManagedCloseState` のスコープ内で save 有無を処理して `workbook.Close(SaveChanges: false)` へ進めます。
7. post-close follow-up
   - dirty path では managed close 内で、clean close では before-close 中に `PostCloseFollowUpScheduler` を予約します。
8. close 後判定
   - `PostCloseFollowUpScheduler` は対象 workbook が残っていないことを確認し、Excel busy なら retry し、visible workbook が 1 つも無い場合だけ Excel 終了を試みます。

dirty path の大まかな順序は `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` です。

## 関連テスト

- `CaseWorkbookLifecycleServicePolicyTests`
  - 初回初期化、before-close、sheet change の policy 判定を確認します。
- `CaseWorkbookLifecycleServiceThinOrchestrationTests`
  - prompt / folder offer / managed close / post-close follow-up の thin orchestration を確認します。
- `ManagedCloseStateTests`
  - managed close の入れ子状態を確認します。
- `KernelCaseCreationServiceTests`
  - 関連する managed close scope 利用を間接的に参照します。

`CaseClosePromptService`、`CaseFolderOpenService`、`KernelNameRuleReader`、`PostCloseFollowUpScheduler` の専用テストは、現行 `dev/CaseInfoSystem.Tests` では未確認です。

## この文書で固定する前提

- 依頼上、この current state は build / test / `DeployDebugAddIn` / 実機確認まで通った安定到達点として記録対象にします。
- ただし、その実行ログやチェックリスト結果自体はリポジトリ内では未確認であり、証跡の保管場所は不明です。
- Compile / build 成功と runtime `Addins\` 反映成功、実機確認成功は別物として扱う前提を維持します。

## 残る注意点

- `CaseWorkbookLifecycleService` は分割後も orchestration hub のままで、close 順序依存と CASE HOME 表示補正が同居しています。
- `PostCloseFollowUpScheduler` の visible workbook 判定、Excel busy retry、Excel 終了判定は lifecycle 安定性に直結するため、TaskPaneManager 整理とは切り離して扱うほうが安全です。
- direct `MessageBox.Show` は `CaseClosePromptService` と `CaseWorkbookLifecycleService` の一部に残っています。詳細は `docs/technical-debt.md` を参照してください。
