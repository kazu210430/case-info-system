# Case Workbook Lifecycle Current State

## 位置づけ

この文書は、`CaseWorkbookLifecycleService` 分割と、その後続整理である `CaseClosePromptService` / `CaseFolderOpenService` 分離後の現在地を、現行 `main` のコードベースに合わせて固定するための補助文書です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- TaskPane 側の現在地: `docs/taskpane-refactor-current-state.md`
- 残課題: `docs/technical-debt.md`

この文書は TaskPaneManager 側の追加整理を扱いません。lifecycle 側の安定到達点だけを明文化します。

今回この文書で固定するのは、close / quit のうち `Kernel HOME close`、`Kernel managed close`、`CASE managed close`、`post-close quit` の安定化到達点だけです。全 workbook close 経路の一般ルールではありません。`KernelUserDataReflectionService`、`MasterWorkbookReadAccessService`、`CaseWorkbookOpenStrategy` などの読み取り専用 / 一時 workbook close は今回の確定範囲に含めません。

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
   - dirty path では `CaseWorkbookLifecycleService` が dispatcher 経由で managed close を予約し、`ManagedCloseState` のスコープ内で save 有無を処理し、今回安定化対象の managed close 経路では `WorkbookCloseInteropHelper.CloseWithoutSave(workbook)` を使って `false, Type.Missing, Type.Missing` の optional 引数を明示した close へ進めます。
7. post-close follow-up
   - dirty path では managed close 内で、clean close では before-close 中に `PostCloseFollowUpScheduler` を予約します。
8. close 後判定
   - `PostCloseFollowUpScheduler` は対象 workbook が残っていないことを確認し、Excel busy なら retry し、visible workbook が 1 つも無い場合だけ Excel 終了を試みます。`Quit` 成功後は終了中 `Application` を restore せず、`DisplayAlerts` の restore は失敗時だけに限定します。

dirty path の大まかな順序は `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` です。

## 今回安定化対象の close / quit 境界

### Kernel HOME close

- `KernelHomeForm` は close 意思表示を受け、`FormClosing` cancel で close 可否を制御します。
- `KernelWorkbookService.RequestCloseHomeSessionFromForm(...)` が backend close を調停します。
- HOME close は fail-closed とし、backend close 成功後にのみ HOME session / binding / visibility を解放します。
- close 失敗時は Form を閉じず、binding / visibility を維持します。
- HOME session の finalization は `FormClosed` 後の `FinalizePendingHomeSessionCloseAfterFormClosed()` に限定します。

### Kernel / CASE managed close

- 今回安定化対象の managed close 経路では `WorkbookCloseInteropHelper` を経由します。
- `Workbook.Close(SaveChanges: false)` のような named argument は使いません。
- save ありの Kernel managed close では `Type.Missing, Type.Missing, Type.Missing`、save なしの Kernel / CASE managed close では `false, Type.Missing, Type.Missing` を明示して渡します。
- close 後に対象 workbook を再参照しません。

### managed close / post-close quit の最小 COM 原則

- 今回安定化対象の managed close / quit 経路では `Save` / `DisplayAlerts` / `Close` / `Quit` 以外の COM 操作を増やしません。
- この経路では `ExcelApplicationStateScope` を使いません。
- `DisplayAlerts` は個別の `try/finally` または同等の局所 restore で扱います。
- `Quit` 成功後は終了中 `Application` を restore しません。
- `Quit` 失敗時だけ `DisplayAlerts` を restore します。

### CASE post-close quit

- `PostCloseFollowUpScheduler` が visible workbook を確認し、残っていなければ `Quit` を試みます。
- 設計目標は CASE close 後に白 Excel を残さないことです。
- これは今回安定化対象の managed close / quit 経路の話であり、全 close 経路の一般ルールではありません。

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

## Shadow copy / 実機反映

- Excel が起動中だと古い shadow copy DLL が使われ続けることがあります。
- 実機確認前は Excel を完全終了します。
- 実行 DLL の確認は `Runtime execution observed` ログの `assemblySha256` を使います。

## 残る注意点

- `CaseWorkbookLifecycleService` は分割後も orchestration hub のままで、close 順序依存と CASE HOME 表示補正が同居しています。
- `PostCloseFollowUpScheduler` の visible workbook 判定、Excel busy retry、Excel 終了判定は lifecycle 安定性に直結するため、TaskPaneManager 整理とは切り離して扱うほうが安全です。
- direct `MessageBox.Show` は `CaseClosePromptService` と `CaseWorkbookLifecycleService` の一部に残っています。詳細は `docs/technical-debt.md` を参照してください。
- helper 非経由 close が `KernelUserDataReflectionService`、`MasterWorkbookReadAccessService`、`CaseWorkbookOpenStrategy` などに残っています。
- `WorkbookPromptSuppressionHelper` の `Workbook.Saved` 操作は今回対象外です。
- これらは別途棚卸し対象であり、今回 docs の確定範囲外です。
