# ThisAddIn Startup Test Checklist

## 位置づけ

この文書は、`ThisAddIn` / `Startup` / `TaskPane` 初期表示まわりのリファクタ後に確認すべき実機テスト観点を、現行 `main` 基準で固定するための checklist です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- Startup 順序固定と add-in 境界の前提: `docs/thisaddin-boundary-inventory.md`
- TaskPane 現在地の前提: `docs/taskpane-refactor-current-state.md`
- workbook / window 境界の補足: `docs/workbook-window-activation-notes.md`

この文書は build / test を置き換えるものではありません。特に `ThisAddIn` / `Startup` / `TaskPane` 初期表示まわりは、build / test OK と実機 OK を別扱いで記録します。

## 基準フロー

実機確認時は、少なくとも次の基準フローを崩していないことを確認します。

1. `ThisAddIn_Startup(...)`
2. `InitializeStartupDiagnostics()`
3. `CreateStartupCompositionRoot()` -> `Compose()` -> `ApplyCompositionRoot()`
4. `InitializeStartupBoundaryCoordinator()`
5. `InitializeAdaptersAfterComposition()`
6. `HookApplicationEvents()`
7. `AddInStartupBoundaryCoordinator.RunAfterApplicationEventsHooked()`（startup HOME 判定、`RefreshTaskPane("Startup", null, null)`、managed-close startup guard）
8. `WorkbookOpen -> WorkbookActivate -> WindowActivate`
9. `ThisAddIn_Shutdown(...)`

Startup 順序の正本は `docs/thisaddin-boundary-inventory.md` の Startup 順序固定メモとします。

## 実機テスト checklist

### 起動直後

- [ ] Excel 起動時に Add-in が例外なく正常起動する。
- [ ] 起動直後に余計な白画面、空 window、残留 window が出ない。
- [ ] `ActiveWorkbook` が `null` でも例外にならず、起動継続できる。
- [ ] startup trace / diagnostics が従来どおり観測できる。

### Kernel HOME

- [ ] `Kernel HOME` が従来どおり表示される。
- [ ] suppression / protection 条件では `Kernel HOME` が表示されない。
- [ ] `Kernel HOME` の表示判定が startup context 判定後に行われている。
- [ ] Kernel workbook open と `Kernel HOME` 表示判定の関係が崩れていない。

### 初回 TaskPane refresh

- [ ] 初回 `RefreshTaskPane("Startup", null, null)` が `Kernel HOME` 表示判定後に行われる。
- [ ] Startup 時に window 依存処理を `WorkbookOpen` 前提で実行していない。
- [ ] `WorkbookActivate` / `WindowActivate` 後に TaskPane が正しく表示、再表示される。
- [ ] CASE / Kernel の role 判定が混線しない。

### CASE open / activate / close

- [ ] CASE open 後に TaskPane が従来どおり表示される。
- [ ] `WorkbookActivate` / `WindowActivate` で pane が再表示される。
- [ ] 複数 CASE / 複数 window で TaskPane が混線しない。
- [ ] CASE close 時に残 pane、白画面、空 window、例外が出ない。

### Excel close / Shutdown

- [ ] Excel close 時に例外が出ない。
- [ ] event unwiring 由来の例外が出ない。
- [ ] TaskPane host dispose が従来どおり行われる。
- [ ] shutdown trace が従来どおり観測できる。

## 守るべき前提

- [ ] `WorkbookOpen` は window 安定境界ではない。
- [ ] window 依存処理は `WorkbookActivate` / `WindowActivate` 以降で扱う。
- [ ] `ActiveWorkbook` は `null` になりうる。
- [ ] Startup 順序固定 docs を正とする。
- [ ] 実機 OK と build / test OK は別扱いにする。
