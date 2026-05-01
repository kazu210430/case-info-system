# Workbook Window Activation Notes

## 目的

この文書は、`WorkbookOpen` / `WorkbookActivate` / `WindowActivate` のイベント順序と、window 依存処理を安全に扱う境界を固定するための補足メモです。

- 全体構成の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- TaskPane 設計の前提: `docs/taskpane-architecture.md`
- UI 制御の前提: `docs/ui-policy.md`

## 結論

1. `WorkbookOpen` は workbook が開いた通知であり、window 安定通知ではありません。
2. `WorkbookOpen` 時点では `ActiveWorkbook` と `ActiveWindow` が未確定な場合があります。
3. window 依存処理の安全境界は `WorkbookActivate` 以降、必要なら `WindowActivate` 以降です。
4. workbook-only 処理と window-dependent 処理を混ぜない方針を優先します。
5. startup context 系の再導入より先に、このイベント境界の安定化を優先する必要があります。

## 確認できたイベント順序

現行コードと実機ログで、次の順序が確認できます。

1. `WorkbookOpen`
2. `WorkbookActivate`
3. `WindowActivate`

補足:

- `WorkbookOpen` では workbook 引数自体は渡されます。
- ただし、その workbook が active workbook になっているとは限りません。
- `WindowActivate` は active workbook の確定後、実 window が activate された段階として扱います。

## ActiveWorkbook / ActiveWindow が未確定になる条件

コードとログから、次の条件が確認できます。

- 新しい Excel セッション起動直後である
- workbook は開いたが、まだ active workbook になっていない
- visible window がまだ確定していない
- `Application.ActiveWorkbook` / `Application.ActiveWindow` getter が一時的に空または例外になる

このため、`WorkbookOpen` だけを根拠に window 前提処理へ進むのは危険です。

## ResolveWorkbookPaneWindow が安全に成功する条件

現行コードでは、window 解決は実質的に次のどちらかが必要です。

1. 対象 workbook から visible window を取得できる
2. active workbook が対象 workbook と一致し、active window を取得できる

この条件を満たさない場合、window 解決失敗は異常ではなく、単に timing 未確定として扱うべきです。

## 単体生成 CASE 再オープン調査結果

調査で確認できた事実:

- 単体生成 CASE の再オープンで、`WorkbookOpen` 時点の `ActiveWorkbook` / `ActiveWindow` が空のケースがありました。
- 同じケースで `ResolveWorkbookPaneWindow` は失敗しました。
- その後 `WorkbookActivate` で `ActiveWorkbook` と window が確定し、後続処理は回復しました。

この結果から、白 Excel 相当の不整合は `WorkbookOpen` 時点での window 未確定を安全に扱えていないことと関係します。

## 設計上の扱い

### Workbook-only 処理

- role 判定
- workbook 単位の lifecycle 初期化
- suppression 登録や既知 path 更新など、window を必要としない処理

これらは `WorkbookOpen` で扱ってよいです。

### Window-dependent 処理

- pane 対象 window の確定
- window 可視化や前面化
- window key に依存する pane host 再利用
- active window を前提とする UI 表示調停

これらは `WorkbookActivate` 以降、必要なら `WindowActivate` 以降を安全境界として扱います。

## 先に安定化すべきこと

- startup context 系の再導入や再分解より前に、`WorkbookOpen` と `WorkbookActivate` の境界を整理すること
- window 未確定を例外扱いせず、正常な timing 差として扱うこと
- workbook-only と window-dependent の責務を混ぜないこと

## 不明として残す事項

- すべての Excel 実行環境で `WorkbookActivate` と `WindowActivate` のどちらを最終安全境界とするのが最適か
- active workbook 未確定が出る環境差や OneDrive 同期状態との相関
- ready-show retry と protection の最適秒数

これらはコードだけでは確定せず、実機観測を前提に扱います。
