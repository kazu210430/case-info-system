# UI / 表示制御

## 概要

案件情報System では、CASE 作成後の表示、文書作成中の表示、会計書類セット表示、TaskPane 表示、Excel ウィンドウ復旧を個別に制御しています。この文書では、コードから確認できる表示制御の方針を整理します。

- TaskPane refresh の retry / protection / ready-show の policy 正本は `docs/taskpane-refresh-policy.md` です。
- フロー横断の現行挙動正本は `docs/current-flow-source-of-truth.md` です。

## UI制御の原則

- WorkbookOpen 直接依存の表示制御を行わない
- 表示は専用サービス経由で行う
- ScreenUpdating は必ず復元する
- Window 状態は復旧処理を前提とする
- TaskPane は遅延表示を前提とする
- `WorkbookOpen` 直後の window-dependent refresh は shared policy で skip 判定する
- 裏Excel / hidden session を表示制御の一般手段として使わない
- hidden session を許容するのは、`KernelUserDataReflectionService` の未 open Base / Accounting 反映と、CASE新規作成専用 hidden create session だけに限定する
- `KernelUserDataReflectionService` の managed hidden reflection session は非表示処理として扱い、表示制御経路に昇格させない
- `AccountingSetKernelSyncService` へ補助処理専用の別 `Excel.Application` fallback を再導入しない

## CASE新規作成専用 managed hidden create session と表示境界

- CASE新規作成の hidden create session は非表示の作業経路であり、表示完了や foreground の最終責務を持ちません。
- interactive な CASE 表示は、hidden create session close 後に shared app の `OpenHiddenForCaseDisplay(...)`、`KernelCasePresentationService`、`WorkbookWindowVisibilityService` が引き継ぎます。
- `KernelHomeForm` / `KernelWorkbookCloseService` は CASE 作成フロー中の Kernel HOME close で display restore を skip し、表示済み CASE より前に Kernel を戻さない契約で動作します。
- interactive route と `CreateCaseBatch` はどちらも save 前に owned workbook window を `normal + visible` へ正規化します。
- interactive route の save 前正規化は保存状態の正規化であり、最終表示責務は shared/current app への handoff 後にだけ成立します。
- `app-cache` は `CaseWorkbookOpenStrategy` が所有する retained hidden app-cache の例外であり、裏Excel一般化の根拠にしません。

## 禁止事項

- WorkbookOpen 直後に直接 UI 表示制御を行う実装は禁止する

## WorkbookOpen 直後の shared policy

- `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` を、`WorkbookOpen` 直後の window-dependent refresh skip 判定の正本とします。
- `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` はこの policy の利用者であり、同じ skip 条件を個別に持ちません。
- この policy は pure 判定のみを持ち、ログ出力・状態変更・COMメンバーアクセス・UI操作を含めません。
- 目的は、`WorkbookOpen` 直後に直接 UI 表示制御を行わないという重要ルールをコード上でも 1 か所に集約し、将来のドリフトを防ぐことです。

## 待機 UI

待機 UI の専用サービスが少なくとも次の用途で存在します。

- CASE 作成後の表示待機
- 文書作成時の待機
- 会計書類セット作成時の待機

待機 UI は、処理完了までの一時的な見せ方を担いますが、詳細な表示文言やデザイン方針まではこの文書では扱いません。

### 不明点

- 待機 UI の文言や見た目の正式仕様は、コードだけでは確定しません。

## TaskPane 表示制御

- TaskPane のタイトルは `案件情報System` です。
- TaskPane は左ドックに配置されます。
- Workbook と Window の状態に応じて、既存 Pane の再利用または再描画が行われます。
- 一時抑止、遅延再試行、準備完了後の表示予約が実装されています。
- `WorkbookOpen` 直後に workbook は取得できても window が未解決な refresh は確定させず skip し、後続イベントへ委ねます。
- TaskPane で使う snapshot / cache は表示補助です。保存・生成・実行判断の正本として扱わない方針を維持します。
- retry / protection / ready-show の詳細 policy は `docs/taskpane-refresh-policy.md` を参照します。

### 不明点

- Pane 再利用判定の全条件は、この文書では列挙しません。

## 画面表示制御

画面表示の安定化のため、少なくとも次の制御が確認できます。

- `ScreenUpdating` の一時停止と復元
- 非表示オープンを使った表示準備
- Workbook Window の可視化
- WindowActivate 後の TaskPane 再調整
- Kernel HOME 表示直後の一時的なイベント抑止
- `KernelUserDataReflectionService` の未 open Base / Accounting 反映は managed hidden reflection session で閉じ、save 前の owned workbook window visibility restore を含めても shared/current app の表示経路へは昇格させない

### Kernel HOME unbound 表示

- valid binding を持たない `unbound` HOME は placeholder-only として表示します。
- `unbound` HOME 表示中は、Kernel が既に open でも Kernel workbook / Kernel window へ触れません。
- `unbound` HOME 表示のために、open していない Kernel workbook を探索・open しません。
- `unbound` HOME close 時も、Kernel window を復元対象として扱いません。
- startup で使う open Kernel workbook の有無は bool の表示判定材料として扱い、表示制御のために 1 冊の Kernel workbook を選ぶ API へ昇格させません。

### Kernel HOME close の UI 制御

close / quit のうち `Kernel HOME close` では、UI と backend close の責務を分離します。

- `KernelHomeForm` は close 意思表示を受け、`FormClosing` cancel で close 可否を制御します。
- `KernelWorkbookService` が backend close と HOME session / binding / visibility 解放を調停します。
- HOME close は fail-closed とし、backend close 成功前に Form を閉じません。
- close 失敗時は binding / visibility を維持し、FormClosed finalization へ進めません。
- finalization は backend close 成功後の `FormClosed` に限定します。

### 不明点

- すべての表示不整合ケースに対する期待挙動は、コードだけでは確定しません。

## 前面化制御

前面化は Excel と Word の両方で行われます。

- Excel 側
  - Workbook Window の表示回復
  - 最小化解除
  - 前面化 API 呼び出し
- Word 側
  - 生成した文書の表示
  - 必要に応じた前面化 API 呼び出し

Kernel HOME も WinForms 側で `Show`、`Activate`、`BringToFront` を使って表示されます。

### 不明点

- 前面化 API 呼び出しの個別条件分岐までは、この文書では列挙しません。

## Excel ウィンドウ復旧

Excel ウィンドウ復旧専用のサービスが存在します。確認できる処理は次のとおりです。

1. Excel Application の可視状態を確認します。
2. `ScreenUpdating` を有効に戻します。
3. 対象 Workbook の Window を解決または再取得します。
4. Window を可視化します。
5. 最小化状態であれば通常表示へ戻します。
6. 条件に応じて前面化します。

この復旧処理は、CASE 表示や表示失敗後の回復と関係します。

### 不明点

- 復旧失敗時の再試行や代替経路の全件は、この文書では列挙しません。

## CASE close 後の白Excel防止

close / quit のうち `CASE managed close` と `post-close quit` では、visible workbook が無ければ `PostCloseFollowUpScheduler` が `Quit` を試み、白 Excel を残さないことを設計目標とします。

- `Quit` 成功後は終了中 `Application` を restore しません。
- `Quit` 失敗時だけ `DisplayAlerts` を restore します。
- これは今回安定化対象の managed close / quit 経路の話であり、全 close 経路の一般ルールではありません。

### 不明点

- `白Excel` 相当の不具合に対する運用上の呼称は、コードだけでは確定しません。

## CASE HOME の見え方維持

- CASE HOME では、左列固定に関する再適用処理があります。
- `FreezePanes`、`SplitColumn`、`ScrollColumn`、`ScrollRow` の制御により、表示位置を安定させようとする実装が確認できます。

### 不明点

- 左列固定の再適用が走る全契機は、この文書では列挙しません。

## Application.DoEvents 使用禁止

- Application.DoEvents() は原則使用禁止とする
- 理由:
  - メッセージキュー内のすべてのイベントを処理してしまうため、処理途中で再入が発生する
  - 再入により状態不整合・二重実行・表示崩れが発生するリスクがある

特にOfficeアドインでは以下の問題を引き起こす:

- Excel/Wordイベントの割り込み
- UI状態の不整合
- 再現性の低い不具合

### 代替方針

- UI更新は以下で行う
  - BeginInvoke
  - await / 非同期処理
  - UIスレッド内の通常描画（Refresh / Invalidate）

- UIスレッドをブロックしない設計を優先する

### 例外

以下をすべて満たす場合のみ使用検討可

- 再入しても安全な処理であることを説明できる
- 呼び出し範囲が完全に限定されている
- 他UI・業務フローへ影響しないことが保証できる
