# A2 Window Visibility Current State

## 位置づけ

この文書は、現行 `main`（基準点: `31b803ddf5f4cc70ddf1df6b9a0856dbb183c8ce`）における A2 の到達点を、新チャット用の current state として固定するための補助文書です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御方針の前提: `docs/ui-policy.md`
- Workbook / Window 境界の補足: `docs/workbook-window-activation-notes.md`
- CASE lifecycle 側の現在地: `docs/case-workbook-lifecycle-current-state.md`
- TaskPane refresh policy の正本: `docs/taskpane-refresh-policy.md`
- service responsibility 棚卸し: `docs/a-priority-service-responsibility-inventory.md`

今回の目的は、A2 の現在地と、今後に意図的に積み残した領域を docs に固定することです。ここでは実装開始を宣言せず、現行 `main` で確認できる ownership と guard だけを記録します。

## A2 現在地

完了済み:

- A2-1 execution ownership
- A2-2 visibility policy ownership
- A2-3-1 ReleaseHomeDisplay branch dispatch

### A2-1

Excel main window / foreground execution ownership 分離

到達点:

- `KernelWorkbookService`
  - 判断 owner
- `ExcelWindowRecoveryService`
  - Win32 / app hwnd / foreground execution owner

### A2-2

HOME visibility policy ownership 分離

到達点:

- `KernelWorkbookService`
  - visibility orchestration owner
- `KernelWorkbookHomeDisplayVisibilityPolicy`
  - visibility decision owner
- `ExcelWindowRecoveryService`
  - execution owner

### A2-3-1

ReleaseHomeDisplay post-decision branch dispatch 明示化

到達点:

- `KernelWorkbookService`
  - lifecycle / orchestration owner
- policy 群
  - decision owner
- `ReleaseHomeDisplay` branch dispatch
  - execution branch readability
- `ExcelWindowRecoveryService`
  - recovery primitive owner

`KernelWorkbookService.ReleaseHomeDisplay(...)` は、release action の決定と実行 branch を分離し、`SkipRestore` / `PromoteAndRestore` / `RestoreWithoutPromotion` の dispatch を明示的に読む構造になっています。

## ownership の現在地

### `KernelWorkbookService`

- visibility orchestration owner
- HOME lifetime owner
- `ReleaseHomeDisplay` lifecycle owner
- branch dispatch owner

### `KernelWorkbookHomeDisplayVisibilityPolicy`

- HOME visibility decision owner

### `KernelWorkbookWindowRestorePolicy`

- restore decision owner

### `KernelWorkbookPromotionPolicy`

- promotion decision owner

### `KernelWorkbookHomeReleaseFallbackPolicy`

- release fallback decision owner

### `ExcelWindowRecoveryService`

- app/window recovery primitive owner
- foreground / app hwnd execution owner

## 維持済みガード

- `WorkbookOpen -> WorkbookActivate -> WindowActivate` 境界維持
- `WorkbookOpen` 直後の window-dependent UI 制御禁止
- `ScreenUpdating` 発火条件不変
- `ShowExcelMainWindow() -> ShowKernelWorkbookWindows(true)` 順序維持
- `RestoreWithoutPromotion` は `ShowKernelWorkbookWindows(false)` の direct restore のまま
- `activateWindow` / `bringToFront` semantics 不変
- HOME close fail-closed / pending close / `FormClosed` handshake 不変
- TaskPane ready-show / protection 未介入
- Word foreground 未介入

## 意図的に未着手の領域

以下は「未対応」ではなく、安全単位を超えるため意図的に未着手とする領域です。

- HOME visibility lifetime
- foreground retry semantics
- visible window resolve ownership
- CASE ready-show / TaskPane protection
- `PostCloseFollowUpScheduler`
- close retry semantics
- HOME close fail-closed handshake
- `ShowKernelWorkbookWindows` の内部分割
- `ExcelWindowRecoveryService` への restore 統一
- Word foreground / foreground API 共通化

## STOP 判断理由

- HOME visibility lifetime:
  pending close / `FormClosed` / close retry と結合しており、visibility policy だけでは切れない。
- foreground retry semantics:
  HOME / CASE / TaskPane で retry の意味が異なり、共通 abstraction 化すると timing ownership が崩れる。
- visible window resolve ownership:
  recovery / ready-show / visibility ensure で fallback semantics が異なり、統合すると `WorkbookOpen` timing 境界へ波及する。
- CASE ready-show / TaskPane protection:
  TaskPane timing と protection 開始条件に結合しており、A2 の window visibility ownership と混ぜると危険。
- `PostCloseFollowUpScheduler`:
  no visible workbook 判定、Excel busy retry、`Quit` 判定まで持っており、HOME visibility ownership とは安全単位が異なる。
- close retry semantics:
  HOME close fail-closed と CASE managed close / post-close follow-up では retry の責務境界が異なる。
- HOME close fail-closed handshake:
  pending close 登録、backend close 完了、`FormClosed` finalization が一体で、window visibility 側だけ先に触ると close safety を壊す。
- `ShowKernelWorkbookWindows` の内部分割:
  app visible ensure、window visible restore、promotion 条件が近接しており、今の A2 範囲では実行責務の再混線を招きやすい。
- `ExcelWindowRecoveryService` への restore 統一:
  `RestoreWithoutPromotion` に `ScreenUpdating` restore や activation semantics を新たに持ち込む危険がある。
- Word foreground / foreground API 共通化:
  Word 側は文書表示フローに属し、Excel HOME / CASE の window visibility ownership と同一単位で扱えない。

## 今後の候補

次に進む場合の候補:

- B1 TaskPane 起動配線の循環解消
- A2 深掘りを続ける場合は、まず別安全単位として棚卸し必須
- HOME lifetime / retry / resolver / TaskPane protection は即実装禁止

## 維持すべき読み方

- A2 は「window visibility / foreground execution ownership の整理」までを到達点とする。
- A2 は TaskPane retry / protection の仕様固定や削減を完了したことを意味しない。
- A2 は HOME close / CASE close / post-close quit の全体整理を完了したことを意味しない。
- `WorkbookOpen` を window 安定境界として扱う読み方へ戻さない。
- open 中 workbook の host 再利用、TaskPane ready-show、CASE protection の policy 群とは切り分けて読む。

## 今回の docs 作業での禁止事項

- コード変更
- build 必須化
- 実装開始
- A2 の未着手領域を「次に必ずやる」と断定すること
- `architecture.md` / `flows.md` / `ui-policy.md` への過剰追記
