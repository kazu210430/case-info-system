# TaskPane Refactor Current State

## 位置づけ

この文書は、TaskPane 側の優先度Aリファクタについて、現行 `main` で確認できる到達点を固定するための現在地文書です。

- TaskPane 設計正本: `docs/taskpane-architecture.md`
- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- 優先度A棚卸し: `docs/a-priority-service-responsibility-inventory.md`
- protection / ready-show 危険領域の補足:
  - `docs/taskpane-protection-ready-show-investigation.md`
  - `docs/taskpane-protection-baseline.md`
  - `docs/taskpane-protection-observation-checklist.md`

この文書は設計正本を置き換えるものではありません。TaskPane 優先度Aで「どこまで main に固定済みか」「どこが未着手・保留か」を明示するための補助文書です。

## 今回固定する到達点

現行 `main` に対して、TaskPane 側の優先度A到達点は次の整理で固定します。

1. TaskPane の runtime 設計正本は `docs/taskpane-architecture.md` とする。
2. 文書ボタン定義の正本、Base 埋込 snapshot、CASE cache、prompt / resolver の責務分離は、`docs/taskpane-architecture.md` の記述を現行到達点として扱う。
3. 優先度Aのうち、production code 変更なしで完了確認できた棚卸し結果は `docs/a-priority-service-responsibility-inventory.md` を基準に読む。
4. protection / ready-show / retry / suppression を含む危険領域は、未着手・保留として扱い、完了済みとはみなさない。
5. 実機観測が必要な論点は未確定のまま残し、コードだけでは断定しない。

## 完了済みとして固定する事項

### 1. TaskPane 設計正本の固定

- TaskPane の正本は Kernel `雛形一覧` と Kernel `TASKPANE_MASTER_VERSION` である。
- Base 埋込 snapshot と CASE snapshot cache は、いずれも派生 cache であり正本ではない。
- TaskPane 表示の解決順は `CASE cache -> Base cache -> Master rebuild` である。
- 開いている CASE は、後から成功した雛形登録・更新へ自動追随しない。
- `DocumentNamePromptService` は CASE cache だけを参照し、master fallback しない。
- `DocumentTemplateResolver` は CASE cache 優先で解決し、miss 時のみ master fallback する。

### 2. TaskPane 周辺で完了済みとして扱う bridge / 境界整理

`docs/a-priority-service-responsibility-inventory.md` を基準に、現行 `main` で完了済みとして扱うのは次です。

- `DocumentCommandService`
  - `ScreenUpdating`、TaskPane refresh suppression、active refresh、Kernel sheet refresh は bridge 経由へ整理済み。
- `WindowActivatePaneHandlingService`
  - `ShouldIgnoreWindowActivateDuringCaseProtection(...)` 判定は bridge 経由へ整理済み。
- 補助境界として確認済みの事項
  - `TaskPaneHost` は `Globals.ThisAddIn` ではなく constructor 注入の `ThisAddIn` を VSTO `CustomTaskPane` の生成・破棄境界として使う。
  - `TaskPaneHost` 自体は表示判断を持たない薄い host ラッパーとして扱う。

### 3. docs 側で固定済みの危険領域棚卸し

次の論点は、すでに docs 上で危険領域として棚卸し済みであることを到達点に含めます。

- ready-show / suppression の順序を壊してはいけないこと
- `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の protection 判定が連動していること
- retry `80ms`、fallback timer `400ms`、`3 attempts` はコード上の事実として確認できるが、仕様根拠は未確認であること
- visible pane early-complete が既存 CASE pane の不要な refresh 回避に使われること

## 未着手・保留として固定する事項

次は優先度Aに含まれるが、現時点では完了済みへ移さない領域です。

- `KernelCasePresentationService`
  - ready-show 要求前後の suppression / release / workbook window 可視化の順序を含む危険領域
- `TaskPaneRefreshOrchestrationService`
  - retry、fallback、protection 最上流判定、visible pane early-complete を含む危険領域
- `TaskPaneRefreshCoordinator`
  - CASE refresh 完了後の foreground 保証と protection 開始を含む危険領域
- `WorkbookLifecycleCoordinator`
  - `WorkbookActivate` 再入抑止の判定境界
- `TaskPaneManager`
  - host ライフサイクル、role 別描画、action 後 refresh を一体で持つ pane 制御本体

これらは「未着手」または「保留」の扱いを維持し、今回の現在地文書で完了扱いへ動かしません。

## 今回の到達点に含めない事項

次は現行 docs / code だけでは確定しないため、到達点として固定しません。

- retry 値や protection 5 秒の正式な仕様根拠
- Pane 再利用判定の全条件
- 実機でのちらつき、二重表示、出遅れの最終観測結果
- `WindowActivate` 固有の体感挙動の完全な期待仕様

## 次の実装着手時に守ること

- `docs/taskpane-architecture.md` を設計正本として維持する
- `WorkbookOpen` 直後に直接 UI 表示制御を追加しない
- snapshot / cache を保存・生成・実行判断の正本へ戻さない
- ready-show / suppression / protection の順序を変える変更は、危険領域として別途確認してから扱う
- host 再利用経路と visible pane early-complete を安易に単純化しない

## 一言まとめ

TaskPane 側の優先度Aは、設計正本・責務棚卸し・危険領域の事実整理までは `main` に固定済みです。

一方で、ready-show / protection / retry / host 再利用を含む本線ロジックは、まだ完了済みとは扱わず、未着手・保留として残します。
