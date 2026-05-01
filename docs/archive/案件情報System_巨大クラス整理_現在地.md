# 注意（アーカイブ）

このドキュメントは過去の設計状態を示すものであり、
現在の実装とは一致しません。

最新の状態は以下を参照してください。
- docs/taskpane-architecture.md

# 案件情報System  
## 巨大クラス分割・未使用コード整理の現在地（2026-04）

---

## ■ 目的
巨大クラスを安全に分割・整理し、
将来の機能追加に耐える構造へ移行する。

- 既存挙動維持最優先
- 危険領域は触らない
- 小さく分割 → 判断 → 次へ
- 一気にやらない

---

## ■ 完了済み（重要）

### KernelWorkbookService
- Resolver 分離済み
- StateService 分離済み
- SettingsService 分離済み
- private leftover 削除済み
- 👉 危険領域専任 orchestrator として完成

---

### TaskPaneManager
- 棚卸し完了
- 分割非推奨と判断
- DI 残骸削除完了（CompositionRoot側）
- 👉 pane制御本体として固定

---

### ThisAddIn
- Workbook 系 → Coordinator 化完了
- Sheet 系 → Coordinator 化完了
- SelectionChange / AfterCalculate → 最小委譲化完了
- 👉 境界クラス（受けて流すだけ）へ到達

---

## ■ 現在の状態（重要）

### Before
巨大クラスが混在

ThisAddIn = 何でもやる  
Kernel = 巨大  
TaskPane = 巨大  

### After
責務分離済み

ThisAddIn = 境界（イベント受信 + guard）  
KernelWorkbookService = 危険領域 orchestration  
TaskPaneManager = pane制御本体  
State / Settings / Resolver = 純粋ロジック  

👉 巨大クラス問題は実質解消済み

---

## ■ 残っている作業

### 次対象
👉 **AccountingWorkbookService**

理由：
- 未使用候補が明確
- COM系だが整理しやすい
- 最後の「削減フェーズ」

---

## ■ 作業方針（固定）

### 分割
- すでに終了（これ以上やらない）

### 削除
- 明確な未使用のみ
- 範囲を広げない

---

## ■ 危険領域（触らない）

- COM操作
- lifecycle（open/close）
- window制御
- pane表示本線
- HOME制御
- save/quit
- event suppression

---

## ■ ゴール

- 巨大クラスを消すことではない
- 👉 **責務が明確で壊れない構造にすること**

---

## ■ 現在地（超重要）

[完了]  
KernelWorkbookService  
TaskPaneManager  
ThisAddIn  

[進行中]  
AccountingWorkbookService ← ここから再開  

---

## ■ 一言まとめ

- 分割フェーズ → 完了
- 構造整理 → 完了
- 👉 これからは「仕上げ（削減）」フェーズ

---

## ■ 作業開始宣言（新チャット用）

AccountingWorkbookService の最終整理から再開します。
