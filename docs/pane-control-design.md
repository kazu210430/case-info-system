# Pane Control Design (案件情報System)

最終更新: 2026-04

---

## 1. 目的

TaskPane の表示制御を、
- 安定させる
- 拡張しやすくする
- 挙動を一意に説明できる状態にする

---

## 2. 基本構造

Pane 制御は以下の3層に分かれる：

1. **入口制御（suppression）**
2. **表示判定（PaneDisplayPolicy）**
3. **実行（TaskPaneManager / orchestration）**

---

## 3. accepted request

### 定義

`accepted request` とは、

> `TaskPaneDisplayRequest` が生成され、かつ  
> `PaneDisplayPolicy` に渡してよい最小前提を満たした request

### 最小条件

- request が存在する
- target window が存在する
- windowKey を解決できる

### 注意

- 「表示対象であること」は含まない
- accepted の後に `Show / Hide / Reject` に分岐する

---

## 4. suppression

### 定義

> request を無効とするのではなく、  
> 入口で request を発行しない / 消費する条件

### 特徴

- accepted request の**外側**
- `Reject` とは別
- 主に HOME / activate 系で使用

---

## 5. 直refresh系（requestを使わない経路）

対象：

- WorkbookOpen
- WorkbookActivate
- SheetActivate など

### 特徴

- `TaskPaneDisplayRequest` を作らない
- `PaneDisplayPolicy` を通らない
- context（role / window）から直接判定

### 位置づけ

> accepted request 系とは別系統

---

## 6. PaneDisplayPolicy の結果

### 4分類

| 結果 | 意味 | 副作用 |
|------|------|--------|
| ShowExisting | 既存 pane を再表示 | target window |
| ShowWithRender | 新規 render | target window |
| Hide | pane を非表示 | target window のみ |
| Reject | 何もしない | なし |

---

## 7. Hide

### 定義

> accepted request かつ  
> 表示対象ではないが managed pane が残っている場合の局所 conceal

### 特徴

- target window 単位
- render しない
- refresh に流さない

---

## 8. HideAll

### 定義

> 既存の managed pane 状態が現在の context と整合せず、  
> そのまま残すと不正になるため、安全側に全退避する処理

### 特徴

- global（全 window）
- policy の結果ではない
- cleanup / retreat

---

## 9. 境界整理（重要）

### Suppressed
- 入口で止める
- request を発行しない

### Reject
- 表示処理に進めない
- no-op

### Hide
- 局所 conceal

### HideAll
- 全体 retreat

---

## 10. HOME 抑止

### 定義

> HOME 表示と activate の競合を避けるための一時的入口消費

### 重要ルール

- policy に入れない
- `HideAll` に寄せない
- request invalid と扱わない

---

## 11. generic policy に入れるもの

- target window の存在
- windowKey 解決
- role 判定
- managed pane 残存
- render 必要性

---

## 12. generic policy に入れないもの

- HOME 抑止
- suppression 状態
- イベント種別依存（Open / Activate）
- startup 特殊処理

---

## 13. 設計ルール（最重要）

- Suppressed と Reject を混ぜない
- Hide と HideAll を混ぜない
- HOME 抑止を policy に入れない
- accepted request と表示対象を混同しない
- cleanup（HideAll）と判定（Policy）を分離する

---

## 14. 今回あえて未確定の論点

- Reject の厳密定義（狭義/広義）
- request 系と直refresh系の統一
- workbook/window 整合チェックの強化
- HideAll 系の完全整理

---

## 15. 一言まとめ

Pane制御は以下で理解する：

- Suppressed → 入口で止める
- accepted request → policyに渡す
- Policy → Show / Hide / Reject
- HideAll → 失敗時の全体退避
