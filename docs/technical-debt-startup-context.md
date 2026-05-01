本ファイルは現時点では実装変更対象ではなく、startup context 周辺に残している technical debt の整理メモである。

# Startup Context Technical Debt

## 1. 概要

`KernelStartupContextInspector` による startup / 初期 HOME 表示判定向けの責務分離自体は完了している。

ただし、事実収集の一部には Window 依存の観測が残っている。

具体的には、可視 non-kernel workbook の有無を把握するために、`HasVisibleNonKernelWorkbook` 相当の処理で次を参照している。

- `workbook.Windows`
- `window.Visible`

この依存は現時点では意図的に残しているものであり、挙動維持のためには正しい。一方で、将来的には startup context 本体から分離可能な関心事として整理候補に置く。

## 2. 現状の構造

### `HasVisibleNonKernelWorkbook` の役割

`HasVisibleNonKernelWorkbook` の役割は、「Kernel 以外の workbook に visible な window が存在するか」という事実だけを観測することにある。

この観測結果は、startup 時の HOME 表示判定に渡す材料の一部として使われる。

### `workbook.Windows` / `window.Visible` を使用している理由

workbook 単位で open かどうかを見るだけでは、「現在 visible な non-kernel workbook があるか」は確定できない。

そのため現行実装では、対象 workbook が持つ `Windows` コレクションを列挙し、その中に `Visible == true` の window があるかどうかを見ている。

### これは window 制御ではなく可視状態の観測である

この処理は次を行っていない。

- `ActiveWindow` の解決
- window の activate
- window の復帰
- window の前面化
- pane 対象 window の確定

つまり、これは「window 制御」ではなく、「可視状態の観測」に限った依存である。

## 3. なぜ今は分離しないか

現時点では、振る舞い不変が最優先である。

`docs/flows.md` にあるとおり、

- `WorkbookOpen` は window 安定境界ではない
- window 依存処理は `WorkbookActivate` / `WindowActivate` 以降で扱う

という原則は維持する必要がある。

一方で、`HasVisibleNonKernelWorkbook` の window 列挙は、UI 制御のためではなく、既存の HOME 表示判定に渡す「可視 workbook の有無」という事実を旧挙動のまま維持するために残している。

ここを今変更すると、次のリスクがある。

- startup 時の HOME 表示判定が変わる
- 他 workbook の表示有無に応じた既存挙動が変わる
- 結果として表示制御や白 Excel 回避の既存バランスに影響する

このため、現段階では安全に切り出せる単位とは扱わず、意図的に残す。

## 4. 将来の分離案

候補としては、`WorkbookVisibilityInspector`（仮）を別サービスとして切り出す。

### 想定責務

- workbook の可視状態の列挙
- `window.Visible` の観測
- visible workbook の有無に関する純粋な事実収集

このサービスは、window 制御責務を持たず、`Visible` 観測だけを担当する純粋観測サービスとして設計する。

### 分離後の理想構造

`KernelStartupContextInspector`
↓
`WorkbookVisibilityInspector`
↓
Window 列挙

この形にできれば、startup context 側は「startup 用事実収集」に集中し、window 可視状態の列挙は別関心事として整理できる。

## 5. 優先度

優先度は B〜C とする。

他の責務分離、特に次の整理より後順位で扱う。

- TaskPane 系の責務整理
- 既存 service 境界の整理

理由は、現行挙動への影響リスクに対して、直近の実益が限定的だからである。

## 6. 注意点

- `ActiveWindow` 依存と `window.Visible` 観測は別概念である
- window 復帰制御とは明確に切り離す
- COM 寿命管理を新たに広げない
- window handle を扱う責務を持たせない
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の境界整理を変更しない

補足として、この負債メモは「今すぐ修正すべき不具合」の記録ではない。現時点では正しい実装として残しつつ、将来の責務分離時に誤って UI 制御や window 復帰責務と混ぜないための記録である。
