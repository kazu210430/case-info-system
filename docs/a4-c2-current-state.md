# A4 C2 Current State

## 位置づけ

この文書は、現行 `main`（基準点: `5af1b8cf6ecb1a82f15862a75d9c35ddac35aa90`）における B1 -> B2-1 -> A4 -> C1 -> C2 完了時点の現在地を、cross-cutting checkpoint として固定するための補助文書です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御方針の前提: `docs/ui-policy.md`
- TaskPane 側の現在地: `docs/taskpane-refactor-current-state.md`
- TaskPane 設計正本: `docs/taskpane-architecture.md`
- metadata / resolver / eligibility の補足: `docs/template-metadata-inventory.md`
- TaskPane deferred items の補足: `docs/taskpane-refactor-deferred-items.md`

今回の目的は、現行 `main` で確認できる到達点と、今後に意図的に積み残した領域を docs に固定することです。ここでは新規実装案を追加せず、ownership、guard、fail-closed 方針だけを記録します。

この文書は docs-only checkpoint です。今回の docs 更新自体が build、test、runtime `Addins\` 反映、実機観測の再実施を意味するものではありません。

## 完了済み安全単位

### A1

HOME session close 系整理

- `Kernel HOME close`
- `Kernel managed close`
- `CASE managed close`
- `post-close quit`

### A2

window visibility / foreground ownership 整理

- execution ownership と visibility policy ownership を分離済み
- `WorkbookOpen -> WorkbookActivate -> WindowActivate` 境界を維持したまま整理済み
- 詳細な現在地は `docs/a2-window-visibility-current-state.md` を参照

### A3

`KernelTemplateSyncService` ownership 整理

- validation preflight ownership
- publication execution ownership
- temporary worksheet state lifetime ownership
- Base snapshot storage ownership
- master list row payload ownership

### B1

TaskPane 起動 wiring の composition-time cycle 解消

- placeholder self-reference 撤去
- retry / ready-show / protection semantics 不変

### B2-1

plain OK notification の `UserErrorService` 集約

- 対象:
  - `KernelCommandService`
  - `WorkbookCaseTaskPaneRefreshCommandService`
  - `UserErrorService`
- 非対象:
  - `CasePaneCacheRefreshNotificationService`
  - `TaskPaneRefresh*` 系

### A4

publication side effects ownership 整理

- publication side effects は `PublicationExecutor` に集約済み
- 順序は次で固定する
  - `WriteToMasterList`
  - `TASKPANE_MASTER_VERSION +1`
  - Kernel save
  - Base snapshot sync
  - `InvalidateCache`
- failure semantics は次で固定する
  - preflight failure: 副作用なし
  - kernel save failure: Base sync / invalidate へ進まない
  - base sync failure: invalidate 実行、success + warning 維持
- `SYSTEM_ROOT` 文脈、invalidate API、cache key 解決方式は不変

### C1

`DocumentExecutionEligibilityService` direct contract test 追加

- `.doc` unsupported
- `.docm` unsupported
- `.dotm` macro reject
- lookup miss fail-closed
- template dir 未導出 fail-closed
- template file 不在 fail-closed
- output folder 不在 fail-closed
- case context null fail-closed
- case snapshot 空 fail-closed
- 実行可能ケースを通す
- `DocumentName` 空は warning のみ

### C2

`DocumentTemplateResolver` direct contract test 追加

- key trim / normalize
- lookup miss -> `null`
- `WORD_TEMPLATE_DIR` 優先
- `SYSTEM_ROOT\雛形` fallback
- dir 未導出時は `null` ではなく empty path spec
- `ResolutionSource` 引き継ぎ
- supported extension helper
- `.dotm` は resolver では supported
- macro reject は eligibility 側責務

## 維持すべき重要ガード

- `WorkbookOpen -> WorkbookActivate -> WindowActivate` 境界
- `WorkbookOpen` 直後の window-dependent UI 制御禁止
- TaskPane ready-show semantics
- protection timing
- retry semantics
- fail-closed 方針
- `SYSTEM_ROOT` 文脈
- foreground semantics
- `雛形一覧` D:F 手修正前提
- 文書実行の許可境界としての `.docm` / `.dotm` 非許可方針
- notification wording / timing

## 意図的に積み残した領域

### B2-2

`CasePaneCacheRefreshNotificationService`

- notification service に見えるが、実際には `workbook.Saved` restore、cache refresh side effect、`WorkbookOpen` / `WorkbookActivate` timing、policy gating を持つ
- TaskPane lifecycle ownership に近い危険領域のため、今回の完了済み範囲には含めない

### TaskPane protection / ready-show 系

- `CASE ready-show policy`
- TaskPane protection timing
- protection retry semantics
- `WorkbookOpen` timing 周辺

この領域は未着手のまま残す。helper 分離済みの箇所があっても、TaskPane lifecycle が整理完了したとは読まない。

### visible window resolve ownership

- visible window resolve 統一
- foreground retry semantics
- `ExcelWindowRecoveryService` への restore 完全統一

この領域は `WorkbookOpen` timing 境界、ready-show、visibility ensure と結合しているため未着手のまま残す。

### HOME visibility lifetime

- HOME lifecycle と visibility lifetime が密結合しているため未着手

### `PostCloseFollowUpScheduler`

- close retry semantics
- lifecycle timing
- fail-closed

上記と強く結合しているため未着手のまま残す。

### snapshot chunk / metadata broader 整理

A4 は publication side effects ownership に限定しているため、次は未着手のまま残す。

- snapshot chunk 分散
- metadata shape 整理
- snapshot serialization 境界

## この現在地の読み方

- A4 は publication side effects ownership の整理であり、transaction / rollback 化を完了したことは意味しない
- C1 / C2 は direct contract test の追加であり、integration test 化を意味しない
- B2-1 完了に `CasePaneCacheRefreshNotificationService` は含めない
- TaskPane lifecycle 全体を整理済みと読むことはできない
- 未着手項目を「対応済み」へ読み替えない

## 全体評価

現在の整理は、lifecycle、visibility、publication side effects、notification、eligibility、resolver、wiring を安全単位で分離してきた状態です。

重要なのは、巨大クラス分割そのものよりも、ownership と変更理由を安全単位で分離して main に固定してきた点です。A4 / C2 完了時点でも、TaskPane lifecycle 本線、visible window resolve、HOME lifetime、post-close follow-up、snapshot / metadata の広域整理は、まだ別単位で慎重に扱うべき積み残しとして残ります。
