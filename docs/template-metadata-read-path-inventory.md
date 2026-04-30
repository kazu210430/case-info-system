# Template Metadata Read Path Inventory

## 1. 目的

この文書は、文書ボタン表示・文書名入力・文書生成で使う `caption / key / file / metadata` の読取経路を、read-only adapter 化や責務整理の前提として棚卸しするための調査メモです。

- 参照前提:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
- 関連資料:
  - `docs/template-metadata-inventory.md`
  - `docs/readonly-api-adoption-status.md`
- 今回の範囲:
  - 調査と docs 整理のみ
  - コード修正なし

## 2. 先に結論

### 2.1 現時点で正本と呼べる情報源

| 項目 | 現時点の正本 | 備考 |
| --- | --- | --- |
| `key` | Kernel `雛形一覧` A列 | 元の起点は雛形ファイル名先頭の `NN_` だが、runtime 参照の正本は同期後の一覧 |
| `caption` | Kernel `雛形一覧` C列 | 元の起点は雛形ファイル名から抽出した `DisplayName` |
| `TemplateFileName` | Kernel `雛形一覧` B列 | 実行時の `file` 正本 |
| `group` / タブ名 | Kernel `雛形一覧` E列 | 空欄は snapshot 構築時に `その他` へ正規化される |
| 文書ボタン色 | Kernel `雛形一覧` D列の塗り色 | snapshot では `DOC.FillColor` へ写る |
| タブ色 | Kernel `雛形一覧` F列の塗り色 | snapshot では `TAB.BackColor` へ写る |
| global master version | Kernel `TASKPANE_MASTER_VERSION` | Base / CASE 側は mirror / provenance |
| `TemplatePath` | 専用の保存正本なし | `DocumentTemplateResolver` が `WORD_TEMPLATE_DIR` または `SYSTEM_ROOT\雛形` から都度導出する |

### 2.2 現時点で補助キャッシュと呼べる情報源

| 情報 | 位置づけ | 保存場所 |
| --- | --- | --- |
| Base 埋込 snapshot | 新規 CASE 初期表示用の派生 cache | `TASKPANE_BASE_SNAPSHOT_*` / `TASKPANE_BASE_MASTER_VERSION` |
| CASE snapshot cache | 表示中 CASE と整合する派生 cache | `TASKPANE_SNAPSHOT_CACHE_*` / CASE 側 `TASKPANE_MASTER_VERSION` |
| `MasterTemplateCatalogService` の一覧 cache | master sheet 読取結果のメモリ cache | Add-in プロセス内 |
| `DocumentExecutionEligibilityService` の eligibility cache | 実行前判定結果のメモリ cache | Add-in プロセス内 |

### 2.3 責務境界

- `DocumentNamePromptService`
  - CASE cache から `caption` を引けた場合だけ prompt 初期値へ使う
  - master catalog には fallback しない
  - `TASKPANE_DOC_NAME_OVERRIDE_*` へ一時保存する入口
- `DocumentTemplateResolver`
  - CASE cache 優先 + master catalog fallback で `key -> DocumentName / TemplateFileName` を解決する
  - `TemplatePath` を都度導出する
  - 実行側の正本確認入口

## 3. サービス別インベントリ

| サービス | 取得している情報 | 情報源 | 区分 | 呼び出し元 | 呼び出し先 | 変更リスク | 今後の整理候補 | 今は触らない理由 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `MasterTemplateSheetReader` | `key` `TemplateFileName` `caption` `TabName` `FillColor` `TabBackColor` | `雛形一覧` A:E 値と D/F 塗り色、3行目以降 | 正本 reader | `MasterTemplateCatalogService` `TaskPaneSnapshotBuilderService`、一部 `KernelTemplateSyncService` | Excel Worksheet | 列意味解釈がここに依存する | read-only reader の共通口として維持 | 既存列構成と色読取を崩すと display / resolver 両方に波及する |
| `MasterTemplateSheetReaderAdapter` | `MasterTemplateSheetReader` の read-only adapter | `MasterTemplateSheetReader` | adapter | `AddInCompositionRoot` から注入 | `MasterTemplateSheetReader.Read` | 低 | 既存 direct call の置換先 | 既に low-risk な adapter なので役割変更不要 |
| `MasterTemplateCatalogService` | `key` `DocumentName` `TemplateFileName` `BackColor` | `SYSTEM_ROOT` 解決後の Master `雛形一覧` | 正本 reader + メモリ cache | `DocumentTemplateLookupService` `KernelTemplateSyncService.InvalidateCache()` | `IMasterTemplateSheetReader` | 中 | `key -> template metadata` の共通 projection へ寄せる | cache invalidation と Master open/read-only 制御が一体 |
| `TaskPaneSnapshotBuilderService` | 文書ボタン表示用 snapshot、`TAB/DOC` 定義、master version | CASE cache / Base cache / Master `雛形一覧` | 派生 snapshot builder | `TaskPaneManager` | `IMasterTemplateSheetReader`、Workbook package / read-only Master open | 高 | snapshot build と storage read の分離 | 表示フロー、stale 判定、CASE cache 更新が密結合 |
| `TaskPaneSnapshotCacheService` | `key -> DocumentName / TemplateFileName` lookup、Base -> CASE promote | CASE `TASKPANE_SNAPSHOT_CACHE_*`、Base `TASKPANE_BASE_*` | 派生 cache reader | `DocumentTemplateLookupService`、CASE cache cleanup | `TaskPaneSnapshotParser` | 中 | Base/CASE snapshot read helper の共通化 | promote / clear / compatibility 判定を持つ |
| `DocumentTemplateLookupService` | CASE cache-only lookup と master fallback lookup | `TaskPaneSnapshotCacheService`、`MasterTemplateCatalogService` | 読取調停 | `DocumentTemplateResolver` `DocumentNamePromptService` | 同左 | 低 | read-only lookup 契約の中心に寄せる | 現行の fallback policy を変えないことが重要 |
| `DocumentTemplateResolver` | `DocumentName` `TemplateFileName` `TemplatePath` `ResolutionSource` | CASE cache、Master catalog、`WORD_TEMPLATE_DIR`、`SYSTEM_ROOT` | 実行側 reader | `DocumentExecutionEligibilityService` | `IDocumentTemplateLookupReader` | 高 | path 導出を残したまま lookup 部だけ共通化 | `TemplatePath` と実ファイル確認の責務を壊せない |
| `DocumentNamePromptService` | prompt 初期値、override 一時保持 | CASE cache、Excel active/visible window | 補助 UI reader | `TaskPaneManager` | `ICaseCacheDocumentTemplateReader` `DocumentNameOverrideScope` | 高 | prompt 初期値 lookup の read-only 依存を維持 | cache-only policy を変えると表示中 Pane とズレる |
| `TaskPaneManager` | 表示用 snapshot、選択タブ、押下 action | `ICaseTaskPaneSnapshotReader` の結果 | 派生 UI | `TaskPaneRefreshOrchestrationService` など | `TaskPaneSnapshotParser` `CaseTaskPaneViewStateBuilder` `DocumentNamePromptService` `DocumentCommandService` | 高 | reader 差し替えのみ | UI 制御責務を docs/ui-policy.md 前提で維持する必要がある |
| `KernelTemplateSyncService` | 雛形登録結果、`雛形一覧` A:C、Base 用 snapshot | 雛形フォルダ、`CaseList_FieldInventory`、Kernel `雛形一覧` | 正本 writer | `KernelCommandService` | `MasterTemplateSheetReader`、`MasterTemplateCatalogService.InvalidateCache()` | 高 | reader 部だけ adapter へ寄せる余地あり | 書込責務と version 更新の正本なので read-only 化対象ではない |
| `AddInCompositionRoot` | reader 境界の配線 | なし | composition | Add-in 起動 | `DocumentTemplateLookupService` `TaskPaneSnapshotBuilderService` | 中 | 依存の向きの固定点として維持 | constructor / DI 変更は今回禁止範囲 |

## 4. 読み取り経路

### 4.1 文書ボタン表示時

1. `TaskPaneManager` が `ICaseTaskPaneSnapshotReader.BuildSnapshotText` を呼ぶ。
2. `TaskPaneSnapshotBuilderService` が次の順で snapshot を解決する。
   - CASE `TASKPANE_SNAPSHOT_CACHE_*`
   - Base `TASKPANE_BASE_SNAPSHOT_*`
   - Master `雛形一覧` 再構築
3. Master 再構築時、`MasterTemplateSheetReader` が `雛形一覧` を読む。
   - A列: `key`
   - B列: `TemplateFileName`
   - C列: `caption`
   - D列: 文書ボタン色
   - E列: タブ名
   - F列: タブ色
4. `TaskPaneSnapshotBuilderService` は snapshot `DOC` 行へ次を積む。
   - `Key`
   - `Caption`
   - `ActionKind=doc`
   - `TabName`
   - `RowIndex`
   - `FillColor`
   - `TemplateFileName`
5. `TaskPaneSnapshotParser` が snapshot を `TaskPaneSnapshot` に変換する。
6. `CaseTaskPaneViewStateBuilder` が UI 用に変換する。
   - 個別タブ内の並び順は `RowIndex`
   - `全て` タブだけは `Key` 数値順

補足:

- `group` に相当する値は `TabName`
- 個別タブの `order` は `雛形一覧` の行走査順から増分される `RowIndex`
- タブ自体の `order` は、最初に出現した順番

### 4.2 文書ボタン押下後の文書生成

1. `TaskPaneManager.ExecuteCaseAction`
2. `DocumentNamePromptService.TryPrepare`
3. `DocumentCommandService.Execute`
4. `DocumentExecutionEligibilityService.Evaluate`
5. `DocumentTemplateResolver.Resolve`
6. `DocumentCreateService.Execute`

### 4.3 `DocumentTemplateResolver` が参照する正本

- まず `DocumentTemplateLookupService.TryResolveWithMasterFallback`
  - `TaskPaneSnapshotCacheService.TryGetDocumentTemplateLookupFromCache`
  - miss 時だけ `MasterTemplateCatalogService.TryGetTemplateByKey`
- その後 `DocumentTemplateResolver` が `TemplatePath` を導出
  - `WORD_TEMPLATE_DIR` 優先
  - 未設定なら `SYSTEM_ROOT\雛形`

つまり、`DocumentTemplateResolver` にとっての metadata 正本は snapshot そのものではなく、`CASE cache` を優先参照したうえで足りない時に読む `MasterTemplateCatalogService` 側です。`TemplatePath` は保存正本を持たず resolver で毎回計算します。

### 4.4 `DocumentNamePromptService` が使う情報

- 入力:
  - 対象 CASE workbook
  - 押下された文書 `key`
- 表示:
  - `ICaseCacheDocumentTemplateReader.TryResolveFromCaseCache`
  - 返った `DocumentTemplateLookupResult.DocumentName` を prompt 初期値に使用
- 解決しないもの:
  - `TemplateFileName`
  - `TemplatePath`
  - master fallback
- 出力:
  - `DocumentNameOverrideScope`
  - `TASKPANE_DOC_NAME_OVERRIDE_ENABLED`
  - `TASKPANE_DOC_NAME_OVERRIDE`

#### 4.4.1 `DocumentNamePromptService` lookup inventory

| 項目 | 現在の事実 |
| --- | --- |
| サービス名 | `DocumentNamePromptService` |
| 現在の責務 | 文書名入力 prompt を開く前に、CASE cache から初期値候補を引き、確定値を `DocumentNameOverrideScope` に渡す補助 UI サービス |
| 入力 | `Excel.Workbook`、押下された文書 `key` |
| 出力 | `bool`、`DocumentNameOverrideScope`、一時 DocProperty `TASKPANE_DOC_NAME_OVERRIDE_ENABLED` / `TASKPANE_DOC_NAME_OVERRIDE` |
| 直接依存 | `ExcelInteropService`、`ICaseCacheDocumentTemplateReader`、`Logger` |
| 参照 metadata | lookup 入力として `key`、lookup 成功時の `DocumentTemplateLookupResult.DocumentName` |
| 間接的に参照成立に効く metadata | `TaskPaneSnapshotCacheService` 側では `TaskPaneDocDefinition.TemplateFileName` が空だと lookup 不成立になるため、prompt 側は `TemplateFileName` を直接使わないが、`file` 情報の有無に間接依存する |
| 参照しない情報 | `TemplatePath`、master catalog、実体テンプレートファイル存在、実行可否 |
| 情報源 | 第一経路は CASE `TASKPANE_SNAPSHOT_CACHE_*`。CASE cache 空または古い場合は `TaskPaneSnapshotCacheService.PromoteBaseSnapshotToCaseCacheIfNeeded` により Base `TASKPANE_BASE_*` が CASE cache へ昇格した後、その CASE cache を読む |
| lookup service 使用状況 | `DocumentNamePromptService` 自身は `ICaseCacheDocumentTemplateReader` 経由。実体は `DocumentTemplateLookupService.TryResolveFromCaseCache` が `TaskPaneSnapshotCacheService.TryGetDocumentTemplateLookupFromCache` へ委譲する |
| cache-only policy の実装上の意味 | `TryResolveFromCaseCache` が失敗した時点で空文字を返し、prompt 初期値を空欄のまま開く。`MasterTemplateCatalogService` への fallback 呼び出しは行わない |
| master fallback を追加しない理由 | `docs/flows.md` が、文書名入力 UI は表示中 Pane と整合する CASE cache 表示状態に従い、master fallback は `DocumentTemplateResolver` 側の実行時解決責務と定義しているため |
| `DocumentTemplateResolver` との違い | `DocumentNamePromptService` は prompt 初期値だけを扱う補助 UI。`DocumentTemplateResolver` は `IDocumentTemplateLookupReader` 経由で CASE cache 優先・master fallback ありの metadata 解決を行い、さらに `TemplatePath` を導出する実行側サービス |
| `TaskPaneSnapshotCacheService` との関係 | prompt 側の cache-only lookup は最終的に `TaskPaneSnapshotCacheService` が返す `DocumentTemplateLookupResult` に依存する。Base promote、snapshot compatibility 判定、CASE cache clear の影響を受ける |
| 既存テスト | `DocumentTemplateLookupServiceTests` が、CASE cache hit 時の prompt 初期値反映、CASE cache miss 時の prompt 空欄維持、resolver 側 master fallback との責務分離、`ICaseCacheDocumentTemplateReader` の no-fallback を担保している |
| 今後の整理余地 | 既に consumer 依存は `ICaseCacheDocumentTemplateReader` に分離済み。今後整理するなら、`DocumentNamePromptService` の constructor 契約を変えず、cache-only lookup 実装の内部委譲や test coverage 拡張を小単位で進める余地がある |
| 変更リスク | prompt 初期値の参照元を master 側へ広げると、表示中 Pane と prompt のズレ、開いている CASE が後から登録された雛形へ勝手に追随する挙動変化、`docs/flows.md` と矛盾するリスクがある |
| 今は触らない理由 | cache-only policy と prompt 挙動は docs とテストで固定点があり、今回の目的は調査と記録のみであるため |

### 4.5 `TaskPaneSnapshotCacheService` / `TaskPaneSnapshotBuilderService`

#### `TaskPaneSnapshotCacheService`

- 保持:
  - CASE `TASKPANE_SNAPSHOT_CACHE_*`
  - Base `TASKPANE_BASE_SNAPSHOT_*`
- 読取:
  - 必要なら Base snapshot を CASE cache へ promote
  - snapshot を parse して `TaskPaneDocDefinition` から `DocumentName` と `TemplateFileName` を返す
- 位置づけ:
  - 表示整合のための補助 cache
  - latest master version との照合はここでは行わない

#### `TaskPaneSnapshotBuilderService`

- 読取:
  - CASE / Base DocProperty cache
  - Master `TASKPANE_MASTER_VERSION`
  - Master `雛形一覧`
- 保持:
  - Master 再構築または Base fallback の結果を CASE cache に保存
- 位置づけ:
  - 表示用 snapshot builder
  - 保存・生成・実行判断の正本ではない

### 4.6 `MasterTemplateCatalogService` / `MasterTemplateSheetReaderAdapter` との関係

- `MasterTemplateSheetReaderAdapter`
  - `IMasterTemplateSheetReader` として `MasterTemplateSheetReader.Read` を包む
- `MasterTemplateCatalogService`
  - CASE workbook の `SYSTEM_ROOT` を手掛かりに Master を read-only で開く
  - `IMasterTemplateSheetReader` を使って `雛形一覧` を `MasterTemplateRecord` へ変換する
  - `DocumentTemplateLookupService` から master fallback 先として使われる
- `TaskPaneSnapshotBuilderService`
  - 同じ `IMasterTemplateSheetReader` で表示用 snapshot を再構築する

## 5. 現時点の注意点

### 5.1 表示 metadata と実行 metadata は完全同一ではない

`TaskPaneSnapshotBuilderService` は `key` と `caption` があれば `DOC` 行を生成します。`TemplateFileName` が空でも表示上の文書ボタンは作られます。一方で:

- `TaskPaneSnapshotCacheService.TryGetDocumentTemplateLookupFromCache` は `TemplateFileName` 空を不成立扱いする
- `MasterTemplateCatalogService` も `key` または `TemplateFileName` が空の行を無視する

したがって、表示できるが実行解決できない行があり得ます。これは現状の事実であり、read-only adapter 化でも安易に均してはいけません。

### 5.2 snapshot cache の位置づけ

- CASE snapshot cache は表示中 CASE に追随する補助 cache
- Base snapshot は新規 CASE 初期状態の配布用 cache
- どちらも正本ではない
- `DocumentNamePromptService` はこの補助 cache の表示整合を使う
- `DocumentTemplateResolver` は補助 cache を優先しつつ、実行時だけ master へ fallback する

### 5.3 master sheet 読み取りの位置づけ

- `MasterTemplateSheetReader` は `雛形一覧` の列意味を解釈する共通 reader
- ただし、現時点では `KernelTemplateSyncService.BuildTaskPaneSnapshot` が static reader を直接呼んでおり、すべてが adapter 経由に統一されているわけではない

## 6. 今後の安全な整理順

1. `MasterTemplateSheetReader` 系
   - `雛形一覧` 列解釈の read-only 入口をさらに統一する
   - 特に `KernelTemplateSyncService` 側の direct call を adapter 寄せ候補として整理する
2. `DocumentTemplateLookupService`
   - `key -> DocumentName / TemplateFileName / ResolutionSource` の read-only 窓口を固定する
   - prompt cache-only と resolver master fallback の両 policy は保持する
   - `DocumentNamePromptService` 側はすでに `ICaseCacheDocumentTemplateReader` 依存なので、将来差し替える場合も consumer 契約は固定したまま内部委譲だけを動かすのが最小単位候補
3. Base / CASE snapshot storage の read helper
   - `TaskPaneSnapshotCacheService` と `CaseTemplateSnapshotService` の読取重複を先に整理する
4. その後で限定的な consumer 差し替え
   - `TaskPaneManager`
   - `DocumentNamePromptService`
   - `DocumentTemplateResolver`

## 7. 今は触らない方がよい箇所

| 箇所 | 理由 |
| --- | --- |
| `TaskPaneSnapshotBuilderService.BuildSnapshotText` | CASE cache / Base cache / Master rebuild、stale 判定、CASE cache 更新が一体 |
| `DocumentNamePromptService` の cache-only policy | 表示中 Pane と prompt 初期値の整合を保つ前提 |
| `DocumentTemplateResolver` の CASE cache 優先 | 開いている CASE の表示状態と実行解決元を揃える前提 |
| `KernelTemplateSyncService` の `TASKPANE_MASTER_VERSION` 更新と Base 書込 | 正本更新の責務そのもの |
| `WorkbookActivate` / `WindowActivate` 前提の Pane 再利用設計 | `docs/flows.md` と `docs/ui-policy.md` が維持対象としている |

## 8. 不明点

- `metadata` という語はコード上の正式型名ではなく、本書での便宜上の総称
- snapshot `META` 行に含まれる master version 自体は parser の主経路で保持されていないため、その保持理由の全量はこの範囲では確定しない
- `MasterTemplateCatalogService` のメモリ cache を複数 master root がどう使い分ける想定かは、この調査範囲では断定しない
