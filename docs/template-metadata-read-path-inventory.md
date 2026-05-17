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
| `TaskPaneSnapshotChunkReadHelper` | snapshot chunk の count 読取と連結 | `TASKPANE_*_COUNT`、`TASKPANE_*_XX` DocProperty | shared primitive reader | `TaskPaneSnapshotCacheService` `CaseTemplateSnapshotService` `TaskPaneSnapshotBuilderService` | `ExcelInteropService.TryGetDocumentProperty` | 低 | Base/CASE 共通の raw chunk read 契約として維持 | ここへ promote / stale / UI policy を混ぜると helper 境界が崩れる |
| `TaskPaneSnapshotChunkStorageHelper` | snapshot chunk の count 更新、分割保存、末尾 chunk 空文字化、clear | `TASKPANE_*_COUNT`、`TASKPANE_*_XX` DocProperty | shared primitive writer | `TaskPaneSnapshotCacheService` `CaseTemplateSnapshotService` `TaskPaneSnapshotBuilderService` | `ExcelInteropService.SetDocumentProperty` | 低 | Base/CASE 共通の raw chunk write 契約として維持 | delete / promote / compatibility policy を持たせない前提で再利用されている |
| `TaskPaneSnapshotBuilderService` | 文書ボタン表示用 snapshot、`TAB/DOC` 定義、master version | CASE cache / Base cache / Master `雛形一覧` | 派生 snapshot builder | `TaskPaneManager` | `IMasterTemplateSheetReader`、Workbook package / read-only Master open | 高 | snapshot build と storage read の分離 | 表示フロー、stale 判定、CASE cache 更新が密結合 |
| `TaskPaneSnapshotCacheService` | `key -> DocumentName / TemplateFileName` lookup、Base -> CASE promote | CASE `TASKPANE_SNAPSHOT_CACHE_*`、Base `TASKPANE_BASE_*` | 派生 cache reader | `DocumentTemplateLookupService`、CASE cache cleanup | `TaskPaneSnapshotParser` | 中 | Base/CASE snapshot read helper の共通化 | promote / clear / compatibility 判定を持つ |
| `CaseTemplateSnapshotService` | 新規 CASE 初期化時の CASE version 同期、Base snapshot の CASE cache 昇格 | Kernel `TASKPANE_MASTER_VERSION`、CASE/Base snapshot DocProperty | initializer 専用 promote service | `CaseWorkbookInitializer` | `TaskPaneSnapshotChunkReadHelper` `TaskPaneSnapshotChunkStorageHelper` | 中 | init 専用 promote と lookup promote の差分を明示したまま重複整理 | 新規 CASE 初期化で既存 CASE cache を上書きする初期化責務を持ち、lookup 時 promote と同一化できない |
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
  - `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup`
  - miss 時だけ `MasterTemplateCatalogService.TryGetTemplateByKey`
- その後 `DocumentTemplateResolver` が `TemplatePath` を導出
  - `WORD_TEMPLATE_DIR` 優先
  - 未設定なら `SYSTEM_ROOT\雛形`

つまり、`DocumentTemplateResolver` にとっての metadata 正本は snapshot そのものではなく、`CASE cache` を優先参照したうえで足りない時に読む `MasterTemplateCatalogService` 側です。`TemplatePath` は保存正本を持たず resolver で毎回計算します。

#### 4.3.1 `DocumentTemplateResolver` lookup inventory

| 項目 | 現在の事実 |
| --- | --- |
| サービス名 | `DocumentTemplateResolver` |
| 現在の責務 | 文書 `key` を正規化し、`IDocumentTemplateLookupReader.TryResolveWithMasterFallback` で `DocumentName` / `TemplateFileName` / `ResolutionSource` を解決し、`WORD_TEMPLATE_DIR` または `SYSTEM_ROOT\雛形` から `TemplatePath` を導出して `DocumentTemplateSpec` を返す。未解決時は `null` を返す |
| 入力 | `Excel.Workbook`、文書 `key` |
| 出力 | `DocumentTemplateSpec` または `null`。補助 API として `TemplateExists`、`IsSupportedWordTemplate` を持ち、`DocumentExecutionEligibilityService` が後段判定に使う |
| 直接依存 | `ExcelInteropService`、`PathCompatibilityService`、`IDocumentTemplateLookupReader`、`Logger` |
| 参照 metadata | `DocumentTemplateLookupResult.Key` / `DocumentName` / `TemplateFileName` / `ResolutionSource`、CASE DocProperty `WORD_TEMPLATE_DIR` / `SYSTEM_ROOT` |
| 情報源 | 1) `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup`、2) miss 時だけ `MasterTemplateCatalogService.TryGetTemplateByKey`、3) path 部分は resolver 自身が DocProperty から導出 |
| lookup service 使用状況 | `AddInCompositionRoot` が `DocumentTemplateLookupService` を `IDocumentTemplateLookupReader` として注入し、現在の consumer は `DocumentExecutionEligibilityService -> DocumentTemplateResolver` |
| CASE cache 優先の実装上の意味 | `TryResolveWithMasterFallback` は先に CASE cache-only lookup を実行し、成功したらその `DocumentName` / `TemplateFileName` をそのまま採用する。resolver 自身は global master version の新旧比較を行わないため、開いている CASE の表示中 Pane と同じ cache 系 metadata を実行側でも使う |
| master fallback の実装上の意味 | fallback は CASE cache lookup が `false` を返した時だけ発火する。対象には cache 空、snapshot 互換性不一致による clear 後、key 不一致、`TemplateFileName` 空で lookup 不成立になったケースが含まれる |
| `TemplatePath` 導出責務 | 保存済み正本はなく、resolver が `WORD_TEMPLATE_DIR` 優先、未設定時は `SYSTEM_ROOT\雛形` を組み立てて都度決める。どちらも取れない場合は空文字 |
| `DocumentNamePromptService` との違い | prompt 側は CASE cache caption の補助 UI。resolver 側は master fallback と path 導出まで含む実行用解決 |
| `TaskPaneSnapshotCacheService` との関係 | cache lookup の成否は `TaskPaneSnapshotCacheService` に依存し、Base -> CASE promote、snapshot compatibility clear、`TemplateFileName` 空時の不成立判定の影響を受ける |
| `MasterTemplateCatalogService` との関係 | fallback 時だけ master reader を使う。master 側は `SYSTEM_ROOT` を手掛かりに Kernel workbook を read-only で開き、`雛形一覧` の `key` と `TemplateFileName` が揃う行だけ `MasterTemplateRecord` に載せる |
| 既存テスト | `DocumentTemplateLookupServiceTests` が CASE cache hit、CASE cache miss + master fallback、key 不在、cache-only reader no-fallback、`WORD_TEMPLATE_DIR` 未設定時の `SYSTEM_ROOT\雛形` path 導出を確認している |
| 今後の整理余地 | 現在の consumer 契約は `IDocumentTemplateLookupReader` で固定されているため、将来の整理は constructor や interface を変えずに `DocumentTemplateLookupService` 内部委譲を差し替えるのが最小境界 |
| 変更リスク | CASE cache 優先や path 導出責務を resolver から外すと、`DocumentExecutionEligibilityService` の実行判定、表示中 Pane との整合、`docs/flows.md` の責務分離に波及する |
| 今は触らない理由 | 実装・docs・既存テストが CASE cache 優先 + master fallback + resolver の path 導出を前提に揃っているため |

#### 4.3.2 `DocumentTemplateLookupService` inventory

| 項目 | 現在の事実 |
| --- | --- |
| サービス名 | `DocumentTemplateLookupService` |
| 現在の責務 | `ICaseCacheDocumentTemplateReader` と `IDocumentTemplateLookupReader` の両方を実装し、CASE cache-only lookup と master fallback lookup の方針差を 1 箇所で調停する |
| 入力 | `Excel.Workbook`、文書 `key` |
| 出力 | `DocumentTemplateLookupResult` または `false` |
| 直接依存 | `TaskPaneSnapshotCacheService`、`MasterTemplateCatalogService` |
| CASE cache-only 経路 | `TryEnsurePromotedCaseCacheThenResolve` は `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` へ委譲する。名前どおり promotion-aware であり、pure read ではない |
| master fallback 経路 | `TryResolveWithMasterFallback` は CASE cache hit なら即 return し、miss 時だけ `MasterTemplateCatalogService.TryGetTemplateByKey` を呼ぶ。fallback 結果も `TemplateFileName` が空なら失敗扱いにする |
| lookup service 使用状況 | 既に `DocumentNamePromptService` は `ICaseCacheDocumentTemplateReader` 経由、`DocumentTemplateResolver` は `IDocumentTemplateLookupReader` 経由で利用している。今回の調査範囲では、この 2 箇所が `DocumentTemplateLookupService` 経由化済みの consumer |
| CASE cache 優先の意味 | prompt と resolver の双方が同じ CASE cache reader を共有しつつ、caller intent に応じて fallback 可否だけを切り分けられる |
| master fallback の意味 | fallback の責務を `DocumentNamePromptService` ではなく `DocumentTemplateResolver` 側へ限定するための policy 境界 |
| `TaskPaneSnapshotCacheService` / `MasterTemplateCatalogService` との境界 | cache の promote / clear / parse は `TaskPaneSnapshotCacheService`、master の open/read-only/cache は `MasterTemplateCatalogService`、caller 向けの fallback policy だけを `DocumentTemplateLookupService` が持つ |
| 今後の安全な最小単位 | 既存 consumer 契約を保ったまま、このサービス内部の委譲先や projection を整理するのが最も狭い変更面 |
| 変更リスク | `TryEnsurePromotedCaseCacheThenResolve` に fallback を混ぜる、または `TryResolveWithMasterFallback` の hit/miss 条件を変えると、prompt と resolver の責務分離が崩れる |
| 今は触らない理由 | `AddInCompositionRoot` の interface 配線、`DocumentNamePromptService` の cache-only policy、`DocumentTemplateResolver` の fallback policy がここを前提に成立しているため |

補足:

- この調査範囲で確認できた既存テストは `DocumentTemplateLookupServiceTests` に集中しており、prompt/resolver の責務分離と `TemplatePath` 導出は担保されている。
- 一方で、`TaskPaneSnapshotCacheService.PromoteBaseSnapshotToCaseCacheIfNeeded` の昇格条件と snapshot 互換性不一致 clear については、この調査範囲では専用テストを確認できていない。

### 4.4 `DocumentNamePromptService` が使う情報

- 入力:
  - 対象 CASE workbook
  - 押下された文書 `key`
- 表示:
  - `ICaseCacheDocumentTemplateReader.TryEnsurePromotedCaseCacheThenResolve`
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
| lookup service 使用状況 | `DocumentNamePromptService` 自身は `ICaseCacheDocumentTemplateReader` 経由。実体は `DocumentTemplateLookupService.TryEnsurePromotedCaseCacheThenResolve` が `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` へ委譲する |
| cache-only policy の実装上の意味 | `TryEnsurePromotedCaseCacheThenResolve` が失敗した時点で空文字を返し、prompt 初期値を空欄のまま開く。`MasterTemplateCatalogService` への fallback 呼び出しは行わない。ただし lookup 前に Base snapshot promotion と CASE DocProperty 更新が起きうる |
| master fallback を追加しない理由 | `docs/flows.md` が、文書名入力 UI は表示中 Pane と整合する CASE cache 表示状態に従い、master fallback は `DocumentTemplateResolver` 側の実行時解決責務と定義しているため |
| `DocumentTemplateResolver` との違い | `DocumentNamePromptService` は prompt 初期値だけを扱う補助 UI。`DocumentTemplateResolver` は `IDocumentTemplateLookupReader` 経由で CASE cache 優先・master fallback ありの metadata 解決を行い、さらに `TemplatePath` を導出する実行側サービス |
| `TaskPaneSnapshotCacheService` との関係 | prompt 側の cache-only lookup は最終的に `TaskPaneSnapshotCacheService` が返す `DocumentTemplateLookupResult` に依存する。Base promote、snapshot compatibility 判定、CASE cache clear の影響を受ける |
| 既存テスト | `DocumentTemplateLookupServiceTests` が、CASE cache hit 時の prompt 初期値反映、CASE cache miss 時の prompt 空欄維持、resolver 側 master fallback との責務分離、`ICaseCacheDocumentTemplateReader` の no-fallback を担保している |
| 今後の整理余地 | 既に consumer 依存は `ICaseCacheDocumentTemplateReader` に分離済み。今後整理するなら、`DocumentNamePromptService` の constructor 契約を変えず、cache-only lookup 実装の内部委譲や test coverage 拡張を小単位で進める余地がある |
| 変更リスク | prompt 初期値の参照元を master 側へ広げると、表示中 Pane と prompt のズレ、開いている CASE が後から登録された雛形へ勝手に追随する挙動変化、`docs/flows.md` と矛盾するリスクがある |
| 今は触らない理由 | cache-only policy と prompt 挙動は docs とテストで固定点があり、今回の目的は調査と記録のみであるため |

### 4.5 `TaskPaneSnapshotCacheService` / `TaskPaneSnapshotBuilderService`

#### 4.5.1 Base snapshot / CASE snapshot cache storage inventory

| 項目 | Base snapshot | CASE snapshot cache |
| --- | --- | --- |
| 意味 | Base ブックに埋め込まれる配布用 snapshot。新規 CASE 作成時と CASE cache 欠損時の初期ソース | 開いている CASE で表示・prompt・resolver が優先参照する作業用 cache |
| 保存先 DocProperty | `TASKPANE_BASE_SNAPSHOT_COUNT`、`TASKPANE_BASE_SNAPSHOT_01..NN`、`TASKPANE_BASE_MASTER_VERSION` | `TASKPANE_SNAPSHOT_CACHE_COUNT`、`TASKPANE_SNAPSHOT_CACHE_01..NN`、CASE 側 `TASKPANE_MASTER_VERSION` |
| 主な writer | `KernelTemplateSyncService.SaveSnapshotToBaseWorkbook` | `CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache`、`TaskPaneSnapshotCacheService.SaveTaskPaneSnapshotCache`、`TaskPaneSnapshotBuilderService.SaveCaseSnapshotCache` |
| 主な reader | `TaskPaneSnapshotCacheService`、`TaskPaneSnapshotBuilderService`、`CaseTemplateSnapshotService` | `TaskPaneSnapshotCacheService`、`TaskPaneSnapshotBuilderService`、その先の `DocumentTemplateLookupService` / `DocumentNamePromptService` / `DocumentTemplateResolver` |
| 用途 | 新規 CASE 初期状態の配布、CASE cache 再生元 | 表示中 Pane の文書定義、prompt 初期値、resolver の CASE cache 優先 lookup |
| 正本性 | 正本ではない | 正本ではない |
| global master との関係 | `TASKPANE_BASE_MASTER_VERSION` に「この Base snapshot を作ったときの master version」を保持 | CASE 側 `TASKPANE_MASTER_VERSION` に「この CASE cache / CASE 表示系が最後に採用した master version」を保持 |

補足:

- Base snapshot と CASE snapshot cache はどちらも `TaskPaneSnapshotFormat.ExportVersion=2` の snapshot text を 240 文字単位で chunk 保存します。
- `TaskPaneSnapshotParser` は `META` 行を読めますが、parser 結果として master version を保持していません。master version は DocProperty 側で判定されます。

#### 4.5.2 `TASKPANE_MASTER_VERSION` 系 DocProperty の役割

| DocProperty | 配置先 | 現在の役割 |
| --- | --- | --- |
| `TASKPANE_MASTER_VERSION` | Kernel | `KernelTemplateSyncService.IncrementTaskPaneMasterVersion` が更新する global master version の正本 |
| `TASKPANE_MASTER_VERSION` | Base | Base 保存時に mirror される値。新規 CASE 作成時にコピーされうる |
| `TASKPANE_MASTER_VERSION` | CASE | CASE cache / 表示系が最後に採用した master version。`TaskPaneSnapshotBuilderService` の stale 判定と `TaskPaneSnapshotCacheService` の promote 判定に使われる |
| `TASKPANE_BASE_MASTER_VERSION` | Base と CASE 内の埋込 Base snapshot 領域 | Base snapshot 自体の provenance。Base snapshot を CASE cache へ promote / fallback するときの比較値 |

補足:

- `CaseWorkbookInitializer` はいったん Kernel の `TASKPANE_MASTER_VERSION` を CASE に写しますが、その直後に `CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache` が `TASKPANE_BASE_MASTER_VERSION` を CASE 側 `TASKPANE_MASTER_VERSION` へ上書きします。したがって新規 CASE の実効 version は、最終的に埋込 Base snapshot 側へ揃います。

#### 4.5.3 snapshot chunk helper の責務境界

| helper | 現在の責務 | 持たない責務 |
| --- | --- | --- |
| `TaskPaneSnapshotChunkReadHelper` | count DocProperty を正数として読み、`TASKPANE_*_XX` を `01..NN` 順で連結して raw snapshot text を返す。引数不足、count 不正、count<=0 は空文字を返す | promote 判断、compatibility 判定、stale 判定、Master rebuild、`TaskPaneSnapshotParser` 呼び出し、UI 制御 |
| `TaskPaneSnapshotChunkStorageHelper.SaveSnapshot` | snapshot text を既定 240 文字で分割し、count 更新、使用中 chunk 書込、余剰旧 chunk の空文字化を行う | どの snapshot を保存するかの選択、master version 更新、property delete、UI 制御 |
| `TaskPaneSnapshotChunkStorageHelper.ClearSnapshot` | count を `0` にし、以前使われていた chunk prop を空文字化する | promote 判断、compatibility 判定、stale 判定、Master rebuild、property delete、UI 制御 |

補足:

- `TaskPaneSnapshotChunkStorageHelper.SaveSnapshot` は `snapshotText` が空のとき count を `0` にして終了します。既存 chunk prop の空文字化まで必要な経路は `ClearSnapshot` を使います。
- helper は raw chunk I/O の shared primitive であり、Base と CASE のどちらを扱うか、いつ clear/promote するかは caller 側サービスに残ります。

helper が持たない責務:

- promote 判断
- compatibility 判定
- stale 判定
- Master rebuild
- UI 制御

#### 4.5.4 `TaskPaneSnapshotCacheService` inventory

| 項目 | 現在の事実 |
| --- | --- |
| サービス名 | `TaskPaneSnapshotCacheService` |
| 現在の責務 | helper で CASE/Base chunk を読み書きしつつ、CASE cache lookup、on-demand Base -> CASE promote、compatibility clear、key 正規化、`DocumentTemplateLookupResult` 生成を担う |
| 入力 | `Excel.Workbook`、文書 `key` |
| 出力 | `DocumentTemplateLookupResult` または `false`。補助 API として `TryEnsurePromotedCaseCacheThenGetDocInfo`、`ClearCaseSnapshotCacheChunks` を持つ |
| 直接依存 | `ExcelInteropService`、`Logger` |
| 読取対象 | CASE `TASKPANE_SNAPSHOT_CACHE_COUNT` / `TASKPANE_SNAPSHOT_CACHE_XX`、Base `TASKPANE_BASE_SNAPSHOT_COUNT` / `TASKPANE_BASE_SNAPSHOT_XX`、CASE `TASKPANE_MASTER_VERSION`、`TASKPANE_BASE_MASTER_VERSION` |
| write path | `PromoteBaseSnapshotToCaseCacheIfNeeded` が CASE cache chunk と CASE `TASKPANE_MASTER_VERSION` を更新する。`ClearSnapshotParts` は count を `0` にし、既存 chunk prop へ空文字を書き戻す。`ClearCaseSnapshotCacheChunks` は `TASKPANE_SNAPSHOT_CACHE_XX` だけを Delete する |
| read path | `DocumentTemplateLookupService.TryEnsurePromotedCaseCacheThenResolve` から呼ばれ、`DocumentNamePromptService` と `DocumentTemplateResolver` の CASE cache lookup 入口になる。promotion-aware なので、この行でいう read path は pure read を意味しない |
| promote 条件 | `PromoteBaseSnapshotToCaseCacheIfNeeded` は、1) Base snapshot が存在し互換、かつ 2) CASE cache が空、または 3) `TASKPANE_BASE_MASTER_VERSION > TASKPANE_MASTER_VERSION`、または 4) CASE 側 version 未設定かつ Base 側 version が正値、のときに Base snapshot を CASE cache へ昇格する |
| compatibility / clear 条件 | CASE cache snapshot text が非互換なら CASE cache count を `0` にし chunk を空文字化。Base snapshot text が非互換なら Base snapshot count を `0` にし chunk を空文字化。lookup 対象 CASE cache が非互換なら lookup 前に CASE cache を clear して `false` を返す |
| しないこと | latest master version の読取や global stale 判定はしない。`TemplateFileName` が空の `DOC` 行を成功扱いしない |
| `PromoteBaseSnapshotToCaseCacheIfNeeded` の意味 | 表示中 CASE の lookup 系が最低限参照できる CASE cache を補充するための on-demand promote。global 最新かどうかではなく、CASE と埋込 Base の version 比較だけで動く |
| `DocumentTemplateLookupService` / `DocumentNamePromptService` / `DocumentTemplateResolver` との接続 | `DocumentTemplateLookupService` がこのサービスを CASE cache reader として包む。`DocumentNamePromptService` は cache-only、`DocumentTemplateResolver` は miss 時のみ master fallback へ進む |
| 今後の整理余地 | `LoadSnapshotParts` / `SaveTaskPaneSnapshotCache` / `ClearSnapshotParts` / promote 判定は `CaseTemplateSnapshotService` と重複しており、read helper 化の最小候補になる |
| 変更リスク | promote 条件を latest master 連動に変える、`TemplateFileName` 空行を成功扱いに変える、Base clear をやめる、`ClearCaseSnapshotCacheChunks` に count 変更を混ぜる、といった変更は prompt / resolver / case-list registration の前提を壊しやすい |
| 今は触らない理由 | lookup のたびに promote と compatibility clear が入るため、表示整合・prompt 初期値・resolver の CASE cache 優先がこの service の副作用を前提にしている |

#### 4.5.5 `TaskPaneSnapshotBuilderService` inventory

| 項目 | 現在の事実 |
| --- | --- |
| サービス名 | `TaskPaneSnapshotBuilderService` |
| 現在の責務 | CASE pane 表示用 snapshot の構築元を選び、必要なら CASE cache を更新して `TaskPaneBuildResult` を返す表示系 builder |
| 入力 | `Excel.Workbook` |
| 出力 | `TaskPaneBuildResult(SnapshotText, UpdatedCaseSnapshotCache)` |
| 直接依存 | `Excel.Application`、`ExcelInteropService`、`PathCompatibilityService`、`IMasterTemplateSheetReader`、`Logger` |
| 読取対象 | CASE `TASKPANE_SNAPSHOT_CACHE_*`、Base `TASKPANE_BASE_SNAPSHOT_*`、CASE `TASKPANE_MASTER_VERSION`、`TASKPANE_BASE_MASTER_VERSION`、Master `TASKPANE_MASTER_VERSION`、Master `雛形一覧` |
| write path | Base fallback または Master rebuild 時に CASE `TASKPANE_SNAPSHOT_CACHE_*` を更新する。Base fallback / Master rebuild 時は CASE `TASKPANE_MASTER_VERSION` も更新する。非互換 CASE/Base cache はそれぞれ count=`0` と chunk 空文字化で clear する |
| read path | `TaskPaneManager.RenderHost` と `RenderCaseHostAfterAction` が `ICaseTaskPaneSnapshotReader.BuildSnapshotText` 経由で使う |
| CASE cache 採用条件 | CASE cache が存在し互換で、さらに `TryReadLatestMasterVersion` が成功し、`latestMasterVersion <= CASE TASKPANE_MASTER_VERSION` のときにのみ source=`CaseCache` で採用する |
| Base snapshot 採用条件 | Base snapshot が存在し互換で、かつ 1) latest master version を読めない、または 2) `latestMasterVersion <= TASKPANE_BASE_MASTER_VERSION` のときに source=`BaseCacheFallback` または `BaseCache` として採用し、同時に CASE cache へ保存する |
| Master rebuild 条件 | CASE cache / Base snapshot のどちらも採用できない場合だけ Master を read-only で開き、`雛形一覧` から snapshot を再構築して CASE cache へ保存する |
| compatibility / clear 条件 | CASE cache text が `ExportVersion=2` 以外なら CASE cache を clear。Base snapshot text が非互換なら Base snapshot を clear。どちらも clear 後は次の候補へ進む |
| `TASKPANE_MASTER_VERSION` の役割 | CASE cache stale 判定では CASE 側 version、Base fallback では `TASKPANE_BASE_MASTER_VERSION`、Master rebuild では Master 側 version を CASE 側へ書き戻す |
| 表示専用の補正 | CASE cache 採用時は `ApplyDynamicSpecialButtonOverrides` で `CASELIST_REGISTERED` に応じた SPECIAL ボタン表示だけを動的補正して返すが、その補正結果を CASE cache へ保存し直しはしない |
| `TaskPaneManager` との関係 | `TaskPaneManager` はこの結果を parse して view state を作る。builder が CASE cache を更新したかどうかは `TaskPaneBuildResult.UpdatedCaseSnapshotCache` と `TaskPaneManager` の通知判定に使われる |
| `TaskPaneSnapshotCacheService` との境界 | builder は表示元選択と latest master 比較を持つ。cache service は lookup と on-demand promote だけを持つ。両者は同じ DocProperty 群を読むが stale 判定材料が異なる |
| 今後の整理余地 | CASE/Base chunk load/save/clear は helper 化余地があるが、latest master 比較と CASE cache 更新通知の責務は builder から外しにくい |
| 変更リスク | CASE cache 採用条件、Base fallback 条件、Master rebuild 条件、`UpdatedCaseSnapshotCache` の返し方を変えると `TaskPaneManager` の再利用・通知・表示更新ポリシーに波及する |
| 今は触らない理由 | display path、cache write、stale 判定、Master read-only open、通知フラグが 1 メソッドに密結合しているため |

#### 4.5.6 `CaseTemplateSnapshotService` inventory

| 項目 | 現在の事実 |
| --- | --- |
| サービス名 | `CaseTemplateSnapshotService` |
| 現在の責務 | 新規 CASE 初期化時に Kernel `TASKPANE_MASTER_VERSION` を CASE へ同期し、埋込 Base snapshot を CASE cache へ初期 promote する initializer 専用 service |
| 入力 | `Excel.Workbook kernelWorkbook`、`Excel.Workbook caseWorkbook` |
| 出力 | なし。副作用として CASE `TASKPANE_MASTER_VERSION`、`TASKPANE_SNAPSHOT_CACHE_*` を更新する |
| 直接依存 | `ExcelInteropService` |
| helper 利用 | Base snapshot 読取は `TaskPaneSnapshotChunkReadHelper`、Base/CASE clear は `TaskPaneSnapshotChunkStorageHelper.ClearSnapshot` を使う |
| 初期 promote の挙動 | `TASKPANE_BASE_SNAPSHOT_COUNT > 0` なら Base snapshot を CASE cache chunk へ物理コピーし、`TASKPANE_BASE_MASTER_VERSION` が正値なら CASE `TASKPANE_MASTER_VERSION` をその値へ揃える |
| compatibility / clear 条件 | Base snapshot が非互換なら Base snapshot と CASE cache の両方を clear する |
| `TaskPaneSnapshotCacheService` と異なる点 | lookup 時 promote と違い、既存 CASE cache がより新しくても初期化時 promote は埋込 Base snapshot で上書きする |
| しないこと | latest master version との stale 判定、Master rebuild、文書 `key` lookup、UI 制御 |
| 今後の整理余地 | raw chunk copy 自体は helper へ寄せられるが、新規 CASE 初期化専用の上書き policy は service 側へ残す必要がある |
| 変更リスク | init promote を lookup promote と同一化すると、新規 CASE 作成直後の CASE cache / version 初期化意味が変わる |
| 今は触らない理由 | `CaseWorkbookInitializer` 配下の初期化フローがこの「初回は Base を正として写す」前提で成立しているため |

#### 4.5.7 CASE cache / Base snapshot の write path 実態

| 起点 | サービス | 現在の write / clear |
| --- | --- | --- |
| 雛形登録・更新成功 | `KernelTemplateSyncService` | Kernel `TASKPANE_MASTER_VERSION` を `+1` し、Base へ `TASKPANE_BASE_SNAPSHOT_*`、`TASKPANE_BASE_MASTER_VERSION`、`TASKPANE_MASTER_VERSION` を書く |
| 新規 CASE 初期化 | `CaseWorkbookInitializer -> CaseTemplateSnapshotService` | Kernel の `TASKPANE_MASTER_VERSION` を CASE に写し、その後 Base snapshot を CASE cache へコピーし、`TASKPANE_BASE_MASTER_VERSION` があれば CASE `TASKPANE_MASTER_VERSION` をその値へ揃える |
| lookup 時 | `TaskPaneSnapshotCacheService` | Base snapshot が CASE より新しい、または CASE cache が空なら CASE cache へ promote する |
| CASE pane 表示時 | `TaskPaneSnapshotBuilderService` | Base fallback または Master rebuild の結果を CASE cache へ保存し、必要に応じて CASE `TASKPANE_MASTER_VERSION` を更新する |
| 案件一覧登録後 | `CaseListRegistrationService` | CASE `TASKPANE_SNAPSHOT_CACHE_COUNT=0` をセットし、`TaskPaneSnapshotCacheService.ClearCaseSnapshotCacheChunks` で `TASKPANE_SNAPSHOT_CACHE_XX` を削除する |

補足:

- CASE cache write path は 1 つではありません。initializer、lookup、display build の 3 経路があり、どこで最新化されたかで比較材料が少し異なります。
- `CaseTemplateSnapshotService` も `TaskPaneSnapshotCacheService` も Base -> CASE promote を持っており、今後 read helper 化するならここが最初の重複解消候補です。

#### 4.5.8 既存テストが担保していること / いないこと

- `SnapshotOutputRegressionTests`
  - Master rebuild で生成される snapshot text 形式
  - CASE cache への保存
  - Master rebuild 後に CASE `TASKPANE_MASTER_VERSION` が更新されること
  - `TemplateFileName` 空でも `DOC` 行自体は snapshot に残ること
- `DocumentTemplateLookupServiceTests`
  - CASE cache hit 時、resolver と prompt が同じ `DocumentName` / `TemplateFileName` 系 metadata を使うこと
  - CASE cache miss 時、resolver だけが master fallback すること
  - prompt 側は cache-only で、master fallback しないこと
  - `WORD_TEMPLATE_DIR` 未設定時に `SYSTEM_ROOT\雛形` へ path fallback すること
- `TaskPaneManager*Tests`
  - `UpdatedCaseSnapshotCache` を使った通知判断
  - `WorkbookActivate` / `WindowActivate` での host 再利用と再描画スキップ方針
- `TaskPaneSnapshotCacheStorageBehaviorTests`
  - `TaskPaneSnapshotCacheService.PromoteBaseSnapshotToCaseCacheIfNeeded` が、CASE cache 空、Base version の方が新しい、CASE version 欠損の各条件で promote すること
  - CASE cache が有効で Base version が新しくないときは promote しないこと
  - CASE cache 非互換時は CASE cache を clear すること
  - Base snapshot 非互換時は lookup promote 経路では Base snapshot だけを clear し、有効 CASE cache は残すこと
  - `CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache` は lookup promote と異なり、初期化時は既存 CASE cache を上書きしうること
  - `CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache` で Base snapshot が非互換なら Base と CASE の両 snapshot を clear すること

現時点で確認できていない専用テスト:

- `TaskPaneSnapshotChunkReadHelper` / `TaskPaneSnapshotChunkStorageHelper` 単体の direct unit test
- `TaskPaneSnapshotChunkStorageHelper.SaveSnapshot(string.Empty)` の count だけを `0` にする境界
- helper 自体が property delete を行わないことの専用固定点

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

- `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` は `TemplateFileName` 空を不成立扱いする
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

1. `TaskPaneSnapshotChunkReadHelper` / `TaskPaneSnapshotChunkStorageHelper`
   - raw chunk I/O 契約だけを対象にし、promote / stale / UI policy を持ち込まない
   - caller ごとの差分は維持したまま、DocProperty 連結・分割・clear の境界だけを固定する
2. `TaskPaneSnapshotCacheService` と `CaseTemplateSnapshotService`
   - helper を共有 primitive として使い続けつつ、Base -> CASE promote の重複を整理する
   - ただし lookup promote と initializer promote の policy 差は残す
3. `MasterTemplateSheetReader` 系
   - `雛形一覧` 列解釈の read-only 入口をさらに統一する
   - 特に `KernelTemplateSyncService` 側の direct call を adapter 寄せ候補として整理する
4. `DocumentTemplateLookupService`
   - `key -> DocumentName / TemplateFileName / ResolutionSource` の read-only 窓口を固定する
   - prompt cache-only と resolver master fallback の両 policy は保持する
   - `DocumentNamePromptService` 側はすでに `ICaseCacheDocumentTemplateReader` 依存なので、将来差し替える場合も consumer 契約は固定したまま内部委譲だけを動かすのが最小単位候補
5. その後で限定的な consumer 差し替え
   - `TaskPaneManager`
   - `DocumentNamePromptService`
   - `DocumentTemplateResolver`

## 7. 今は触らない方がよい箇所

| 箇所 | 理由 |
| --- | --- |
| `TaskPaneSnapshotBuilderService.BuildSnapshotText` | CASE cache / Base cache / Master rebuild、stale 判定、CASE cache 更新が一体 |
| helper へ promote / stale / UI policy を寄せる変更 | shared primitive の責務を超え、`TaskPaneSnapshotCacheService` / `CaseTemplateSnapshotService` / `TaskPaneSnapshotBuilderService` の境界を壊しやすい |
| `CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache` | 新規 CASE 初期化専用で、lookup promote とは意図的に上書き policy が異なる |
| `DocumentNamePromptService` の cache-only policy | 表示中 Pane と prompt 初期値の整合を保つ前提 |
| `DocumentTemplateResolver` の CASE cache 優先 | 開いている CASE の表示状態と実行解決元を揃える前提 |
| `KernelTemplateSyncService` の `TASKPANE_MASTER_VERSION` 更新と Base 書込 | 正本更新の責務そのもの |
| `WorkbookActivate` / `WindowActivate` 前提の Pane 再利用設計 | `docs/flows.md` と `docs/ui-policy.md` が維持対象としている |

## 8. 不明点

- `metadata` という語はコード上の正式型名ではなく、本書での便宜上の総称
- snapshot `META` 行に含まれる master version 自体は parser の主経路で保持されていないため、その保持理由の全量はこの範囲では確定しない
- `MasterTemplateCatalogService` のメモリ cache を複数 master root がどう使い分ける想定かは、この調査範囲では断定しない
