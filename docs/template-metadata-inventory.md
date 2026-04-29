# Template Metadata Inventory

## 1. 目的と前提

この文書は、雛形一覧 / TaskPane snapshot / 文書ボタン metadata の現状を、将来の統合修正前に棚卸しするための調査メモです。

- 参照前提:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
- 対象:
  - 雛形一覧
  - TaskPane snapshot
  - 文書ボタン metadata
  - `template key / caption / file / version / cache`
- 対象サービス:
  - `KernelTemplateSyncService`
  - `TaskPaneSnapshotBuilderService`
  - `MasterTemplateCatalogService`
  - `DocumentTemplateResolver`
  - `DocumentNamePromptService`
  - `TaskPaneSnapshotCacheService`
  - `TaskPaneManager`
- 関連補助として確認したサービス:
  - `CaseTemplateSnapshotService`
  - `CaseWorkbookInitializer`
  - `DocumentExecutionEligibilityService`
  - `DocumentCreateService`
  - `WordTemplateRegistrationValidationService`

この文書は現状整理のみを目的とし、仕様変更提案やコード変更は含みません。

## 2. 現状のデータ流れ

### 2.1 雛形登録・更新から雛形一覧まで

1. `WordTemplateRegistrationValidationService` が `SYSTEM_ROOT\雛形` 直下の Word テンプレート候補を走査する。
2. 各ファイルについて:
   - `key` はファイル名先頭の `NN_` から抽出・2桁正規化される。
   - `caption` の元値はファイル名の `_` 以降から `DisplayName` として抽出される。
   - ContentControl Tag 検証は `CaseList_FieldInventory` を基準に行われる。
3. `KernelTemplateSyncService` が Kernel の `雛形一覧` に反映する。
4. 反映時に `KernelTemplateSyncService` が書き換えるのは `雛形一覧` の A:C 列である。
   - A列: `key`
   - B列: `template file name`
   - C列: `caption`
5. 同サービスは `TASKPANE_MASTER_VERSION` を更新し、その後 Base 用 snapshot を生成して Base に埋め込む。
6. `雛形一覧` D:F は、人間による手修正で更新する運用である。
   - E列の tab 名と D/F の色は runtime で読取対象だが、自動生成・自動更新される前提ではない。
   - `KernelTemplateSyncService` が自動更新する対象は A:C に限られる。

### 2.2 雛形一覧から Base 埋込 snapshot まで

1. `KernelTemplateSyncService.BuildTaskPaneSnapshot` が `雛形一覧` を読み、snapshot 文字列を組み立てる。
2. snapshot には `META / SPECIAL / TAB / DOC` 行が含まれる。
3. Base には以下の DocProperty が保存される。
   - `TASKPANE_BASE_SNAPSHOT_COUNT`
   - `TASKPANE_BASE_SNAPSHOT_XX`
   - `TASKPANE_BASE_MASTER_VERSION`
   - `TASKPANE_MASTER_VERSION`

### 2.3 新規 CASE 作成時

1. `KernelCaseCreationService` が Base を物理コピーして CASE を作成する。
2. `CaseWorkbookInitializer` が CASE の基本 DocProperty を設定する。
3. `CaseWorkbookInitializer` は `CaseTemplateSnapshotService` を使って:
   - Kernel の `TASKPANE_MASTER_VERSION` を CASE に同期する。
   - Base 埋込 snapshot を CASE cache に昇格する。

### 2.4 CASE 表示・TaskPane 構築時

1. `TaskPaneManager` が CASE pane 描画時に `TaskPaneSnapshotBuilderService.BuildSnapshotText` を呼ぶ。
2. `TaskPaneSnapshotBuilderService` は次の優先順で snapshot を解決する。
   - CASE cache
   - Base 埋込 snapshot
   - Kernel `雛形一覧` から再構築
3. 取得した snapshot 文字列を `TaskPaneSnapshotParser` で `TaskPaneSnapshot` に変換する。
4. `CaseTaskPaneViewStateBuilder` が `TaskPaneSnapshot` を UI 用 ViewState に変換し、`TaskPaneManager` が描画する。

### 2.5 文書ボタン押下時

1. `TaskPaneManager` が押下された `actionKind` と `key` を受け取る。
2. `DocumentNamePromptService` が初期文書名を取得し、必要なら一時 override を CASE DocProperty に書く。
3. `DocumentCommandService` が `DocumentExecutionEligibilityService` を呼ぶ。
4. `DocumentExecutionEligibilityService` は `DocumentTemplateResolver` で `DocumentTemplateSpec` を解決する。
5. `DocumentTemplateResolver` は次の優先順で `key -> file / caption` を解決する。
   - `TaskPaneSnapshotCacheService` 経由の CASE cache
   - `MasterTemplateCatalogService` 経由の Kernel `雛形一覧`
6. `DocumentCreateService` が `DocumentTemplateSpec.DocumentName` と override 情報から最終文書名を決め、`TemplateFileName` から `TemplatePath` を導出して文書作成を実行する。

## 3. metadata 項目整理

| 項目 | 現状の正本 / 起点 | 主な生成元 | 主な読取先 | 主な変換 |
| --- | --- | --- | --- | --- |
| `template key` | 起点は雛形ファイル名の `NN_`、runtime 正本は Kernel `雛形一覧` A列 | `WordTemplateRegistrationValidationService`、`KernelTemplateSyncService` | `KernelTemplateSyncService`、`TaskPaneSnapshotBuilderService`、`MasterTemplateCatalogService`、`TaskPaneSnapshotCacheService`、`DocumentTemplateResolver` | 数値 key を 2桁文字列へ正規化 |
| `caption` | 起点は雛形ファイル名の `_` 以降、runtime 正本は Kernel `雛形一覧` C列 | `WordTemplateRegistrationValidationService`、`KernelTemplateSyncService` | `TaskPaneSnapshotBuilderService`、`MasterTemplateCatalogService`、`TaskPaneSnapshotCacheService`、`DocumentNamePromptService`、`DocumentCreateService` | `DisplayName` / `DocumentName` / snapshot `Caption` として受け渡し |
| `file name` | Kernel `雛形一覧` B列 | `WordTemplateRegistrationValidationService`、`KernelTemplateSyncService` | `TaskPaneSnapshotBuilderService`、`MasterTemplateCatalogService`、`TaskPaneSnapshotCacheService`、`DocumentTemplateResolver` | snapshot `TemplateFileName`、`DocumentTemplateSpec.TemplateFileName` へ変換 |
| `file path` | 正本としては保持せず、`WORD_TEMPLATE_DIR` または `SYSTEM_ROOT\雛形` から都度導出 | `DocumentTemplateResolver` | `DocumentExecutionEligibilityService`、`DocumentCreateService` | `TemplateFileName + template directory -> TemplatePath` |
| `document kind` | 雛形一覧には持たず、runtime では `doc` が派生的に付与される | `KernelTemplateSyncService`、`TaskPaneSnapshotBuilderService`、`DocumentTemplateResolver` | `TaskPaneManager`、`DocumentCommandService` | snapshot `ActionKind=doc`、`DocumentTemplateSpec.ActionKind=doc` |
| `version` | 正本は Kernel `TASKPANE_MASTER_VERSION` | `KernelTemplateSyncService` | `TaskPaneSnapshotBuilderService`、`CaseTemplateSnapshotService`、`TaskPaneSnapshotCacheService`、`DocumentExecutionEligibilityService` | Base / CASE 側には mirror としてコピーされる |
| `cache` | Base 埋込 cache と CASE cache はいずれも派生情報 | `KernelTemplateSyncService`、`CaseTemplateSnapshotService`、`TaskPaneSnapshotBuilderService`、`TaskPaneSnapshotCacheService` | `TaskPaneSnapshotBuilderService`、`TaskPaneSnapshotCacheService`、`DocumentTemplateResolver`、`DocumentNamePromptService` | chunk 化して DocProperty 保存 |
| `snapshot` | 雛形一覧から生成される派生表現 | `KernelTemplateSyncService`、`TaskPaneSnapshotBuilderService` | `TaskPaneSnapshotParser`、`TaskPaneSnapshotCacheService` | `META / SPECIAL / TAB / DOC` のTSV文字列 |
| Base 埋込情報 | Base の `TASKPANE_BASE_*` と `TASKPANE_MASTER_VERSION` | `KernelTemplateSyncService` | `CaseTemplateSnapshotService`、`TaskPaneSnapshotBuilderService`、`TaskPaneSnapshotCacheService` | 新規 CASE に物理コピーされる |
| CASE 側 cache | CASE の `TASKPANE_SNAPSHOT_CACHE_*` と CASE側 `TASKPANE_MASTER_VERSION` | `CaseTemplateSnapshotService`、`TaskPaneSnapshotBuilderService`、`TaskPaneSnapshotCacheService` | `TaskPaneSnapshotBuilderService`、`TaskPaneSnapshotCacheService`、`DocumentTemplateResolver`、`DocumentNamePromptService` | 表示中 Pane と実行時解決元をそろえるために利用 |

### 3.1 `template key`

- 作成:
  - `WordTemplateRegistrationValidationService.ValidateKey` がファイル名から抽出する。
  - `KernelTemplateSyncService.WriteToMasterList` が `雛形一覧` A列へ書く。
- 読取:
  - `TaskPaneSnapshotBuilderService`
  - `MasterTemplateCatalogService`
  - `TaskPaneSnapshotCacheService`
  - `DocumentTemplateResolver`
  - `TaskPaneManager`
- 変換:
  - 複数サービスが独自に 2桁正規化を持つ。

### 3.2 `caption`

- 作成:
  - `WordTemplateRegistrationValidationService.ExtractDisplayName` がファイル名から生成する。
  - `KernelTemplateSyncService.WriteToMasterList` が `雛形一覧` C列へ書く。
- 読取:
  - snapshot builder 群は `DOC` 行の `Caption` に積む。
  - `MasterTemplateCatalogService` は `DocumentName` として返す。
  - `TaskPaneSnapshotCacheService` は snapshot の `Caption` を `documentName` として返す。
  - `DocumentNamePromptService` は prompt 初期値に使う。
  - `DocumentCreateService` は最終文書名の既定値に使う。
- 補足:
  - 現行コードでは UI 表示名と既定文書名に同じ値が使われている。

### 3.3 `file name / file path`

- `file name`
  - 正本は `雛形一覧` B列。
  - snapshot では `DOC` 行の末尾に積まれる。
  - `MasterTemplateCatalogService` でも `TemplateFileName` として返る。
- `file path`
  - 正本として保存されず、`DocumentTemplateResolver` が都度導出する。
  - 優先順:
    - `WORD_TEMPLATE_DIR`
    - `SYSTEM_ROOT\雛形`

### 3.4 `document kind`

- 文書ボタンについては `doc` が hard-coded で付与される。
- `雛形一覧` の列には保存されない。
- `TaskPaneSnapshot` と `DocumentTemplateSpec` の両方で派生的に持つ。

### 3.5 `version`

- Kernel:
  - `TASKPANE_MASTER_VERSION` が現行 master version の正本。
- Base:
  - `TASKPANE_BASE_MASTER_VERSION` は Base 埋込 snapshot がどの master version 由来かを示す。
  - `TASKPANE_MASTER_VERSION` も mirror として保存される。
- CASE:
  - `TASKPANE_MASTER_VERSION` は CASE cache の provenance と freshness 判定用であり、global 正本ではない。
  - 開いている CASE が最新 master に追随しないことは現行仕様である。

### 3.6 `cache / snapshot`

- Base 埋込 snapshot:
  - `TASKPANE_BASE_SNAPSHOT_COUNT`
  - `TASKPANE_BASE_SNAPSHOT_XX`
- CASE cache:
  - `TASKPANE_SNAPSHOT_CACHE_COUNT`
  - `TASKPANE_SNAPSHOT_CACHE_XX`
- snapshot format:
  - `META`
  - `SPECIAL`
  - `TAB`
  - `DOC`

## 4. 正本・派生・補助・一時情報の分類

### 4.1 正本と判断できた情報

| 分類対象 | 正本と判断した場所 | 根拠 |
| --- | --- | --- |
| runtime 用 `template key / file name / caption` | Kernel `雛形一覧` A:C | `MasterTemplateCatalogService` と snapshot builder 群がここを直接読むため |
| runtime 用 tab 名 / 色 | Kernel `雛形一覧` D:F / E:F | snapshot builder 群がここを読む一方、sync は A:C しか書かず、D:F は人間による手修正運用で維持されるため |
| global master version | Kernel `TASKPANE_MASTER_VERSION` | 更新元が `KernelTemplateSyncService` に集約されているため |
| 実体テンプレートファイル | `SYSTEM_ROOT\雛形` 配下ファイル | 文書作成時に最終的に参照される実ファイルであるため |

### 4.2 正本から生成される派生情報

- Base 埋込 snapshot (`TASKPANE_BASE_*`)
- CASE snapshot cache (`TASKPANE_SNAPSHOT_CACHE_*`)
- `TaskPaneSnapshot`
- `TaskPaneDocDefinition`
- `MasterTemplateRecord`
- `DocumentTemplateSpec`
- `TemplatePath`
- `ActionKind=doc`

### 4.3 表示用に加工された情報

- `CaseTaskPaneViewState`
- `CaseTaskPaneActionViewState`
- `CaseTaskPaneTabPageViewState`
- CASELIST 状態に応じた special button の caption / backcolor 上書き

### 4.4 実行時にだけ使う一時情報

- `TASKPANE_DOC_NAME_OVERRIDE_ENABLED`
- `TASKPANE_DOC_NAME_OVERRIDE`
- `TASKPANE_SUPPRESS_CASE_REVEAL`
- `DocumentNameOverrideScope`

### 4.5 古い仕様由来で残っている可能性がある情報

以下は「現行コードで生成または保持はされるが、主経路での利用が限定的または未確認」な情報です。

| 情報 | 観測できた事実 | 現時点の扱い |
| --- | --- | --- |
| snapshot `META` 行の埋込 master version | builder は `META` に version を書くが、`TaskPaneSnapshotParser.ParseMeta` はそれを `TaskPaneSnapshot` に保持しない | 保持理由は未確認 |
| snapshot `META` 行の workbook 名 / path | parser は読むが、現行コード検索では主に `ERROR` 判定以外の利用を確認できなかった | 補助情報の可能性 |
| snapshot `PreferredPaneWidth` | parser は読むが、現行 CASE UI は `DocTaskPaneControl` 側で再計算した `PreferredPaneWidthHint` を使う | 現行表示主経路では未使用に見える |
| `CaseTemplateSnapshotService` | 新規 CASE 初期化では使われる一方、同種の cache / promote ロジックは別サービスにもある | 残存補助かどうかは未確認 |

## 5. サービス別責務表

| サービス | 入力 | 出力 | 主に読む情報種別 | 重複している責務 | 将来寄せる候補 |
| --- | --- | --- | --- | --- | --- |
| `KernelTemplateSyncService` | 雛形フォルダ、`CaseList_FieldInventory`、Kernel `雛形一覧` | `雛形一覧` A:C 更新、Kernel version、Base snapshot/version | 生ファイル入力 + master sheet | master sheet 解釈、snapshot 生成、chunk 保存、key 正規化 | Master metadata reader + snapshot serializer |
| `TaskPaneSnapshotBuilderService` | CASE workbook、CASE cache、Base 埋込、Kernel `雛形一覧` | snapshot text、CASE cache 更新 | 派生 cache 優先、必要時に正本 | master sheet 解釈、snapshot 生成、cache load/save、version 判定 | snapshot build/storage coordinator |
| `MasterTemplateCatalogService` | CASE workbook から解決した Kernel path、Kernel `雛形一覧` | `MasterTemplateRecord` 一覧 / key lookup | 正本の master sheet | master sheet 解釈、key 正規化 | shared master metadata reader |
| `DocumentTemplateResolver` | CASE workbook、doc key | `DocumentTemplateSpec` | CASE cache 優先、master fallback | key->caption/file 解決 | shared template metadata resolver |
| `DocumentNamePromptService` | CASE workbook、doc key | `DocumentNameOverrideScope` | CASE cache のみ | key->caption 解決 | `DocumentTemplateResolver` と統合または共通化 |
| `TaskPaneSnapshotCacheService` | CASE workbook | Base->CASE promote、cache lookup、cache clear | Base/CASE 派生 cache | chunk load/save、base promote、key 正規化 | shared snapshot storage service |
| `TaskPaneManager` | `WorkbookContext`、snapshot builder 結果、UI action | CASE pane 描画、action dispatch | 派生 snapshot / UI state | action 前 prompt と action 後 refresh の調停 | 表示責務は維持、metadata 所有は持たせない |

### 5.1 関連補助サービス

| サービス | 入力 | 出力 | 観測できた位置づけ |
| --- | --- | --- | --- |
| `CaseTemplateSnapshotService` | Kernel workbook、CASE workbook | CASE version 同期、Base 埋込 snapshot の CASE cache 昇格 | 新規 CASE 初期化専用の補助。`TaskPaneSnapshotCacheService` と役割が近い |
| `CaseWorkbookInitializer` | Kernel workbook、CASE workbook、作成 plan | CASE 基本 DocProperty、顧客名反映、snapshot 初期化 | 新規 CASE 作成フローの入口 |
| `DocumentExecutionEligibilityService` | CASE workbook、actionKind、key | eligibility、`DocumentTemplateSpec`、`CaseContext` | resolver の結果を実行前チェックに接続 |
| `DocumentCreateService` | CASE workbook、`DocumentTemplateSpec`、`CaseContext` | Word 文書作成 | `DocumentName` を既定文書名として使用 |

## 6. 重複している解釈ロジック

### 6.1 `雛形一覧` の読み取り解釈が複数サービスに分散

- `KernelTemplateSyncService`
  - `ReadMasterSheetSnapshot`
  - `BuildTaskPaneSnapshot`
- `TaskPaneSnapshotBuilderService`
  - `ReadMasterSheetSnapshot`
  - `AppendTemplateDefinitions`
- `MasterTemplateCatalogService`
  - `ReadMasterTemplateList`

重複している内容:

- A:C / D:F の列意味解釈
- key の 2桁正規化
- 行スキップ条件
- 色取得

### 6.2 snapshot 文字列の組立が複数サービスに分散

- `KernelTemplateSyncService.BuildTaskPaneSnapshot`
- `TaskPaneSnapshotBuilderService.BuildSnapshotText`

重複している内容:

- `META / SPECIAL / TAB / DOC` の行構造
- field escape
- タブ順と row index の組立
- `全て` タブの補完
- preferred pane width 算出

### 6.3 cache chunk 読み書きが複数サービスに分散

- `TaskPaneSnapshotBuilderService`
- `TaskPaneSnapshotCacheService`
- `CaseTemplateSnapshotService`

重複している内容:

- `*_COUNT` と `*_XX` の chunk 保存
- Base -> CASE promote
- cache clear
- version copy

### 6.4 `key -> caption/file` 解決経路が複数ある

- `TaskPaneSnapshotCacheService.TryGetDocInfoFromCache`
- `MasterTemplateCatalogService.TryGetTemplateByKey`
- `DocumentTemplateResolver.Resolve`
- `DocumentNamePromptService.FindDocumentCaptionByKey`

特に差分がある点:

- `DocumentTemplateResolver` は CASE cache に無ければ master catalog にフォールバックする。
- `DocumentNamePromptService` は CASE cache のみを参照し、master catalog にはフォールバックしない。
- 現行の CASE pane 操作では通常 cache がある前提で動くが、lookup ポリシー自体は一致していない。

### 6.5 version の意味付けが複数ある

- Kernel `TASKPANE_MASTER_VERSION`
  - global master version
- Base `TASKPANE_BASE_MASTER_VERSION`
  - Base 埋込 snapshot の由来 version
- CASE `TASKPANE_MASTER_VERSION`
  - CASE cache の freshness / provenance
- snapshot `META` 埋込 version
  - 現行 parser では利用先未確認

## 7. 将来の統合候補

### 7.1 `雛形一覧` の single reader 化

候補:

- `key / caption / file / tab / colors` を一度だけ解釈する shared reader を持つ
- `MasterTemplateCatalogService`、`TaskPaneSnapshotBuilderService`、`KernelTemplateSyncService` はその reader を使う形へ寄せる
- ただし `雛形一覧` D:F の手修正運用は前提として維持し、tab 名や色を自動生成・自動更新する設計に変えない

期待効果:

- 列意味解釈の分散削減
- key 正規化の二重実装削減
- tab / color を含む完全な master projection を一箇所で定義できる

### 7.2 snapshot serializer / parser 契約の一本化

候補:

- snapshot の build・parse・format version を一つの契約に寄せる
- `META` 行で保持する項目と、実際に使う項目を一致させる

期待効果:

- builder 間の重複削減
- `META` 埋込情報の生死が明確になる

### 7.3 Base / CASE cache の storage API 統合

候補:

- `TaskPaneSnapshotCacheService` と `CaseTemplateSnapshotService` が持つ chunk 保存・promote・clear を一本化する

期待効果:

- 新規 CASE 作成時と表示時の cache 操作差分を減らせる
- Base/CASE/version の更新責務を整理しやすい

### 7.4 `key -> template metadata` 解決の単一路線化

候補:

- `DocumentNamePromptService` と `DocumentTemplateResolver` が共通の lookup を使う

期待効果:

- prompt 初期値と実行時 resolver の乖離を避けやすい
- caption / file 解釈ロジックの重複削減

## 8. 今回あえて変更しないこと

- `TASKPANE_MASTER_VERSION` の更新方針
- snapshot format
- Base 埋込 snapshot の存在
- CASE cache 優先の解決順
- `WorkbookActivate` / `WindowActivate` 時の host 再利用方針
- 文書ボタン実行フロー
- 文書名 prompt の UI 仕様
- 雛形一覧の列構成

## 9. 未確認事項

- snapshot `META` の version / workbook name / workbook path / preferred width が外部利用されるか
  - `dev\CaseInfoSystem.ExcelAddIn` 内では主経路利用を確認できない項目がある。
- `CaseTemplateSnapshotService` が現役設計なのか、移行途中の残存補助なのか
  - 新規 CASE 初期化では使われているが、同系統処理は別サービスにも存在する。
- `DocumentNamePromptService` の cache-only lookup が意図的設計かどうか
  - 実装事実は確認できるが、設計意図まではコードだけでは断定できない。

## 10. まとめ

- runtime の雛形 metadata 正本は、テンプレートファイルそのものではなく、同期後の Kernel `雛形一覧` と判断できる。
- Base 埋込 snapshot と CASE cache は、どちらも正本ではなく派生 cache である。
- global version の正本は Kernel `TASKPANE_MASTER_VERSION` であり、Base / CASE 側の同名または関連 version は mirror / provenance として読むのが現状に合う。
- 重複の中心は `雛形一覧` 解釈、snapshot 生成、cache 保存、`key -> caption/file` 解決である。
- 特に `TaskPaneSnapshotBuilderService`、`KernelTemplateSyncService`、`MasterTemplateCatalogService`、`TaskPaneSnapshotCacheService`、`CaseTemplateSnapshotService` の間に、統合余地が大きい。
