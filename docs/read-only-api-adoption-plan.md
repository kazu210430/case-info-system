# Read-Only API Adoption Plan

## 調査条件

- 基準点確認:
  - `main` = `15462c7f62d85d557010dd513e34c24bdc56b944`
  - `origin/main` = `15462c7f62d85d557010dd513e34c24bdc56b944`
- 作業ツリー確認:
  - 調査開始時点で `git status --short` は空
- 参照した文書:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - 補助資料として `docs/template-metadata-inventory.md`
- 今回の作業:
  - コード変更なし
  - build 未実行
  - 本ファイル作成のみ

## 1. 現在の参照経路の棚卸し

| 対象 | 現在の参照経路 | 主サービス | 事実 |
| --- | --- | --- | --- |
| TaskPane 表示 | `CASE cache -> Base 埋込 snapshot -> Kernel 雛形一覧再構築` | `TaskPaneSnapshotBuilderService` | `BuildSnapshotText` がこの優先順で読取し、必要時に CASE cache を更新する |
| prompt 初期値 | `CASE cache only` | `TaskPaneManager` -> `DocumentNamePromptService` -> `ICaseCacheDocumentTemplateReader` -> `DocumentTemplateLookupService.TryEnsurePromotedCaseCacheThenResolve` | CASE cache miss 時は空欄で prompt を開く。master fallback はしない。ただし lookup 前の promotion により CASE DocProperty が更新される場合がある |
| 文書実行時 metadata 解決 | `CASE cache -> master catalog fallback` | `DocumentExecutionEligibilityService` -> `DocumentTemplateResolver` -> `DocumentTemplateLookupService.TryResolveWithMasterFallback` | CASE cache miss 時のみ `MasterTemplateCatalogService` を読む |
| `TemplatePath` 解決 | `WORD_TEMPLATE_DIR` 優先、未設定なら `SYSTEM_ROOT\雛形` | `DocumentTemplateResolver` | `TemplatePath` は保存されず、resolver が都度導出する |
| 文書作成時の最終文書名 | `templateSpec.DocumentName` を既定値とし、prompt override があれば上書き | `DocumentCreateService` | prompt の入力値は `TASKPANE_DOC_NAME_OVERRIDE_*` 経由で一時反映される |
| Base -> CASE cache 昇格 | `TASKPANE_BASE_*` を `TASKPANE_SNAPSHOT_CACHE_*` に昇格 | `TaskPaneSnapshotCacheService` / `CaseTemplateSnapshotService` | 両サービスに近い責務がある |
| 雛形一覧 A:F の読取 | `MasterTemplateSheetReader.Read` | `KernelTemplateSyncService` / `TaskPaneSnapshotBuilderService` / `MasterTemplateCatalogService` | Master sheet の行解釈自体は既に共通 reader 化されている |
| 文書ボタン metadata の UI 化 | `snapshot DOC 行 -> parser -> view state` | `TaskPaneSnapshotParser` / `CaseTaskPaneViewStateBuilder` / `TaskPaneManager` | 表示は snapshot 依存だが、実行正本ではない |

## 2. caption / file / key / TemplatePath の取得元一覧

| 項目 | 現在の取得元 | 主な利用先 | 補足 |
| --- | --- | --- | --- |
| `key` | Kernel `雛形一覧` A列、snapshot `DOC.Key`、UI action `e.Key` | `TaskPaneSnapshotBuilderService` `MasterTemplateCatalogService` `TaskPaneSnapshotCacheService` `DocumentTemplateResolver` `DocumentNamePromptService` | 2桁正規化は複数箇所にある |
| `caption` | Kernel `雛形一覧` C列、snapshot `DOC.Caption` | `MasterTemplateCatalogService` `TaskPaneSnapshotCacheService` `DocumentNamePromptService` `DocumentCreateService` | コード上は `DocumentName` 名義で流れる |
| `file` (`TemplateFileName`) | Kernel `雛形一覧` B列、snapshot `DOC.TemplateFileName` | `MasterTemplateCatalogService` `TaskPaneSnapshotCacheService` `DocumentTemplateResolver` | prompt では未使用、実行側で使用 |
| `TemplatePath` | `TemplateFileName + WORD_TEMPLATE_DIR` または `TemplateFileName + SYSTEM_ROOT\雛形` | `DocumentTemplateResolver` が生成し、`DocumentExecutionEligibilityService` と `DocumentCreateService` が消費 | path 自体を snapshot / cache / master sheet に保持していない |

追加の取得元メモ:

- Master sheet の A:F 解釈は `MasterTemplateSheetReader` に集約済み
- CASE cache の `caption/file/key` 解釈は `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` に集約済み
- 文書実行用の `caption/file/key/source` 解釈は `DocumentTemplateLookupService` が調停している

## 3. read-only API に差し替えてよい候補

低リスク候補のみを列挙する。

| 候補 | 現状 | 差し替え先候補 | 危険度 | 根拠 |
| --- | --- | --- | --- | --- |
| prompt 系の新規参照 | CASE cache を直接読む実装を今後増やす余地がある | `ICaseCacheDocumentTemplateReader` | 低 | 既に `DocumentNamePromptService` が採用済みで、cache-only 仕様を表現できている |
| 文書実行系の新規 metadata 参照 | `MasterTemplateCatalogService` 直接呼び出しを増やすと CASE cache 優先順を壊しやすい | `DocumentTemplateLookupService.TryResolveWithMasterFallback` | 低 | 既に `DocumentTemplateResolver` が採用済みで、CASE cache 優先 + master fallback を保持している |
| Master sheet A:F の新規 read-only 参照 | 個別サービスが worksheet を独自解釈すると列意味の重複が増える | `MasterTemplateSheetReader` | 低 | 既に 3 サービスで使用中。既存責務を壊さず読取だけ共通化できる |
| CASE cache から `caption/file/key/source` を取得する promotion-aware 参照 | `TaskPaneSnapshotParser` を各所で再利用すると snapshot 依存が広がる | `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` または `ICaseCacheDocumentTemplateReader` | 低 | 既存の CASE cache reader に結果型がある。pure read ではない |
| 旧 `TryEnsurePromotedCaseCacheThenGetDocInfo` の利用拡大抑止 | 文字列 out 2本の旧 API | `DocumentTemplateLookupResult` ベース API へ寄せる | 低 | 現行 tree 上では未使用。今後の増殖を止めやすい |

## 4. 差し替えを禁止または保留すべき箇所

| 箇所 | 判定 | 理由 |
| --- | --- | --- |
| `KernelTemplateSyncService` の `雛形一覧` 更新、`TASKPANE_MASTER_VERSION` 更新、Base snapshot 書込 | 禁止 | 書込責務そのもの。read-only API 化の対象ではない |
| `TaskPaneSnapshotBuilderService` の `CASE cache -> Base -> Master rebuild` 判断 | 保留 | 表示フロー・stale 判定・CASE cache 更新が一体化しており、読取だけ切り出すと挙動差異が出やすい |
| `CaseTemplateSnapshotService` の Base -> CASE 昇格 | 保留 | 新規 CASE 初期化フローと密結合。read-only 差し替えの最小単位に向かない |
| `DocumentExecutionEligibilityService` の `TemplatePath` 実在確認、拡張子確認、macro 判定 | 禁止 | 実行可否の最終判断であり、snapshot / cache を正本化してはいけない |
| `DocumentCreateService` の `TemplatePath` 使用 | 禁止 | 実ファイル生成処理。resolver の結果を使うだけに留めるべき |
| `AccountingTemplateResolver` | 保留 | Word 文書 metadata lookup とは別責務。会計 Excel テンプレート探索を混ぜない |
| `TaskPaneManager` の UI 表示制御 | 禁止 | `docs/ui-policy.md` 前提。read-only API 採用のために UI 制御責務へ波及させない |
| snapshot `META` を正本扱いする変更 | 禁止 | parser は `META` の version を保持しておらず、表示補助以上の役割が確認できない |

不明:

- `TaskPaneSnapshotBuilderService` の cache 読取部分だけを独立 read-only API に分離した場合、既存の host 再利用・再描画頻度に副作用が出ないかは docs 上では不明

## 5. prompt cache-only ルールに関係する箇所

- `TaskPaneManager.ExecuteCaseAction`
  - `ActionKind=doc` のときだけ `DocumentNamePromptService.TryPrepare` を先に呼ぶ
- `DocumentNamePromptService`
  - `FindDocumentCaptionByKey` は `ICaseCacheDocumentTemplateReader.TryEnsurePromotedCaseCacheThenResolve` だけを呼ぶ
- `ICaseCacheDocumentTemplateReader`
  - interface 自体が `CASE cache だけを読み取り` と明記している
- `DocumentTemplateLookupService.TryEnsurePromotedCaseCacheThenResolve`
  - `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` への薄い委譲
- `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup`
  - CASE cache / Base 昇格後の snapshot から `DocumentName` と `TemplateFileName` を返す
  - Base 昇格時に CASE cache chunk と `TASKPANE_MASTER_VERSION` を DocProperty に書き戻す
- `DocumentNameOverrideScope`
  - prompt 入力値を一時 DocProperty に保存
- `DocumentCreateService.ResolveDocumentName`
  - override が有効なときだけ prompt 入力値を最終文書名へ反映

確認済みテスト:

- `dev/CaseInfoSystem.SnapshotRegressionTests/DocumentTemplateLookupServiceTests.cs`
  - CASE cache hit 時は resolver と prompt が同じ caption を使う
  - CASE cache miss 時は resolver は master fallback するが、prompt は空欄のまま

## 6. 文書実行 CASE cache 優先 + master fallback ルールに関係する箇所

- `DocumentCommandService.ExecuteDocumentAction`
  - 実行前に `DocumentExecutionEligibilityService.Evaluate` を呼ぶ
- `DocumentExecutionEligibilityService.Evaluate`
  - `DocumentTemplateResolver.Resolve` の結果で eligibility を判定
- `DocumentTemplateResolver.Resolve`
  - `DocumentTemplateLookupService.TryResolveWithMasterFallback` を呼ぶ
- `DocumentTemplateLookupService.TryResolveWithMasterFallback`
  - 先に `TryEnsurePromotedCaseCacheThenResolve`
  - miss 時だけ `MasterTemplateCatalogService.TryGetTemplateByKey`
- `MasterTemplateCatalogService`
  - `SYSTEM_ROOT` を基点に Master を read-only で開き、`雛形一覧` を読む
- `DocumentCreateService`
  - eligibility 済み `templateSpec` を使って Word 作成を実行

このルールにより、prompt と違って文書実行側だけが master fallback を持つ。

## 7. TemplatePath resolver 責務に関係する箇所

- `DocumentTemplateResolver.Resolve`
  - `TemplateFileName` と template directory から `TemplatePath` を構築
- `DocumentTemplateResolver.ResolveTemplateDirectory`
  - `WORD_TEMPLATE_DIR` 優先
  - 未設定時は `SYSTEM_ROOT\雛形`
- `DocumentTemplateResolver.TemplateExists`
  - `TemplatePath` の実在確認
- `DocumentExecutionEligibilityService`
  - `TemplatePath` を読んで supported / macro / exists を判定するだけ
- `DocumentCreateService`
  - `TemplatePath` を読んで `CreateDocumentFromTemplate` するだけ

結論:

- `TemplatePath` を導出する責務は `DocumentTemplateResolver` に留めるべき
- 他サービスが `SYSTEM_ROOT` や `WORD_TEMPLATE_DIR` から独自に Word 文書 `TemplatePath` を組み立てる差し替えは避ける

## 8. snapshot が正本化しないよう注意すべき箇所

- `TaskPaneSnapshotBuilderService`
  - snapshot 生成時に `CASELIST_REGISTERED` で special button の caption / 色を動的上書きする
  - 表示用の加工が混ざっているため、正本ではない
- `TaskPaneSnapshotParser`
  - `META.WorkbookName` `META.WorkbookPath` `PreferredPaneWidth` は保持するが、version は保持しない
  - `META` 全体を正本前提で使う設計にはなっていない
- `TaskPaneManager`
  - render signature に `TASKPANE_SNAPSHOT_CACHE_COUNT` と `CASELIST_REGISTERED` を使う
  - 表示更新判定の都合で snapshot 関連状態を見ているだけ
- `CaseListRegistrationService`
  - 案件一覧登録後に `TASKPANE_SNAPSHOT_CACHE_COUNT=0` と chunk clear を実施する
  - CASE 状態変化で cache を無効化しており、cache が正本ではないことを示している
- `TaskPaneSnapshotCacheService`
  - Base snapshot を CASE cache に昇格する
  - Base / CASE いずれも派生 cache であり、global 正本ではない

不明:

- snapshot `META` に version を書いている理由のうち、表示以外の利用先は今回の範囲では確認できず不明

## 9. 次に実装する場合の最小単位案

### 案A: CASE cache-only reader の利用拡大

- 変更対象:
  - prompt 系または CASE 表示整合が必要な新規 read-only 参照
- 使う口:
  - `ICaseCacheDocumentTemplateReader`
- 狙い:
  - cache-only 仕様を型で固定する
- 危険度:
  - 低

### 案B: 実行系 lookup の入口統一

- 変更対象:
  - 文書実行系で今後追加される `key -> caption/file` 参照
- 使う口:
  - `DocumentTemplateLookupService.TryResolveWithMasterFallback`
- 狙い:
  - `CASE cache 優先 + master fallback` を再実装させない
- 危険度:
  - 低

### 案C: Master sheet 読取の追加参照を `MasterTemplateSheetReader` に限定

- 変更対象:
  - `雛形一覧` A:F を新たに読む必要がある read-only 処理
- 使う口:
  - `MasterTemplateSheetReader`
- 狙い:
  - 列解釈の重複を増やさない
- 危険度:
  - 低

### 案D: snapshot storage read helper の抽出

- 変更対象:
  - `TaskPaneSnapshotBuilderService` と `TaskPaneSnapshotCacheService` と `CaseTemplateSnapshotService` に分散した読取ロジック
- 狙い:
  - Base/CASE cache の読み取り境界整理
- 危険度:
  - 中
- 備考:
  - 書込・昇格・invalidate を触るとリスクが上がるため、最初は read-only 部分だけに限定する

## 10. 推奨されるテスト追加案

- `DocumentTemplateLookupService` / `DocumentNamePromptService`
  - CASE cache stale でも prompt は master fallback しないこと
- `DocumentTemplateResolver`
  - CASE cache stale または空のとき、master fallback しても `TemplatePath` 解決責務が resolver に留まること
- `TaskPaneSnapshotCacheService`
  - Base 昇格後の `TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` が `DocumentTemplateLookupResult` を安定して返し、CASE cache / version DocProperty 更新を維持すること
- `CaseListRegistrationService` + snapshot
  - `CASELIST_REGISTERED=1` 後に CASE cache が無効化され、再構築時に special button だけが動的反映されること
- `MasterTemplateSheetReader`
  - A:F の空行 / 不完全行 / tab 色 / fill 色の解釈が既存挙動と同値であること
- 同値確認テスト
  - 新 read-only API を導入する場合、現行呼び出し元と `key / caption / file / source` が同値になること

既存で土台になるテスト:

- `dev/CaseInfoSystem.SnapshotRegressionTests/DocumentTemplateLookupServiceTests.cs`
- `dev/CaseInfoSystem.Tests/MasterTemplateSheetReaderTests.cs`
- `dev/CaseInfoSystem.SnapshotRegressionTests/SnapshotOutputRegressionTests.cs`

## まとめ

- 既に low-risk に寄せられている read-only 入口は `ICaseCacheDocumentTemplateReader` `DocumentTemplateLookupService` `MasterTemplateSheetReader` の3つである
- 次に広げるなら、この3つの利用拡大を優先し、`KernelTemplateSyncService` `TaskPaneSnapshotBuilderService` の書込・昇格・UI表示責務には踏み込まない方が安全である
- `prompt cache-only`、`文書実行は CASE cache 優先 + master fallback`、`TemplatePath は resolver 責務`、`snapshot は正本ではない` の4点は固定条件として維持する
