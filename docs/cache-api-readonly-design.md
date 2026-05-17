# Cache API Readonly Design

## 1. 目的

この文書は、[docs/template-metadata-inventory.md](docs/template-metadata-inventory.md) の棚卸し結果を前提に、将来の cache API 統合へ向けた読み取り専用 API 候補を整理するための設計メモである。

今回の対象は docs 整理のみであり、実装変更は行わない。

### 1.1 非目的

- `DocumentNamePromptService` の cache-only 仕様変更
- `DocumentTemplateResolver` の master fallback 仕様変更
- `TaskPaneSnapshotBuilderService` の cache 更新挙動変更
- snapshot を正本扱いすること
- cache を保存・生成・実行判断の正本扱いにすること

## 2. 前提となる境界

- runtime の template metadata の正本は Kernel `雛形一覧` と `TASKPANE_MASTER_VERSION` である。
- Base `TASKPANE_BASE_*` と CASE `TASKPANE_SNAPSHOT_CACHE_*` は派生 cache である。
- `TaskPaneSnapshot` は表示用 snapshot であり、保存・生成・実行判断の正本ではない。
- `DocumentNamePromptService` は CASE cache hit 時だけ caption を prompt 初期値に使い、cache miss 時に master fallback しない。
- 文書実行は `DocumentTemplateResolver -> DocumentTemplateLookupService.TryResolveWithMasterFallback` を通り、CASE cache 優先、master fallback ありである。
- Base snapshot は新規 CASE 初期配布と TaskPane 表示再構築の補助であり、prompt lookup や文書実行 lookup の直接正本ではない。
- 現行の CASE cache lookup は pure read ではない。`TryEnsurePromotedCaseCacheThenResolve` / `TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` は、必要時に Base snapshot promotion と CASE DocProperty 更新を行う。

## 3. 既存の読み取り口の分類

| 読み取り口 | 現在の主な呼び出し元 | 一次参照元 | fallback | 副作用 | 備考 |
| --- | --- | --- | --- | --- | --- |
| prompt 用 lookup | `DocumentNamePromptService` | `TaskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup` | 禁止 | Base snapshot promote が起きうる | `caption` だけを prompt 初期値に使う |
| 文書実行用 lookup | `DocumentTemplateResolver` | CASE cache | `MasterTemplateCatalogService` へ fallback あり | CASE cache leg で Base snapshot promote が起きうる | `TemplatePath` 解決は resolver の責務 |
| TaskPane 表示用 snapshot | `TaskPaneManager` / `TaskPaneSnapshotBuilderService` | CASE cache | Base snapshot, さらに MasterList rebuild | CASE cache 更新あり | 表示用 read model を返すが read-only ではない |
| CASE cache 参照 | `DocumentTemplateLookupService`, `TaskPaneSnapshotBuilderService` | `TASKPANE_SNAPSHOT_CACHE_*` | なし | format 非互換時 clear, promote あり | `TaskPaneSnapshotCacheService` が入口 |
| Base snapshot 参照 | `CaseTemplateSnapshotService`, `TaskPaneSnapshotBuilderService`, `TaskPaneSnapshotCacheService` | `TASKPANE_BASE_*` | CASE cache miss 時の補助 | CASE cache へ copy あり | 新規 CASE 初期配布と stale 時 fallback 用 |
| master fallback ありの参照 | `DocumentTemplateResolver`, `TaskPaneSnapshotBuilderService` | `MasterTemplateCatalogService` または Master workbook/package property | あり | snapshot rebuild 時は CASE cache 更新あり | 実行系と表示系で用途が異なる |
| master fallback 禁止の参照 | `DocumentNamePromptService` | CASE cache のみ | 禁止 | Base snapshot promote が起きうる | prompt の誤補完防止が優先 |

## 4. API 候補

公開候補は、prompt 用と実行用を別 API として分ける。共通化が必要でも、呼び出し側に raw bool を渡す設計ではなく、mode / policy / caller intent を明示できる内部 API に限定する。

| API候補 | 主目的 | 入力 | 戻り値 | 参照元 | fallback可否 | 利用してよいサービス | 利用禁止の場面 | 備考 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `ICaseCacheDocumentTemplateReader` | CASE cache-only で `key -> caption/file` を解決する | `Workbook`, `key` | `DocumentTemplateLookupResult` | `TaskPaneSnapshotCacheService` | 不可 | `DocumentNamePromptService` | `DocumentTemplateResolver`, 保存・生成・実行判断 | prompt 用 facade 候補。cache miss は空で返す。現行実装は promotion-aware で pure read ではない |
| `IDocumentTemplateLookupReader` | 文書実行用の metadata lookup を解決する | `Workbook`, `key`, `LookupIntent` | `DocumentTemplateLookupResult` | CASE cache, 必要時 `MasterTemplateCatalogService` | intent 次第 | `DocumentTemplateResolver` | `DocumentNamePromptService` からの直接利用 | `LookupIntent.ExecutionAllowMasterFallback` のような明示が必要。CASE cache leg は promotion-aware |
| `IMasterTemplateCatalogReader` | Kernel `雛形一覧` の read-only 参照を閉じ込める | `Workbook`, `key` または list 要求 | `MasterTemplateRecord` / list | `MasterTemplateCatalogService` | 不可 | `IDocumentTemplateLookupReader`, snapshot rebuild 系 | prompt UI の直接利用 | `SYSTEM_ROOT` 解決と master workbook read-only open を内包 |
| `ICaseSnapshotStorageReader` | CASE / Base の chunk storage を読み分ける | `Workbook`, `SnapshotStorageKind` | `SnapshotStorageReadResult` | DocProperty `TASKPANE_SNAPSHOT_CACHE_*`, `TASKPANE_BASE_*` | 不可 | `TaskPaneSnapshotCacheService`, `CaseTemplateSnapshotService`, 将来の snapshot reader | 実行判断、prompt caption 決定の正本扱い | raw snapshot text, count, version, compatibility を返す低レベル reader |
| `ITaskPaneSnapshotReader` | 表示用 `TaskPaneSnapshot` を読む | `Workbook`, `SnapshotReadIntent` | `TaskPaneSnapshot` または `TaskPaneBuildResult` 相当 | CASE cache, Base snapshot, 必要時 MasterList | 読み取り intent 次第 | `TaskPaneManager` 系 | 文書実行 lookup, prompt lookup | public API 化は後段。現行 builder は CASE cache 更新を伴うため read-only 化前提の整理が必要 |
| `ITemplateMetadataResolutionPolicy` | lookup policy を mode で固定する内部境界 | `LookupIntent` | policy object | 呼び出し側 intent | 可否を policy で固定 | `IDocumentTemplateLookupReader` 内部 | UI や app service からの直接利用 | 共通化しても caller が `allowFallback=true` のように直接指定しないようにする |

### 4.1 候補 API の責務境界

- `ICaseCacheDocumentTemplateReader`
  - CASE cache hit / miss を返すだけに留める。
  - master fallback を入れない。
- `IDocumentTemplateLookupReader`
  - `DocumentTemplateResolver` の前段として metadata を解決する。
  - `TemplatePath` 解決や Word template 拡張子判定は持たない。
- `IMasterTemplateCatalogReader`
  - Master workbook の open / close / hidden read-only を吸収する。
  - 正本参照だが、呼び出し側に workbook open を漏らさない。
- `ICaseSnapshotStorageReader`
  - snapshot text と version provenance を返す低レベル reader に限定する。
  - UI 用 `TaskPaneSnapshot` 変換までは持たない。
- `ITaskPaneSnapshotReader`
  - UI 表示用 snapshot 専用とし、文書 template metadata の正本 API と混同しない。

## 5. prompt と文書実行の差分

### 5.1 固定仕様

- `DocumentNamePromptService` は CASE cache hit 時だけ caption を prompt 初期値に使う。
- `DocumentNamePromptService` は cache miss 時に master fallback しない。
- 文書実行は `DocumentTemplateResolver` 経由で CASE cache 優先・master fallback ありである。
- この 2 つを単一 API に無理に統合しない。

### 5.2 統合する場合の条件

統合が必要でも、公開面は少なくとも次のどちらかにする。

1. prompt 用 facade と実行用 facade を分け、内部で policy を共有する。
2. 単一 reader にする場合でも `LookupIntent.PromptCacheOnly` / `LookupIntent.ExecutionAllowMasterFallback` のように caller intent を必須入力にする。

次のような設計は不可とする。

- `allowMasterFallback: bool` を呼び出し側が自由に指定するだけの API
- prompt 呼び出しが誤って master fallback あり mode を通れる API
- snapshot reader をそのまま文書実行 lookup に使う API

## 6. 統一してよい候補 / まだ統一してはいけない候補

| 項目 | 統一可否 | 理由 | 統一条件 | 先に必要なテスト |
| --- | --- | --- | --- | --- |
| CASE cache の `key -> caption/file` 読取入口 | 統一可 | `TaskPaneSnapshotCacheService` が実質 single reader になっている | prompt と実行で facade を分ける | CASE cache hit / miss, caption-key-file 一致 |
| Master `雛形一覧` の read-only 参照 | 統一可 | `MasterTemplateCatalogService` が単独責務を持つ | open/read-only/close と cache invalidate を維持 | master fallback あり, master version stale |
| CASE/Base snapshot chunk 読取 | 統一可 | DocProperty 読取が複数箇所に散っている | raw storage 読取だけに限定する | snapshot format version 非互換, Base 初期配布, CASE cache clear 後再構築 |
| prompt 用 lookup と文書実行用 lookup | 統一不可 | fallback policy が異なる | facade 分離または intent 強制 | DocumentNamePromptService cache-only, master fallback 禁止 |
| `TaskPaneSnapshotBuilderService` と文書 template lookup | 統一不可 | snapshot builder は表示用かつ CASE cache 更新副作用がある | read-only 版 snapshot reader を別に設計できた後 | CASE cache stale, Base fallback, MasterList rebuild |
| Base snapshot 参照と文書実行 lookup | 統一不可 | Base snapshot は初期配布 / 表示補助であり実行正本ではない | 実行判断で Base snapshot を使わないと明文化 | Base 由来 snapshot 初期配布, 実行時 master 正本確認 |
| `TemplatePath` 解決と metadata lookup | 統一不可 | path 解決は `WORD_TEMPLATE_DIR` / `SYSTEM_ROOT` に依存し、metadata lookup と責務が異なる | `DocumentTemplateResolver` の境界維持 | caption/key/file 一致, path 解決一致 |

## 7. 同値確認テスト案

詳細なケース表と停止条件は `docs/cache-api-equivalence-test-design.md` を正とする。本節は read-only API 候補設計に接続するための要約だけを残す。

将来 API を実装する前提で、最低限次を確認する。

| テスト観点 | 期待結果 |
| --- | --- |
| CASE cache hit | prompt 初期値と実行用 metadata が同じ `caption` / `file` を返す |
| CASE cache miss | prompt は空のまま、実行用 lookup だけが fallback 先を評価する |
| master fallback あり | `DocumentTemplateResolver` 相当の結果が現行と同じ `key` / `caption` / `file` になる |
| master fallback 禁止 | prompt 系 caller では master catalog に進まない |
| `caption` / `key` / `file` 一致 | 現行 `DocumentTemplateLookupServiceTests` と同値である |
| snapshot format version 非互換 | 互換性なし CASE cache / Base snapshot は clear され、互換な経路だけが残る |
| master version stale | 表示系は stale を見て rebuild 経路へ進む。一方で prompt / 実行 lookup は CASE cache hit 中に stale だけで master fallback へ切り替えない |
| Base 由来 snapshot 初期配布 | 新規 CASE 初期化後、Base snapshot が CASE cache に promote される |
| CASE cache clear 後の再構築 | `TASKPANE_SNAPSHOT_CACHE_COUNT=0` 後に Base または MasterList から期待通り戻る |
| `DocumentNamePromptService` の cache-only 仕様 | cache miss でも prompt 初期値は空、master fallback は発生しない |
| `TemplatePath` 解決の責務維持 | metadata reader 導入後も `WORD_TEMPLATE_DIR` 優先、なければ `SYSTEM_ROOT\\雛形` を使う |
| Base snapshot を正本化していないこと | 実行判断・保存判断では Base snapshot 由来値だけで確定しない |

### 7.1 既存回帰テストとの対応

既存の `DocumentTemplateLookupServiceTests` は次の固定点を既に持つ。

- CASE cache hit 時、resolver と prompt が同じ caption を使う。
- CASE cache miss 時、resolver は master fallback するが prompt は空のまま。
- key 不在時、resolver は null、prompt 初期値は空。

将来の API 実装では、まずこの 3 件を green のまま維持することを前提条件とする。

## 8. 推奨移行順序

### 第1段階: docs 設計のみ

- 正本 / cache / snapshot 境界を `template-metadata-inventory` と本書で固定する。
- prompt と文書実行の policy 差分を docs で明文化する。
- 実装変更はしない。

### 第2段階: read-only API 追加

- `ICaseCacheDocumentTemplateReader`
  - 現行 `main` では導入済み。`DocumentNamePromptService` はこの cache-only reader に依存し、CASE cache miss 時に master fallback しない。
- `IDocumentTemplateLookupReader`
  - 現行 `main` では導入済み。`DocumentTemplateResolver` はこの reader に依存し、`DocumentTemplateLookupService.TryResolveWithMasterFallback(...)` 経由で CASE cache 優先 + master fallback を維持する。
- `DocumentTemplateLookupService`
  - 現行 `main` では `ICaseCacheDocumentTemplateReader` と `IDocumentTemplateLookupReader` の両方を実装し、prompt 用 cache-only lookup と文書実行用 fallback lookup を分けている。
- `IMasterTemplateCatalogReader`
- `ICaseSnapshotStorageReader`

現行 `main` で導入済みの reader は consumer 依存の分離に留め、`TaskPaneSnapshotCacheService` の promote / clear / compatibility 判定や `TaskPaneSnapshotBuilderService` の rebuild / Base fallback / CASE cache 更新は完了扱いにしない。

### 第3段階: 既存処理との同値確認

- `DocumentTemplateLookupService`
- `DocumentNamePromptService`
- `DocumentTemplateResolver`
- `TaskPaneSnapshotBuilderService`

上記と新 API の返り値・fallback 経路・stale 判定が一致することをテストで確認する。

### 第4段階: 低リスク箇所からの参照差し替え

優先度は次の順とする。

1. `DocumentNamePromptService` と `DocumentTemplateResolver` の consumer 依存差し替え（現行 `main` で実施済み）
2. `DocumentTemplateLookupService` 内部での reader 利用
3. CASE cache の raw 読取集約
4. `MasterTemplateCatalogService` を直接使う read-only 参照の集約

`TaskPaneSnapshotBuilderService` の差し替えは後回しにする。表示系は CASE cache 更新、副作用つき Base fallback、dynamic special button override を含むため、低リスクとは扱わない。

### 第5段階: 旧経路削除の判断条件

次を満たすまで旧経路削除は行わない。

- prompt cache-only 仕様の回帰が 0 件
- master fallback あり / なしの両系統で同値確認が完了
- snapshot format 非互換と stale version の挙動差分が 0 件
- Base snapshot promote と CASE cache clear 後再構築の挙動差分が 0 件
- `DocumentTemplateResolver` の `TemplatePath` 解決責務が変わっていない
- TaskPane 表示系で CASE cache 更新タイミングの差分が出ていない

## 9. 未確認事項

- `TaskPaneSnapshot` の `META` 行に含まれる workbook name / path / preferred width を read-only API の正式戻り値にどこまで含めるべきかは未確認。
- `TaskPaneSnapshotBuilderService` を完全 read-only reader に分離する場合、現行の CASE cache 保存副作用を別サービスへどこまで移すべきかは要確認。
- `DocumentTemplateLookupResult` と `MasterTemplateRecord` のどちらを API 契約の共通戻り値に寄せるべきかは未確認。
