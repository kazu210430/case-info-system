# Cache API Equivalence Test Design

## 1. 目的

この文書は、`docs/template-metadata-inventory.md` と `docs/cache-api-readonly-design.md` を前提に、将来 read-only cache API を実装する前に必要な「既存経路との同値確認テスト設計」を整理するための設計メモである。

今回の対象は docs 整理のみであり、コード変更・テスト実装は行わない。

## 2. 前提と参照

- 参照 docs
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/template-metadata-inventory.md`
  - `docs/cache-api-readonly-design.md`
- 調査対象サービス
  - `DocumentNamePromptService`
  - `DocumentTemplateLookupService`
  - `DocumentTemplateResolver`
  - `TaskPaneSnapshotCacheService`
  - `TaskPaneSnapshotBuilderService`
  - `CaseTemplateSnapshotService`
  - `MasterTemplateCatalogService`
  - `KernelTemplateSyncService`
  - `CaseWorkbookInitializer`

## 3. 対象フローの要約

### 3.1 prompt 用 lookup

1. `DocumentNamePromptService` が `DocumentTemplateLookupService.TryResolveFromCaseCache` を呼ぶ。
2. `TaskPaneSnapshotCacheService` が CASE cache を参照する。
3. CASE cache が空の場合でも、互換な Base snapshot があれば CASE cache へ昇格してから参照する。
4. `caption` を取得できた場合だけ prompt 初期値に使う。
5. CASE cache / 昇格後 CASE cache で解決できない場合、master fallback は行わない。

### 3.2 文書実行用 lookup

1. `DocumentExecutionEligibilityService` が `DocumentTemplateResolver.Resolve` を呼ぶ。
2. `DocumentTemplateResolver` が `DocumentTemplateLookupService.TryResolveWithMasterFallback` を呼ぶ。
3. 先に CASE cache を参照する。
4. CASE cache miss 時だけ `MasterTemplateCatalogService` に fallback する。
5. 解決した `TemplateFileName` を使って、`DocumentTemplateResolver` が別責務として `TemplatePath` を導出する。

### 3.3 Base / CASE / snapshot

1. `KernelTemplateSyncService` が Kernel `雛形一覧` を更新し、Base snapshot と `TASKPANE_MASTER_VERSION` を更新する。
2. `CaseWorkbookInitializer` / `CaseTemplateSnapshotService` が新規 CASE に Base snapshot を昇格する。
3. `TaskPaneSnapshotBuilderService` は `CASE cache -> Base snapshot -> MasterList rebuild` の順で表示用 snapshot を解決する。
4. `TaskPaneSnapshot` は表示用断面であり、保存・生成・実行判断の正本ではない。

## 4. 同値確認の基本方針

- 新 API 候補は、既存経路と同じ結果を返すことを確認してから差し替える。
- prompt 用 lookup と文書実行用 lookup は別仕様として確認する。
- cache miss 時の挙動は caller intent ごとに確認する。
- fallback 可否はテスト名と期待値に明示する。
- snapshot は表示用断面であり、正本ではないことを確認する。
- TemplatePath 解決は metadata lookup とは別責務として確認する。
- Base snapshot は新規 CASE への初期配布用であり、metadata lookup の正本ではないことを確認する。
- CASE cache hit 中は、master version が新しくても prompt / 実行 lookup が自動で master fallback へ切り替わらないことを、現行コード準拠の期待値として扱う。

### 4.1 安全と言える条件

「既存経路と新 API 候補で何が同じなら安全と言えるか」は、少なくとも次を満たすこととする。

- caller intent ごとに同じ解決元を使うこと
- 同じ `key / caption / file` を返すこと
- fallback の可否が同じこと
- CASE cache hit / miss 時の結果が同じこと
- Base snapshot / CASE snapshot / TaskPane snapshot の責務境界を崩さないこと
- `TemplatePath` 解決を metadata lookup に混ぜないこと
- snapshot / cache を保存・生成・実行判断の正本扱いしないこと

## 5. テストケース表

| No | 観点 | 既存経路 | 新API候補 | 入力条件 | 期待値 | fallback | 確認対象 | 備考 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 01 | prompt: CASE cache hit | `DocumentNamePromptService -> DocumentTemplateLookupService.TryResolveFromCaseCache` | `ICaseCacheDocumentTemplateReader` | CASE cache に `key=01` の `caption/file` がある | prompt 初期値に `caption` を使う | 禁止 | `initialDocumentName`、解決元 | prompt 契約は `caption` 中心 |
| 02 | prompt: CASE cache miss | 同上 | 同上 | CASE cache なし、master には該当 key あり | prompt 初期値は空のまま | 禁止 | 空文字、master 未参照 | 文書実行 lookup と同じ期待値にしない |
| 03 | prompt: Base snapshot 初期配布経由 | 同上 | `ICaseCacheDocumentTemplateReader` + `ICaseSnapshotStorageReader` | CASE cache 空、互換な Base snapshot あり | Base snapshot が CASE cache に昇格した後、prompt 初期値に `caption` を使う | 禁止 | CASE cache 昇格後の lookup 結果 | Base を直接正本扱いしない |
| 04 | prompt: cache-only guard | `DocumentNamePromptService` | `ICaseCacheDocumentTemplateReader` | CASE cache miss、master hit | master fallback しない | 禁止 | 呼出経路、解決元 | テスト名に `CacheOnly` を明示 |
| 05 | 実行: CASE cache hit | `DocumentTemplateResolver -> TryResolveWithMasterFallback` | `IDocumentTemplateLookupReader` | CASE cache に `key/caption/file` あり | CASE cache を優先して同じ `key/caption/file` を返す | 許可だが未使用 | `DocumentTemplateLookupResult`、`ResolutionSource` | master に同 key があっても CASE cache 優先 |
| 06 | 実行: CASE cache miss + master fallback | 同上 | 同上 | CASE cache miss、master hit | master fallback で同じ `key/caption/file` と template record を返す | 許可 | `DocumentTemplateLookupResult`、`MasterTemplateRecord` | `TemplatePath` は別テストで確認 |
| 07 | 実行: key 不在 | 同上 | 同上 | CASE cache miss、master miss | `null` / 未解決で一致する | 許可 | 未解決時の戻り値 | prompt 側は空欄開始を別途確認 |
| 08 | `caption / key / file` 一致 | CASE cache 経路 / master fallback 経路 | `ICaseCacheDocumentTemplateReader` / `IDocumentTemplateLookupReader` | 同じ key を CASE cache と master の両方に用意 | caller intent ごとに既存経路と同じ `key/caption/file` | caller intent 依存 | 正規化後 key、caption、file | prompt では file を返却契約にしない |
| 09 | `DocumentNamePromptService` の cache-only 仕様 | `DocumentNamePromptService` | `ICaseCacheDocumentTemplateReader` | CASE cache miss、Base snapshot も不使用、master hit | prompt 初期値は空、master lookup は走らない | 禁止 | prompt 初期値、参照経路 | 確定済み仕様 |
| 10 | `DocumentTemplateResolver` の CASE cache 優先・master fallback 仕様 | `DocumentTemplateResolver` | `IDocumentTemplateLookupReader` | hit / miss の両条件 | hit では CASE cache 優先、miss でのみ master fallback | 許可 | 解決元、戻り値 | stale だけで fallback しない |
| 11 | `TemplatePath` 解決責務の維持 | `DocumentTemplateResolver.ResolveTemplateDirectory` | `IDocumentTemplateLookupReader` + 既存 resolver | `WORD_TEMPLATE_DIR` あり / なしの両条件 | metadata lookup は `TemplateFileName` まで、`TemplatePath` は resolver が導出 | 該当なし | `TemplatePath`、`TemplateFileName` | `WORD_TEMPLATE_DIR` 優先、なければ `SYSTEM_ROOT\\雛形` |
| 12 | snapshot format version 非互換 | `TaskPaneSnapshotCacheService` / `TaskPaneSnapshotBuilderService` | `ICaseSnapshotStorageReader` | CASE cache または Base snapshot が非互換形式 | 非互換 snapshot は clear され、互換な経路だけを残す | caller intent 依存 | clear 動作、次の参照先 | 非互換 snapshot を正本扱いしない |
| 13 | master version stale | lookup 系 + `TaskPaneSnapshotBuilderService` | `IDocumentTemplateLookupReader` / `ITaskPaneSnapshotReader` | CASE cache hit、Kernel master version が新しい | 表示系は rebuild / Base fallback を評価、prompt / 実行 lookup は CASE cache hit を維持 | prompt 禁止 / 実行 許可だが miss 時のみ | 経路別の差分 | 開いている CASE が自動追随しない現行仕様を維持 |
| 14 | Base 由来 snapshot 初期配布 | `CaseWorkbookInitializer` + `CaseTemplateSnapshotService` | `ICaseSnapshotStorageReader` | 新規 CASE 作成直後 | Base snapshot が CASE cache に複写され、CASE master version が同期される | なし | `TASKPANE_SNAPSHOT_CACHE_*`、`TASKPANE_MASTER_VERSION` | 初期配布用であり正本ではない |
| 15 | CASE cache clear 後の再構築 | `TaskPaneSnapshotCacheService.ClearCaseSnapshotCacheChunks` + `TaskPaneSnapshotBuilderService` | `ICaseSnapshotStorageReader` / `ITaskPaneSnapshotReader` | `TASKPANE_SNAPSHOT_CACHE_COUNT=0`、Base 互換あり / なし | 互換 Base があれば CASE cache 復元、なければ MasterList rebuild | caller intent 依存 | 復元元、再構築結果 | build 前提ではなく設計上の期待値 |
| 16 | TaskPane 表示 snapshot と lookup の責務分離 | `TaskPaneSnapshotBuilderService` / lookup 系 | `ITaskPaneSnapshotReader` / template readers | 表示更新と文書 lookup を同一 CASE で実行 | snapshot は表示用情報、lookup は `key/caption/file` 解決のみ | 該当なし | 返却モデル、責務境界 | `SPECIAL/TAB/DOC` と metadata lookup を混ぜない |
| 17 | Base snapshot を正本化していないこと | Base snapshot 参照経路全般 | `ICaseSnapshotStorageReader` | Base snapshot に旧値、master に新値 | 保存・生成・実行判断は Base snapshot だけで確定しない | 該当なし | 正本確認箇所 | Base は初期配布 / 表示補助 |

## 6. prompt 用 lookup の同値確認

### 6.1 分けて確認すること

- CASE cache hit 時
  - `caption` を prompt 初期値に使う。
- CASE cache miss 時
  - master fallback しない。
  - prompt 初期値を無理に補完しない。
- Base snapshot 初期配布時
  - Base snapshot が CASE cache に昇格した後だけ prompt に使われる。
  - Base snapshot を prompt lookup の正本扱いにはしない。

### 6.2 prompt で必要な確認範囲

- 必須
  - `caption`
  - CASE cache hit / miss の差
  - master fallback 禁止
- 補助的に確認するもの
  - key 正規化が CASE cache lookup と一致すること
- prompt 契約に含めないもの
  - `TemplatePath`
  - 実行用 `template record`
  - master fallback 結果

### 6.3 文書実行用 lookup と同じ期待値にしないこと

- prompt は `caption` を UI 初期値に使う補助であり、実体テンプレート解決の正本ではない。
- 実行系が `file` と `TemplatePath` を必要とすることを、そのまま prompt の同値条件に持ち込まない。
- prompt の安全条件は「caption が一致すること」だけでなく、「cache miss で補完しないこと」も含む。

## 7. 文書実行用 lookup の同値確認

### 7.1 分けて確認すること

- CASE cache hit 時
  - CASE cache を優先する。
  - master version が新しくても、CASE cache hit 中は stale だけを理由に fallback しない。
- CASE cache miss 時
  - master fallback 可。
- fallback 後に確認すべき値
  - `key`
  - `caption`
  - `file`
  - `template record`

### 7.2 `TemplatePath` 解決は別責務

- metadata lookup の同値確認は `TemplateFileName` までで止める。
- `TemplatePath` は `DocumentTemplateResolver` が:
  - `WORD_TEMPLATE_DIR`
  - なければ `SYSTEM_ROOT\\雛形`
  の順で導出する。
- read-only cache API 候補が path 解決まで引き受ける設計は不可とする。

## 8. Base / CASE / snapshot の同値確認

- Base snapshot は新規 CASE への初期配布用である。
- CASE snapshot は実行時・表示中 Pane と整合する cache である。
- `TaskPaneSnapshot` は表示用断面であり、保存・生成・実行判断の正本ではない。
- snapshot format version 非互換時は、非互換 cache を clear し、互換な経路だけを残す。
- master version stale 時は、表示系の `TaskPaneSnapshotBuilderService` と lookup 系で期待値を混同しない。
- CASE cache clear 後は、互換 Base があればそこから再構築し、なければ MasterList rebuild を期待値とする。

### 8.1 Base snapshot / CASE snapshot の境界

- Base snapshot が持つ責務
  - 新規 CASE 初期配布
  - CASE cache 再構築時の補助
- Base snapshot が持たない責務
  - 保存判断の正本
  - 文書実行判断の正本
  - `TemplatePath` 解決
- CASE snapshot が持つ責務
  - 表示中 Pane と整合する `key/caption/file` の cache
- CASE snapshot が持たない責務
  - global master version の正本
  - すべての CASE に対する共通 metadata の決定

## 9. 将来のテスト実装優先順位

実装はしない。将来テスト実装する場合の推奨順だけを固定する。

1. `DocumentNamePromptService` の cache-only 仕様
2. `DocumentTemplateResolver` の fallback 仕様
3. `DocumentTemplateLookupService` の CASE cache 優先仕様
4. `MasterTemplateCatalogService` の read-only 参照
5. CASE cache clear / 再構築
6. Base snapshot 初期配布
7. `TaskPaneSnapshotBuilderService` 関連は後回し

## 10. 実装に進む前の停止条件

- 期待値が確定できない場合は実装しない。
- prompt と文書実行の仕様差が曖昧な場合は実装しない。
- snapshot を正本扱いしそうな場合は実装しない。
- `TemplatePath` 解決責務が混ざる場合は実装しない。
- docs とコードが矛盾する場合はコードを優先し、docs に要確認として残す。
- CASE cache hit 中の stale master の扱いを caller intent ごとに分離できない場合は実装しない。

## 11. 未確認事項

- `TaskPaneSnapshot` の `META` 行に含まれる workbook 名 / path / preferred width を、将来の read-only API 戻り値にどこまで含めるべきかは未確認。
- `TaskPaneSnapshotBuilderService` の表示用 rebuild と、read-only snapshot reader の責務境界をどこで切るかは要確認。
