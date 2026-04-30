# Readonly Snapshot Refactor Checkpoint

## 1. この checkpoint の対象

- 作成日: 2026-04-30
- 基準 commit: `ab06a64e77cd75a59c5ca2be115d26c71541e40e`
- 範囲: docs-only
- 参照 docs:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/readonly-api-adoption-status.md`
  - `docs/template-metadata-read-path-inventory.md`
- 参照 code:
  - `MasterTemplateCatalogService`
  - `TaskPaneSnapshotBuilderService`
  - `DocumentTemplateLookupService`
  - `DocumentNamePromptService`
  - `DocumentTemplateResolver`
  - `TaskPaneSnapshotCacheService`
  - `CaseTemplateSnapshotService`
  - `TaskPaneSnapshotChunkReadHelper`
  - `TaskPaneSnapshotChunkStorageHelper`
- 参照 tests:
  - `DocumentTemplateLookupServiceTests`
  - `TaskPaneSnapshotCacheStorageBehaviorTests`

## 2. 今回固定する到達点

- `MasterTemplateCatalogService` は `IMasterTemplateSheetReader` 経由で Master `雛形一覧` を read-only に読む構成で固定する。
- `TaskPaneSnapshotBuilderService` は `IMasterTemplateSheetReader` 経由で Master sheet 読取を行い、`CASE cache -> Base snapshot -> Master rebuild` の表示用 snapshot 解決順を維持する。
- `DocumentTemplateLookupService` は prompt 用 cache-only lookup と resolver 用 master fallback lookup の調停点として扱う。
- `DocumentNamePromptService` は CASE cache から `DocumentName` を引けた場合だけ prompt 初期値に使い、master fallback しない cache-only policy で固定する。
- `DocumentTemplateResolver` は `DocumentTemplateLookupService.TryResolveWithMasterFallback` を通じて CASE cache 優先 + master fallback で解決し、`TemplatePath` 導出責務を持つ構成で固定する。
- metadata read path は、正本を Master `雛形一覧`、派生 cache を Base snapshot / CASE cache、表示用断面を TaskPane snapshot として区分する。
- `TaskPaneSnapshotCacheService` には lookup 時 promote、compatibility clear、CASE cache lookup の責務が残る。
- `TaskPaneSnapshotChunkReadHelper` は raw chunk read、`TaskPaneSnapshotChunkStorageHelper` は raw chunk write / clear の shared primitive として導入済みと扱う。
- helper 境界として、promote 条件、compatibility 判定、stale 判定、Master rebuild、UI 制御を helper へ持ち込まない前提を docs に反映済みとする。

## 3. 現在の責務境界

- master sheet read の入口:
  - `MasterTemplateCatalogService` と `TaskPaneSnapshotBuilderService` が `IMasterTemplateSheetReader` を使って Master `雛形一覧` を読む。
- key -> metadata lookup の入口:
  - `DocumentTemplateLookupService` が `TaskPaneSnapshotCacheService` と `MasterTemplateCatalogService` を束ねる。
- prompt 用 lookup:
  - `DocumentNamePromptService` -> `ICaseCacheDocumentTemplateReader` -> `DocumentTemplateLookupService.TryResolveFromCaseCache`
- resolver 用 lookup:
  - `DocumentTemplateResolver` -> `IDocumentTemplateLookupReader` -> `DocumentTemplateLookupService.TryResolveWithMasterFallback`
- CASE cache / Base snapshot の位置づけ:
  - CASE cache は表示中 CASE と整合する派生 cache。
  - Base snapshot は新規 CASE 初期状態の配布用 cache。
  - どちらも正本ではない。
- read helper の責務:
  - `TaskPaneSnapshotChunkReadHelper` は `COUNT` と `XX` chunk を読んで raw snapshot text を連結する。
- storage helper の責務:
  - `TaskPaneSnapshotChunkStorageHelper` は chunk 分割保存、count 更新、余剰 chunk 空文字化、clear を行う。
- helper が持たない責務:
  - promote 判断
  - compatibility 判定
  - stale 判定
  - Master rebuild
  - UI 表示制御

## 4. 現在守られているテスト固定点

- `DocumentTemplateLookupServiceTests` が固定していること:
  - CASE cache hit 時、resolver と prompt が同じ `DocumentName` / `TemplateFileName` 系 metadata を使うこと
  - CASE cache miss 時、resolver だけが master fallback すること
  - prompt 側は cache-only で、master fallback しないこと
  - `WORD_TEMPLATE_DIR` 未設定時、resolver が `SYSTEM_ROOT\雛形` へ `TemplatePath` fallback すること
- `TaskPaneSnapshotCacheStorageBehaviorTests` が固定していること:
  - `TaskPaneSnapshotCacheService.PromoteBaseSnapshotToCaseCacheIfNeeded` の promote 条件
  - CASE cache 非互換時の CASE clear
  - Base snapshot 非互換時の Base clear と、有効 CASE cache を残す lookup promote 挙動
  - `CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache` が lookup promote と異なり初期化時上書きを行いうること
  - `CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache` で Base 非互換時に Base / CASE の両 snapshot を clear すること
- build / test の前回成功状況:
  - `docs/readonly-api-adoption-status.md` には、2026-04-30 時点の read-only API adoption 作業ブランチで `.\build.ps1 -Mode Compile` 成功、`.\build.ps1 -Mode Test` 成功、`CaseInfoSystem.SnapshotRegressionTests` 8 件成功の記録がある。
  - この checkpoint 作成では build / test を再実行していない。

## 5. 今後触ってよい候補

1. `TaskPaneSnapshotChunkReadHelper` / `TaskPaneSnapshotChunkStorageHelper` の単体テスト追加
2. raw chunk I/O 契約の補強
3. `TaskPaneSnapshotCacheService` / `CaseTemplateSnapshotService` の小さな重複整理
4. `TaskPaneSnapshotBuilderService` の整理
5. `DocumentNamePromptService` / `DocumentTemplateResolver` のさらなる整理

補足:

- 上ほど安全度が高い。
- 下へ行くほど表示フロー、stale 判定、fallback policy に近づくため変更面が広い。

## 6. 今は触らない方がよい箇所

- promote 条件
- compatibility 判定
- stale 判定
- Master rebuild 条件
- UI 表示制御
- `WorkbookActivate` / `WindowActivate` 前提の Pane 再利用
- `KernelTemplateSyncService` の version / Base 書込
- `DocumentNamePromptService` の cache-only policy
- `DocumentTemplateResolver` の master fallback

## 7. 次作業に入る前の注意

- main の基準点確認を必ず行う。
- 未コミット変更があれば停止する。
- Deploy / manifest / `.vsto` 差分を commit しない。
- build 成功と実機確認成功を分けて扱う。
- docs-only / test-only / production code 変更を混ぜない。
- 変更範囲を 1 サービスまたは 1 責務に絞る。
- 迷ったら実装せず docs 調査で止める。

## 8. 未確認事項

- 実機確認結果はリポジトリ内で確認できず、不明。
- `TaskPaneSnapshotChunkReadHelper` / `TaskPaneSnapshotChunkStorageHelper` の direct unit test は確認できていない。
- helper の契約補強は候補だが、現時点では docs 固定のみで未着手。
