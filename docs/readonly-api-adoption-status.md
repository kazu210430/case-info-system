# read-only API adoption status

- 関連 checkpoint: `docs/readonly-snapshot-refactor-checkpoint.md`

## 1. 現在の状態（事実のみ）

### read-only API 導入済み箇所

- `DocumentTemplateResolver` 周り
  - `a92834c` で `DocumentTemplateLookupService` に `IDocumentTemplateLookupReader` を追加した。
  - consumer 側は `DocumentTemplateResolver` から `DocumentTemplateLookupService` 直接依存を外し、`IDocumentTemplateLookupReader` 経由へ差し替えた。
- `DocumentNamePromptService` 周り
  - 現行 `main` では `DocumentNamePromptService` は `ICaseCacheDocumentTemplateReader` に依存している。
  - prompt 初期値は `DocumentTemplateLookupService.TryResolveFromCaseCache(...)` 経由の CASE cache-only lookup で取得する。
  - CASE cache miss 時に master catalog へ fallback しない契約は維持されている。
- `DocumentTemplateLookupService` 周り
  - 現行 `main` では `DocumentTemplateLookupService` が `ICaseCacheDocumentTemplateReader` と `IDocumentTemplateLookupReader` の両方を実装する。
  - `TryResolveFromCaseCache(...)` は CASE cache-only、`TryResolveWithMasterFallback(...)` は CASE cache 優先 + master fallback の実行側 lookup として分かれている。
- `TaskPaneManager` 周り
  - `91d0777` で `TaskPaneSnapshotBuilderService` に `ICaseTaskPaneSnapshotReader` を追加した。
  - consumer 側は `TaskPaneManager` から `TaskPaneSnapshotBuilderService` 直接依存を外し、`ICaseTaskPaneSnapshotReader` 経由へ差し替えた。
- `TaskPaneSnapshotBuilderService` 周り
  - 2026-04-30 時点の作業で `IMasterTemplateSheetReader` を constructor 注入した。
  - `MasterList rebuild` と pane 幅計算で使う master sheet 読み取りを、`MasterTemplateSheetReader` 直接呼び出しから既存 adapter 経由へ差し替えた。

### 差し替え方針

- consumer 側のみ差し替える。
- 既存ロジックは変更しない。
  - `DocumentTemplateLookupService.TryResolveWithMasterFallback` の解決順は維持する。
  - `TaskPaneSnapshotBuilderService.BuildSnapshotText` の snapshot 解決順は維持する。

### テスト状態

- 2026-04-30 時点の read-only API adoption 作業ブランチで `.\build.ps1 -Mode Compile` は成功。
- 2026-04-30 時点の read-only API adoption 作業ブランチで `.\build.ps1 -Mode Test` は成功。
- 2026-04-30 時点の read-only API adoption 作業ブランチで `CaseInfoSystem.SnapshotRegressionTests` は 8 件すべて成功。

### 実機確認結果

- `TaskPane` 表示問題なし: 不明
- CASE 切替問題なし: 不明
- snapshot 更新問題なし: 不明
- 表示内容差分なし: 不明

リポジトリ内では、上記実機確認結果の記録までは確認できていない。

## 2. 設計ルール（重要）

- prompt は cache-only とし、master fallback を入れない。
- 文書実行は CASE cache 優先 + master fallback とする。
- `TemplatePath` は resolver 責務とする。
- snapshot は正本ではない。
- read-only API は参照経路の整理のみを目的とする。

## 3. 実装ルール

- consumer 側のみ差し替える。
- service / resolver の責務は変えない。
- fallback の意味を変えない。
- snapshot を正本化しない。

## 4. まだ完了扱いしないこと（重要）

- `TaskPaneSnapshotBuilderService` の snapshot rebuild / Base fallback / CASE cache 更新を pure read-only API へ寄せること。
- `TaskPaneSnapshotCacheService` の promote / clear / compatibility 判定を read-only lookup から完全分離すること。
- prompt / lookup 系 consumer 全件が `DocumentNamePromptService` と `DocumentTemplateResolver` だけで十分かの全量保証。
- `TemplatePath` 解決責務を `DocumentTemplateResolver` から外すこと。
- 実機での TaskPane 表示、CASE 切替、snapshot 更新、表示内容差分なしの確認。
- 大規模リファクタリング。

## 5. 今後の進め方

- low-risk な参照差し替えを小さく進める。
- 1〜2箇所単位で実施する。
- 毎回 build / test / 実機確認を分けて確認する。
- 問題なければ次へ進む。
