# read-only API adoption status

## 1. 現在の状態（事実のみ）

### read-only API 導入済み箇所

- `DocumentTemplateResolver` 周り
  - `a92834c` で `DocumentTemplateLookupService` に `IDocumentTemplateLookupReader` を追加した。
  - consumer 側は `DocumentTemplateResolver` から `DocumentTemplateLookupService` 直接依存を外し、`IDocumentTemplateLookupReader` 経由へ差し替えた。
- `TaskPaneManager` 周り
  - `91d0777` で `TaskPaneSnapshotBuilderService` に `ICaseTaskPaneSnapshotReader` を追加した。
  - consumer 側は `TaskPaneManager` から `TaskPaneSnapshotBuilderService` 直接依存を外し、`ICaseTaskPaneSnapshotReader` 経由へ差し替えた。

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

## 4. 今回やらなかったこと（重要）

- prompt 系の変更
- 文書実行系の変更
- `TemplatePath` 解決の変更
- snapshot 再構築ロジックの変更
- 新規テスト追加
- 大規模リファクタリング

## 5. 今後の進め方

- low-risk な参照差し替えを小さく進める。
- 1〜2箇所単位で実施する。
- 毎回 build / test / 実機確認を分けて確認する。
- 問題なければ次へ進む。
