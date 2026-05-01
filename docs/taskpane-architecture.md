# TaskPane Architecture

## 位置づけ

この文書は、TaskPane リファクタリング後の現行設計を固定するための正本です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- 優先度A現在地の補足: `docs/taskpane-refactor-current-state.md`
- metadata / cache / snapshot の補足: `docs/template-metadata-inventory.md`
- 読取経路の補足: `docs/template-metadata-read-path-inventory.md`

この文書で扱う対象は、CASE 向け TaskPane の表示、TaskPane 用 metadata の正本と派生 cache、文書ボタン実行時の解決責務です。

## 結論

TaskPane 設計の現行正本は、次の整理で固定します。

1. 文書ボタン定義の runtime 正本は Kernel `雛形一覧` と Kernel `TASKPANE_MASTER_VERSION` です。
2. Base 埋込 snapshot と CASE snapshot cache は、どちらも正本ではなく派生 cache です。
3. TaskPane の表示は `CASE cache -> Base cache -> Master rebuild` の順で解決します。
4. 表示中の CASE は、後から成功した雛形登録・更新に自動追随しません。
5. 文書名入力は CASE cache にのみ従い、実行時 template 解決だけが master fallback を持ちます。
6. UI 制御は専用サービス経由で行い、`WorkbookOpen` 直後に直接表示制御しません。

## 正本と派生情報

### 正本

- Kernel `雛形一覧`
  - A列: `key`
  - B列: `TemplateFileName`
  - C列: `caption`
  - D列: 文書ボタン色
  - E列: タブ名
  - F列: タブ色
- Kernel `TASKPANE_MASTER_VERSION`

### 派生 cache

- Base 埋込 snapshot
  - `TASKPANE_BASE_SNAPSHOT_COUNT`
  - `TASKPANE_BASE_SNAPSHOT_XX`
  - `TASKPANE_BASE_MASTER_VERSION`
  - Base 側 `TASKPANE_MASTER_VERSION`
- CASE snapshot cache
  - `TASKPANE_SNAPSHOT_CACHE_COUNT`
  - `TASKPANE_SNAPSHOT_CACHE_XX`
  - CASE 側 `TASKPANE_MASTER_VERSION`

### 扱いの原則

- snapshot / cache は表示補助です。
- snapshot / cache を保存・生成・実行判断の正本にしてはいけません。
- `TemplatePath` は保存正本を持たず、`DocumentTemplateResolver` が都度導出します。

## 主要責務

### `KernelTemplateSyncService`

- `SYSTEM_ROOT\雛形` と `CaseList_FieldInventory` を前提に雛形登録・更新を行う
- Kernel `雛形一覧` を更新する
- `TASKPANE_MASTER_VERSION` を成功時に無条件で `+1` する
- Base 用 snapshot と version を再生成して埋め込む

### `TaskPaneSnapshotBuilderService`

- CASE 表示時の snapshot 解決元を選ぶ
- 優先順は `CASE cache -> Base cache -> Master rebuild`
- 必要時に CASE cache を更新する

### `TaskPaneSnapshotCacheService`

- 文書ボタン実行時の CASE cache lookup を担う
- Base snapshot の on-demand promote を担う
- cache 互換性不一致時の clear を担う

### `TaskPaneManager`

- snapshot を UI 表示用 state に変換して描画する
- 文書ボタン押下を `DocumentNamePromptService` と `DocumentCommandService` へ接続する

### `TaskPaneRefreshOrchestrationService` / `WindowActivatePaneHandlingService`

- TaskPane 再描画要求、遅延表示、Window 単位の表示調停を担う
- host 再利用と再表示の方針を維持する

### `DocumentNamePromptService`

- CASE cache から `caption` を引けた場合だけ prompt 初期値へ使う
- master fallback は行わない

### `DocumentTemplateResolver`

- CASE cache 優先で `key -> DocumentName / TemplateFileName` を解決する
- CASE cache miss 時だけ master catalog に fallback する
- `WORD_TEMPLATE_DIR` または `SYSTEM_ROOT\雛形` から `TemplatePath` を導出する

## フロー固定

### 1. 雛形登録・更新成功時

1. `KernelTemplateSyncService` が Kernel `雛形一覧` を更新する
2. `TASKPANE_MASTER_VERSION` を `+1` する
3. Base 用 snapshot を再生成する
4. Base に `TASKPANE_BASE_SNAPSHOT_*`、`TASKPANE_BASE_MASTER_VERSION`、`TASKPANE_MASTER_VERSION` を保存する
5. master catalog cache を無効化する

注意:

- version 更新は内容差分比較ではなく、成功時無条件 `+1` です
- この方針は変えません

### 2. 新規 CASE 作成時

1. Base を物理コピーして CASE を作成する
2. Base 埋込 snapshot / version を新規 CASE が引き継ぐ
3. `CaseTemplateSnapshotService` が初期 promote を行う
4. 新規 CASE は原則として最新 snapshot を持った状態で開始する

### 3. 既存 CASE 表示時

1. `TaskPaneSnapshotBuilderService` が CASE cache を確認する
2. CASE cache が使えなければ Base snapshot を確認する
3. どちらも使えない場合だけ Master から rebuild する
4. いったん生成した Pane / host / control は、CASE を閉じるまで維持する

注意:

- `WorkbookActivate` / `WindowActivate` のたびに version 比較して Pane を作り直す設計ではありません
- 開いている CASE が最新雛形へ自動追随しないことは現行仕様です

### 4. 文書ボタン押下時

1. `TaskPaneManager` がボタン押下を受ける
2. `DocumentNamePromptService` が CASE cache の `caption` だけを使って prompt 初期値を準備する
3. `DocumentCommandService` が実行に進む
4. `DocumentExecutionEligibilityService` が `DocumentTemplateResolver` を通して実行前確認を行う
5. `DocumentTemplateResolver` は CASE cache 優先、miss 時のみ master fallback で `TemplateFileName` を解決する
6. `DocumentCreateService` が文書生成する

責務分離:

- prompt 初期値は表示中 CASE に整合する補助情報
- 実行時解決は保存・生成判断に必要な正本確認入口

### 5. 案件一覧登録後

1. CASE 側 `TASKPANE_SNAPSHOT_CACHE_COUNT` を `0` に戻す
2. `TASKPANE_SNAPSHOT_CACHE_XX` を削除する
3. Base 側 `TASKPANE_BASE_*` には触れない

## UI 制御方針

- TaskPane は左ドック固定です
- 表示制御は専用サービス経由です
- `WorkbookOpen` 直後に直接 UI 表示制御を追加しません
- 遅延表示、一時抑止、Window 単位の再利用を前提にします
- `ScreenUpdating` を変更した場合は必ず復元します

## 触ってはいけない固定点

- `TASKPANE_MASTER_VERSION` の成功時無条件 `+1`
- Base snapshot 埋め込み
- CASE cache 優先の表示・lookup
- `DocumentNamePromptService` の cache-only policy
- `DocumentTemplateResolver` の CASE cache 優先 + master fallback
- `WorkbookActivate` / `WindowActivate` の host 再利用経路
- snapshot / cache を正本扱いする変更
- `WorkbookOpen` 直接依存の表示制御

## 不明として残す事項

- Pane 再利用判定の全条件
- retry 間隔や protection 秒数の正式な仕様根拠
- 実機観測を伴うちらつきや ready-show の最終挙動詳細

コードだけで確定できない事項は、この文書でも断定しません。
