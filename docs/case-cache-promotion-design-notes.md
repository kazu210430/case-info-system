# CASE Cache Promotion Design Notes

## 1. 目的

この文書は、CASE cache 参照経路に Base 埋込 snapshot からの昇格副作用が混在している現状を、コード変更前の設計メモとして固定するためのものです。

- 対象:
  - `ICaseCacheDocumentTemplateReader`
  - `IDocumentTemplateLookupReader`
  - `ICaseTaskPaneSnapshotReader`
  - `IMasterTemplateCatalogReader`
  - `DocumentTemplateLookupService`
  - `TaskPaneSnapshotCacheService`
  - `TaskPaneSnapshotBuilderService`
  - `MasterTemplateCatalogService`
  - `CaseTemplateSnapshotService`
- 前提:
  - [architecture.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\architecture.md)
  - [flows.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\flows.md)
  - [ui-policy.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\ui-policy.md)
  - [template-metadata-inventory.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\template-metadata-inventory.md)

この文書は現状整理と将来方針の固定を目的とし、実装変更は含みません。

## 2. 現状の問題整理

### 2.1 read-only と実際の副作用のズレ

- [ICaseCacheDocumentTemplateReader.cs](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\ICaseCacheDocumentTemplateReader.cs:7) は「CASE cache だけを読み取り、文書テンプレート metadata を返す read-only API」と説明している。
- [IDocumentTemplateLookupReader.cs](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\IDocumentTemplateLookupReader.cs:7) は「CASE cache を優先しつつ、必要時のみ master catalog へフォールバックする read-only 参照口」と説明している。
- しかし実装先の [DocumentTemplateLookupService.cs](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentTemplateLookupService.cs:22) は、[TaskPaneSnapshotCacheService.cs](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:87) を経由して lookup を行う。
- [TaskPaneSnapshotCacheService.TryGetDocumentTemplateLookupFromCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:87) は、lookup 前に [PromoteBaseSnapshotToCaseCacheIfNeeded](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:30) を実行する。
- [PromoteBaseSnapshotToCaseCacheIfNeeded](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:70) は CASE cache chunk と `TASKPANE_MASTER_VERSION` を DocProperty に書き戻す。

このため、read-only と説明されている interface の利用が、実際には CASE workbook の DocProperty 更新を伴う場合があります。

### 2.2 責務混在

- `lookup`
  - key から caption / template file name を解決する責務
- `promotion`
  - Base 埋込 snapshot を CASE cache へ昇格する責務

現状はこの 2 つが [TaskPaneSnapshotCacheService](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:12) に同居し、そのサービスを [DocumentTemplateLookupService](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentTemplateLookupService.cs:9) が read 系 interface の実装として公開しています。

### 2.3 類似責務の重複

- 新規 CASE 初期化では [CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\CaseTemplateSnapshotService.cs:32) が Base 埋込 snapshot を CASE cache へ昇格する。
- 表示中や lookup 時には [TaskPaneSnapshotCacheService.PromoteBaseSnapshotToCaseCacheIfNeeded](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:30) が同系統の昇格を行う。

Base から CASE への昇格責務は 1 箇所に集約されていません。

## 3. 影響範囲

### 3.1 prompt

- [DocumentNamePromptService](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentNamePromptService.cs:8) は [ICaseCacheDocumentTemplateReader](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\ICaseCacheDocumentTemplateReader.cs:9) に依存する。
- [FindDocumentCaptionByKey](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentNamePromptService.cs:64) は [TryResolveFromCaseCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentNamePromptService.cs:73) を呼ぶ。
- その実装は前述のとおり CASE cache 昇格を含み得る。

影響として確認できる事実:

- 文書名入力 prompt 用の caption 参照は、純粋 read とは限らない。
- prompt 準備中に CASE DocProperty が更新される経路が存在する。

### 3.2 document execution

- [DocumentTemplateResolver](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentTemplateResolver.cs:9) は [IDocumentTemplateLookupReader](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\IDocumentTemplateLookupReader.cs:9) に依存する。
- [Resolve](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentTemplateResolver.cs:32) は [TryResolveWithMasterFallback](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentTemplateResolver.cs:50) を呼ぶ。
- [DocumentExecutionEligibilityService](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentExecutionEligibilityService.cs:8) は [Evaluate](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentExecutionEligibilityService.cs:35) の中で:
  - [BuildEligibleCacheKey](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentExecutionEligibilityService.cs:149) により `TASKPANE_MASTER_VERSION` を含む cache key を作る
  - その後 [DocumentTemplateResolver.Resolve](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentExecutionEligibilityService.cs:71) を呼ぶ

影響として確認できる事実:

- document execution の template resolve は、CASE cache lookup の名目で CASE cache 昇格を伴い得る。
- `Evaluate` は resolve 前に cache key を作っているため、resolve 中の `TASKPANE_MASTER_VERSION` 更新が eligibility cache の扱いに関係し得る構造になっている。

### 3.3 TaskPane

- [ICaseTaskPaneSnapshotReader](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\ICaseTaskPaneSnapshotReader.cs:9) の実装は [TaskPaneSnapshotBuilderService](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:14)。
- [TaskPaneManager](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\TaskPaneManager.cs:16) は [BuildSnapshotText](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\TaskPaneManager.cs:741) を呼ぶ。
- [TaskPaneSnapshotBuilderService.BuildSnapshotText](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:86) は:
  - CASE cache 読取
  - Base cache 読取
  - Base cache を CASE cache へ保存
  - `TASKPANE_MASTER_VERSION` 更新
  - Master ブック open/close
  - MasterList rebuild
  を行う。

影響として確認できる事実:

- TaskPane 側 interface は read-only を名乗っていないが、snapshot build が CASE cache 更新と Master 読取を伴う。
- TaskPane 側は元から build / rebuild 責務として副作用を持つ構造であり、prompt / document execution の問題とは性質が異なる。

## 4. 現在の昇格発生箇所一覧

### 4.1 新規 CASE 初期化時

- [CaseWorkbookInitializer.InitializeCore](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\CaseWorkbookInitializer.cs:45)
- 呼び出し:
  - [SyncMasterVersionFromKernel](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\CaseWorkbookInitializer.cs:49)
  - [PromoteEmbeddedSnapshotToCaseCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\CaseWorkbookInitializer.cs:50)
- 実装:
  - [CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\CaseTemplateSnapshotService.cs:32)

### 4.2 CASE cache lookup 時

- [DocumentTemplateLookupService.TryResolveFromCaseCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\App\DocumentTemplateLookupService.cs:22)
- 実体:
  - [TaskPaneSnapshotCacheService.TryGetDocumentTemplateLookupFromCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:87)
- 事前呼出:
  - [PromoteBaseSnapshotToCaseCacheIfNeeded](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:96)

### 4.3 TaskPane snapshot build 時

- [TaskPaneSnapshotBuilderService.BuildSnapshotText](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:86)
- Base cache から CASE cache へ保存:
  - [SaveCaseSnapshotCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:130)
  - [SetDocumentProperty(TASKPANE_MASTER_VERSION)](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:132)
  - [SaveCaseSnapshotCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:139)
  - [SetDocumentProperty(TASKPANE_MASTER_VERSION)](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:141)
- Master rebuild 後の CASE cache 保存:
  - [SetDocumentProperty(TASKPANE_MASTER_VERSION)](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:159)
  - [SaveCaseSnapshotCache](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotBuilderService.cs:171)

## 5. 設計案

### 5.1 A案: 最小修正

内容:

- interface 名は維持する。
- `read-only` という説明を、実装実態に合わせて修正する。
- `CASE cache lookup は Base 埋込 snapshot の CASE cache 昇格を伴うことがある` と明記する。

特徴:

- 実装差し替えなし
- 呼出し順序変更なし
- 設計上のズレは残る

### 5.2 B案: interface 分離

内容:

- 純粋 read の参照口と、promotion を含む lookup 参照口を分ける。
- prompt は純粋 read 経路を使う。
- document execution は promotion-aware lookup を使う。

特徴:

- read-only と副作用の境界は明確になる
- ただし promotion の所有者は別途定義が必要
- 実装差し替え時は依存配線変更が発生する

### 5.3 C案: promotion 専用 service 新設

内容:

- Base 埋込 snapshot から CASE cache への昇格責務を専用 service に分離する。
- reader / lookup は昇格責務を持たない形にする。
- prompt / document execution / TaskPane build は、必要なら先に promotion を明示実行し、その後 read / build を行う構成にする。

特徴:

- `lookup` と `promotion` の責務境界が最も明確になる
- [CaseTemplateSnapshotService](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\CaseTemplateSnapshotService.cs:7) と [TaskPaneSnapshotCacheService](C:\Users\kazu2\Documents\案件情報System\開発用\dev\CaseInfoSystem.ExcelAddIn\Infrastructure\TaskPaneSnapshotCacheService.cs:12) の重複責務整理にもつながる
- 実装時の影響範囲は最も広い

## 6. 推奨方針

現時点の推奨方針は次のとおりです。

- 今は実装しない
- まずこの問題を文書として固定する
- 将来の設計変更目標は C案とする

根拠:

- 現状の問題は単なる interface 名の違和感ではなく、read-only と説明される経路に CASE cache 昇格副作用が混在している設計上のズレである
- prompt / document execution にも影響し得るため、low-risk な interface 差し替えとは扱わない
- C案が最も責務境界を明確にできる

ただし、現時点では次の事項は未実施とする。

- interface 追加
- interface 改名
- interface 削除
- 実装差し替え
- 昇格責務の再配線

## 7. 非対象

この文書では次を扱いません。

- TaskPane host 再利用方針の変更
- `WorkbookActivate` / `WindowActivate` の表示制御変更
- `TASKPANE_MASTER_VERSION` 更新方針の変更
- Base 埋込 snapshot 廃止
- prompt UI 仕様変更
- document execution の業務仕様変更

## 8. 参照

- [architecture.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\architecture.md)
- [flows.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\flows.md)
- [ui-policy.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\ui-policy.md)
- [template-metadata-inventory.md](C:\Users\kazu2\Documents\案件情報System\開発用\docs\template-metadata-inventory.md)

## 付録: 本流外の low-risk 候補

### Word warm-up 判定の参照差し替え

対象:
ThisAddIn.ScheduleWordWarmup()

内容:
_documentExecutionModeService.CanAttemptVstoExecution() の参照を
interface 化または依存整理する候補

理由:
- promotion 副作用なし
- prompt / document execution 本線に影響しない
- TemplatePath / snapshot 再構築に非関与
- consumer 側のみで完結

評価:
- 安全性: 高い
- 実装コスト: 低
- 優先度: 低〜中（本流外）

方針:
- 今回は実装しない
- read-only API adoption 本流完了後にまとめて実施する