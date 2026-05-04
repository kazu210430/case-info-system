本ファイルは現時点では修正対象ではなく、安定状態維持のための記録である。

# 技術的負債メモ

確認時点のコードでは、指定の `FolderWindowService.cs (line 80)` は `WaitForFolderWindow` の開始位置で、`Thread.Sleep` 自体は line 88 にあります。指定の `WorkbookClipboardPreservationService.cs (line 144)` は retry ループの開始位置で、`Thread.Sleep` 自体は line 156 にあります。

## Kernel workbook 選択境界

### KernelCommandService.cs:57 / KernelTemplateSyncService.cs:129 / KernelOpenWorkbookLocator.cs:36

* 内容: 雛形登録・更新は `KernelCommandService.Execute(context, actionId)` の `reflect-template` 分岐から始まるが、`context` は `ExecuteReflectTemplate()` / `KernelTemplateSyncService.Execute()` へ渡されない。`KernelTemplateSyncService.Execute()` は `GetOpenKernelWorkbook()` に依存し、`KernelOpenWorkbookLocator.GetOpenKernelWorkbook()` は開いている workbook を列挙して Kernel と判定された最初の workbook を返す。
* 見ていない文脈: active workbook、visible workbook、`WorkbookContext.SystemRoot`、表示中の CASE workbook。
* 下流との関係: `MasterTemplateCatalogService.InvalidateCache(workbook)` は resolved master path 単位で cache を無効化するため、cache 境界自体は改善済み。残課題はその upstream にある「どの Kernel workbook を対象として処理するか」の境界。
* 影響範囲: 雛形登録・更新、Base snapshot 更新、master catalog cache invalidate、複数 `SYSTEM_ROOT` 共存時の操作対象選択。
* 危険度: 中
* 現状: 単一 Kernel workbook 運用では問題化しにくいため未対応。複数 Kernel workbook や hidden workbook が同時に存在する場合は、利用者の意図と異なる root を操作対象にする余地がある。
* 将来案: command / UI / CASE 文脈から `SYSTEM_ROOT` を明示的に渡し、その文脈で Kernel workbook を確定する。`GetOpenKernelWorkbook()` は単一 root 前提の convenience に限定し、複数 root を跨ぐ経路では使用範囲を絞る。
* 補足方針: HOME unbound は placeholder-only に固定する。`GetOpenKernelWorkbook()` を HOME 表示補助や HOME close 時の復元補助として「1冊選ぶ API」に使う経路は縮退対象とし、残す場合も startup の open 有無判定などの convenience に限定する。

## CASE workbook lifecycle orchestration 境界

### CaseWorkbookLifecycleService.cs:143 / CaseWorkbookLifecycleService.cs:455 / WorkbookLifecycleCoordinator.cs:154

* 内容: `CaseWorkbookLifecycleService` は分割後も lifecycle orchestration の中心であり、dirty session 管理、created case folder offer pending、managed close scheduling、post-close follow-up 予約、CASE HOME 表示補正の順序依存を抱える。prompt UI、folder open、name rule 読取、managed close 入れ子管理、post-close scheduler は別サービスへ分離済みだが、close 契約自体は複数クラスに跨って維持されている。
* 影響範囲: UI / Office操作 / lifecycle
* 危険度: 中
* 現状: 現行 `main` はこの分割後構成を現在地として固定するが、追加 refactor では `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` の順序保持が必要。

## Thread.Sleep 依存

### TaskPaneRefreshOrchestrationService.cs:300

* 内容: `ResolveWorkbookPaneWindow` で `Application.DoEvents()` の後に固定待機している。
* 何を待っているか: 対象 workbook に対応する `Excel.Window` が解決可能になるまでの UI 更新。
* 影響範囲: UI / Office操作
* 危険度: 高
* 現状: 実機では安定動作しているため未対応

### TaskPaneRefreshOrchestrationService.cs:472

* 内容: `WaitForTaskPaneReadyRetry` で再試行前に固定待機している。
* 何を待っているか: task pane を再表示する前提となる workbook window の可視化と ready 状態の反映。
* 影響範囲: UI / Office操作
* 危険度: 高
* 現状: 実機では安定動作しているため未対応

### FolderWindowService.cs:88

* 内容: `WaitForFolderWindow` で Explorer ウィンドウ探索を 100ms 間隔でポーリングしている。対象メソッドの開始位置は line 80。
* 何を待っているか: `explorer.exe` 起動後に対象フォルダのウィンドウハンドルが見つかること。
* 影響範囲: UI / その他
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### PathCompatibilityService.cs:739

* 内容: `WaitRetryTickMs` で固定待機し、ファイル昇格処理の retry 間隔として使っている。
* 何を待っているか: `PromoteFileToDestinationSafely` と `PromoteAdjacentStagingFileToDestinationSafely` の再試行前に、ファイル状態やロック状態が変わること。
* 影響範囲: Office操作 / その他
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WordInteropService.cs:558

* 内容: `SleepSafe` で固定待機し、Word ウィンドウを前面化する retry 間隔として使っている。
* 何を待っているか: `TryBringWindowToFront` 実行後に対象 Word ウィンドウが foreground になること。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookClipboardPreservationService.cs:156

* 内容: `WriteClipboardTextWithRetry` で `ExternalException` 捕捉後に 40ms 待機して再試行している。retry ループの開始位置は line 144。
* 何を待っているか: `Clipboard.SetDataObject` 実行時のクリップボード競合やロックが解消されること。
* 影響範囲: Clipboard
* 危険度: 高
* 現状: 実機では安定動作しているため未対応

---

## MessageBox.Show 直書き

### AccountingFormHelperService.cs:374

* 内容: `EnsurePaymentHistoryInputVisible` で、お支払い履歴クリア確認のダイアログを直接表示している。
* 呼び出し箇所の役割: 支払履歴入力 UI を開く前の確認処理。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:199

* 内容: `Reset` で、分割払い予定表を全消去する確認ダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表リセット前の確認。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:240

* 内容: `TryValidateCreateRequest` で、請求額 0 円を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の作成前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:244

* 内容: `TryValidateCreateRequest` で、1 回目期限の日付不正を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の作成前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:248

* 内容: `TryValidateCreateRequest` で、分割払い額未入力を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の作成前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:252

* 内容: `TryValidateCreateRequest` で、分割回数上限超過を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の作成前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:267

* 内容: `TryValidateChangeRequest` で、請求額 0 円を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の途中変更前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:271

* 内容: `TryValidateChangeRequest` で、変更回未入力を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の途中変更前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:275

* 内容: `TryValidateChangeRequest` で、変更回の入力範囲不正を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の途中変更前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:279

* 内容: `TryValidateChangeRequest` で、変更後の分割払い額未入力を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の途中変更前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:288

* 内容: `TryValidateChangeRequest` で、対象回が既に完済済みであることを通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の途中変更前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:293

* 内容: `TryValidateChangeRequest` で、分割回数上限超過を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い予定表の途中変更前バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleCommandService.cs:525

* 内容: `ShowLoadFormStateNumericReadWarning` で、数値読取失敗項目の警告ダイアログを直接表示している。
* 呼び出し箇所の役割: フォーム状態読み込み時の警告表示。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInternalCommandService.cs:81

* 内容: `ShowPendingMessage` で、対象外フローを後回しにしている旨のダイアログを直接表示している。
* 呼び出し箇所の役割: 内部コマンドの未対応フロー通知。
* 影響範囲: UI / その他
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### AccountingPaymentHistoryCommandService.cs:634

* 内容: `ShowLoadFormStateNumericReadWarning` で、数値読取失敗項目の警告ダイアログを直接表示している。
* 呼び出し箇所の役割: 支払履歴フォーム状態読み込み時の警告表示。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingPaymentHistoryCommandService.cs:639

* 内容: `ShowInformationMessage` で、任意メッセージの情報ダイアログを直接表示している。
* 呼び出し箇所の役割: 支払履歴操作時の情報通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingSaveAsService.cs:123

* 内容: `Execute` で、保存先パスを含む保存完了ダイアログを直接表示している。
* 呼び出し箇所の役割: 会計系 Save As 実行後の完了通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingSetCreateService.cs:117

* 内容: `Execute` で、入力できなかった代理人がいることを通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 会計書類セット生成時の部分失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### AccountingSheetCommandService.cs:57

* 内容: `ResetSheet` で、対象シートのリセット確認ダイアログを直接表示している。
* 呼び出し箇所の役割: 会計シートリセット前の確認。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingSheetControlService.cs:373

* 内容: `ApplyBaseAmountHighlight` で、経済的利益額入力を促すダイアログを直接表示している。
* 呼び出し箇所の役割: チェックボックス連動時の入力促し。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### CaseClosePromptService.cs:24

* 内容: `ShowClosePrompt` で、「保存しますか？」のダイアログを直接表示している。
* 呼び出し箇所の役割: CASE workbook クローズ時の保存確認。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### CaseWorkbookLifecycleService.cs:494

* 内容: `ExecuteManagedSessionClose` で、保存または終了失敗時の警告ダイアログを直接表示している。
* 呼び出し箇所の役割: CASE workbook 管理クローズ処理の例外通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### CaseClosePromptService.cs:39

* 内容: `ShowCreatedCaseFolderOfferPrompt` で、作成済み案件フォルダを開くか確認するダイアログを直接表示している。
* 呼び出し箇所の役割: 案件作成完了後のフォルダオファー確認。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:59

* 内容: `Execute` で、未対応 actionId を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: Kernel コマンドの未対応操作通知。
* 影響範囲: UI / その他
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:90

* 内容: `ExecuteSheetNavigation` で、シートを開けなかったことを通知するダイアログを直接表示している。
* 呼び出し箇所の役割: Kernel からのシート遷移失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:98

* 内容: `ExecuteRegisterUserInfo` で、ユーザー情報反映完了のダイアログを直接表示している。
* 呼び出し箇所の役割: ユーザー情報登録後の完了通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:101

* 内容: `ExecuteRegisterUserInfo` で、ユーザー情報登録失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: ユーザー情報登録失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:109

* 内容: `ExecuteReflectTemplate` で、`kernelTemplateSyncResult.Message` をそのまま表示するダイアログを直接表示している。
* 呼び出し箇所の役割: 雛形登録・更新結果の通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:112

* 内容: `ExecuteReflectTemplate` で、雛形登録・更新失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: 雛形登録・更新失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:122

* 内容: `ExecuteReflectAccountingSetOnly` で、会計書類セット転記エラーのダイアログを直接表示している。
* 呼び出し箇所の役割: 会計書類セット反映失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelCommandService.cs:132

* 内容: `ExecuteReflectBaseHomeOnly` で、Base ホーム転記エラーのダイアログを直接表示している。
* 呼び出し箇所の役割: Base ホーム反映失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelWorkbookLifecycleService.cs:211

* 内容: `HandleWorkbookBeforeClose` で、「保存しますか？」のダイアログを直接表示している。
* 呼び出し箇所の役割: Kernel workbook クローズ時の保存確認。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelWorkbookLifecycleService.cs:258

* 内容: `RequestManagedCloseFromHomeExit` で、「保存しますか？」のダイアログを直接表示している。
* 呼び出し箇所の役割: HOME 画面からの管理クローズ時の保存確認。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelWorkbookLifecycleService.cs:376

* 内容: `ExecuteManagedClose` で、保存または終了失敗時の警告ダイアログを直接表示している。
* 呼び出し箇所の役割: Kernel workbook 管理クローズ処理の例外通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### TaskPaneManager.cs:785

* 内容: `NotifyCasePaneUpdatedIfNeeded` で、文書ボタンパネル更新完了のダイアログを直接表示している。
* 呼び出し箇所の役割: pane 更新結果の通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookCaseTaskPaneRefreshCommandService.cs:34

* 内容: `Refresh` で、対象ブック取得失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: 手動 pane 更新コマンドの事前条件エラー通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookCaseTaskPaneRefreshCommandService.cs:40

* 内容: `Refresh` で、pane 更新サービス利用不可のダイアログを直接表示している。
* 呼び出し箇所の役割: 手動 pane 更新コマンドの依存サービス不足通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookCaseTaskPaneRefreshCommandService.cs:47

* 内容: `Refresh` で、CASE ブック以外での実行を通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 手動 pane 更新コマンドの実行対象チェック。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookCaseTaskPaneRefreshCommandService.cs:55

* 内容: `Refresh` で、文書ボタンパネル更新失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: 手動 pane 更新コマンドの失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookResetCommandService.cs:49

* 内容: `Execute` で、リセット確認メッセージを直接表示している。
* 呼び出し箇所の役割: workbook リセット前の確認。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookResetCommandService.cs:79

* 内容: `ShowResult` で、`result.Message` をそのまま表示するダイアログを直接表示している。
* 呼び出し箇所の役割: workbook リセット結果の通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:56

* 内容: `ShowCustomDocumentProperties` で、対象ブック取得失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: カスタムドキュメントプロパティ表示前の事前条件エラー通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:80

* 内容: `ShowCustomDocumentProperties` で、一覧表示失敗の警告ダイアログを直接表示している。
* 呼び出し箇所の役割: カスタムドキュメントプロパティ表示失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:98

* 内容: `SelectAndSaveSystemRoot` で、対象ブック取得失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: `SYSTEM_ROOT` 更新前の事前条件エラー通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:128

* 内容: `SelectAndSaveSystemRoot` で、更新後のパスを含む完了ダイアログを直接表示している。
* 呼び出し箇所の役割: `SYSTEM_ROOT` 更新完了通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:137

* 内容: `SelectAndSaveSystemRoot` で、更新失敗の警告ダイアログを直接表示している。
* 呼び出し箇所の役割: `SYSTEM_ROOT` 更新失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:155

* 内容: `CopySampleColumnBToHome` で、対象ブック取得失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: `shSample` から `shHOME` への転記前の事前条件エラー通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:171

* 内容: `CopySampleColumnBToHome` で、対象シート取得失敗のダイアログを直接表示している。
* 呼び出し箇所の役割: `shSample` / `shHOME` 存在チェック。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:189

* 内容: `CopySampleColumnBToHome` で、転記対象なしと `shHOME` B 列クリア完了のダイアログを直接表示している。
* 呼び出し箇所の役割: 転記元データなし時の結果通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:213

* 内容: `CopySampleColumnBToHome` で、転記範囲を含む完了ダイアログを直接表示している。
* 呼び出し箇所の役割: `shSample` から `shHOME` への転記完了通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### WorkbookRibbonCommandService.cs:222

* 内容: `CopySampleColumnBToHome` で、転記失敗の警告ダイアログを直接表示している。
* 呼び出し箇所の役割: `shSample` から `shHOME` への転記失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### CreatedCaseNoticeService.cs:24

* 内容: `ShowCreatedCaseCompleted` で、案件情報System 作成完了のダイアログを直接表示している。
* 呼び出し箇所の役割: 案件作成完了通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### UserErrorService.cs:21

* 内容: `ShowUserError` で、ユーザー向けエラーメッセージを直接表示している。
* 呼び出し箇所の役割: 例外発生時の共通エラー通知。
* 影響範囲: UI / その他
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### AccountingImportRangePromptForm.cs:163

* 内容: `BtnConfirm_Click` で、対象範囲を数字で指定するよう求めるダイアログを直接表示している。
* 呼び出し箇所の役割: 取込範囲入力フォームのバリデーション。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingImportRangePromptForm.cs:165

* 内容: `BtnConfirm_Click` で、60 回目までの範囲指定を求めるダイアログを直接表示している。
* 呼び出し箇所の役割: 取込範囲入力フォームのバリデーション。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingImportRangePromptForm.cs:167

* 内容: `BtnConfirm_Click` で、終期は始期以上にするよう求めるダイアログを直接表示している。
* 呼び出し箇所の役割: 取込範囲入力フォームのバリデーション。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingImportRangePromptForm.cs:181

* 内容: `OnFormClosing` で、ボタンで閉じるよう求めるダイアログを直接表示している。
* 呼び出し箇所の役割: 取込範囲入力フォームのクローズ制御。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingInstallmentScheduleInputForm.cs:119

* 内容: `ShowInvoiceEditRestrictedMessage` で、入力フォームからは変更できないことを通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 分割払い入力フォームの編集制限通知。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingPaymentHistoryInputForm.cs:167

* 内容: `ShowInvoiceEditRestrictedMessage` で、入力フォームからは変更できないことを通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 支払履歴入力フォームの編集制限通知。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingReverseGoalSeekForm.cs:140

* 内容: `BtnCalculate_Click` で、目標金額を数字で入力するよう求めるダイアログを直接表示している。
* 呼び出し箇所の役割: 逆算フォームの入力バリデーション。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### AccountingReverseGoalSeekForm.cs:170

* 内容: `OnFormClosing` で、ボタンで閉じるよう求めるダイアログを直接表示している。
* 呼び出し箇所の役割: 逆算フォームのクローズ制御。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### DocumentNamePromptForm.cs:86

* 内容: `BtnOk_Click` で、文書名入力を求めるダイアログを直接表示している。
* 呼び出し箇所の役割: 文書名入力フォームのバリデーション。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### KernelHomeForm.cs:237

* 内容: `BtnCreate_Click` で、顧客名入力を求めるダイアログを直接表示している。
* 呼び出し箇所の役割: Kernel HOME からの案件作成前バリデーション。
* 影響範囲: UI
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### KernelHomeForm.cs:257

* 内容: `OpenSheet` で、シートを開けなかったこととログ確認を促すダイアログを直接表示している。
* 呼び出し箇所の役割: Kernel HOME からのシート遷移失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelHomeForm.cs:266

* 内容: `OpenSheet` で、シートを開けなかったこととログ確認を促すダイアログを直接表示している。
* 呼び出し箇所の役割: `ShowKernelSheetAndRefreshPane` 実行失敗時の通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelHomeForm.cs:293

* 内容: `HandleCaseCreationResult` で、案件作成結果を取得できなかったことを通知するダイアログを直接表示している。
* 呼び出し箇所の役割: 案件作成結果 null 時の通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### KernelHomeForm.cs:298

* 内容: `HandleCaseCreationResult` で、`result.UserMessage` をそのまま表示するダイアログを直接表示している。
* 呼び出し箇所の役割: 案件作成結果に付随するユーザー向けメッセージ通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### ThisAddIn.cs:815

* 内容: `RefreshActiveWorkbookCaseTaskPane` で、pane 更新サービス利用不可のダイアログを直接表示している。
* 呼び出し箇所の役割: add-in 直下の pane 更新失敗通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### Program.cs:115

* 内容: `ShowErrorMessage` で、ランチャーのエラーダイアログを直接表示している。
* 呼び出し箇所の役割: Excel ランチャー起動失敗時などのエラー通知。
* 影響範囲: UI / その他
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### ContentControlBatchReplaceForm.cs:110

* 内容: `ExecuteButton_Click` で、旧タグまたは旧タイトル入力を求めるダイアログを直接表示している。
* 呼び出し箇所の役割: Word 一括置換フォームの入力バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### ContentControlBatchReplaceForm.cs:122

* 内容: `ReplaceNextButton_Click` で、旧タグまたは旧タイトル入力を求めるダイアログを直接表示している。
* 呼び出し箇所の役割: Word 一括置換フォームの入力バリデーション。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### ContentControlBatchReplaceForm.cs:129

* 内容: `ReplaceNextButton_Click` で、例外メッセージをそのまま表示するダイアログを直接表示している。
* 呼び出し箇所の役割: Word 一括置換の実行エラー通知。
* 影響範囲: UI / Office操作
* 危険度: 中
* 現状: 実機では安定動作しているため未対応

### CaseInfoSystem.WordAddIn/ThisAddIn.cs:94

* 内容: `ToggleStylePaneForActiveDocument` で、文書を開いてから実行するよう促すダイアログを直接表示している。
* 呼び出し箇所の役割: Word スタイルペイン切替前の事前条件エラー通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### CaseInfoSystem.WordAddIn/ThisAddIn.cs:223

* 内容: `ShowContentControlBatchReplaceForm` で、アクティブな文書がないことを通知するダイアログを直接表示している。
* 呼び出し箇所の役割: Word 一括置換フォーム起動前の事前条件エラー通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応

### CaseInfoSystem.WordAddIn/ThisAddIn.cs:236

* 内容: `ShowContentControlBatchReplaceForm` で、`BuildCompletionMessage(result)` の結果を表示するダイアログを直接表示している。
* 呼び出し箇所の役割: Word 一括置換処理の完了通知。
* 影響範囲: UI / Office操作
* 危険度: 低
* 現状: 実機では安定動作しているため未対応
