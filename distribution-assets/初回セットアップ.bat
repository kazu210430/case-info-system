@echo off
setlocal

set "EXIT_CODE=1"
set "SCRIPT_DIR=%~dp0"
set "CERT_FILE=%SCRIPT_DIR%CaseInfoSystem.Internal.cer"
set "EXCEL_VSTO=%SCRIPT_DIR%Addins\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.vsto"
set "WORD_VSTO=%SCRIPT_DIR%Addins\CaseInfoSystem.WordAddIn\CaseInfoSystem.WordAddIn.vsto"

echo.
echo ========================================
echo 案件情報System 初回セットアップを開始します。
echo 管理者として実行してください。
echo ========================================
echo.
echo [1/5] 必要なファイルを確認しています...

if not exist "%CERT_FILE%" (
    echo エラー: 証明書ファイルが見つかりません。
    echo 対象パス: "%CERT_FILE%"
    goto :fail
)

if not exist "%EXCEL_VSTO%" (
    echo エラー: Excel Add-in の .vsto が見つかりません。
    echo 対象パス: "%EXCEL_VSTO%"
    goto :fail
)

if not exist "%WORD_VSTO%" (
    echo エラー: Word Add-in の .vsto が見つかりません。
    echo 対象パス: "%WORD_VSTO%"
    goto :fail
)

echo [2/5] 証明書を Trusted Root に登録しています...
certutil -f -addstore Root "%CERT_FILE%"
if errorlevel 1 (
    echo.
    echo エラー: Trusted Root への証明書登録に失敗しました。
    echo 管理者権限で実行しているか確認してください。
    goto :fail
)

echo.
echo [3/5] 証明書を TrustedPublisher に登録しています...
certutil -f -addstore TrustedPublisher "%CERT_FILE%"
if errorlevel 1 (
    echo.
    echo エラー: TrustedPublisher への証明書登録に失敗しました。
    echo 管理者権限で実行しているか確認してください。
    goto :fail
)

echo.
echo [4/5] Excel Add-in のセットアップを起動しています...
start "" "%EXCEL_VSTO%"
if errorlevel 1 (
    echo.
    echo エラー: Excel Add-in の .vsto を起動できませんでした。
    goto :fail
)
echo Excel Add-in のインストーラーが起動しました。
echo 画面の案内に従って完了したら、キーを押して次へ進んでください。
pause

echo.
echo [5/5] Word Add-in のセットアップを起動しています...
start "" "%WORD_VSTO%"
if errorlevel 1 (
    echo.
    echo エラー: Word Add-in の .vsto を起動できませんでした。
    goto :fail
)
echo Word Add-in のインストーラーが起動しました。
echo 画面の案内に従って完了したら、キーを押してください。
pause

echo.
echo 初回セットアップが完了しました。
echo Excel と Word を再起動してから利用を開始してください。
set "EXIT_CODE=0"
goto :end

:fail
echo.
echo 初回セットアップを中断しました。内容を確認してから画面を閉じてください。

:end
echo.
pause
exit /b %EXIT_CODE%
