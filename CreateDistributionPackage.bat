@echo off
setlocal

set "EXIT_CODE=1"
set "PUSHD_DONE="
set "SCRIPT_DIR=%~dp0"

pushd "%SCRIPT_DIR%" >nul
if errorlevel 1 (
    echo.
    echo エラー: バッチファイルの配置フォルダへ移動できませんでした。
    goto :end
)
set "PUSHD_DONE=1"

set "REPO_ROOT=%CD%"
set "BUILD_SCRIPT=%REPO_ROOT%\build.ps1"
set "PACKAGE_SCRIPT=%REPO_ROOT%\scripts\Build-DistributionPackage.ps1"
set "CERT_PFX=C:\Users\kazu2\AppData\Local\CaseInfoSystem\Certificates\CaseInfoSystem.InternalRelease.pfx"
set "CERT_THUMBPRINT=820B4E775B9AA8CB21E64A894A176F75A9F48CE4"

echo.
echo ========================================
echo 配布用 ZIP 作成を開始します。
echo ========================================
echo.
echo [1/4] 必要なファイルと証明書を確認しています...

if not exist "%BUILD_SCRIPT%" (
    echo エラー: build.ps1 が見つかりません。
    echo 対象パス: "%BUILD_SCRIPT%"
    goto :fail
)

if not exist "%PACKAGE_SCRIPT%" (
    echo エラー: scripts\Build-DistributionPackage.ps1 が見つかりません。
    echo 対象パス: "%PACKAGE_SCRIPT%"
    goto :fail
)

if not exist "%CERT_PFX%" (
    echo エラー: Release 用証明書ファイルが見つかりません。
    echo 対象パス: "%CERT_PFX%"
    goto :fail
)

echo [2/4] Release Add-in package を生成しています...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BUILD_SCRIPT%" -Mode DeployReleaseAddIn -Project All -ReleaseCertificateKeyFile "%CERT_PFX%" -ReleaseCertificateThumbprint "%CERT_THUMBPRINT%"
if errorlevel 1 (
    echo.
    echo エラー: Release Add-in package の生成に失敗しました。
    goto :fail
)

echo.
echo [3/4] 配布用 ZIP を生成しています...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PACKAGE_SCRIPT%"
if errorlevel 1 (
    echo.
    echo エラー: 配布用 ZIP の生成に失敗しました。
    goto :fail
)

echo.
echo [4/4] 配布用 ZIP の生成が完了しました。
echo 結果を確認してから画面を閉じてください。
set "EXIT_CODE=0"
goto :end

:fail
echo.
echo 処理を中断しました。内容を確認してから画面を閉じてください。

:end
if defined PUSHD_DONE (
    popd >nul
)
echo.
pause
exit /b %EXIT_CODE%
