@echo off
chcp 65001 >nul
REM ============================================================
REM  一键构建 ScreenOCR.exe(自动合并 Umi-OCR 引擎, 输出自包含包)
REM  在任意一台装有 Python 的 Windows 电脑上运行本脚本即可。
REM   ★ 要让 exe 在 Win7 上运行, 请使用 Python 3.8 (3.9+ 不支持 Win7)。
REM  产物: dist\ScreenOCR\  (内含 ScreenOCR.exe 和 Umi-OCR 引擎)
REM ============================================================

setlocal
cd /d "%~dp0\.."
set "PACK="%~dp0..\dist\ScreenOCR""

echo [1/5] 检查 Python ...
where python >nul 2>nul
if errorlevel 1 (
    echo 未找到 Python, 请先安装 Python 3.8 (勾选 "Add to PATH")。
    pause
    exit /b 1
)
python --version

echo [2/5] 安装依赖 ...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖安装失败, 请检查网络后重试。
    pause
    exit /b 1
)

echo [3/5] 清理旧构建 ...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [4/5] 打包 exe ...
python -m PyInstaller --noconfirm --clean --onedir --windowed --name ScreenOCR app\main.py
if errorlevel 1 (
    echo 打包失败。
    pause
    exit /b 1
)

echo [5/5] 合并 Umi-OCR 引擎 ...
if exist vendor\Umi-OCR (
    echo 使用本地 vendor\Umi-OCR 目录 ...
    xcopy /e /y /q vendor\Umi-OCR dist\ScreenOCR\Umi-OCR
) else (
    echo 下载官方 Umi-OCR Rapid 自解压包并解压 ...
    powershell -NoProfile -Command ^
      "$r = Invoke-RestMethod 'https://api.github.com/repos/hiroi-sora/Umi-OCR/releases/latest' -Headers @{'User-Agent'='build'}; ^
       $a = $r.assets | Where-Object { $_.name -match 'Rapid' -and $_.name -like '*.7z.exe' } | Select-Object -First 1; ^
       if(-not $a){ throw '未找到 Rapid 包' }; ^
       Invoke-WebRequest $a.browser_download_url -OutFile '%TEMP%\umi.7z.exe'; ^
       & '%TEMP%\umi.7z.exe' -y -o'%TEMP%\umi_extract'; ^
       $u = Get-ChildItem '%TEMP%\umi_extract' -Recurse -Filter 'Umi-OCR.exe' | Select-Object -First 1; ^
       if(-not $u){ throw '解压失败' }; ^
       Copy-Item -Recurse $u.DirectoryName 'dist\ScreenOCR\Umi-OCR'"
    if errorlevel 1 (
        echo 自动合并失败。请在能联网的电脑下载 Umi-OCR_Rapid 自解压包,
        echo 解压得到 Umi-OCR 文件夹后放入 %~dp0..\vendor\Umi-OCR, 重新运行本脚本。
        pause
        exit /b 1
    )
)

echo.
echo ============================================================
echo  构建完成: %PACK%
echo  把整个 ScreenOCR 文件夹拷贝到 Win7 电脑, 双击 ScreenOCR.exe 即可。
echo ============================================================
pause