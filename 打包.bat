@echo off
chcp 65001 >nul
echo ========================================
echo    DOCX to Markdown 转换器打包脚本
echo ========================================
echo.

:: 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

:: 安装依赖
echo [1/4] 安装依赖...
pip install python-docx Pillow pyinstaller -q
if errorlevel 1 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)
echo       完成!
echo.

:: 清理旧文件
echo [2/4] 清理旧文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo       完成!
echo.

:: 打包
echo [3/4] 开始打包...
echo.

:: 执行PyInstaller
pyinstaller gui_app.spec --clean -y

if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo [4/4] 复制文档...
copy requirements.txt "dist\DOCXtoMarkdown\" >nul 2>&1
echo       完成!
echo.

echo ========================================
echo    打包完成！
echo ========================================
echo.
echo 生成的文件在: dist\DOCXtoMarkdown\ 目录下
echo   - DOCXtoMarkdown.exe
echo.
echo 是否打开dist目录? (Y/N)
set /p choice=
if /i "%choice%"=="Y" (
    start explorer "%~dp0dist"
)
pause
