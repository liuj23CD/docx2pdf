@echo off
REM Windows 打包脚本
REM 使用方法：双击运行或在命令提示符中执行

setlocal enabledelayedexpansion

echo ========================================
echo DOCX2PDF Windows 打包脚本
echo ========================================
echo.

REM 检查虚拟环境
if not exist ".venv\Scripts\activate" (
    echo 错误: 未找到虚拟环境，请先创建虚拟环境
    echo 运行: python -m venv .venv
    pause
    exit /b 1
)

REM 激活虚拟环境
call .venv\Scripts\activate

REM 安装 PyInstaller（如果尚未安装）
echo 正在安装 PyInstaller...
pip install pyinstaller --quiet

REM 清理之前的构建
echo 清理之前的构建文件...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "DOCX2PDF.spec" del /q DOCX2PDF.spec

REM 打包
echo.
echo 开始打包...
echo.

pyinstaller --name DOCX2PDF ^
  --onefile ^
  --windowed ^
  --add-data "docx2pdf_gui.py;." ^
  --hidden-import=docx2pdf ^
  --hidden-import=PyQt5 ^
  --hidden-import=PyQt5.QtWidgets ^
  --hidden-import=PyQt5.QtCore ^
  --hidden-import=PyQt5.QtGui ^
  docx2pdf_gui.py

REM 检查结果
if exist "dist\DOCX2PDF.exe" (
    echo.
    echo ========================================
    echo 打包成功！
    echo ========================================
    echo 输出文件: dist\DOCX2PDF.exe
    echo.
    dir dist\DOCX2PDF.exe
    echo.
    echo 请将 dist 目录中的文件分发给用户
) else (
    echo.
    echo ========================================
    echo 打包失败！
    echo ========================================
    exit /b 1
)

REM 保持命令提示符窗口打开
pause
