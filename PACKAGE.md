# DOCX2PDF 打包说明

## 简介

本项目已经成功打包成独立的可执行文件。支持多种转换引擎，可根据需要选择。

## 依赖

- Python 3.7+
- docx2pdf 0.1.8+
- PyQt5 5.15.11+
- pywin32 (可选，用于 Windows Word COM 引擎)
- PyInstaller 6.22.3+

## 使用说明

1. 双击运行 `DOCX2PDF.exe`（Windows）或 `DOCX2PDF`（Linux）
2. 在图形界面中选择要转换的文件
3. 设置输出目录
4. 选择转换引擎（docx2pdf 或 win32com）
5. 点击"开始转换"按钮
6. 转换完成后，程序会自动打开输出目录

## 注意事项

1. **docx2pdf 引擎**: 需要安装 LibreOffice 进行文档转换：
   - Windows: 从 https://www.libreoffice.org 下载并安装
   - Linux: `sudo apt-get install libreoffice`

2. **win32com 引擎**: 需要安装 Microsoft Word：
   - Windows: 安装 Microsoft Office

3. **杀毒软件**: 某些杀毒软件可能会将打包后的可执行文件误报为病毒，这是正常现象，可以将文件添加到白名单。

4. **文件大小**: 打包后的文件较大（约 50MB），这是因为包含了 Python 运行时和所有依赖库。这是正常现象。

5. **性能**: 第一次启动可能稍慢，因为需要解压资源文件。

## 转换引擎

### 1. docx2pdf (默认, 跨平台)
- **优点**：跨平台支持，不需要 Microsoft Word
- **缺点**：需要安装 LibreOffice

### 2. win32com (Windows Word COM)
- **优点**：转换质量高，不需要额外系统依赖
- **缺点**：仅限 Windows，需要安装 Microsoft Word
## 在 Windows 上打包

由于 PyInstaller 在 Linux 上只能打包 Linux 可执行文件，在 Windows 上只能打包 exe 文件，所以您需要在 Windows 系统上执行以下步骤来打包 Windows exe 文件：

### 方法一：使用打包脚本（推荐）

1. 将项目文件复制到 Windows 系统
2. 在 Windows 命令提示符或 PowerShell 中运行：

