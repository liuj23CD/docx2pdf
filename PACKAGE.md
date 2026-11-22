# DOCX2PDF 打包说明

## 简介

本项目已经成功打包成独立的可执行文件。在 Linux 系统上打包的可执行文件位于 `dist/DOCX2PDF`。

## 在 Windows 上打包

由于 PyInstaller 在 Linux 上只能打包 Linux 可执行文件，在 Windows 上只能打包 exe 文件，所以您需要在 Windows 系统上执行以下步骤来打包 Windows exe 文件：

### 方法一：使用打包脚本（推荐）

1. 将项目文件复制到 Windows 系统
2. 在 Windows 命令提示符或 PowerShell 中运行：

```powershell
# 激活虚拟环境
.venv\Scripts\activate

# 安装 PyInstaller（如果尚未安装）
pip install pyinstaller

# 执行打包
pyinstaller --name DOCX2PDF --onefile --windowed ^
  --add-data "docx2pdf_gui.py;." ^
  --hidden-import=docx2pdf ^
  --hidden-import=PyQt5 ^
  --hidden-import=PyQt5.QtWidgets ^
  --hidden-import=PyQt5.QtCore ^
  --hidden-import=PyQt5.QtGui ^
  docx2pdf_gui.py
```

打包完成后，exe 文件将位于 `dist/DOCX2PDF.exe`。

### 方法二：使用 package.py 脚本

```bash
# 在 Windows 上
python package.py --win32
```

## 打包说明

- **console=False**: 程序运行时不显示控制台窗口
- **windowed**: 窗口模式运行
- **onefile**: 打包成单个可执行文件
- **hidden-imports**: 包含所有必要的隐藏导入

## 依赖

- Python 3.7+
- docx2pdf 0.1.8+
- PyQt5 5.15.11+
- PyInstaller 6.22.3+

## 使用说明

1. 双击运行 `DOCX2PDF.exe`（Windows）或 `DOCX2PDF`（Linux）
2. 在图形界面中选择要转换的文件
3. 设置输出目录
4. 点击"开始转换"按钮
5. 转换完成后，程序会自动打开输出目录

## 注意事项

1. **LibreOffice**: docx2pdf 依赖 LibreOffice 进行文档转换。请确保系统已安装 LibreOffice：
   - Windows: 从 https://www.libreoffice.org 下载并安装
   - Linux: `sudo apt-get install libreoffice`

2. **杀毒软件**: 某些杀毒软件可能会将打包后的可执行文件误报为病毒，这是正常现象，可以将文件添加到白名单。

3. **文件大小**: 打包后的文件较大（约 50MB），这是因为包含了 Python 运行时和所有依赖库。这是正常现象。

4. **性能**: 第一次启动可能稍慢，因为需要解压资源文件。

## 技术栈

- **GUI**: PyQt5
- **转换库**: docx2pdf
- **打包工具**: PyInstaller
- **语言**: Python 3.12