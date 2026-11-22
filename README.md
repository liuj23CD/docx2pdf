# DOCX/DOC 到 PDF 转换器

使用 docx2pdf 库实现的图形界面 DOC/DOCX 到 PDF 转换工具，支持批量转换和进度显示。

## 功能特点

- ✅ 支持 DOC 和 DOCX 格式转换
- ✅ 图形界面操作简单
- ✅ 批量转换支持
- ✅ 实时进度显示
- ✅ 自定义输出目录
- ✅ 转换日志记录
- ✅ 转换完成后自动打开输出目录

## 安装依赖

```bash
pip install -r requirements.txt
```

或者手动安装：

```bash
pip install docx2pdf PyQt5
```

## 运行方式

```bash
python docx2pdf_gui.py
```

## 使用说明

1. **添加文件**: 点击"浏览"按钮选择要转换的 DOC/DOCX 文件
2. **管理文件列表**: 使用"添加文件"、"移除选择"、"清空列表"按钮管理待转换文件
3. **设置输出目录**: 选择转换后 PDF 文件的保存位置
4. **开始转换**: 点击"开始转换"按钮，等待转换完成
5. **批量转换**: 点击"批量转换"按钮，自动转换指定目录下的所有 DOC/DOCX 文件

## 技术栈

- **GUI**: PyQt5
- **转换库**: docx2pdf
- **语言**: Python 3.7+

## 注意事项

- 确保系统已安装 LibreOffice（docx2pdf 依赖）
- 在 Linux 系统上，可能需要手动安装 LibreOffice：
  ```bash
  sudo apt-get install libreoffice
  ```
- 转换大文件可能需要一些时间，请耐心等待