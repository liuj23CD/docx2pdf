# DOCX/DOC 到 PDF 转换器

使用 docx2pdf 库实现的图形界面 DOC/DOCX 到 PDF 转换工具，支持批量转换和进度显示。

## 功能特点

- ✅ 支持 DOC 和 DOCX 格式转换
- ✅ 图形界面操作简单
- ✅ 多引擎支持（docx2pdf/Word COM）
- ✅ 批量转换支持
- ✅ 实时进度显示
- ✅ 自定义输出目录
- ✅ 转换日志记录
- ✅ 转换完成后自动打开输出目录

## 转换引擎

程序支持多种转换引擎，可根据需要选择：

1. **docx2pdf** (默认, 跨平台)
   - 优点：跨平台支持，不需要 Microsoft Word
   - 缺点：需要安装 LibreOffice

2. **win32com** (Windows Word COM)
   - 优点：转换质量高，不需要额外系统依赖
   - 缺点：仅限 Windows，需要安装 Microsoft Word

## 安装依赖

