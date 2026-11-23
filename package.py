#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX2PDF 打包脚本
用于将程序打包成 Windows 可执行文件

使用方法:
    python package.py --win32    # 打包 Windows 32/64位 exe
    python package.py --linux    # 打包 Linux 可执行文件
    python package.py --macos    # 打包 macOS 应用
"""

import os
import sys
import shutil
import argparse
from pathlib import Path

# 获取虚拟环境路径
PROJECT_DIR = Path(__file__).parent
VENV_DIR = PROJECT_DIR / '.venv'


def check_dependencies():
    """检查依赖是否安装"""
    print("检查依赖...")
    
    # 检查 pyinstaller
    try:
        import PyInstaller
        print("✓ PyInstaller 已安装")
    except ImportError:
        print("✗ PyInstaller 未安装")
        print("请运行: pip install pyinstaller")
        return False
    
    # 检查 docx2pdf
    try:
        from docx2pdf import convert
        print("✓ docx2pdf 已安装")
    except ImportError:
        print("✗ docx2pdf 未安装")
        print("请运行: pip install docx2pdf")
        return False
    
    # 检查 win32com
    try:
        import win32com
        print("✓ win32com 已安装")
    except ImportError:
        print("✗ win32com 未安装")
        print("如需使用 Windows Word COM 引擎，请运行: pip install pywin32")

    # 检查 PyQt5
    try:
        from PyQt5.QtWidgets import QApplication
        print("✓ PyQt5 已安装")
    except ImportError:
        print("✗ PyQt5 未安装")
        print("请运行: pip install PyQt5")
        return False
    
    return True


def get_python_libs_path():
    """获取 Python 库路径"""
    import site
    return site.getsitepackages()[0]


def package_win32(output_dir='dist'):
    """打包 Windows exe"""
    print("\n开始打包 Windows exe...")
    
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)
    
    # PyInstaller 命令
    cmd = [
        'pyinstaller',
        '--name=DOCX2PDF',
        '--onefile',
        '--windowed',  # 无控制台窗口
        '--icon=None',  # 可以添加图标
        '--add-data', f'{PROJECT_DIR / "docx2pdf_gui.py"};.',
        '--hidden-import=docx2pdf',
        '--hidden-import=win32com',
        '--hidden-import=PyQt5',
        '--hidden-import=PyQt5.QtWidgets',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        '--hidden-import=concurrent.futures',
        f'--distpath={output_dir}',
        f'--workpath={PROJECT_DIR / "build"}',
        f'--specpath={PROJECT_DIR}',
        str(PROJECT_DIR / 'docx2pdf_gui.py'),
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    
    # 执行打包
    import subprocess
    result = subprocess.run(cmd, cwd=str(PROJECT_DIR))
    
    if result.returncode == 0:
        print(f"\n✓ 打包成功!")
        print(f"输出文件: {output_dir}/DOCX2PDF.exe")
        return True
    else:
        print(f"\n✗ 打包失败! (返回码: {result.returncode})")
        return False


def package_linux(output_dir='dist'):
    """打包 Linux 可执行文件"""
    print("\n开始打包 Linux 可执行文件...")
    
    os.makedirs(output_dir, exist_ok=True)
    
    cmd = [
        'pyinstaller',
        '--name=DOCX2PDF',
        '--onefile',
        '--add-data', f'{PROJECT_DIR / "docx2pdf_gui.py"}:.',
        '--hidden-import=docx2pdf',
        '--hidden-import=win32com',
        '--hidden-import=PyQt5',
        '--hidden-import=PyQt5.QtWidgets',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        f'--distpath={output_dir}',
        f'--workpath={PROJECT_DIR / "build"}',
        str(PROJECT_DIR / 'docx2pdf_gui.py'),
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    
    import subprocess
    result = subprocess.run(cmd, cwd=str(PROJECT_DIR))
    
    if result.returncode == 0:
        print(f"\n✓ 打包成功!")
        print(f"输出文件: {output_dir}/DOCX2PDF")
        return True
    else:
        print(f"\n✗ 打包失败! (返回码: {result.returncode})")
        return False


def clean_build():
    """清理构建文件"""
    print("\n清理构建文件...")
    
    build_dir = PROJECT_DIR / 'build'
    dist_dir = PROJECT_DIR / 'dist'
    spec_file = PROJECT_DIR / 'DOCX2PDF.spec'
    
    if build_dir.exists():
        shutil.rmtree(build_dir)
        print("✓ 已删除 build 目录")
    
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
        print("✓ 已删除 dist 目录")
    
    if spec_file.exists():
        spec_file.unlink()
        print("✓ 已删除 DOCX2PDF.spec 文件")


def main():
    parser = argparse.ArgumentParser(description='DOCX2PDF 打包工具')
    parser.add_argument('--win32', action='store_true', help='打包 Windows exe')
    parser.add_argument('--linux', action='store_true', help='打包 Linux 可执行文件')
    parser.add_argument('--macos', action='store_true', help='打包 macOS 应用')
    parser.add_argument('--clean', action='store_true', help='清理构建文件')
    parser.add_argument('--output', '-o', default='dist', help='输出目录')
    
    args = parser.parse_args()
    
    if args.clean:
        clean_build()
        return
    
    # 检查依赖
    if not check_dependencies():
        print("\n依赖检查失败，请先安装依赖")
        sys.exit(1)
    
    # 执行打包
    if args.win32:
        success = package_win32(args.output)
    elif args.linux:
        success = package_linux(args.output)
    elif args.macos:
        print("macOS 打包功能待实现")
        success = False
    else:
        parser.print_help()
        success = False
    
    if success:
        print("\n打包完成!")
    else:
        sys.exit(1)


if __name__ == '__main__':
    main()

