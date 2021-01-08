#!/usr/bin/env python3
# -\- coding: utf-8 -\-
from PyInstaller.__main__ import run
import resource
import os

# -F:打包成一个EXE文件
# -w:不带console输出控制台，window窗体格式 
# --paths：依赖包路径 
# --icon：图标 
# --noupx：不用upx压缩 
# --clean：清理掉临时文件

if __name__ == '__main__':
    opts = ['-F', '-w', r'--paths=C:\Users\10820\.virtualenvs\10820-8QlLVL_3\Scripts', r'--icon=..\Icon\window.ico',
            '--noupx', '--clean',
            r'./XmindToExcel.py']
    run(opts)
