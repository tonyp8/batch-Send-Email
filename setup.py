# -*- coding: utf-8 -*-
"""
Author: @coco29
Time: 2025/07/20 
Project: cx_Freeze打包batchSendMain
"""

from cx_Freeze import setup, Executable

# 1、基础配置(可不进行配置)
build_exe_options = {
    "packages": ['os'],  # 需要包含的额外包,不需要全部写上
    "excludes": [],  # 需要排除的包
    "include_files": ['email/','DefaultSheet.sheet','config.cfg'],  # 需要包含的额外文件或文件夹
    "optimize": 2,  # 优化级别，0（不优化），1（简单优化）， 2（最大优化）
}

# 2、创建可执行文件的配置
executables = [
    Executable(
        script="sendMailUI3.py",  # 需要打包的.py文件
        base= 'Win32GUI',  # 控制台应用程序使用 None，GUI应用程序使用 'Win32GUI'
        icon="icon.ico",  # 可执行文件的图标
        target_name="sendMailGUI.exe",  # 生成的可执行文件名称
    )
]

# 3、调用 setup 函数
setup(
    name="batchSendMail-GUI",  # 应用程序名称
    version="3.8.0",  # 应用程序版本
    description="批量发送邮件GUI",  # 应用程序描述
    options={"build_exe": build_exe_options},  # 构建选项，若没有配置，可不写
    executables=executables,  # 可执行文件配置
)
