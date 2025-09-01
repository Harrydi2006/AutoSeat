#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
启动脚本 - 检查环境并运行应用程序
"""

import sys
import os

def check_python_version():
    """检查Python版本"""
    print(f"Python版本: {sys.version}")
    if sys.version_info < (3, 7):
        print("警告: 建议使用Python 3.7或更高版本")
        return False
    return True

def check_dependencies():
    """检查依赖包"""
    required_packages = [
        'streamlit',
        'pandas', 
        'numpy',
        'matplotlib',
        'ortools',
        'openpyxl',
        'PIL'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            if package == 'PIL':
                __import__('PIL')
            elif package == 'ortools':
                __import__('ortools.sat.python.cp_model')
            else:
                __import__(package)
            print(f"✓ {package} 已安装")
        except ImportError:
            print(f"✗ {package} 未安装")
            missing_packages.append(package)
    
    return missing_packages

def install_missing_packages(packages):
    """安装缺失的包"""
    if not packages:
        return True
        
    print(f"\n正在安装缺失的包: {', '.join(packages)}")
    
    # 创建安装命令
    package_map = {
        'PIL': 'Pillow',
        'ortools': 'ortools'
    }
    
    install_list = []
    for pkg in packages:
        install_list.append(package_map.get(pkg, pkg))
    
    import subprocess
    try:
        cmd = [sys.executable, '-m', 'pip', 'install'] + install_list
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print("✓ 依赖包安装成功")
            return True
        else:
            print(f"✗ 安装失败: {result.stderr}")
            return False
    except Exception as e:
        print(f"✗ 安装过程出错: {e}")
        return False

def main():
    """主函数"""
    print("=" * 50)
    print("智能分座位系统 - 启动检查")
    print("=" * 50)
    
    # 检查Python版本
    if not check_python_version():
        return False
    
    print("\n检查依赖包...")
    missing = check_dependencies()
    
    if missing:
        print(f"\n发现 {len(missing)} 个缺失的依赖包")
        if input("是否自动安装? (y/n): ").lower().startswith('y'):
            if not install_missing_packages(missing):
                print("\n依赖包安装失败，请手动安装后重试")
                return False
            print("\n重新检查依赖包...")
            missing = check_dependencies()
            if missing:
                print(f"仍有缺失的包: {missing}")
                return False
        else:
            print("请手动安装缺失的依赖包后重试")
            return False
    
    print("\n✓ 所有依赖包检查通过")
    
    # 检查utils模块
    try:
        from utils import load_names_from_excel
        print("✓ Utils模块导入成功")
    except ImportError as e:
        print(f"✗ Utils模块导入失败: {e}")
        return False
    
    print("\n=" * 50)
    print("环境检查完成，准备启动应用程序...")
    print("=" * 50)
    
    # 启动Streamlit应用
    try:
        import subprocess
        cmd = [sys.executable, '-m', 'streamlit', 'run', 'app.py']
        subprocess.run(cmd)
    except Exception as e:
        print(f"启动应用程序失败: {e}")
        return False
    
    return True

if __name__ == "__main__":
    main()