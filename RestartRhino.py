# -*- coding: utf-8 -*-
"""
RestartRhino.py - Restart Rhino application

Save current file and restart Rhino to apply updates.
"""
import os
import sys
import subprocess

try:
    import rhinoscriptsyntax as rs
    import Rhino
    RHINO_ENV = True
except ImportError:
    RHINO_ENV = False
    print("This script must be run in Rhino.")


def get_rhino_exe():
    """Get Rhino executable path"""
    if not RHINO_ENV:
        return None
    
    try:
        # 获取当前 Rhino 进程路径
        import System.Diagnostics as Diagnostics
        process = Diagnostics.Process.GetCurrentProcess()
        return process.MainModule.FileName
    except Exception:
        pass
    
    # 备用方案：常见安装路径
    common_paths = [
        r"C:\Program Files\Rhino 8\System\Rhino.exe",
        r"C:\Program Files\Rhino 7\System\Rhino.exe",
    ]
    for path in common_paths:
        if os.path.exists(path):
            return path
    
    return None


def restart_rhino():
    """Restart Rhino application"""
    if not RHINO_ENV:
        print("This script must be run in Rhino.")
        return
    
    print("=" * 50)
    print("Restart Rhino")
    print("=" * 50)
    
    # 获取当前打开的文件
    doc = Rhino.RhinoDoc.ActiveDoc
    current_file = doc.Path if doc and doc.Path else None
    
    if current_file:
        print(f"Current file: {current_file}")
    
    # 检查是否有未保存的更改
    if doc and doc.Modified:
        result = rs.MessageBox(
            "Save changes before restart?",
            3,  # Yes/No/Cancel
            "Restart Rhino"
        )
        if result == 6:  # Yes
            if current_file:
                rs.Command("_Save", False)
                print("File saved.")
            else:
                rs.Command("_SaveAs", False)
        elif result == 2:  # Cancel
            print("Restart cancelled.")
            return
    
    # 获取 Rhino 路径
    rhino_exe = get_rhino_exe()
    if not rhino_exe:
        print("Error: Cannot find Rhino executable.")
        rs.MessageBox("Cannot find Rhino executable.", 16, "Error")
        return
    
    print(f"Rhino: {rhino_exe}")
    
    # 构建启动命令
    args = [rhino_exe]
    if current_file and os.path.exists(current_file):
        # 使用 /nosplash 避免启动画面，直接打开文件
        args.append("/nosplash")
        args.append(current_file)
    
    print(f"Restarting: {' '.join(args)}")
    
    # 启动新 Rhino 实例（不用 shell=True，避免中文路径问题）
    try:
        subprocess.Popen(args)
    except Exception as e:
        print(f"Error starting Rhino: {e}")
        return
    
    # 关闭当前 Rhino
    rs.Command("_Exit", False)


if __name__ == "__main__":
    restart_rhino()
