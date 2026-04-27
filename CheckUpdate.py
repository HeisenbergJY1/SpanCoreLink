# -*- coding: utf-8 -*-
"""
CheckUpdate.py - Check and upgrade PySap2000 library
"""
import sys
import os
import subprocess

try:
    import rhinoscriptsyntax as rs
    RHINO_ENV = True
except ImportError:
    RHINO_ENV = False


def get_rhino_python():
    """获取 Rhino 内置 Python 解释器路径"""
    # Rhino 8 的 Python 环境路径
    user_home = os.path.expanduser("~")
    rhino_python = os.path.join(user_home, ".rhinocode", "py39-rh8", "python.exe")

    if os.path.exists(rhino_python):
        return rhino_python

    # 备用：尝试其他版本
    rhinocode_dir = os.path.join(user_home, ".rhinocode")
    if os.path.exists(rhinocode_dir):
        for folder in os.listdir(rhinocode_dir):
            if folder.startswith("py") and "-rh" in folder:
                py_path = os.path.join(rhinocode_dir, folder, "python.exe")
                if os.path.exists(py_path):
                    return py_path

    return None


def get_rhino_site_envs():
    """获取 Rhino 的 site-envs 目录（Rhino 8 优先加载的位置）"""
    user_home = os.path.expanduser("~")
    site_envs_dir = os.path.join(user_home, ".rhinocode", "py39-rh8", "site-envs")

    if os.path.exists(site_envs_dir):
        # 查找 default-xxx 目录
        for folder in os.listdir(site_envs_dir):
            if folder.startswith("default-"):
                return os.path.join(site_envs_dir, folder)

    return None


def get_installed_location():
    """获取 PySap2000 实际安装位置（用于确定升级目标）"""
    try:
        import PySap2000
        if hasattr(PySap2000, '__file__') and PySap2000.__file__:
            # 返回包所在的父目录（如 site-packages 或 site-envs/default-xxx）
            pkg_dir = os.path.dirname(PySap2000.__file__)
            return os.path.dirname(pkg_dir)
    except ImportError:
        pass
    return None


def get_installed_version():
    """Get installed PySap2000 version"""
    try:
        import PySap2000
        return PySap2000.__version__
    except ImportError:
        return None


def get_pypi_version():
    """Get latest version from PyPI"""
    try:
        import urllib.request
        import json
        url = "https://pypi.org/pypi/pysap2000/json"
        with urllib.request.urlopen(url, timeout=10) as response:
            data = json.loads(response.read().decode())
            return data["info"]["version"]
    except Exception:
        return None


def compare_versions(v1: str, v2: str) -> int:
    """Compare versions: -1 if v1 < v2, 0 if equal, 1 if v1 > v2"""
    def parse(v):
        return [int(x) for x in v.split('.')]
    try:
        p1, p2 = parse(v1), parse(v2)
        max_len = max(len(p1), len(p2))
        p1.extend([0] * (max_len - len(p1)))
        p2.extend([0] * (max_len - len(p2)))
        for a, b in zip(p1, p2):
            if a < b: return -1
            if a > b: return 1
        return 0
    except Exception:
        return 0


def get_pip_installed_version(python_exe: str, package: str) -> str:
    """通过 pip show 获取已安装包的版本（绕过 import 缓存）"""
    cmd = [python_exe, "-m", "pip", "show", package]
    
    env = os.environ.copy()
    for key in list(env.keys()):
        if "PYTHON" in key.upper() or key in ["VIRTUAL_ENV", "CONDA_PREFIX"]:
            del env[key]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, env=env)
        if result.returncode == 0:
            for line in result.stdout.split('\n'):
                if line.startswith('Version:'):
                    return line.split(':', 1)[1].strip()
    except Exception:
        pass
    return None


def run_pip(python_exe: str, args: list) -> bool:
    """执行 pip 命令，清除环境变量避免冲突"""
    cmd = [python_exe, "-m", "pip"] + args
    
    # 清除可能导致冲突的环境变量
    env = os.environ.copy()
    for key in list(env.keys()):
        if "PYTHON" in key.upper() or key in ["VIRTUAL_ENV", "CONDA_PREFIX"]:
            del env[key]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            env=env
        )
        if result.returncode == 0:
            return True
        else:
            if result.stderr:
                print(f"Error: {result.stderr[:200]}")
            return False
    except subprocess.TimeoutExpired:
        print("Error: Command timed out")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def check_dependencies():
    """Check dependencies status"""
    deps = ["comtypes", "orjson"]
    missing = []
    
    print("Dependencies:")
    for pkg in deps:
        try:
            __import__(pkg)
            print(f"  {pkg}: OK")
        except ImportError:
            print(f"  {pkg}: MISSING")
            missing.append(pkg)
    
    return missing


def check_update():
    """Main function: check version and upgrade if needed"""
    print("=" * 50)
    print("PySap2000 Update Check")
    print("=" * 50)
    
    # 获取 Rhino Python 路径
    python_exe = get_rhino_python()
    if python_exe:
        print(f"Python: {python_exe}")
    else:
        print("Python: Not found (will show manual command)")
    print("")
    
    # Get installed version
    installed = get_installed_version()
    
    if installed is None:
        print("Status: NOT INSTALLED")
        print("")
        if python_exe:
            print("Installing PySap2000...")
            site_envs = get_rhino_site_envs()
            if site_envs:
                pip_args = ["install", "--target", site_envs, "pysap2000"]
            else:
                pip_args = ["install", "pysap2000"]
            
            if run_pip(python_exe, pip_args):
                print("Install successful! Restart Rhino to use.")
            else:
                print("Auto-install failed. Run manually in CMD:")
                if site_envs:
                    print(f'  "{python_exe}" -m pip install --target "{site_envs}" pysap2000')
                else:
                    print(f'  "{python_exe}" -m pip install pysap2000')
        else:
            print("Run in CMD:")
            print("  pip install pysap2000")
        print("=" * 50)
        return
    
    print(f"Installed: {installed}")
    
    # Get PyPI version
    pypi = get_pypi_version()
    if pypi is None:
        print("Latest:    (network error)")
        if RHINO_ENV:
            rs.MessageBox(
                f"Cannot connect to PyPI to check updates.\n\n"
                f"Current version: {installed}\n"
                f"Please check your network connection.",
                48, "Network Error"
            )
    else:
        print(f"Latest:    {pypi}")
    print("")
    
    # Check dependencies first
    missing = check_dependencies()
    print("")
    
    # Compare and upgrade
    need_upgrade = False
    upgrade_success = False
    
    if pypi:
        cmp = compare_versions(installed, pypi)
        if cmp < 0:
            print(f"Status: Update available ({installed} -> {pypi})")
            need_upgrade = True
            
            if python_exe:
                print("Upgrading...")
                # 根据实际 import 位置决定安装目标
                install_location = get_installed_location()
                if install_location and "site-envs" in install_location:
                    # 如果当前是从 site-envs 加载的，就装到那里
                    pip_args = ["install", "--upgrade", "--no-cache-dir", "--target", install_location, "pysap2000"]
                else:
                    # 否则用默认方式
                    pip_args = ["install", "--upgrade", "--no-cache-dir", "pysap2000"]

                if run_pip(python_exe, pip_args):
                    # 验证升级是否真的成功
                    new_version = get_pip_installed_version(python_exe, "pysap2000")
                    if new_version and compare_versions(new_version, installed) > 0:
                        print(f"Upgrade successful! ({installed} -> {new_version})")
                        upgrade_success = True
                    else:
                        print(f"Upgrade may have failed. Version still: {new_version or installed}")
                        if RHINO_ENV:
                            rs.MessageBox(
                                f"Upgrade may have failed.\n"
                                f"Version still: {new_version or installed}\n\n"
                                f"Try running in CMD:\n"
                                f'"{python_exe}" -m pip install --upgrade pysap2000',
                                48, "Upgrade Warning"
                            )
                else:
                    print("Auto-upgrade failed.")
                    if RHINO_ENV:
                        rs.MessageBox(
                            f"Auto-upgrade failed.\n\n"
                            f"Try running in CMD:\n"
                            f'"{python_exe}" -m pip install --upgrade pysap2000',
                            16, "Upgrade Failed"
                        )
        elif cmp == 0:
            print("Status: Up to date")
        else:
            print("Status: Development version")
    
    # Install missing dependencies
    deps_success = False
    if missing and python_exe:
        print(f"Installing: {', '.join(missing)}...")
        if run_pip(python_exe, ["install"] + missing):
            print("Dependencies installed!")
            deps_success = True
        else:
            print("Failed to install dependencies.")
            if RHINO_ENV:
                rs.MessageBox(
                    f"Failed to install dependencies: {', '.join(missing)}\n\n"
                    f"Try running in CMD:\n"
                    f'"{python_exe}" -m pip install {" ".join(missing)}',
                    16, "Dependency Install Failed"
                )
    
    # Show manual command if auto failed
    if (need_upgrade and not upgrade_success) or (missing and not deps_success):
        print("")
        print("-" * 50)
        print("Run manually in CMD:")
        packages = []
        if need_upgrade and not upgrade_success:
            packages.append("pysap2000")
        if missing and not deps_success:
            packages.extend(missing)
        
        if python_exe:
            print(f'  "{python_exe}" -m pip install --upgrade {" ".join(packages)}')
        else:
            print(f'  pip install --upgrade {" ".join(packages)}')
    
    if upgrade_success or deps_success:
        print("")
        print("Please restart Rhino to apply changes.")
    
    print("=" * 50)


if __name__ == "__main__":
    check_update()
