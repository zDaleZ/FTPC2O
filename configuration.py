import winreg
import sys
import os
import ctypes
import shutil

def is_admin():
    """检查当前是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def elevate_privileges(arg):
    """请求管理员权限并重启脚本"""
    if not is_admin():
        if getattr(sys, 'frozen', False):
            print("请求管理员权限以操作注册表...")
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, arg, None, 1
            )
            sys.exit(0)
        else:
            print("安装需要管理员权限，请使用管理员身份再试一次。当前程序将退出。")
            sys.exit(0)

def refresh_system():
    """
    通知 Windows 系统文件关联已更改，需要刷新。
    """
    # 事件 ID: SHCNE_ASSOCCHANGED (0x08000000) - 文件关联变更
    SHCNE_ASSOCCHANGED = 0x08000000
    # 标志: SHCNF_FLUSH (0x1000) - 立即刷新
    SHCNF_FLUSH = 0x1000
    
    # 调用 SHChangeNotify
    # 后两个参数通常为 None (IntPtr.Zero in C#)，这里用 None 或 ctypes.c_void_p(0) 均可
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_FLUSH, None, None)
    print("已通知 Windows 系统刷新文件关联。")

def create_registry_key(key_path, access=winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY):
    """创建注册表键，如果已存在则打开"""
    try:
        key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, key_path, 0, access)
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"创建注册表键 {key_path} 时出错: {e}")
        return False

def set_registry_value(key_path, value_name, value_data, value_type=winreg.REG_SZ, 
                      access=winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY):
    """设置注册表值"""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, access) as key:
            winreg.SetValueEx(key, value_name, 0, value_type, value_data)
        return True
    except Exception as e:
        print(f"设置注册表值 {key_path}\\{value_name} 时出错: {e}")
        return False

def delete_registry_key(key_path, access=winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY):
    """删除注册表键及其所有子键"""
    try:
        # 先尝试删除所有子键
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, access) as key:
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, 0)
                    subkey_path = f"{key_path}\\{subkey_name}"
                    delete_registry_key(subkey_path, access)
                except OSError:
                    break
        
        # 删除当前键
        winreg.DeleteKeyEx(winreg.HKEY_LOCAL_MACHINE, key_path, access)
        return True
    except Exception as e:
        print(f"删除注册表键 {key_path} 时出错: {e}")
        return False

def delete_registry_value(key_path, value_name, access=winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY):
    """删除注册表值"""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, access) as key:
            winreg.DeleteValue(key, value_name)
        return True
    except Exception as e:
        print(f"删除注册表值 {key_path}\\{value_name} 时出错: {e}")
        return False

def install_registry_entries():
    """安装所有需要的注册表项"""
    # 获取当前Python解释器路径
    exe_path = sys.executable
    
    # 创建主键
    keys_to_create = [
        r"SOFTWARE\DaleZ",
        r"SOFTWARE\DaleZ\FTPC2O",
        r"SOFTWARE\DaleZ\FTPC2O\Capabilities",
        r"SOFTWARE\DaleZ\FTPC2O\Capabilities\URLAssociations",
        r"SOFTWARE\Classes\ftp\OpenWithProgids",
        r"SOFTWARE\Classes\FTPDOWNLOADER",
        r"SOFTWARE\Classes\FTPDOWNLOADER\Application",
        r"SOFTWARE\Classes\FTPDOWNLOADER\DefaultIcon",
        r"SOFTWARE\Classes\FTPDOWNLOADER\shell",
        r"SOFTWARE\Classes\FTPDOWNLOADER\shell\open",
        r"SOFTWARE\Classes\FTPDOWNLOADER\shell\open\command"
    ]
    
    for key in keys_to_create:
        if not create_registry_key(key):
            return False
    
    # 设置注册表值
    values_to_set = [
        (r"SOFTWARE\DaleZ\FTPC2O\Capabilities", "ApplicationDescription", "快速下载 FTP 服务器中的文件并打开"),
        (r"SOFTWARE\DaleZ\FTPC2O\Capabilities\URLAssociations", "ftp", "FTPDOWNLOADER"),
        (r"SOFTWARE\Classes\ftp\OpenWithProgids", "FTPDOWNLOADER", ""),
        (r"SOFTWARE\Classes\FTPDOWNLOADER", None, "FTPC2O"),  # 默认值
        (r"SOFTWARE\Classes\FTPDOWNLOADER\Application", "ApplicationIcon", f"{exe_path},0"),
        (r"SOFTWARE\Classes\FTPDOWNLOADER\Application", "ApplicationName", "FTP 即点即开"),
        (r"SOFTWARE\Classes\FTPDOWNLOADER\Application", "ApplicationDescription", "快速下载 FTP 服务器中的文件并打开"),
        (r"SOFTWARE\Classes\FTPDOWNLOADER\DefaultIcon", None, f"{exe_path},0"),  # 默认值
        (r"SOFTWARE\Classes\FTPDOWNLOADER\shell\open\command", None, f'"{exe_path}" "%1"'),  # 默认值
        (r"SOFTWARE\RegisteredApplications", "FTPC2O", r"SOFTWARE\DaleZ\FTPC2O\Capabilities")
    ]
    
    for key_path, value_name, value_data in values_to_set:
        if not set_registry_value(key_path, value_name, value_data):
            return False
    
    print("注册表项安装成功！")
    refresh_system()
    return True

def uninstall_registry_entries():
    """卸载所有相关的注册表项"""
    # 删除 RegisteredApplications 中的值
    delete_registry_value(r"SOFTWARE\RegisteredApplications", "FTPC2O")
    
    # 删除 FTPDOWNLOADER 相关键
    delete_registry_key(r"SOFTWARE\Classes\FTPDOWNLOADER")
    
    # 删除 ftp\OpenWithProgids 中的值
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\ftp\OpenWithProgids", 
                           0, winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY) as key:
            winreg.DeleteValue(key, "FTPDOWNLOADER")
    except Exception as e:
        print(f"删除 ftp OpenWithProgids 值时出错: {e}")
    
    # 删除 DaleZ 相关键
    delete_registry_key(r"SOFTWARE\DaleZ\FTPC2O")
    
    # 检查 DaleZ 是否还有其他子键，如果没有则删除 DaleZ
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\DaleZ", 
                           0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            try:
                winreg.EnumKey(key, 0)  # 尝试获取第一个子键
                # 如果有子键，则不删除 DaleZ
            except OSError:
                # 如果没有子键，则删除 DaleZ
                delete_registry_key(r"SOFTWARE\DaleZ")
    except Exception:
        pass  # DaleZ 键可能不存在
    
    print("注册表项卸载成功！")
    refresh_system()
    return True

def main():
    # 检查并获取管理员权限
    elevate_privileges()
    
    # 主菜单
    while True:
        print("\n=== FTPC2O 注册表安装工具 ===")
        print("1. 安装注册表项")
        print("2. 卸载注册表项")
        print("3. 退出")
        
        choice = input("请选择操作 (1-3): ").strip()
        
        if choice == "1":
            if install_registry_entries():
                print("安装完成！")
            else:
                print("安装过程中出现错误。")
        elif choice == "2":
            if uninstall_registry_entries():
                print("卸载完成！")
            else:
                print("卸载过程中出现错误。")
        elif choice == "3":
            print("退出程序。")
            break
        else:
            print("无效选择，请重新输入。")

if __name__ == "__main__":
    main()