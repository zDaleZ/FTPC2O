# coding=utf-8
import os
import sys
import ftplib
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import json
from ctypes import windll
import time
import re
import urldecoder
import pathlib
from winui_ify import make_it_winui

# os.startfile('ms-settings:defaultapps')

# 修复 DPI
windll.shcore.SetProcessDpiAwareness(1)

# def set_DPI(root):
#    root.tk.call('tk', 'scaling', Scale/75)

def set_DPI(root):
    import tkinter.font
    scaling = float(root.tk.call('tk', 'scaling'))
    if scaling > 1.4:
        for name in tkinter.font.names(root):
            font = tkinter.font.Font(root=root, name=name, exists=True)
            size = int(font['size'])
            if size < 0:
                font['size'] = round(-0.75*size)

# 设置文件保存位置和每个服务器的用户名和密码
SETTINGS_FILE = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'settings.json')
DEFAULT_SAVE_PATH = os.path.expanduser('~/Downloads')  # 默认保存目录
ICO_PATH = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'ftp.ico')
ftp = None
settings = {}
stopping = False
trycount = 0
taskbar_mode = 0
progress_value = 0
speed = 0
elapsed_time = 0
remaining_time = 0
daemon = False
set_determin = False
downloaded_size = 0
total_size = 0
remaining_time = 0

def load_settings():
    global settings
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, 'r') as f:
            settings = json.load(f)
    else:
        settings = {'save_path': DEFAULT_SAVE_PATH, 'servers': {}, 'timeout': 60, 'blocksize': 8192, 'oldui': False, 'daemon': True, 'debug': False}

def save_settings(save_path, server_entries, timeout, blocksize, oldui, daemon, debug):
    settings['save_path'] = save_path if save_path else DEFAULT_SAVE_PATH  # 提供默认保存目录
    settings['servers'] = {}
    settings['timeout'] = timeout
    settings['blocksize'] = blocksize
    settings['oldui'] = oldui
    settings['daemon'] = daemon
    settings['debug'] = debug
    for i in range(0, len(server_entries), 4):
        server = server_entries[i].get()
        username = server_entries[i+1].get()
        password = server_entries[i+2].get()
        if not server or not username or not password:
            messagebox.showerror('错误', '服务器、用户名和密码不能为空')
            return
        if not validate_server_format(server):
            messagebox.showerror('错误', f'服务器格式不正确：{server}')
            return
        settings['servers'][server] = {'username': username, 'password': password}
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f)

        messagebox.showinfo(title="成功保存", message="设置已保存。")
    except Exception as e:
        messagebox.showerror(title="未能保存", message=f"保存设置时出现错误：{e}")
    

def validate_server_format(server):
    # 使用正则表达式验证 ftp:// 格式的 URL
    pattern = re.compile(
        r'^ftp://(?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9]?\.)*[a-zA-Z]{2,6}'  # 域名部分
        r'|(?:\d{1,3}\.){3}\d{1,3}'  # IP 地址部分
        r'(:[0-9]{1,5})?'  # 端口号部分（可选）：1到5位数字
        r'(/.*)?$'  # 路径部分（可选）
    )

    return pattern.match(server)

def debug_print(message):
    if settings['debug']:
        print(message)

def pathlib_join(*parts):
    """使用 pathlib 进行安全的路径拼接"""
    path = pathlib.Path(parts[0])
    for part in parts[1:]:
        path /= part.lstrip('/\\')
    return path

def update_progress(status_label, progress):
    global taskbar_mode, set_determin
    if set_determin:
        progress.config(mode='determinate', maximum=100, value=progress_value)
        set_determin = False
        progress.stop()
    if total_size and not stopping:
        if taskbar_mode != 2:
            taskbar_mode = 2
            taskbar_progress.set_mode(2)
        taskbar_progress.set_progress(int(downloaded_size), int(total_size))
        status_label.config(text=f'进度：{int(progress_value)}%（{downloaded_size / 1024 / 1024:.2f} MB，共 {total_size / 1024 / 1024:.2f} MB），剩余时间：{time.strftime("%H:%M:%S", time.gmtime(remaining_time))}')
        progress['value'] = progress_value
    if not taskbar_progress._is_init:
        taskbar_progress.init()

    if not stopping: root.after(100, lambda: update_progress(status_label, progress))

    if stopping:
        force_exit()
        return

def download_file(ftp_url, status_label):
    global trycount, stopping, daemon, set_determin, total_size
    try:
        # 解析 FTP URL
        url_parts = urldecoder.decode_ftp_url(ftp_url)
        port = url_parts['Port']
        server = url_parts['HostName']
        full_server_str = f"ftp://{server}"
        file_path = url_parts['UrlPath']
        username = url_parts['UserName']

        # 确保 settings 字典中包含所有必要的键
        if 'save_path' not in settings:
            settings['save_path'] = DEFAULT_SAVE_PATH
        if 'servers' not in settings:
            settings['servers'] = {}
        if 'timeout' not in settings:
            settings['timeout'] = 60
        if 'blocksize' not in settings:
            settings['blocksize'] = 8192
        if 'debug' not in settings:
            settings['debug'] = False

        print(settings['save_path'])

        # 创建本地文件夹结构
        local_dir = pathlib_join(settings['save_path'], server, os.path.dirname(file_path))
        os.makedirs(local_dir, exist_ok=True)
        local_filename = os.path.join(local_dir, os.path.basename(file_path))
        local_filename_temp = f"{local_filename}.ftpdownload"

        debug_print(f"准备下载文件到：{local_filename}")
        if os.path.exists(local_filename_temp):
            debug_print(f"{local_filename} 文件已在下载")
            os._exit(0)
            return

        if os.path.exists(local_filename):
            # 如果文件已经存在...
            choose = messagebox.askquestion('直接从本地打开或继续下载', '此文件已保存到本地。是否要直接打开本地文件？\n按“否”可继续从服务器下载文件。')
            if choose == 'yes':
                debug_print(f"文件已存在，打开文件：{local_filename}")
                os.startfile(local_filename)
                force_exit()
                return

            if choose == 'no':
                debug_print("文件已存在，但继续下载")
                os.remove(local_filename)

        blocksize = settings['blocksize']

        while True:
            trycount += 1
            try:
                global ftp
                refresh_ftp_object(server, port, timeout=settings['timeout'])  # 使用解析后的端口号
                try:
                    if full_server_str in settings['servers']:
                        ftp.login(settings['servers'][full_server_str]['username'], settings['servers'][full_server_str]['password'])
                        debug_print(f"登录成功：{settings['servers'][full_server_str]['username']}")
                    else:
                        ftp.login()
                except ftplib.error_perm as e:
                    ftp.close()
                    refresh_ftp_object(server, port, timeout=settings['timeout'])
                    log_in_ftp(server, port, username, timeout=settings['timeout'])

                try:
                    ftp.voidcmd('OPTS UTF8 ON')
                except Exception as e:
                    debug_print(e)

                # 获取文件大小
                ftp.voidcmd('TYPE I')
                total_size = ftp.size(file_path)

                downloaded_size = 0
                start_time = time.time()

                set_determin = True

                # 下载文件
                try:
                    daemon = True
                    with open(local_filename_temp, 'wb') as f:
                        def callback(data):
                            global progress_value, speed, elapsed_time, remaining_time, downloaded_size
                            f.write(data)
                            downloaded_size += len(data)
                            progress_value = (downloaded_size / total_size) * 100
                            elapsed_time = time.time() - start_time
                            speed = downloaded_size / elapsed_time
                            remaining_time = (total_size - downloaded_size) / speed
                            # update_status(downloaded_size, total_size, remaining_time, progress=(downloaded_size / total_size) * 100)

                        debug_print(f"开始下载文件：{file_path}")
                        sock = ftp.transfercmd('RETR ' + file_path) if f.tell() == 0 else \
                                   ftp.transfercmd('RETR ' + file_path, rest=f.tell())
                        while True:
                            if stopping:
                                break
                            block = sock.recv(blocksize)
                            if not block: break
                            callback(block)
                        
                        sock.close()

                        if total_size > f.tell():
                            # 没下完就跳出循环，肯定有问题
                            # 不过如果是因为上边要求取消的话就算了
                            if not stopping:
                                ftp.close()
                                continue

                        debug_print(f"文件下载完成：{local_filename}")
                except (Exception, ftplib.error_temp) as e:
                    # 如果下载失败，先再试几次
                    if trycount < 5: continue
                    # 删除未完成的文件并显示错误消息
                    os.remove(local_filename_temp)
                    messagebox.showerror('错误', str(e))
                    force_exit()
                    return

                if stopping:
                    os.remove(local_filename_temp)
                    return

                # 打开文件
                status_label.config(text='正在打开文件...')
                os.rename(local_filename_temp, local_filename)
                debug_print(f"打开文件：{local_filename}")
                os.startfile(local_filename)
                stopping = True
                break

            except (EOFError, ftplib.error_temp):
                ftp.close()
                continue

            finally:
                ftp.quit()

        # 关闭窗口
        root.quit()

    except Exception as e:
        debug_print(f'{e},download_file')
        force_exit()

def refresh_ftp_object(host, port, timeout):
    global ftp
    ftp = ftplib.FTP()
    ftp.connect(host, port, timeout)

def log_in_ftp(host, port, username, timeout):
    global taskbar_mode, daemon
    login_successful = False
    while not login_successful:  # 循环直到登录成功或取消
        if not taskbar_progress._is_init:
            taskbar_progress.init()
        if taskbar_mode != 8:
            taskbar_mode = 8
            taskbar_progress.set_mode(8)
        
        # 如果需要用户名和密码，弹出对话框让用户输入
        daemon = True
        dialog = tk.Toplevel()
        dialog.title('请输入用户名和密码')

        namestr = tk.StringVar()
        namestr.set(username)
        ttk.Label(dialog, text='用户名：').grid(row=0, column=0, padx=20, pady=10)
        username_entry = ttk.Entry(dialog, textvariable=namestr)
        username_entry.grid(row=0, column=1, padx=20, pady=10)

        ttk.Label(dialog, text='密码：').grid(row=1, column=0, padx=20, pady=10)
        password_entry = ttk.Entry(dialog, show='*')
        password_entry.grid(row=1, column=1, padx=20, pady=10)

        def on_ok(event=None):
            global ftp, daemon
            daemon = False
            username = username_entry.get()
            password = password_entry.get()
            try:
                ftp.login(username, password)  # 尝试登录
                debug_print(f"登录成功：{username}")
                nonlocal login_successful
                login_successful = True
                dialog.destroy()
            except ftplib.error_perm as e:  # 如果登录失败，显示错误消息
                messagebox.showerror('错误', f'用户名或密码错误，请重试。错误信息：{e}')
                ftp.close()
                refresh_ftp_object(host, port, timeout)
            except EOFError:
                messagebox.showerror('错误', '服务器关闭了连接，请检查凭据或服务器状态。')
                ftp.close()
                refresh_ftp_object(host, port, timeout)

        def on_cancel():
            global stopping
            daemon = False
            dialog.destroy()
            stopping = True
            return

        password_entry.bind("<Return>", on_ok)
        ttk.Button(dialog, text='确定', command=on_ok, style='TButton').grid(row=2, column=0, padx=20, pady=10)
        ttk.Button(dialog, text='取消', command=on_cancel, style='TButton').grid(row=2, column=1, padx=20, pady=10)

        dialog.wait_window()

        if not login_successful:
            return

def keep_alive():
    while True:
        if daemon:
            try:
                ftp.voidcmd('NOOP')
            except Exception as e:
                messagebox.showerror(title="连接已断开", message=f"由于某些原因，连接已断开：{e}\r您的操作可能过慢，或者当前的网络不稳定。您或许可以点击“取消”关闭当前程序，然后再试一次。")
        time.sleep(30)

def download_and_open_file(ftp_url):
    from easy_progressbar import EasyProgressBar

    global taskbar_progress
    taskbar_progress = EasyProgressBar()

    # 创建 GUI
    global root
    root = tk.Tk()
    set_DPI(root)
    root.title('正在下载并打开文件...')
    root.resizable(False, False)  # 设置窗口不可调整大小
    
    status_label = ttk.Label(root, text='请稍候，正在下载文件。完成后文件将自动打开。')
    status_label.pack(padx=20, pady=10)
    progress = ttk.Progressbar(root, length=100, mode='indeterminate', maximum=100, value=0)
    progress.pack(padx=20, pady=10, fill='x', expand=True)
    progress.start()
    status_label = ttk.Label(root, text='进度：0%（0 MB，共 0 MB），剩余时间：正在计算...')
    status_label.pack(padx=20, pady=10)

    def cancel_download():
        global stopping, daemon
        root.title('正在取消...')
        status_label.config(text='正在取消下载。无需尝试关闭此窗口，我们会自动关闭它。您可以最小化。')
        stopping = True
        daemon = False

    cancel_button = ttk.Button(root, text='取消', command=cancel_download)
    cancel_button.pack(padx=20, pady=10)

    def init_taskbar_status():
        global taskbar_mode
        taskbar_progress.init()
        taskbar_mode = 1
        taskbar_progress.set_mode(1)

    # 启动下载线程
    if settings['daemon']: threading.Thread(target=keep_alive).start()
    threading.Thread(target=download_file, args=(ftp_url, status_label), daemon=True).start()

    # 显示 GUI

    if not settings['oldui']: make_it_winui(root)
    set_DPI(root)

    root.protocol("WM_DELETE_WINDOW", cancel_download)

    root.iconbitmap(ICO_PATH)

    root.after(100, init_taskbar_status)

    root.after(100, lambda: update_progress(status_label, progress))

    root.mainloop()

def add_server_entry(frame, server_entries):
    row = len(frame.grid_slaves()) // 4
    ttk.Label(frame, text='服务器：').grid(row=row, column=0, padx=20, pady=10)
    server_entry = ttk.Entry(frame)
    server_entry.grid(row=row, column=1, padx=20, pady=10)
    server_entries.append(server_entry)

    ttk.Label(frame, text='用户名：').grid(row=row, column=2, padx=20, pady=10)
    username_entry = ttk.Entry(frame)
    username_entry.grid(row=row, column=3, padx=20, pady=10)
    server_entries.append(username_entry)

    ttk.Label(frame, text='密码：').grid(row=row, column=4, padx=20, pady=10)
    password_entry = ttk.Entry(frame, show='*')
    password_entry.grid(row=row, column=5, padx=20, pady=10)
    server_entries.append(password_entry)

    remove_button = ttk.Button(frame, text='移除', command=lambda: remove_server_entry(frame, server_entries, row), style='TButton')
    remove_button.grid(row=row, column=6, padx=20, pady=10)
    server_entries.append(remove_button)

def remove_server_entry(frame, server_entries, row):
    for widget in frame.grid_slaves(row=row):
        widget.grid_forget()
    del server_entries[row*4:row*4+4]

def create_gui():
    root = tk.Tk()
    set_DPI(root)
    root.title('FTP 即点即开（预览）')
    root.resizable(False, False)

    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    ttk.Label(main_frame, text="设置", font=("Microsoft YaHei UI", 40, "bold")).grid(row=0, column=0, padx=20, pady=12)

    ttk.Label(main_frame, text='保存路径：').grid(row=1, column=0, padx=20, pady=10)
    save_path_entry = ttk.Entry(main_frame)
    save_path_entry.grid(row=1, column=1, padx=20, pady=10, sticky="ew")
    save_path_entry.insert(0, settings['save_path'])

    ttk.Button(main_frame, text='浏览...', command=lambda: browse_save_path(save_path_entry)).grid(row=1, column=2, padx=20, pady=10)

    ttk.Separator(main_frame).grid(row=2, column=0, columnspan=4, padx=(20, 10), pady=10, sticky="ew")

    server_frame = ttk.Frame(main_frame, padding="10")
    server_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))

    server_entries = []
    for server, creds in settings['servers'].items():
        add_server_entry(server_frame, server_entries)
        server_entries[-4].insert(0, server)
        server_entries[-3].insert(0, creds['username'])
        server_entries[-2].insert(0, creds['password'])

    ttk.Button(main_frame, text='添加服务器', command=lambda: add_server_entry(server_frame, server_entries)).grid(row=4, column=0, columnspan=3, padx=20, pady=10)

    ttk.Separator(main_frame).grid(row=5, column=0, columnspan=4, padx=(20, 10), pady=10, sticky="ew")

    ttk.Label(main_frame, text='超时时间（秒）：').grid(row=6, column=0, padx=20, pady=10)
    timeout_entry = ttk.Entry(main_frame)
    timeout_entry.grid(row=6, column=2, columnspan=2, padx=20, pady=10)
    timeout_entry.insert(0, settings['timeout'])

    ttk.Separator(main_frame).grid(row=7, column=0, columnspan=4, padx=(20, 10), pady=10, sticky="ew")

    ttk.Label(main_frame, text='传输块大小：').grid(row=8, column=0, padx=20, pady=10)
    blocksize_entry = ttk.Entry(main_frame)
    blocksize_entry.grid(row=8, column=2, columnspan=2, padx=20, pady=10)
    blocksize_entry.insert(0, settings['blocksize'] or 8192)

    ttk.Separator(main_frame).grid(row=9, column=0, columnspan=4, padx=(20, 10), pady=10, sticky="ew")

    oldui_var = tk.BooleanVar(value=settings['oldui'])
    ttk.Checkbutton(main_frame, text='回退到传统外观（重新打开生效）', style="Switch.TCheckbutton", variable=oldui_var).grid(row=10, column=0, columnspan=3, padx=20, pady=10)

    ttk.Separator(main_frame).grid(row=11, column=0, columnspan=4, padx=(20, 10), pady=10, sticky="ew")

    daemon_var = tk.BooleanVar(value=settings['daemon'])
    ttk.Checkbutton(main_frame, text='保持连接打开（推荐）', style="Switch.TCheckbutton", variable=daemon_var).grid(row=12, column=0, columnspan=3, padx=20, pady=10)

    ttk.Separator(main_frame).grid(row=13, column=0, columnspan=4, padx=(20, 10), pady=10, sticky="ew")

    debug_var = tk.BooleanVar(value=settings['debug'])
    ttk.Checkbutton(main_frame, text='调试模式', style="Switch.TCheckbutton", variable=debug_var).grid(row=14, column=0, columnspan=3, padx=20, pady=10)

    ttk.Button(main_frame, text='保存设置', style="Accent.TButton", command=lambda: save_settings(save_path_entry.get(), server_entries, int(timeout_entry.get()), int(blocksize_entry.get()), oldui_var.get(), daemon_var.get(), debug_var.get())).grid(row=15, column=1, padx=20, pady=10)
    ttk.Button(main_frame, text='关于', command=about_me).grid(row=15, column=2, padx=20, pady=10)

    if not settings['oldui']: make_it_winui(root)
    set_DPI(root)

    root.iconbitmap(ICO_PATH)

    root.mainloop()

def browse_save_path(entry):
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def force_exit():
    global stopping, taskbar_mode
    try:
        debug_print('尝试退出')
        def stop_connect():
            try:
                ftp.abort()
            except Exception:
                ftp.close()
        if ftp and not stopping:
            stopping = True
            debug_print('Yes')
            threading.Thread(target=stop_connect, daemon=True).start()
        taskbar_progress.set_mode(0)
        taskbar_mode = 0
        taskbar_progress.end()
        root.quit()
        root.destroy()
    except Exception as e:
        debug_print(f'{e},尝试退出')
    debug_print('n')
    os._exit(0)
    return

def about_me():
    messagebox.showinfo(title="关于", message="FTPC2O 版本 PlaceHolder\n开源项目。欢迎参与此项目：zDaleZ\FTPC2O")

if __name__ == '__main__':
    load_settings()
    if len(sys.argv) > 1:
        match sys.argv[1]:
            case "install":
                import configuration
                configuration.elevate_privileges(sys.argv[1])
                if configuration.install_registry_entries():
                    print("安装完成！")
                else:
                    print("安装过程中出现错误。")
            case "uninstall":
                import configuration
                configuration.elevate_privileges(sys.argv[1])
                if configuration.uninstall_registry_entries():
                    print("卸载完成！")
                else:
                    print("卸载过程中出现错误。")
            case "about":
                about_me()
            case _:
                download_and_open_file(sys.argv[1])
    else:
        create_gui()
