# 飞书消息弹窗语音通知程序
# 作者：李欢
# Github:https://github.com/muzihuaner/Feishu_Windows_speak_text
import time
import psutil 
import pywinauto
import sys
import ctypes
import os
import threading
from pywinauto import Application
import pyttsx3
from pystray import MenuItem as item, Icon, Menu
from PIL import Image

# 全局控制变量
exit_event = threading.Event()
monitor_thread = None

# 系统兼容性处理
if sys.getwindowsversion().major < 10:
    ctypes.windll.user32.MessageBoxW(0, "需要Windows 10及以上系统", "系统不兼容", 0x10)
    sys.exit()

# 权限处理
if ctypes.windll.shell32.IsUserAnAdmin() == 0:
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

# 初始化语音引擎
engine = pyttsx3.init()

# 配置参数
target_process = "Feishu.exe"  # 飞书进程名
processed_windows = set()      # 已处理窗口句柄

def get_feishu_windows():
    feishu_windows = []
    for window in pywinauto.Desktop(backend="uia").windows():
        try:
            # 通过进程ID获取进程名
            pid = window.process_id()
            process = psutil.Process(pid)
            if process.name().lower() == target_process.lower():
                feishu_windows.append(window)
        except (psutil.NoSuchProcess, PermissionError):
            continue  # 忽略已关闭的进程
    return feishu_windows

def speak_text(text):
    engine.say(text)
    engine.runAndWait()

def monitor_feishu_popups():
    global processed_windows
    print("开始监控飞书弹窗...")
    while not exit_event.is_set():
        try:
            windows = get_feishu_windows()
            for window in windows:
                if window.handle not in processed_windows:
                    if window.class_name() == "Chrome_WidgetWin_1":
                        content = window.window_text()
                        print("检测到弹窗内容：", content)
                        speak_text(f"您有新的运维通知，请及时处理：{content}")
                        processed_windows.add(window.handle)
            
            # 清理无效句柄
            current_handles = {w.handle for w in windows}
            processed_windows = processed_windows & current_handles
            
        except Exception as e:
            print("监控异常：", str(e))
        
        time.sleep(2)
    print("监控服务已停止")

def on_exit(icon, item):
    global engine
    exit_event.set()
    engine.stop()
    icon.stop()
    os._exit(0)

def on_about(icon, item):
    ctypes.windll.user32.MessageBoxW(0, 
        "飞书消息语音通知程序\n\n版本：1.0\n作者：李欢\nGithub: https://github.com/muzihuaner", 
        "关于", 0x40)

def setup_tray_icon():
    # 加载图标文件
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    icon_path = os.path.join(base_path, 'feishu.ico')
    
    # 创建系统托盘图标
    image = Image.open(icon_path)
    menu = Menu(
        item('关于', on_about),
        item('退出', on_exit)
    )
    icon = Icon('Feishu Notifier', image, menu=menu)
    icon.run()

if __name__ == "__main__":
    # 启动监控线程
    monitor_thread = threading.Thread(target=monitor_feishu_popups, daemon=True)
    monitor_thread.start()
    
    # 启动系统托盘
    setup_tray_icon()
