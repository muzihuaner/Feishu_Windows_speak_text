# 飞书消息弹窗语音通知程序
# 作者：李欢
# Github:https://github.com/muzihuaner/Feishu_Windows_speak_text
import time
import psutil 
import pywinauto
from pywinauto import Application
import pyttsx3
import sys
import ctypes
#系统兼容性处理
if sys.getwindowsversion().major < 10:
    import ctypes
    ctypes.windll.user32.MessageBoxW(0, "需要Windows 10及以上系统", "系统不兼容", 0x10)
    sys.exit()


#权限处理
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
    while True:
        try:
            windows = get_feishu_windows()
            for window in windows:
                if window.handle not in processed_windows:
                    # 更精确的弹窗识别条件
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

if __name__ == "__main__":
    monitor_feishu_popups()
