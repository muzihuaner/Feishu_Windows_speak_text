@echo off
pyinstaller --onefile --noconsole ^
--name=FeishuNotifier ^
--add-data="feishu.ico;." ^
--hidden-import=pyttsx3.drivers ^
--hidden-import=pyttsx3.drivers.sapi5 ^
--hidden-import=pystray._win32 ^
--hidden-import=win32timezone ^
feishu_notifier.py