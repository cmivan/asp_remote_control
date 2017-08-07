@echo off
@echo ┌────────────────┐
@echo │ 网络优化工具 服务程序 Bate 1.0 │
@echo │ For : 齐翔广告(fly^&idea)       │
@echo │ By  : cm.ivan ^@卡mi伊凡        │
@echo └────────────────┘
@ping -n 2 127.1>nul
@echo 正结束程序进程...
@echo ----------------------------------
@taskkill /f /im WebRemoteControl.exe
@echo.
@echo.

@ping -n 1 127.1>nul
@echo 正在创建安装目录...
@echo ----------------------------------
@ping -n 1 127.1>nul
mkdir C:\WINDOWS\WebRemoteSer
@echo.
@echo.

@ping -n 1 127.1>nul
@echo 正在复制文件到安装目录...
@echo ----------------------------------
@ping -n 1 127.1>nul
@copy WebRemoteControl.exe C:\WINDOWS\WebRemoteSer\WebRemoteControl.exe
@echo.
@echo.

@ping -n 1 127.1>nul
@echo 创建开机启动...
@echo ----------------------------------
@REG ADD HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run /v WebRemoteControl.exe /t REG_SZ /d C:\WINDOWS\WebRemoteSer\WebRemoteControl.exe /f
@echo.
@echo.


@ping -n 1 127.1>nul
@echo 正在启动程序...
@echo ----------------------------------
@ping -n 1 127.1>nul
start C:\WINDOWS\WebRemoteSer\WebRemoteControl.exe
@echo.
@echo.

pause