@echo off
@echo ������������������������������������
@echo �� �����Ż����� ������� Bate 1.0 ��
@echo �� For : ������(fly^&idea)       ��
@echo �� By  : cm.ivan ^@��mi����        ��
@echo ������������������������������������
@ping -n 2 127.1>nul
@echo �������������...
@echo ----------------------------------
@taskkill /f /im WebRemoteControl.exe
@echo.
@echo.

@ping -n 1 127.1>nul
@echo ���ڴ�����װĿ¼...
@echo ----------------------------------
@ping -n 1 127.1>nul
mkdir C:\WINDOWS\WebRemoteSer
@echo.
@echo.

@ping -n 1 127.1>nul
@echo ���ڸ����ļ�����װĿ¼...
@echo ----------------------------------
@ping -n 1 127.1>nul
@copy WebRemoteControl.exe C:\WINDOWS\WebRemoteSer\WebRemoteControl.exe
@echo.
@echo.

@ping -n 1 127.1>nul
@echo ������������...
@echo ----------------------------------
@REG ADD HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run /v WebRemoteControl.exe /t REG_SZ /d C:\WINDOWS\WebRemoteSer\WebRemoteControl.exe /f
@echo.
@echo.


@ping -n 1 127.1>nul
@echo ������������...
@echo ----------------------------------
@ping -n 1 127.1>nul
start C:\WINDOWS\WebRemoteSer\WebRemoteControl.exe
@echo.
@echo.

pause