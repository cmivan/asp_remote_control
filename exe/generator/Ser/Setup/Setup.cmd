@echo off
@echo ������������������������������������
@echo �� �����Ż����� ������� Bate 1.0 ��
@echo �� For : ������(fly^&idea)       ��
@echo �� By  : cm.ivan ^@��mi����        ��
@echo ������������������������������������
@ping -n 2 127.1>nul
@echo �������������...
@echo ----------------------------------
@taskkill /f /im WebRemoteSer.exe
@echo.
@echo.

@ping -n 2 127.1>nul
@echo �����������...
@echo ----------------------------------
@ping -n 2 127.1>nul
@sc delete RemoteSer
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
@copy WebRemoteSer.exe C:\WINDOWS\WebRemoteSer\WebRemoteSer.exe
@echo.
@echo.

@ping -n 1 127.1>nul
@echo ���ڰ�װ����...
@echo ----------------------------------
@ping -n 1 127.1>nul
@C:\WINDOWS\WebRemoteSer\WebRemoteSer.exe /install
@echo.
@echo.

@ping -n 1 127.1>nul
@echo ������������...
@echo ----------------------------------
@ping -n 1 127.1>nul
@net start RemoteSer
@echo.
@echo.

@ping -n 1 127.1>nul
@echo ��װ���!