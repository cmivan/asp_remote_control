Echo Off
title Delphi CmServer Брвы v1.0
@mode con cols=115 lines=8
color 9f 
cls
     
echo make by cmivan

del WebRemote.Generator\CMSer.RES

echo ResFlag RCDATA WebRemoteControl.exe>CmSer.rc

Rem path=C:\Program Files\Borland\Delphi7\Bin
path=E:\Soft\Delphi7\Borland\Delphi7\Bin

Brcc32 CmSer.rc

del CmSer.rc

copy CmSer.RES WebRemote.Generator\CMSer.RES

del CMSer.RES
pause