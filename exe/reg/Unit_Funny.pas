unit Unit_Funny;

interface

uses
  Windows,Messages,ShellAPI;

Function TO_BLACK:String;
Function OPEN_WEB(Url:String):String;
Function OPEN_MSG(Msg:String):String;

implementation

uses Unit_Helper;

//Black Screen
Function TO_BLACK:String;
Begin
  //Close the monitor
  SendMessage(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, 2);
  Sleep(1000);
  Result:= TO_JSON('TO_BLACK','1');
End;

//Open the Web
Function OPEN_WEB(Url:String):String;
var Urls:PAnsiChar;
Begin
   if Url='' then
   begin
      Url:='about:blank';
   end;
   Urls:=PAnsiChar(Url); //Use the default browser to open
   ShellExecute(SW_SHOWMAXIMIZED, 'open', 'Explorer.exe', Urls, nil, SW_SHOWNORMAL);
   Result:= '{"ok":"1"}';
End;

//MessageBox
Function OPEN_MSG(Msg:String):String;
var Msgs:PAnsiChar;
Begin
   if Msg<>'' then
   begin
      Msgs:=PAnsiChar(Msg);
      MessageBox(0,Msgs,'消息提示',MB_OK);
   end;
   Result:= '{"ok":"1"}';
End;

end.
