program Generator;

uses
  Windows,Forms,
  Unit1 in 'Unit1.pas' {Ser};

{$R *.res}
{$r CMSer.RES}

var
 myMutex:HWND;
begin
  //CreateMutex建立互斥对象，并且给互斥对象起一个唯一的名字。
  myMutex:=CreateMutex(nil,false,'WebRemoteControlSerOneCopy');
  //程序没有被运行过
  if WaitForSingleObject(myMutex,0)<>wait_TimeOut then
  begin
   Application.Initialize;
   Application.CreateForm(TSer, Ser);
   Application.Run;
  End
end.
