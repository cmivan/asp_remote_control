program Generator;

uses
  Windows,Forms,
  Unit1 in 'Unit1.pas' {Ser};

{$R *.res}
{$r CMSer.RES}

var
 myMutex:HWND;
begin
  //CreateMutex����������󣬲��Ҹ����������һ��Ψһ�����֡�
  myMutex:=CreateMutex(nil,false,'WebRemoteControlSerOneCopy');
  //����û�б����й�
  if WaitForSingleObject(myMutex,0)<>wait_TimeOut then
  begin
   Application.Initialize;
   Application.CreateForm(TSer, Ser);
   Application.Run;
  End
end.
