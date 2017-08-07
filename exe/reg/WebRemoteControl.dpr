program WebRemoteControl;

uses
  Windows,Forms,
  Unit1 in 'Unit1.pas' {FormBox},
  Unit_Helper in 'Unit_Helper.pas',
  Unit_WinSys in 'Unit_WinSys.pas',
  Unit_Tasklist in 'Unit_Tasklist.pas',
  Unit_Http in 'Unit_Http.pas',
  Unit_Create in 'Unit_Create.pas',
  Unit_Funny in 'Unit_Funny.pas';

{$R *.res}

var
 myMutex:HWND;
begin
  //CreateMutex����������󣬲��Ҹ����������һ��Ψһ�����֡�
  myMutex:=CreateMutex(nil,false,'CmSerOneCopy');
  //����û�б����й�
  if WaitForSingleObject(myMutex,0)<>wait_TimeOut then
  begin
   Application.Initialize;
   Application.Title := '';
   Application.ShowMainForm:=false;
   Application.CreateForm(TFormBox, FormBox);
   Application.Run;
  End
end.
