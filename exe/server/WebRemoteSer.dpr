program WebRemoteSer;

uses
  SvcMgr,
  Unit1 in 'Unit1.pas' {Service1: TService},
  Unit_Helper in 'Unit_Helper.pas',
  Unit_WinSys in 'Unit_WinSys.pas',
  Unit_Tasklist in 'Unit_Tasklist.pas',
  Unit_Http in 'Unit_Http.pas',
  Unit_Create in 'Unit_Create.pas',
  Unit_Funny in 'Unit_Funny.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TService1, Service1);
  Application.Run;
end.
