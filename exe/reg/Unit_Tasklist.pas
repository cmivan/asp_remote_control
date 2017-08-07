unit Unit_Tasklist;

interface

uses
  Windows,SysUtils,Classes,TLHelp32;

Function TASK_LIST:String;
Function TASK_KILL(ExeFileName:String):String;


implementation

//Get ther Tasklist
Function TASK_LIST:String;
var
  ProcessName:string;
  ProcessID:string;
  ContinueLoop:Bool;
  FSnapshotHandle:THandle;
  FProcessEntry32:TProcessEntry32;
  TasklistStr:String;
Begin
  TasklistStr:='id,name';
  FSnapshotHandle:=CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS,0);
  FProcessEntry32.dwSize:=Sizeof(FProcessEntry32);
  ContinueLoop:=Process32First(FSnapshotHandle,FProcessEntry32);
  while ContinueLoop  do
  Begin
    ProcessName := FProcessEntry32.szExeFile;
    ProcessID := inttostr(FProcessEntry32.th32ProcessID);
    TasklistStr:= TasklistStr+'|' + ProcessID + ',' + ProcessName;
    ContinueLoop:=Process32Next(FSnapshotHandle,FProcessEntry32);
  End;
  CloseHandle(FSnapshotHandle);
  Result:= TasklistStr;
End;


//Kill tasklist
Function TASK_KILL(ExeFileName:String):String;
const
PROCESS_TERMINATE = $0001;
var
ContinueLoop: BOOL;
FSnapshotHandle: THandle;
FProcessEntry32: TProcessEntry32;
begin
Result:= '0';
FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
while Integer(ContinueLoop) <> 0 do
begin
if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) = UpperCase(ExeFileName))) then
   Result := BoolToStr(TerminateProcess(OpenProcess(PROCESS_TERMINATE,BOOL(0),FProcessEntry32.th32ProcessID),0));
   ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
end;
CloseHandle(FSnapshotHandle);
End;



end.
