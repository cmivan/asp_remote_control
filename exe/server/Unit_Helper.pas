unit Unit_Helper;

interface
uses
  Windows,SysUtils,Classes;

Function RE_PATH(PATH:String):String;
Function RE_FILE(PATH:String):String;
Function TO_JSON(cmd:String;back:String):String;
Function TEXT_SPLIT(s,s1:string):TStringList;
Function TO_CMD(cmd:String):String;


implementation


//Processing the directory path
Function RE_PATH(PATH:String):String;
var Dirs:String;
begin
  Dirs:=StringReplace(PATH ,'/', '\', [rfReplaceAll, rfIgnoreCase]);
  if Dirs[length(Dirs)]<>'\' then Dirs:=Dirs + '\';
  Dirs:=StringReplace(Dirs ,'\\', '\', [rfReplaceAll, rfIgnoreCase]);
  Result:=Dirs;
End;


//Processing the file path
Function RE_FILE(PATH:String):String;
var Dirs:String;
begin
  Dirs:=StringReplace(PATH ,'/', '\', [rfReplaceAll, rfIgnoreCase]);
  Dirs:=StringReplace(Dirs ,'\\', '\', [rfReplaceAll, rfIgnoreCase]);
  Result:=Dirs;
End;


//Converted to json
Function TO_JSON(cmd:String;back:String):String;
var TempS:String;
begin
    TempS := StringReplace(back ,#13#10, '{br}', [rfReplaceAll, rfIgnoreCase]);
    Result:= '{"cmd":"' + cmd + '","result":"' + TempS + '"}';
End;


//String Split
Function TEXT_SPLIT(s,s1:string):TStringList;
begin
Result:=TStringList.Create;
while Pos(s1,s)>0 do
begin
  Result.Add(Copy(s,1,Pos(s1,s)-1));
  Delete(s,1,Pos(s1,s));
end;
Result.Add(s);
End;


//Analysis and implementation
Function TO_CMD(cmd:String):String;
begin
End;


end.
