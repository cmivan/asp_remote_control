unit Unit_Create;

interface

uses
  SysUtils, UrlMon;


Function CREATE_FILE(Paths:String;Texts:String):String;
Function CREATE_DOWN(Path:String;Url:String):String;
Function CREATE_PATH(Path:String):String;

implementation

//Create file
Function CREATE_FILE(Paths:String;Texts:String):String;
Var F:TextFile;
Begin
  AssignFile(F,Paths);
  IF FileExists(Paths) Then
     Append(F)
  Else
     ReWrite(F);
  try
    Writeln(F,Texts);
  finally
    CloseFile(F); 
  end;
  Result:= '{"ok":"1"}';
End;

//Download to create file
Function CREATE_DOWN(Path:String;Url:String):String;
var Rback:Boolean;
begin 
  try 
    Rback := UrlDownloadToFile(nil, PChar(Url), PChar(Path), 0, nil) = 0;
    if Rback then
      begin
        Result:= '{"ok":"1"}';
      end
    else
      begin
        Result:= '{"ok":"0"}';
    end;
  except
    Result:= '{"ok":"0"}';
  end;
end;

//Create folder
Function CREATE_PATH(Path:String):String;
Begin
  ForceDirectories(Path);
  Result:= '{"ok":"1"}';
End;

end.
