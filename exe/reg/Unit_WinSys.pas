unit Unit_WinSys;

interface
uses
  Windows,SysUtils,Classes,nb30,TLHelp32,ShellAPI, Registry;

Function GET_HOST:String;
Function GET_MAC:String;
Function GET_DISK:String;
Function GET_RPATH(Dir: String):String;
Function RUN_FILE(Handle:HWND;Path:String):String;
Function SET_AUTORUN(RTitle,RPath:String):String;

implementation

uses Unit_Helper;


//Get computer name
Function GET_HOST:String;
var
  ComputerName: array[0..MAX_COMPUTERNAME_LENGTH+1] of char;
  Size: Cardinal;
begin
  Size := MAX_COMPUTERNAME_LENGTH+1;
  GetComputerName(ComputerName, Size);
  Result:= StrPas(ComputerName);
end;


//Get Mac
Function GET_MAC:String;
var ncb:TNCB;
   status:TAdapterStatus;
   lanenum:TLanaEnum;
   procedure ResetAdapter (num : char);
   begin
     fillchar(ncb,sizeof(ncb),0);
     ncb.ncb_command:=char(NCBRESET);
     ncb.ncb_lana_num:=num;
     Netbios(@ncb);
   end;
var
   i:integer;
   lanNum : char;
   address : record
   part1 : Longint;
   part2 : Word;
end absolute status;
begin
   Result:='';
   fillchar(ncb,sizeof(ncb),0);
   ncb.ncb_command:=char(NCBENUM);
   ncb.ncb_buffer:=@lanenum;
   ncb.ncb_length:=sizeof(lanenum);
   Netbios(@ncb);
   if lanenum.length=#0 then exit;
   lanNum:=lanenum.lana[0];
   ResetAdapter(lanNum);
   fillchar(ncb,sizeof(ncb),0);
   ncb.ncb_command:=char(NCBASTAT);
   ncb.ncb_lana_num:=lanNum;
   ncb.ncb_callname[0]:='*';
   ncb.ncb_buffer:=@status;
   ncb.ncb_length:=sizeof(status);
   Netbios(@ncb);
   ResetAdapter(lanNum);
   for i:=0 to 5 do
   begin
     result:=result+inttoHex(integer(Status.adapter_address[i]),2);
     if (i<5) then
     result:=result+'-';
   end;
   Result:= result;
end;



//Getting the Driver
Function GET_DISK:String;
var
    buf:array [0..MAX_PATH-1] of char;
    I:Integer;
    I_Result:Integer;
    DISK:string;
    DISKS:string;
begin
    I_Result:=GetLogicalDriveStrings(MAX_PATH,buf);
    for I:=0 to (I_Result div 4)-1 do
    begin
      DISK:= string(buf[I*4]+buf[I*4+1]+buf[I*4+2]);
      DISKS:= DISKS + DISK;
    end;
    Result:= DISKS;
end;



//Get files from the Folder
Function GET_RPATH(Dir: String):String;
var sr:TSearchRec;
var str1,str2:String;
var Dirs:String;
begin
  Dirs:=RE_PATH(Dir);
  Dirs:=ExtractFilePath(Dirs);
  Dirs:=Trim(Dirs+'*.*');
  //FormBox.Edit1.Text:= Dirs;
  str1:='0,0';str2:='0,0';
  Dirs:=Trim(Dirs);
  if FindFirst(Dirs,faAnyFile,sr) = 0 then
  begin
    repeat
      if ((sr.Attr and faDirectory) <> 0) then
      begin
        if ((sr.Name<> '.') and (sr.Name <> '..')) then  //Folder
        str1:=str1 + '|' + sr.Name+','+intToStr(sr.Attr);
      end;
      if ((sr.Attr and faDirectory)=0) then //file
      str2:=str2 + '|' + sr.Name+','+intToStr(sr.Size);
    until FindNext(sr) <> 0;
    FindClose(sr);
  end;
  //FormBox.Memo1.text:=str1+'{:}'+str2;
  Result:=str1+'{:}'+str2;
end;



//Run file
Function RUN_FILE(Handle:HWND;Path:String):String;
begin
    ShellExecute(Handle,'Open',PChar(Path),nil,nil,SW_SHOWNORMAL);
    Result:= '{"ok":"1"}';
end;



//Reg run
Function SET_AUTORUN(RTitle,RPath:String):String;
var
    hReg: TRegIniFile;
begin
    hReg := TRegIniFile.Create(''); //TregIniFile类的对象需要创建
    hReg.RootKey := HKEY_LOCAL_MACHINE; //设置根键,程序名称，可以为自定义值
    hReg.WriteString('Software\Microsoft\Windows\CurrentVersion\Run' + #0,RTitle,RPath);
    //命令行数据，必须为该程序的绝对路径＋程序完整名称
    hReg.destroy; //释放创建的hReg
    Result:= '{"ok":"1"}';
end;




end.
