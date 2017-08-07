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

  ResoStr:TResourceStream;//��Դ��
  MemyStr:TmemoryStream;//�ڴ���

  CmOffset:Integer;
  BufferS:Pchar;
begin

  Num:=5;

  SetLength(SerVal,Num+1);
  SetLength(SerTxt,Num+1);

  CMser:='CMser';  //�ļ�����
  CmOffset:=548268;  //��ʼ��

  if SerUrl.Text='' then begin
	     showmessage('������д�����ַ!');
	   end
  else
	   begin
     
       {��ס��ť}
       SerUrl.Enabled:=false;
       ShengCheng.Enabled:=false;

	     {�����ڴ���}
       MemyStr:=TmemoryStream.Create;
       {������ڳ����ļ�����ȡ�ļ����������ȡ��Դ��}
       if FileExists(CMser+'.dat') then begin
          MemyStr.LoadFromFile(CMser+'.dat');
          end
       else
          begin
          //����Դ�ͷų����������ڴ�����;1,������Դ��ģ��2,��Դ��ʵ,3��Դ����
          ResoStr:=TResourceStream.Create(Hinstance,'ResFlag',RT_RCDATA);
          {����Դ�������ڴ������޸�}
          ResoStr.SaveToStream(MemyStr);
          ResoStr.Free;
       end;


       {������ʾ��Ϣ}
       SerVal[0]:='{TIP}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {������ַurl}
       SerVal[1]:='{URL}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {�Ƿ񴴽����ļ�����}
       SerVal[2]:='{FIL}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {����key}
       SerVal[3]:='{KEY}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {�Ƿ񿪻�����}
       SerVal[4]:='{RUN}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';
       {�Ƿ�ƻ�����}
       SerVal[5]:='{RAT}ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ';



       {ȥ�����ҿո�}
       SerTxt[0]:=Trim(SerTip.Text);
       SerTxt[1]:=Trim(SerUrl.Text);
       SerTxt[2]:=Trim(SerFil.Text);
       {����key}
       SerTxt[3]:=Trim(SerKey.Text);
       {�Ƿ񿪻�����}
       SerTxt[4]:=GetChecked(SerRun.Checked);
       {�Ƿ񿪻�����}
       SerTxt[5]:=GetChecked(SerRat.Checked);

       

       //ָ��һ�ַ���,�ַ���������ΪSerUrl.text��#0(����ո�) �ո�ĸ���Ϊ7-����ĳ�
       for I:=0 to Num do
         begin
         BufferS:=pchar(SerTxt[I]+stringofchar(#0,length(SerVal[I])-length(SerTxt[I])));
         CmOffset:=OffsetStream(MemyStr,SerVal[I]); //��λҪ�޸ĵ��ַ�����λ��,��ͷ����
         if CmOffset>0 then
         begin
            MemyStr.Seek(CmOffset,soFromBeginning);
            MemyStr.WriteBuffer(BufferS^,length(SerVal[I]));
         end
       end;

       {д��Urlbuffer�е�����,������ƶ�7���ֽ�λ��,����}
       MemyStr.SaveToFile('CMSer.exe');
       MemyStr.Free;
       Showmessage('��������ɳɹ�!');
       {������ť}
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


//�ж��Ƿ��Ѿ����
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
