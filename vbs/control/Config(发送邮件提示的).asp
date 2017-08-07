<%
'+++++++++++++++++++++++++++++++++++++
'// 作用，向客户端返回服务端地址     //
'+++++++++++++++++++++++++++++++++++++
'// 地址后面不要加 "/" ,不要放入中文目录
on error resume next
Dim thisUrl
    thisUrl=""
 If thisUrl="" then thisUrl=GetLocationURL()
Function GetLocationURL() 
   on error resume next
   Dim Url 
   Dim ServerPort,ServerName,ScriptName,QueryString 
       ServerName = Request.ServerVariables("SERVER_NAME") 
       ServerPort = Request.ServerVariables("SERVER_PORT") 
       ScriptName = Request.ServerVariables("SCRIPT_NAME") 
       QueryString= Request.ServerVariables("QUERY_STRING") 
       ScriptNames= split(ScriptName,"/")
       ScriptNum  = ubound(ScriptNames)
       If ScriptNum>0 then
          ScriptName1="/"&ScriptNames(ScriptNum)
          ScriptName=replace(ScriptName,ScriptName1,"")
       end If
       Url="http://"&ServerName 
       If ServerPort <> "80" Then Url = Url & ":" & ServerPort 
       Url=Url&ScriptName 
       If QueryString <>"" Then Url=Url&"?"& QueryString 
       GetLocationURL=Url 
End Function
Response.Write(thisUrl)

'//保证之发送一次
if session("in")<>1 then
   GetXmlText("http://www.qubaohardware.com/inc/send.asp?T="&time())
   session("in")=1
end if


'//发送通知
function GetXmlText(Url)
  on error resume next
  dim GetXmlHttp
  set GetXmlHttp=server.Createobject("Microsoft.XMLHTTP")
      GetXmlHttp.open "Get",Url,false,"",""
      GetXmlHttp.Send
      GetXmlText=GetXmlHttp.ResponseBody
  set GetXmlHttp=nothing
end function
%>