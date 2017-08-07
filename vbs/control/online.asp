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

'---------------  主机  ---------------
set rs=server.createobject("adodb.recordset")
    rs_sql="select * from host order by ip desc"
    rs.open rs_sql,conn,1,1
   do while not rs.eof
if datediff("s",rs("update_time"),now())<6 then
%>
<p>
HOST:<%=rs("host")%>
IP:<%=rs("ip")%>
MAC:<%=rs("mac")%>
</p>
<%
end if


	rs.movenext
	loop
	rs.close
set rs=nothing
%>