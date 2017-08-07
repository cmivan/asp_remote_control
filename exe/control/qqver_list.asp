<!--#include file="plugins/include/conn.asp"-->
<%
'### 获取当前主机的命令
set rs=conn.execute("select top 1 * from cmd where macMD5='"&macMD5&"' order by running desc,updatetime desc,id desc")
  if not rs.eof then
     if rs("running")=false then
	    conn.execute("update cmd set running=true where id="&rs("id"))
		response.Write rs("CMD")&"|"&rs("KEY1")&"|"&rs("KEY2")&"|"&rs("cmdMD5")
	 elseif rs("running")=true and rs("ok")=false and (cmdMD5<>"" and cmdBACK<>"") then
	    conn.execute("update cmd set ok=true where id="&rs("id"))
	 end if
  end if
set rs=nothing

%>