<!--#include file="plugins/include/conn.asp"-->
<%

' 2011-8-1
' 获取主机信息
' By :  cmivan

 '<><><><><><><><><>
  CMD  = ucase(request("CMD"))
  KEY1 = request("KEY1")
  KEY2 = request("KEY2")

  if KEY1="" then KEY1="0"
  if KEY2="" then KEY2="0"
  
  macMD5 = session("host_md5")
 '<><>本次命令的md5值
  cmdMD5 = md5(CMD&KEY1&KEY2)

 '<>部分操作需要选择主机后才可以执行
  if (CMD<>"HOST_LIST" and CMD<>"HOST_INFO") and macMD5="" then
     response.Write("{""CMD"":""ERR_MSG"",""INFO"":"""&LANG_ONLINE_OFF&"""}"):response.End()
  elseif CMD<>"HOST_LIST" and CMD<>"HOST_INFO" then
	 '//判断并记录提交的命令
	 set rs=conn.execute("select top 1 * from cmd where macMD5='"&macMD5&"' and cmdMD5='"&cmdMD5&"' order by id desc")
		 if not rs.eof then
			conn.execute("update cmd set CMD='"&CMD&"',KEY1='"&KEY1&"',KEY2='"&KEY2&"',running=false,ok=false,end=false,times=times+1,updatetime='"&now()&"' where id="&rs("id"))
		 else
			conn.execute("insert into cmd(macMD5,cmdMD5,CMD,KEY1,KEY2) values('"&macMD5&"','"&cmdMD5&"','"&CMD&"','"&KEY1&"','"&KEY2&"')")
		 end if
	 set rs=nothing
  end if

  '//筛选命令并执行
  select case CMD
      case "HOST_LIST" '获取主机列表
	       call HOST_LIST()
      case "HOST_INFO" '获取主机信息
	       call HOST_INFO(KEY1) 
  end select
  
%>