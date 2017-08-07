<!--#include file="plugins/include/conn.asp"-->
<%

' 2011-8-1
' 提交命令操作并记录在数据库中
' By :  cmivan

 '<><><><><><><><><>
  dim macMD5,CMD,KEY1,KEY2
      macMD5 = session("host_md5")
  
 '<>部分操作需要选择主机后才可以执行
  if macMD5<>"" then
	 '//判断并记录提交的命令
	 set rs=conn.execute("select top 1 * from cmd where macMD5='"&macMD5&"' and running=true and ok=true and end=false order by updatetime desc")
		 if not rs.eof then
		    conn.execute("update cmd set running=true,ok=true,end=true where id="&rs("id"))
		   '<><><><><><><><><><><>
		    CMD  = rs("CMD")
			KEY1 = rs("KEY1")
			KEY2 = rs("KEY2")
			cmdBACK = rs("cmdBACK")
		   '<><><><><><><><><><><>
			'//筛选命令并执行
			select case CMD
				case "GET_DISK"
					 call GET_DISK()
				case "GET_RPATH"
					 call GET_RPATH(KEY1)
				case "GET_RFILE"
					 call GET_RFILE(KEY1,cmdBACK) 
				case "GET_SCREEN"
					 call GET_SCREEN(cmdBACK)
				case "SYS_LOGOFF"
					 call SYS_LOGOFF()
				case "SYS_REBOOT"
					 call SYS_REBOOT()
				case "SYS_SHUTDOWN"
					 call SYS_SHUTDOWN()
				case "TASK_LIST"
					 call TASK_LIST()
				case "TASK_KILL"
					 call TASK_KILL(KEY2,cmdBACK)
				case "OPEN_WEB"
					 call OPEN_WEB(KEY1)
				case "OPEN_MSG"
					 call OPEN_MSG(KEY1)
				case "RUN_CMD"
					 call RUN_CMD(KEY1,cmdBACK)
				case "TO_BLACK"
					 call TO_BLACK()
			end select
		 end if
	 set rs=nothing
  end if
 %>