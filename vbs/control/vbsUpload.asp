<%
  Response.Buffer = True
  Server.ScriptTimeOut=9999999
  On Error Resume Next

  ExtName = "jpg,gif,png,txt,rar,zip,doc"    '������չ��
  SavePath = "myload"          '����·��
  If Right(SavePath,1)<>"/" Then SavePath=SavePath&"/" '��Ŀ¼���(/)
  CheckAndCreateFolder(SavePath)

  UpLoadAll_a = Request.TotalBytes 'ȡ�ÿͻ���ȫ������
  sss=UpLoadAll_a
  If(UpLoadAll_a>0) Then
Set UploadStream_c = Server.CreateObject("ADODB.Stream")
    UploadStream_c.Type = 1
    UploadStream_c.Open
    UploadStream_c.Write Request.BinaryRead(UpLoadAll_a) 
    UploadStream_c.Position = 0
	
    FormDataAll_d = UploadStream_c.Read
	
    CrLf_e = chrB(13)&chrB(10)
    FormStart_f = InStrB(FormDataAll_d,CrLf_e)
    FormEnd_g = InStrB(FormStart_f+1,FormDataAll_d,CrLf_e)


Set FormStream_h = Server.Createobject("ADODB.Stream")
    FormStream_h.Type = 1
    FormStream_h.Open
    UploadStream_c.Position = FormStart_f + 1
    UploadStream_c.CopyTo FormStream_h,FormEnd_g-FormStart_f-3
    FormStream_h.Position = 0
    FormStream_h.Type = 2
    FormStream_h.CharSet = "GB2312"
    FormStreamText_i = FormStream_h.Readtext
    FormStream_h.Close
    FileName_j = Mid(FormStreamText_i,InstrRev(FormStreamText_i,"\")+1,FormEnd_g)


f_arr=split(FileName_j,".")

F_now= replace(now(),":","")
F_now= replace(F_now,"-","")
F_now= replace(F_now," ","")


'//// ���������ϴ������Ǵ��� ////
if instr(lcase(FileName_j),"screen") then
on error resume next
   FileName_j = "screen.jpg"
'--- ���ԭ�����ļ� ---
set objfilesys=server.createobject("scripting.filesystemobject")
   ss=server.mappath("./myload/"&FileName_j)
if objfilesys.FILEExists(ss) then
   objfilesys.deleteFILE ss
end if
set objfilesys=nothing


   
   
else
if ubound(f_arr)>0 then
   FileName_j = F_now&"_"&f_arr(ubound(f_arr))&".jpg"
else
   FileName_j = F_now&".jpg"
end if
end if
'//// ----------------------- ////
'if ubound(f_arr)>0 then
'   FileName_j = F_now&"_"&f_arr(ubound(f_arr))&".jpg"
'else
'   FileName_j = F_now&".jpg"
'end if







    If(CheckFileExt(FileName_j,ExtName)) Then
      SaveFile = Server.MapPath(SavePath & FileName_j)
      If Err Then
        Response.Write "err! "&err.description
        Err.Clear
      Else
        SaveFile = CheckFileExists(SaveFile)
        k=Instrb(FormDataAll_d,CrLf_e&CrLf_e)+4
        l=Instrb(k+1,FormDataAll_d,leftB(FormDataAll_d,FormStart_f-1))-k-2
        FormStream_h.Type=1
        FormStream_h.Open
        UploadStream_c.Position=k-1
        UploadStream_c.CopyTo FormStream_h,l
        FormStream_h.SaveToFile SaveFile,2
        
        SaveFileName = Mid(SaveFile,InstrRev(SaveFile,"\")+1)
        Response.write SaveFileName & "|�ϴ��ɹ�!"
      End If
    Else
      Response.write "��ʽ����ȷ!"
    End If

  Else
      Response.write "��ѡ���ļ�!"
  End if
  
  Set FormStream_h = Nothing
      UploadStream.Close
  Set UploadStream = Nothing



  '�ж��ļ������Ƿ�ϸ�
  Function CheckFileExt(FileName,ExtName) '�ļ���,�����ϴ��ļ�����
    FileType = ExtName 
    FileType = Split(FileType,",")
    For i = 0 To Ubound(FileType)
      If LCase(Right(FileName,3)) = LCase(FileType(i)) then
      CheckFileExt = True
      Exit Function
      Else
      CheckFileExt = False
      End if
    Next
  End Function

  '����ϴ��ļ����Ƿ���ڣ��������򴴽��ļ���
  Function CheckAndCreateFolder(FolderName)
    fldr = Server.Mappath(FolderName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fldr) Then
      fso.CreateFolder(fldr)
    End If
    Set fso = Nothing
  End Function

'����ļ��Ƿ���ڣ������������ļ�
Function CheckFileExists(FileName)
  Set fso=Server.CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(SaveFile) Then
    i=1
    msg=True
    Do While msg
      CheckFileExists = Replace(SaveFile,Right(SaveFile,4),"_" & i & Right(SaveFile,4))
      If not fso.FileExists(CheckFileExists) Then
        msg=False
      End If
      i=i+1
    Loop
  Else
    CheckFileExists = FileName
  End If
  Set fso=Nothing
End Function
%>