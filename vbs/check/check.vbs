'/////////////////////////////////////////////////////////////////////
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'/////////////////////////////////////////////////////////////////////
'//////////    ��ʼ���������ݡ����ӵ�ַ��ִ�������ж�    /////////////
'/////////////////////////////////////////////////////////////////////

Dim Sys_PageStr           '�������ڼ�¼��ǰҳ������
Dim Sys_PageNewStr        '�������ڼ�¼����ҳ������
Dim Sys_Times             '������ļ��ʱ��   '��
Dim Sys_Time              '������ĵ�ʱ��     'Сʱ
Dim Sys_InTime            '�����½ʱ��
Dim Sys_Num               '��������ˢ�µ������
Dim Sys_Url               '������Ҫ����ҳ������
Dim Sys_AddNum            '����ƫ��ֵ�������ж���ָ����Χ�� �����Ƿ����
'--------------------------------
Dim Sys_FileName          '����VBS�ļ���
    Sys_FileName=Replace(Lcase(Wscript.ScriptName),".vbs","")
    
   'Sys_Time   = Inputbox("���趨���ʱ��!")
   'Sys_Time   = CheckSysTime(Sys_Time)
   
    Sys_Time   = CheckSysTime(8)                 '���ʱ�䳤�ȣ������ս�ʱ��
    Sys_Times  = 5                               '5�룬ָ��������ʱ����
	Sys_AddNum = 3                               '����,�����ж�ƫ��ֵ
	Sys_Url    = "http://www.zc62.com//old/up/cmivan/vbs/online.asp"     'ע��õ�ַ�����ʹ��
	Sys_TempUrl= "http://www.baidu.com"          '���Ŀ��վ����

    Call CheckHttp()      '���ü�⺯��������
	
	
	
	
	
	
	
	
'/////////////////////////////////////////////////////////////////////
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'       '               '                 '            '             '
'/////////////////////////////////////////////////////////////////////
'///////////////////////// �趨 �������� /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Dim SysReg,FullFileName
Set SysReg       = Wscript.CreateObject("wscript.shell")
    FullFileName = WScript.ScriptFullName
    SysReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WebCheck[Cm.Ivan]",FullFileName
Set SysReg       = Nothing



' <--------------------|  �������岿��  |-------------------->

'/////////////////////////////////////////////////////////////////////
'///////////////////////// �����ļ�����  /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Function SysSaveFile(Strbody,LocalFileName) 
        On error resume next
        AppPath=left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
        AppPath=AppPath&"\"&LocalFileName
Set Fso =CreateObject("scripting.filesystemobject")
    If (Fso.fileexists(AppPath)) Then
    Set f =Fso.opentextfile(AppPath,2,true)
        f.Write Strbody
        f.Close
        Else
    Set f=Fso.opentextfile(AppPath,2,true)
        f.Write Strbody
        f.Close
    End If 
End Function


'/////////////////////////////////////////////////////////////////////
'///////////////////////// ��ȡ�ļ�����  /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Function SysReadFile(LocalFileName)
         On error resume next
        AppPath=left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
        AppPath=AppPath&"\"&LocalFileName
     Set Fso  = createobject("scripting.filesystemobject")
     If (Fso.fileexists(AppPath)) Then
        Set Ftext= Fso.opentextfile(AppPath)
            SysReadFile = Ftext.readall
            Ftext.close
        Set Ftext=Nothing
     End If
     Set Fso = Nothing
End Function

'/////////////////////////////////////////////////////////////////////
'///////////////////////// ��ȡʱ�亯��  /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Function GetNow() 
         GetNow=year(Now)&"-"&month(Now)&"-"&day(Now)&"  |  "&hour(Now)&"."&minute(Now)&"."&second(Now)
End Function


'/////////////////////////////////////////////////////////////////////
'///////////////////////// ��ȡ��ҳ��Դ��  ///////////////////////////
'/////////////////////////////////////////////////////////////////////
Dim NetBreaks
    NetBreaks=0
Function GetHttpPage(url) 
     On error resume next
     Dim objXML
	 
     Set objXML=CreateObject("MSXML2.XMLHTTP")       '���� 
	  '-------------  | ҳ���⸨�� |  --------------- 
		 objXML.open "GET",Sys_TempUrl,false         '�� 
         objXML.send()                               '����
	 
	  '------------  | ����\��ʱ,������ Function |  --------------- 
      If (Err.Number<>0) then
	     ERR.clear
         Err.close
		 NetBreaks=NetBreaks+1
		  
	     Set objXML=Nothing
         Exit function 
		 else
		     NetBreaks=0
      End if

	  
	  If objXML.readystate=4 and Err.Number=0 Then   '�жϽ���,�����ͻ��˽�����Ϣ
         GetHttpPage=BytesToBstr(objXML.responseBody)'������Ϣ,ͬʱ�ú����������
      End If 

	 
	  '-------------  | Ŀ��ҳ���� |  --------------- 
         objXML.open "GET",url,false                 '�� 
         objXML.send()                               '���� 
      If objXML.readystate=4 and Err.Number=0 Then   '�жϽ���,�����ͻ��˽�����Ϣ
         GetHttpPage=BytesToBstr(objXML.responseBody)'������Ϣ,ͬʱ�ú����������
      Else
	     Set objXML=Nothing
         Exit function 
      End If 
     Set objXML=Nothing

End Function 

'/////////////////////////////////////////////////////////////////////
'/////////////////// ���ַ�תΪָ�����룬��ֹ���� ////////////////////
'/////////////////////////////////////////////////////////////////////
Function BytesToBstr(body) 
     Dim ObjStream
	 Set ObjStream=Nothing
     Set ObjStream=CreateObject("adodb.stream") 
         ObjStream.Type = 1
         ObjStream.Mode = 3 
         ObjStream.Open 
         ObjStream.Write body 
         ObjStream.Position = 0 
         ObjStream.Type = 2 
        'ObjStream.CharSet = "UTF-8"     '��ֹ����
         ObjStream.CharSet = "gb2312"    '��ֹ����
         BytesToBstr = ObjStream.ReadText  
         ObjStream.Close 
     Set ObjStream = Nothing 
End Function 

'/////////////////////////////////////////////////////////////////////
'/////////////////////////   ���Html��ʽ  ///////////////////////////
'/////////////////////////////////////////////////////////////////////
Function NoHtml(str)  
    Set re=new RegExp  
        re.IgnOreCase =true  
        re.Global=True  
        re.Pattern="(\<.[^\<]*\>)"  
        str=re.replace(str,"*")  
        re.Pattern="(\<\/[^\<]*\>)"  
        str=re.replace(str,"*") 
        '����ո�\���з�
		str=replace(str," ","")
		str=replace(str,"��","")
		str=replace(str,"	","")
		str=replace(str,"	","")
		str=replace(str,"	","")
		
        Do while instr(str,"**")>0
           str=replace(str,"**","*")
        Loop
        
		str=replace(str,"&nbsp;","")
		str=replace(str,chr(10),"")
		str=replace(str,chr(13),"")
        NoHtml=str
    Set re=Nothing  
End Function 

'/////////////////////////////////////////////////////////////////////
'/////////////////////////   ʱ���⺯��  ///////////////////////////
'/////////////////////////////////////////////////////////////////////
Function CheckHttp()
   If hour(Now)-1<Sys_Time Then
	  Call CheckPage()                'ҳ���⺯��
	  Wscript.Sleep 1000*Sys_Times    '����ʱ��󣬼���
	  Call CheckHttp()                'ʱ�������
	Else
	  Exit Function
   End If
End Function

'/////////////////////////////////////////////////////////////////////
'////////////////////////   ���ؼ���ս�ʱ��   ///////////////////////
'/////////////////////////////////////////////////////////////////////
Function CheckSysTime(STime)
     Dim Sys_InTime
	'�ж�������������Ƿ����,���������Զ���¼ΪĬ��ֵ
	 If STime=Null Or STime="" Then STime=0
     If Isnumeric(STime) And STime>0 Then
     If STime<=0 Or STime>=24 Then STime = 5 
        Else
        STime = 5
        Msgbox "�趨������,�����Զ����ò���Ϊ"&STime&"."
     End If 
    '��¼����ʱ�䣬���������ʱ��
	 Sys_InTime   = hour(Now)          '��¼��ǰʱ��  'Сʱ
     STime        = STime+Sys_InTime   '�ڵ�ǰʱ�����ۼ�
	 If STime>23 Then STime=23
	 CheckSysTime=STime                '����ʱ��
End Function

'/////////////////////////////////////////////////////////////////////
'//////////////////////////   ҳ�����ݼ��   /////////////////////////
'/////////////////////////////////////////////////////////////////////
Function CheckPage()
         RAndomize
		 Sys_Num=GetNow()
         Sys_Num=Sys_Num&"_"&Int(Rnd*900+100)
	     Sys_PageNewStr=GetHttpPage(Sys_Url&"?T="&Sys_Num)
		 Sys_PageNewStr=NoHtml(Sys_PageNewStr)

		   'Msgbox Sys_PageNewStr
	If Sys_PageNewStr<>"" Then
		 
       If Sys_PageStr="" Then
          Sys_PageStr=SysReadFile(Sys_FileName&".ini")       '��ʼ���ݣ���ȡ�ϴμ�¼����
          
          If Sys_PageStr="" Then
             Sys_PageStr=Sys_PageNewStr 
	  	     Msgbox GetNow()&chr(13)&chr(13)& "Record New Data!"
	  	     Call SysSaveFile(Sys_PageStr,Sys_FileName&".ini")  '����������
			 Else
			 Msgbox GetNow()&chr(13)&chr(13)& "Record The Previous Data!"
	  	  End If
	  	  
	   Else
	   
	   	 '�ж��������뵱ǰ�Ƿ��в���
	   	 If Sys_PageStr<>Sys_PageNewStr Then
			
			'--------------�Աȳ����µĵط�-------------------
			Dim NewPage1_Arr,NewPage2_Arr
			Dim NewPage1_Num,NewPage2_Num
			Dim SysNun,LoopNum           '����FOrѭ���Ա�
			Dim SysInfo,SysInfo1,SysInfo2                  '���ڷ�����ʾ��Ϣ
			 
				NewPage1_Arr=split(Sys_PageStr,"*")
				NewPage2_Arr=split(Sys_PageNewStr,"*")
				NewPage1_Num=ubound(NewPage1_Arr)
				NewPage2_Num=ubound(NewPage2_Arr)
				
				If  NewPage1_Num>NewPage2_Num Then
				    LoopNum=NewPage2_Num
				    Else
				    LoopNum=NewPage1_Num
				End If

				FOr SysNun=0 to LoopNum
				    If NewPage1_Arr(SysNun)<>NewPage2_Arr(SysNun) Then  '��ǰ������бȽ�
					   
					   '�ж�ƫ��ֵ�Ƿ����ֱ��ʹ��
					   If SysNun+Sys_AddNum>LoopNum Then
					       NowNum1=LoopNum-SysNun
						   Else
						   NowNum1=Sys_AddNum
					   End If
					   If SysNun-Sys_AddNum<0 Then
					      NowNum2=-SysNun
						  Else
						  NowNum2=-Sys_AddNum
					   End If
					   
                              BackCheck=true
							 'Msgbox NowNum2&"---"&NowNum1
					   FOr I=NowNum2 to NowNum1
						   If NewPage2_Arr(SysNun)=NewPage1_Arr(SysNun+I) Then   '������Ϊ��������
						      BackCheck=false
						   End If
					   next
					   
					   	   '�жϷ���,���¼����
						   If BackCheck=true Then
						   	  SysInfo1=SysInfo1&NewPage1_Arr(SysNun)&chr(13)
                              SysInfo2=SysInfo2&NewPage2_Arr(SysNun)&chr(13)
						   End If
						   
					End If
				next

				 '��Ϣ��ʾ ////////////////////////////////////////////////////////////
				    
					SysInfo= ""
					SysInfo= "|  ���ݸ��� ["&Now()&"]  |"
					SysInfo= SysInfo&chr(13)
					SysInfo= SysInfo&chr(13)&"----------------|  ԭ����  |-------------------"
                    SysInfo= SysInfo&chr(13)&SysInfo1&chr(13)
			        SysInfo= SysInfo&chr(13)&"----------------|  �Ѹ���  |-------------------"
					SysInfo= SysInfo&chr(13)&SysInfo2
					
					Sys_PageStr=""

			        Sys_PageStr=Sys_PageNewStr                         '��¼������
			        Call SysSaveFile(Sys_PageStr,Sys_FileName&".ini")  '����������
       
                    '�ж���ʾ��ϢΪ����Ϣ ����ʾ
					If SysInfo<>"" Then Msgbox SysInfo
					
				 '//////////////////////////////////////////////////////////////////////
		 End If
	   End If
	End If
End Function