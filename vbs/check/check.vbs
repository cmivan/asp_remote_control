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
'//////////    开始，定义内容、链接地址、执行内容判断    /////////////
'/////////////////////////////////////////////////////////////////////

Dim Sys_PageStr           '定义用于记录当前页面数据
Dim Sys_PageNewStr        '定义用于记录最新页面数据
Dim Sys_Times             '定义监测的间隔时间   '秒
Dim Sys_Time              '定义监测的的时间     '小时
Dim Sys_InTime            '定义登陆时间
Dim Sys_Num               '定义用于刷新的随机数
Dim Sys_Url               '定义需要检测的页面链接
Dim Sys_AddNum            '定义偏移值，用于判断在指定范围内 数据是否更新
'--------------------------------
Dim Sys_FileName          '定义VBS文件名
    Sys_FileName=Replace(Lcase(Wscript.ScriptName),".vbs","")
    
   'Sys_Time   = Inputbox("请设定监测时间!")
   'Sys_Time   = CheckSysTime(Sys_Time)
   
    Sys_Time   = CheckSysTime(8)                 '监测时间长度，监测的终结时间
    Sys_Times  = 5                               '5秒，指定程序监测时间间隔
	Sys_AddNum = 3                               '整数,用于判断偏于值
	Sys_Url    = "http://www.zc62.com//old/up/cmivan/vbs/online.asp"     '注意该地址的灵活使用
	Sys_TempUrl= "http://www.baidu.com"          '清除目标站数据

    Call CheckHttp()      '调用监测函数，运行
	
	
	
	
	
	
	
	
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
'///////////////////////// 设定 开机运行 /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Dim SysReg,FullFileName
Set SysReg       = Wscript.CreateObject("wscript.shell")
    FullFileName = WScript.ScriptFullName
    SysReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WebCheck[Cm.Ivan]",FullFileName
Set SysReg       = Nothing



' <--------------------|  函数定义部分  |-------------------->

'/////////////////////////////////////////////////////////////////////
'///////////////////////// 保存文件内容  /////////////////////////////
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
'///////////////////////// 读取文件内容  /////////////////////////////
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
'///////////////////////// 读取时间函数  /////////////////////////////
'/////////////////////////////////////////////////////////////////////
Function GetNow() 
         GetNow=year(Now)&"-"&month(Now)&"-"&day(Now)&"  |  "&hour(Now)&"."&minute(Now)&"."&second(Now)
End Function


'/////////////////////////////////////////////////////////////////////
'///////////////////////// 读取定页面源码  ///////////////////////////
'/////////////////////////////////////////////////////////////////////
Dim NetBreaks
    NetBreaks=0
Function GetHttpPage(url) 
     On error resume next
     Dim objXML
	 
     Set objXML=CreateObject("MSXML2.XMLHTTP")       '定义 
	  '-------------  | 页面监测辅助 |  --------------- 
		 objXML.open "GET",Sys_TempUrl,false         '打开 
         objXML.send()                               '发送
	 
	  '------------  | 掉线\超时,则跳过 Function |  --------------- 
      If (Err.Number<>0) then
	     ERR.clear
         Err.close
		 NetBreaks=NetBreaks+1
		  
	     Set objXML=Nothing
         Exit function 
		 else
		     NetBreaks=0
      End if

	  
	  If objXML.readystate=4 and Err.Number=0 Then   '判断解析,完成则客户端接受消息
         GetHttpPage=BytesToBstr(objXML.responseBody)'返回信息,同时用函数定义编码
      End If 

	 
	  '-------------  | 目标页面监测 |  --------------- 
         objXML.open "GET",url,false                 '打开 
         objXML.send()                               '发送 
      If objXML.readystate=4 and Err.Number=0 Then   '判断解析,完成则客户端接受消息
         GetHttpPage=BytesToBstr(objXML.responseBody)'返回信息,同时用函数定义编码
      Else
	     Set objXML=Nothing
         Exit function 
      End If 
     Set objXML=Nothing

End Function 

'/////////////////////////////////////////////////////////////////////
'/////////////////// 将字符转为指定编码，防止乱码 ////////////////////
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
        'ObjStream.CharSet = "UTF-8"     '防止乱码
         ObjStream.CharSet = "gb2312"    '防止乱码
         BytesToBstr = ObjStream.ReadText  
         ObjStream.Close 
     Set ObjStream = Nothing 
End Function 

'/////////////////////////////////////////////////////////////////////
'/////////////////////////   清除Html格式  ///////////////////////////
'/////////////////////////////////////////////////////////////////////
Function NoHtml(str)  
    Set re=new RegExp  
        re.IgnOreCase =true  
        re.Global=True  
        re.Pattern="(\<.[^\<]*\>)"  
        str=re.replace(str,"*")  
        re.Pattern="(\<\/[^\<]*\>)"  
        str=re.replace(str,"*") 
        '清除空格\换行符
		str=replace(str," ","")
		str=replace(str,"　","")
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
'/////////////////////////   时间监测函数  ///////////////////////////
'/////////////////////////////////////////////////////////////////////
Function CheckHttp()
   If hour(Now)-1<Sys_Time Then
	  Call CheckPage()                '页面监测函数
	  Wscript.Sleep 1000*Sys_Times    '休眠时间后，继续
	  Call CheckHttp()                '时间监测程序
	Else
	  Exit Function
   End If
End Function

'/////////////////////////////////////////////////////////////////////
'////////////////////////   返回监测终结时间   ///////////////////////
'/////////////////////////////////////////////////////////////////////
Function CheckSysTime(STime)
     Dim Sys_InTime
	'判断所输入的内容是否符合,不符合则自动记录为默认值
	 If STime=Null Or STime="" Then STime=0
     If Isnumeric(STime) And STime>0 Then
     If STime<=0 Or STime>=24 Then STime = 5 
        Else
        STime = 5
        Msgbox "设定不符合,程序自动设置参数为"&STime&"."
     End If 
    '记录运行时间，并计算结束时间
	 Sys_InTime   = hour(Now)          '记录当前时间  '小时
     STime        = STime+Sys_InTime   '在当前时间上累加
	 If STime>23 Then STime=23
	 CheckSysTime=STime                '返回时间
End Function

'/////////////////////////////////////////////////////////////////////
'//////////////////////////   页面内容监测   /////////////////////////
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
          Sys_PageStr=SysReadFile(Sys_FileName&".ini")       '初始数据，读取上次记录内容
          
          If Sys_PageStr="" Then
             Sys_PageStr=Sys_PageNewStr 
	  	     Msgbox GetNow()&chr(13)&chr(13)& "Record New Data!"
	  	     Call SysSaveFile(Sys_PageStr,Sys_FileName&".ini")  '保存新数据
			 Else
			 Msgbox GetNow()&chr(13)&chr(13)& "Record The Previous Data!"
	  	  End If
	  	  
	   Else
	   
	   	 '判断新数据与当前是否有差异
	   	 If Sys_PageStr<>Sys_PageNewStr Then
			
			'--------------对比出更新的地方-------------------
			Dim NewPage1_Arr,NewPage2_Arr
			Dim NewPage1_Num,NewPage2_Num
			Dim SysNun,LoopNum           '用于FOr循环对比
			Dim SysInfo,SysInfo1,SysInfo2                  '用于返回提示信息
			 
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
				    If NewPage1_Arr(SysNun)<>NewPage2_Arr(SysNun) Then  '当前数组进行比较
					   
					   '判断偏移值是否可以直接使用
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
						   If NewPage2_Arr(SysNun)=NewPage1_Arr(SysNun+I) Then   '成立则为非新数据
						      BackCheck=false
						   End If
					   next
					   
					   	   '判断符合,则记录数据
						   If BackCheck=true Then
						   	  SysInfo1=SysInfo1&NewPage1_Arr(SysNun)&chr(13)
                              SysInfo2=SysInfo2&NewPage2_Arr(SysNun)&chr(13)
						   End If
						   
					End If
				next

				 '信息提示 ////////////////////////////////////////////////////////////
				    
					SysInfo= ""
					SysInfo= "|  数据更新 ["&Now()&"]  |"
					SysInfo= SysInfo&chr(13)
					SysInfo= SysInfo&chr(13)&"----------------|  原数据  |-------------------"
                    SysInfo= SysInfo&chr(13)&SysInfo1&chr(13)
			        SysInfo= SysInfo&chr(13)&"----------------|  已更新  |-------------------"
					SysInfo= SysInfo&chr(13)&SysInfo2
					
					Sys_PageStr=""

			        Sys_PageStr=Sys_PageNewStr                         '记录新数据
			        Call SysSaveFile(Sys_PageStr,Sys_FileName&".ini")  '保存新数据
       
                    '判断提示信息为新信息 则提示
					If SysInfo<>"" Then Msgbox SysInfo
					
				 '//////////////////////////////////////////////////////////////////////
		 End If
	   End If
	End If
End Function