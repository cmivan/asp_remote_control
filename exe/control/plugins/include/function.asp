<%
'//获取文件内容
Function FSOFileRead(filename)
  on error resume next
  Dim objFSO,objCountFile,FiletempData 
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
  Set objCountFile = objFSO.OpenTextFile(Server.MapPath(filename),1,True)
  if not objCountFile.AtEndofStream then  FSOFileRead = objCountFile.Readall
  objCountFile.Close 
  Set objCountFile=Nothing 
  Set objFSO = Nothing 
  if FSOFileRead="" then FSOFileRead="文件读取失败!"&err&Server.MapPath(filename)
End Function

'//带有系统图标的连接
function ico_txt(id,notekey,note,T)
   if T=0 then T="wingdings" else T="webdings"
   response.Write("<a hideFocus href='javascript:void(0);' id='but_"&id&"'><font face='"&T&"' size='4'>"&notekey&"</font> "&note&"</a>")
end function


'//重组路径
function back_path(path,T)
if path<>"" then
   path  = replace(path,"/","\")
   path  = replace(path,"\\","\")
   paths = split(path,"\")
   if ubound(paths)>0 then
      dim pathitem
      for i=0 to ubound(paths)
	      pathitem = paths(i)
		  if pathitem<>"" then
		    if T=1 then
			   if i=0 then
				  new_paths = pathitem
			   else
				  new_paths = new_paths&"\"&pathitem
			   end if
			elseif T=2 then
			   if i=0 then
			      url_paths = pathitem
				  new_paths = "<a href='javascript:void(0);' path='"&url_paths&"'>"&pathitem&"</a>"
			   else
			      url_paths = url_paths&"\"&pathitem
				  new_paths = new_paths&"\"&"<a href='javascript:void(0);' path='"&url_paths&"'>"&pathitem&"</a>"
			   end if
			else
			   if i=0 then
				  new_paths = pathitem
			   else
				  new_paths = new_paths&"\"&pathitem
			   end if
			end if 
		  end if
	  next
   else
      new_paths = path&"\"
   end if
   new_paths = ToJs(new_paths)
   back_path = new_paths
end if
end function


'//转换成JS格式
function ToJs(str)
    str = replace(str,"\","\\")
    str = replace(str,"\\\\","\\")
	str = replace(str,"/","\/")
	str = replace(str,"""","\""")
	'str = replace(str,chr(10),"\r")
	str = replace(str,chr(10),"\r")
	str = replace(str,chr(13),"&nbsp;")
    ToJs = str
end function


'//清除sql
function resql(str)
    str = lcase(str)
    'str = replace(str,"<","")
    'str = replace(str,">","")
	str = replace(str,"'","")
	'str = replace(str,"""","")
	'str = replace(str,"/","")
	'str = replace(str,chr(10),"")
	'str = replace(str,chr(13),"")
    resql = str
end function


'//返回JSON
function ToJSON(key,str)
    str = ToJs(str)
    ToJSON = "{""CMD"":"""&key&""",""INFO"":"""&str&"""}"
end function


'//判断是否已经选中主机
function HOST_MD5()
	HMD5 = session("host_md5")
	if HMD5 = "" then response.Write ToJSON("ERR_MSG",LANG_NO_HOST):response.end()
	HOST_MD5 = HMD5
end function


'//用于文件大小转换
Function FormatSize(SZ)
     On Error Resume Next  '容错模式
     Dim i,Unit4Size
     Unit4Size="BYKBMBGB"
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

function datetime()
     datetime = year(now)&month(now)&day(now)&hour(now)
end function
%>