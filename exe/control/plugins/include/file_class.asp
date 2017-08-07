<%
'#### 解析文件格式
function file_ext(filename)
    dim fileext,fileextArr
	fileext="txt"
    if filename<>"" then
	   fileextArr=split(filename,".")
	   if ubound(fileextArr)>=1 then
	      fileext=fileextArr(ubound(fileextArr))
	   end if
	end if
	file_ext=lcase(fileext)
end function
%>