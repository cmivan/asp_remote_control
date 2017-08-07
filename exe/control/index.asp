<!--#include file="plugins/include/conn.asp"-->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=LANG_SYS_MAIN%></title>
<script src="plugins/style/jquery1.4.js" language="javascript"></script>
<script src="plugins/style/main.js?t=<%=datetime()%>" language="javascript"></script>
<script src="plugins/style/action.js?t=<%=datetime()%>" language="javascript"></script>
<script src="plugins/style/jquery.contextmenu.r2.js"></script>
<link href="plugins/style/css.css" rel="stylesheet" type="text/css" />

<link rel="stylesheet" href="plugins/xy_tips/style.css" type="text/css" media="all" />
<script src="plugins/xy_tips/jquery.XYTipsWindow.2.8.js?t=<%=datetime()%>" language="javascript"></script>
</head><body>
<!--down list-->
<!--#include file="plugins/include/mod_downlist.asp"-->
<div class="windows">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
<!--button--><tr><td class="windows_top">
<div class="top_nav"><!--#include file="plugins/include/mod_button.asp"--></div>
<div class="top_host_box"><%call ico_txt("host",":8",LANG_ONLINE_TIP&" ( <span>0</span> )",0)%></div>
<div class="top_host_info"></div><input id="on_macMD5" type="hidden"/></td></tr>
<!--stat bar-->
<tr><td class="windows_nav">
<table border="0" cellpadding="0" cellspacing="2">
<tr><td width="30" align="center">&nbsp;<font face='wingdings' size='2' class="folder_color">1</font>&nbsp;</td>
<td><div id="navpath_link"></div></td>
<td width="30"><img src="plugins/images/folder1.png" width="16" height="16"></td><td id="STAT_NOTE"></td></tr></table></td></tr>
<!--main body-->
<tr><td class="windows_body" valign="top">
<div id="window_box"><div id="welcome"><%=SYS_WELCOME%></div></div>
</td></tr></table></div><div></div>
<div id="welcomebox" style="display:none"><div id="welcome"><%=SYS_WELCOME%></div></div>
</body>
</html>
