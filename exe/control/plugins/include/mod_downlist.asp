<!--HOST LIST-->
<div id="list_host" class="downbox"><div style="text-align:center; padding:5px;"><img src="plugins/images/loading.gif" /></div></div>

<!--BOOT-->
<div id="list_boot" class="downbox">
<table width="100%" border="0" cellpadding="0" cellspacing="9"><tr><td class="reboot_but_item">
<%call ico_txt("SYS_LOGOFF","y",LANG_BUT_LOGOFF,1)%>
<%call ico_txt("SYS_REBOOT","i",LANG_BUT_REBOOT,1)%>
<%call ico_txt("SYS_SHUTDOWN","x",LANG_BUT_SHUTDOWN,1)%>
</td></tr><tr><td colspan="2" class="down_note"><%=LANG_BOOT_TIP%></td></tr></table>
</div>

<!--CMD-->
<div id="list_cmd" class="downbox">
<table border="0" cellpadding="0" cellspacing="10">
<tr><td><div id='cmd_box' contentEditable='true'>net user</div></td>
<td><a id='cmd_but' href="javascript:void(0);"><%=LANG_BUT_CMD%>(R)...</a></td>
</tr><tr><td colspan="3" class="down_note"><%=LANG_CMD_TIP%></td></tr></table>
</div>

<!--other-->
<div id="list_other" class="downbox">
<table border="0" cellpadding="0" cellspacing="10"><tr><td>
<label><input name="other_buts" type="radio" checked="checked" /><%=LANG_BUT_OPEN_WEB%></label>&nbsp;&nbsp;
<label><input name="other_buts" type="radio" /><%=LANG_BUT_OPEN_MSG%></label>&nbsp;&nbsp;
<span style="color:#000;">|</span>&nbsp;<%call ico_txt("black","n",LANG_BUT_TO_BLACK,0)%>
</td></tr><tr><td>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="other_tab_item">
<tr><td><div id='web_box' contentEditable='true'><%=LANG_BUT_OPEN_WEB_TIP%></div></td>
<td><a id='web_but' href="javascript:void(0);"><%=LANG_BUT_OPEN_WEB_INPUT%></a></td></tr></table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="other_tab_item" style="display:none">
<tr><td><div id='msg_box' contentEditable='true'><%=LANG_BUT_OPEN_MSG_TIP%></div></td>
<td><a id='msg_but' href="javascript:void(0);"><%=LANG_BUT_OPEN_MSG_INPUT%></a></td></tr></table>
</td></tr><tr><td colspan="2" class="down_note"><%=LANG_BUT_OPEN_TIP%></td></tr></table>
</div>



