$(function(){
	//显示主机列表
	Bind_Meun('host');
	//其他按钮
	Bind_Meun('other');
	//重启关机
	Bind_Meun('boot');
	//CMD
	Bind_Meun('cmd');

	//鼠标移过文件目录
	$('#window_files a').live('mouseover',function(){$(this).parent().parent().find('li').attr('class','');$(this).parent().attr('class','on');});
	//其他操作切换
	$('#list_other label').click(function(){
		var index = $(this).index();
		$('#list_other .other_tab_item').css({"display":"none"});
		$('#list_other .other_tab_item').eq(index).css({"display":"block"});
		});
	//切换主机按钮
	$('#list_host .item a').live('click',function(){
		$("#window_box").html($("#welcomebox").html());
		$(".top_host_info").html('<img src="plugins/images/loading_small.gif" />');
		$("#window_box").fadeOut(0).fadeIn(400);
	    //获取当前主机信息
		var HOST_ID = $(this).parent().attr("id");
	    SEND_DATA('HOST_INFO',HOST_ID,'');
		});
	//切换导航栏
	 $("#window_files a").live('dblclick',function(){
		 var CMD  = $(this).attr("CMD");
		 var KEY1 = $(this).attr("KEY1");
		 var KEY2 = $(this).attr("KEY2");
		 GET_DATA(CMD,KEY1,KEY2,'正在执行操作...',1);
		 });
	 //获取文件夹数据
//	 $("#window_files a.folder").live('dblclick',function(){
//		 var CMD  = $(this).attr("CMD");
//		 var KEY1 = $(this).attr("KEY1");
//		 var KEY2 = $(this).attr("KEY2");
//		 GET_DATA(CMD,KEY1,KEY2,'正打开目录...',1);
//		 });
	 //标签都会绑定此右键菜单
	 $('#window_files .item').contextMenu('FilesMenu',{
		 bindings:{
		  'open': function(t) {alert('打开文件夹 '+t.id+'\nAction was Open');},
		  'email': function(t) {alert('Trigger was '+t.id+'\nAction was Email');},
		  'save': function(t) {alert('保存文件 '+t.id+'\nAction was Save');},
		  'delete': function(t) {alert('删除文件 '+t.id+'\nAction was Delete');}
		  }
	 });
	 
			
	//检测在线主机
	setInterval("HOST_TIMER()",500);
	setInterval("DATA_TIMER()",1000);	

});


//定时检测主机
function HOST_TIMER(){
	//获取主机列表
	SEND_DATA('HOST_LIST','','');
	//获取当前主机信息
	SEND_DATA('HOST_INFO','X','');
	}

//定时检测返回的数据
function DATA_TIMER(){ CMD_DATA(); }

//绑定下拉菜单
function Bind_Meun(items){
	$("#but_"+items).mouseover(function(e){$(this).attr("class","on");$("#list_"+items).css({"display":"block"});});
	$("#but_"+items).mouseout(function(){$(this).attr("class","");$("#list_"+items).css({"display":"none"});});
	$("#list_"+items).mouseover(function(){$("#but_"+items).attr("class","on");$(this).css({"display":"block"});});
	$("#list_"+items).mouseout(function(){$("#but_"+items).attr("class","");$(this).css({"display":"none"});});
	}