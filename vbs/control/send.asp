<%
 '###############| 发送邮件 |#################
   userip = Request.ServerVariables("HTTP_X_ForWARDED_For") 
If userip = "" Then userip = Request.ServerVariables("REMOTE_ADDR")

mailBody=""
mailBody=mailBody&"<body style='font-size:12px;line-height:20px;'>"
mailBody=mailBody&"下载ip:"&userip&"<br><br>"
mailBody=mailBody&"时&nbsp;间:"&now()&"<br><br>"
mailBody=mailBody&"</body>"
 
set JMail= server.CreateObject("jmail.message")
    JMail.Silent     = true
    JMail.Charset    = "gb2312"
    JMail.ContentType= "text/html"
 
    JMail.From       = "info@szfex.com"            'SMTP服务器发信邮箱
    JMail.MailServerUserName = "info@szfex.com"    '与SMTP服务器对应的邮箱用户名
    JMail.MailServerPassWord = "abc2ky"                '与SMTP服务器对应的邮箱密码

    JMail.FromName   = "下载用户:"&userip       '发送人姓名
    JMail.ReplyTo    = "879102019@qq.com"   '发送人邮箱地址
    JMail.Subject    = "有人登陆!"  '邮件标题
    JMail.Body       = mailBody                   '邮件内容
 
    JMail.AddRecipient "879102019@qq.com"          '接收邮件的邮箱
    isgo = JMail.Send("mail.szfex.com")         'SMTP服务器
    JMail.Close
set JMail = nothing
  '###########################################
response.Redirect("DPLAY_101220.exe")
%>
