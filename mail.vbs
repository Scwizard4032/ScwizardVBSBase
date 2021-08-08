NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
set Email = CreateObject("CDO.Message")
Email.From = "scwizard@scw-mail.cn" 'send_host
Email.To = "2994843619@qq.com" 'receiver
Email.Subject = "vbs脚本发邮件" 'title
x="E:\vbsbase\mail.txt" '内容
y="E:\vbsbase\mail.txt" '附件
Set fso=CreateObject("Scripting.FileSystemObject")
Set myfile=fso.OpenTextFile(x,1,Ture)
c=myfile.readall
myfile.Close
Email.Textbody = c
Email.AddAttachment y
with Email.Configuration.Fields
.Item(NameSpace&"sendusing") = 2
.Item(NameSpace&"smtpserver") = "test.scw-mail.cn" 'hosts
.Item(NameSpace&"smtpserverport") = 25
.Item(NameSpace&"smtpauthenticate") = 1
.Item(NameSpace&"sendusername") = "Scwizard" 'sender
.Item(NameSpace&"sendpassword") = "wei66179" 'pwd
.Update
end with
dim sendmail
sendmail = Email.Send
msgbox send
Set Email=Nothing
