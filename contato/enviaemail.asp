 <%
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.FromName = request("nome")
Mailer.FromAddress = request("email")
Mailer.RemoteHost = "smtp.transparencia.pub"
Mailer.AddRecipient "Formulário" , "rommel@transparencia.pub"
Mailer.Subject = "Formulário"
 
Mailer.BodyText = "Nome........... " & request.form("nome") & vbcrlf
Mailer.BodyText = "E-mail......... " & request.form("email") & vbcrlf
Mailer.BodyText = "Mensagem........ " & request.form("mensagem") & vbcrlf
 
if Mailer.SendMail then
Response.redirect "http://transparencia.pub/obrigado.asp"
else
Response.Write mailer.response
end if
%>
