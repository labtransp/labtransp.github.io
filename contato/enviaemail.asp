<!-----Segunda parte do formulario, pegue um arquivo em branco e Copie os códigos abaixo
e salve com o nome "enviaemail.asp". Esse será o segundo arquivo.
Para o funcionamento correto do script, altere apenas os campos abaixo:
Na linha Mailer.RemoteHost = "smtp.seudominio.com.br" substitua a parte "seudominio.com.br" pelo nome correspondente ao de seu domínio.
Na linha Mailer.AddRecipient "Formulário" , "seudominio@seudominio.com.br", substitua pelo endereço de e-mail que receberá os dados do formulário.
Na linha Response.redirect "http://www.seudominio.com.br/", preencha com a URL que deve ser apresentada após o envio do formulário.
Exemplo: pegue um arquivo em branco, escreva um texto de agredecimento e salve com o nome obrigado.asp
faça o upload dos arquivos no mesmo diretório (dentro do www).
 
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
Response.redirect "http://transparencia.pub/contato/sucesso.html"
else
Response.Write mailer.response
end if
%>
