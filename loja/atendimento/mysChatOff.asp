<%
'Aquivo para funcionar com varios componentes de e-mail dubairro@yahoo.com.br
%>
<!--#include file="db.asp"-->
<!--#include file="mysConfiguracoes.asp"-->
<%
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.addHeader "pragma","no-cache"
	Response.addHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	Session.LCID = 1046
	
strStatus = OK(Request.QueryString("status"))
	
If strStatus = "TimeOut" AND Session("dubConversaID") <> "" Then
	Call AbreDB
	strSQL = "UPDATE conversas SET status = False WHERE id = " & Session("dubConversaID")
	Conexao.Execute(strSQL)
	Call FechaDB
End If
%>
<%
'Email para onde serão enviados os dados do formulário
email_form = "web@studio7seven.com"

'Título do email
assunto_form = "Formulario Atendimento OnLine"

'Servidor de email
servidor = "mail.studio7seven.com" 'Configurar aqui seu Servidor de e-mail

'Componente para envio do email: CDONTS | AspMail | AspEmail | AspQmail
componente = "AspEmail"

'Cabeçalho do texto do email recebido
cabecalho_email = "Atendimento OnLine Formulário"

'*******************************************************************
%>

<html>
<head>
<title><%=strConfigNome%></title>
<style>
<!--
.texto_pagina { font-family: arial; font-size: 11px; color: #000000 }
.tabela_formulario
{
width: 200;
background-color: white;
}

.titulo_campos { font-family: arial; font-size: 11px; color: #000000; background-color: #FFFFFF }
.campos_formulario { font-family: Tahoma, Verdana, Arial; font-size: 11px; color: #000000; 
               background-color: #FFFFFF; border-style: solid; border-width: 
               1px }
-->
</style>

<body topmargin="0" leftmargin="0" bottommargin="0">
<table height="100%" width="100%" cellspacing="0">
<tr>
<td height="75%" valign="top" style="border: 1 solid #DDDDDD" align="center"><p class="texto_pagina">

<br>

<%
If Not IsEmpty(Request.Form) Then
  strMsg = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
  strMsg = strMsg & "<HTML><HEAD>"
  strMsg = strMsg & "<META content='text/html; charset=iso-8859-1' http-equiv=Content-Type>"
  strMsg = strMsg & "<META content='MSHTML 5.00.2614.3500' name=GENERATOR></HEAD>"
  strMsg = strMsg & "<BODY><FONT face=Tahoma size=2><B>" & cabecalho_email & "</B><BR><BR>"
  strMsg = strMsg & "<B>Nome*</B><BR><BR>"  & Trim(Request.Form("Campo1")) & "<HR size=1 color=gainsboro>"
  strMsg = strMsg & "<B>Telefone*</B><BR><BR>"  & Trim(Request.Form("Campo2")) & "<HR size=1 color=gainsboro>"
  strMsg = strMsg & "<B>e-mail*</B><BR><BR>"  & Trim(Request.Form("Campo3")) & "<HR size=1 color=gainsboro>"
  strMsg = strMsg & "<B>Celular</B><BR><BR>"  & Trim(Request.Form("Campo4")) & "<HR size=1 color=gainsboro>"
  strMsg = strMsg & "<B>Mensagem*</B><BR><BR>"  & Trim(Request.Form("Campo5")) & "<HR size=1 color=gainsboro>"
  strMsg = strMsg & "</FONT></BOBY>"
  strMsg = strMsg & "</HTML>"
  EnviaEmail servidor, componente, email_form, email_form, email_form, assunto_form, strMsg
Function EnviaEmail(Host,Componente,Email,NomeEmail,ParaEmail,Assunto,Mensagem)
Select Case Componente
Case "AspMail"

on error resume next
Set eObjMail = Server.CreateObject("SMTPsvg.Mailer")
eObjMail.FromName = NomeEmail
eObjMail.FromAddress = Email
eObjMail.RemoteHost = Host
eObjMail.AddRecipient "", ParaEmail
eObjMail.Subject = Assunto
eObjMail.ContentType = "text/html"
eObjMail.BodyText = Mensagem	    
eObjMail.SendMail
Set eObjMail = nothing

Case "AspEmail"

on error resume next
Set eObjMail = Server.CreateObject("Persits.MailSender")
eObjMail.Host = Host
eObjMail.From = Email
eObjMail.FromName = NomeEmail
eObjMail.AddReplyTo Email
eObjMail.AddAddress ParaEmail
eObjMail.Subject = Assunto
eObjMail.isHTML = true
eObjMail.Body = Mensagem	 	
eObjMail.Send
Set eObjMail = nothing

Case "AspQmail"

on error resume next
Set eObjMail = Server.CreateObject("SMTPsvg.Mailer")
eObjMail.QMessage = 1
eObjMail.FromName = NomeEmail
eObjMail.FromAddress = Email
eObjMail.RemoteHost = Host
eObjMail.AddRecipient "", ParaEmail
eObjMail.Subject = Assunto
eObjMail.BodyText = Mensagem
objNewMail.SendMail
Set eObjMail = nothing
		
Case "CDONTS"

on error resume next
Set eObjMail = Server.CreateObject("CDONTS.NewMail")
eObjMail.to = ParaEmail
eObjMail.from = NomeEmail & "<" & Email & ">"
eObjMail.subject = Assunto
eObjMail.Importance = 1
eObjMail.BodyFormat = 0
eObjMail.MailFormat = 0
eObjMail.body = Mensagem		
eObjMail.send
Set eObjMail = nothing


Case "DundasMailer"

Set eObjMail = Server.CreateObject("Dundas.Mailer")
eObjMail.TOs.Add ParaEmail
eObjMail.Priority = 1
eObjMail.FromName =  "Fcosta Imobiliaria"
eObjMail.FromAddress = ParaEmail
eObjMail.Subject = Assunto
eObjMail.HTMLBody = Mensagem
eObjMail.SendMail
Set eObjMail= nothing


		
End Select
End Function

%>
<font face="Arial" style="font-size: 13px"><font color="#FF0000"><b>Dados enviados</b></font><BR>
<font color="#FF0000"><b>Obrigado, seus dados foram enviados com sucesso</b>.
</font></font><font color="#FF0000">

<%
Else
%>

&nbsp;
</font>
<form name="form_incluir" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onSubmit="return verifica_form(this);">
<p>
<font color="#FF0000" face="Arial">
<font size="2"><b>No momento nosso suporte está indisponível...</b></font><b><span style="font-size: 13px"><br>
</span></b></font><font face="Arial" style="font-size: 12px">Deixe uma mensagem, em breve um de nossos atendentes entrará em contato.<br>
Utilize o formulário abaixo para deixar sua mensagem.<br>
</font>
<font color="#FF0000" style="font-size: 10px" face="Arial">* Campos obrigatórios</font></p>
<TABLE border=0 cellpadding=2 cellspacing=1 class=tabela_formulario width="243">
<TR class=titulo_campos><TD width="111">Nome*&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</TD><TD width="122">
<INPUT style="border-style: solid; border-width: 1" type="text" name="Campo1" maxlength="60" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario size="20">
</TD></TR>
<TR class=titulo_campos><TD width="111">Telefone*
</TD><TD width="122">
<INPUT style="border-style: solid; border-width: 1" type="text" name="Campo2" maxlength="20" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario size="20">
</TD></TR>
<TR class=titulo_campos><TD width="111">e-mail*
</TD><TD width="122">
<INPUT style="border-style: solid; border-width: 1" type="text" name="Campo3" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" df_validar="email" class=campos_formulario size="20">
</TD></TR>
<TR class=titulo_campos><TD width="111">Celular
</TD><TD width="122">
<INPUT style="border-style: solid; border-width: 1" type="text" name="Campo4" maxlength="20" value="" onKeyPress="desabilita_cor(this)" class=campos_formulario size="20">
</TD></TR>
<TR class=titulo_campos><TD valign="top" width="111">Mensagem*&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</TD><TD width="122">
<TEXTAREA style="border-style: solid; border-width: 1" name="Campo5"  df_verificar="sim" onKeyPress="desabilita_cor(this)" class=campos_formulario rows="4" cols="34"></TEXTAREA></TD></TR>
</TABLE>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;
<input type="submit" name="submit" value="Enviar" class=botao_enviar style="font-family: arial; font-size: 11">
</form>
<%
End If
%>
</td></tr>
</table>
</td>
</tr>
<tr bgcolor="<%=strConfigTopo%>">
<td valign="middle" colspan="2" align="center">
</body>
</html>