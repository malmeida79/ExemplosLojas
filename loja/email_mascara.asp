<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
'#
'#  C�DIGO: VirtuaStore Vers�o OPEN - Copyright 2001-2004 VirtuaStore
'#  URL: http://comunidade.virtuastore.com.br
'#  E-MAIL: jusivansjc@yahoo.com.br
'#  AUTORES: Ot�vio Dias(Desenvolvedor)
'# 
'#     Este programa � um software livre; voc� pode redistribu�-lo e/ou 
'#     modific�-lo sob os termos do GNU General Public License como 
'#     publicado pela Free Software Foundation.
'#     � importante lembrar que qualquer altera��o feita no programa 
'#     deve ser informada e enviada para os criadores, atrav�s de e-mail 
'#     ou da p�gina de onde foi baixado o c�digo.
'#
'#  //-------------------------------------------------------------------------------------
'#  // LEIA COM ATEN��O: O software VirtuaStore OPEN deve conter as frases
'#  // "Powered by VirtuaStore OPEN" ou "Software derivado de VirtuaStore 1.2" e 
'#  // o link para o site http://comunidade.virtuastore.com.br no cr�ditos da 
'#  // sua loja virtual para ser utilizado gratuitamente. Se o link e/ou uma das 
'#  // frases n�o estiver presentes ou visiveis este software deixar� de ser
'#  // considerado Open-source(gratuito) e o uso sem autoriza��o ter� 
'#  // penalidades judiciais de acordo com as leis de pirataria de software.
'#  //--------------------------------------------------------------------------------------
'#      
'#     Este programa � distribu�do com a esperan�a de que lhe seja �til,
'#     por�m SEM NENHUMA GARANTIA. Veja a GNU General Public License
'#     abaixo (GNU Licen�a P�blica Geral) para mais detalhes.
'# 
'#     Voc� deve receber a c�pia da Licen�a GNU com este programa, 
'#     caso contr�rio escreva para
'#     Free Software Foundation, Inc., 59 Temple Place, Suite 330, 
'#     Boston, MA  02111-1307  USA
'# 
'#     Para enviar suas d�vidas, sugest�es e/ou contratar a VirtuaWorks 
'#     Internet Design entre em contato atrav�s do e-mail 
'#     jusivansjc@yahoo.com.br ou pelo endere�o abaixo: 
'#     Rua Ven�ncio Aires, 1001 - Niter�i - Canoas - RS - Brasil. Cep 92110-000.
'#
'#     Para ajuda e suporte acesse um dos sites abaixo:
'#     http://comunidade.virtuastore.com.br
'#     http://www.bondhost.com.br
'#
'#     BOM PROVEITO!          
'#     Equipe VirtuaStore
'#     []'s
'#
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################

'IN�CIO DO C�DIGO

'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP

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
		eObjMail.Username = "corpsjc@megapaper.com.br"
		eObjMail.Password = "20061234"
		eObjMail.Host = "mail.megapaper.com.br"
		eObjMail.From = "corpsjc@megapaper.com.br"
		eObjMail.FromName = NomeEmail
		eObjMail.AddReplyTo Email
		eObjMail.AddAddress ParaEmail
		eObjMail.Subject = Assunto
		eObjMail.isHTML = true
		eObjMail.Body = Mensagem
		eObjMail.Send
		Set eObjMail = nothing

		
Case "CDONTS"
	    on error resume next

	' Texto html
  		msgHTML = Mensagem

	' Definindo uma variavel auxiliar
  		sch = "http://schemas.microsoft.com/cdo/configuration/"

  	' Criando o objeto de configura��o do CDO
  		Set cdoConfig = Server.CreateObject("CDO.Configuration")

  	' Definindo as configura��es
  		cdoConfig.Fields.Item(sch & "sendusing") = 2
  		cdoConfig.Fields.Item(sch & "smtpauthenticate") = 1
  		cdoConfig.Fields.Item(sch & "smtpserver") = "mail.megapaper.com.br"
  		cdoConfig.Fields.Item(sch & "sendusername") = "corpsjc@megapaper.com.br"
 		cdoConfig.Fields.Item(sch & "sendpassword") = "SENHA" 
 		cdoConfig.fields.update

 	' Criando o objeto de msg do CDO
  		Set cdoMessage = Server.CreateObject("CDO.Message")

  	' Associando as configura��es ao obj Mensagem
  		Set cdoMessage.Configuration = cdoConfig

  	' Definido variaveis da msg
  		cdoMessage.From = "corpsjc@megapaper.com.br"
  		cdoMessage.To = ParaEmail
  		cdoMessage.Subject = Assunto

 	    cdoMessage.HTMLBody =  msgHTML
  		if msgHTML <> "" then
    	cdoMessage.AutoGenerateTextBody = false
    	cdoMessage.TextBody =  msgHTML	
  		end if

  		cdoMessage.Send
  		Set cdoMessage = Nothing
  		Set cdoConfig = Nothing
	
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
		
Case "JMail"
		Set jmail = Server.CreateObject("JMail.Message") 
		jmail.Logging = true 
		jmail.silent = true 
		jmail.AddRecipient trim(ParaEmail) 
		jmail.Subject = Assunto 
		jmail.From = Application("emailloja")
		jmail.FromName = NomeEmail
		jmail.AppendHTML Mensagem
		jmail.Maildomain = Replace(Replace(Replace(request.servervaraibles("REMOTE_HOST"), "www.", ""), "http://", ""), "https://", "")
		jmail.Send(Host)
		set jmail = nothing	
		
End Select
End Function
%>