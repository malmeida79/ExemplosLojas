<%


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
	    Set Mail = Server.CreateObject("Persits.MailSender")
	    Mail.Host = "mail.compuwaretecno.com.br"
		Mail.Username = "mail@compuwaretecno.com.br"
        Mail.Password = "mail"
	 	Mail.From = Email
	 	Mail.FromName = NomeEmail
	 	Mail.AddReplyTo Email
	 	Mail.AddAddress ParaEmail
	    Mail.Subject = Assunto
	    Mail.isHTML = true
	 	Mail.Body = Mensagem	 	
	 	Mail.Send
	 	Set Mail = nothing

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