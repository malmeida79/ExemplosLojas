<%
'
'This is a human-readable summary of the Legal Code (the full license).
%>
<!--#include file="db.asp"-->
<%
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.addHeader "pragma","no-cache"
	Response.addHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	Session.LCID = 1046

	If inStr(LCase(Request.ServerVariables("HTTP_REFERER")), "http://" & LCase(Request.ServerVariables("HTTP_HOST"))) = 0 Then Response.End

Call AbreDB

Dim logChat

logChat = False

strSQL = "SELECT operador FROM conversas WHERE id = " & Session("mysConversaID")
Set rs = Conexao.Execute(strSQL)

If CInt(rs("operador")) > 0 Then
	logChat = True

strUltMsg 		  = FormataData(now,"yyyy/mm/dd hh:nn:ss")
If Request.ServerVariables("HTTP_METHOD") = "POST" Then
	strMensagem = OK(Request.Form("mensagem"))
	strMensagem = "<font color=""#BB0000""><b>"& Session("mysNome") &"&nbsp;&raquo;&nbsp;</b>" & setLink(Server.HTMLEncode(strMensagem)) & "</font><br><br>"
	
	strSQL = "UPDATE conversas SET tecla_us = False, texto = texto & '"& strMensagem &"', ult_msg = #"& strUltMsg &"# WHERE id = " & Session("mysConversaID")
	
	Conexao.Execute(strSQL)
	
	Response.Redirect "mysConversa.asp?m=POST"
End If

strSQL = "SELECT texto FROM conversas WHERE id = " & Session("mysConversaID")

Set rs = Conexao.Execute(strSQL)

strTexto = rs("texto")

End If

rs.close
Call FechaDB
%>
<html>
<style>
a:link { font-family: Tahoma, Verdana; font-size: 8pt;
	color: #000000;
	text-decoration: none;
	}
a:visited { font-family: Tahoma, Verdana; font-size: 8pt;
	color: #000000;
	text-decoration: none;
	}
a:hover { font-family: Tahoma, Verdana; font-size: 8pt;
	color: #000000;
	text-decoration: underline;
	}

body { font-family: Tahoma, Verdana; font-size: 8pt; }

body 			{ scrollbar-face-color: #E2E5EA;}
body 			{ scrollbar-shadow-color: #687888;}
body 			{ scrollbar-highlight-color: #ffffff;}
body 			{ scrollbar-3dlight-color: #687888;}
body 			{ scrollbar-darkshadow-color: #E2E5EA;}
body 			{ scrollbar-track-color: #bcbfc0;}
body 			{ scrollbar-arrow-color: #6e7e88;}
</style>
<script language="JavaScript">
function scrollChat() {
  window.scroll(0,50000)
}
<%If Request("m")="POST"Then%>
parent.form.mensagem.value="";
<%Else%>
parent.tecla.innerHTML = "&nbsp;";
<%End If%>
</script>
<body topmargin="10" leftmargin="10" onload="scrollChat()">
<%
If logChat Then
	Response.Write strTexto
End If

If inStr(strTexto,"mysChatFinalizado") Then
	Response.Write "<script>"
	Response.Write "alert('Atendente finalizou atendimento');"
	Response.Write "parent.window.close();"
	Response.Write "</script>"
End If
%>
</body>
</html>
