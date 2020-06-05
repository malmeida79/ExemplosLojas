<%
'This is a human-readable summary of the Legal Code (the full license).
%>
<%
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.addHeader "pragma","no-cache"
	Response.addHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	Session.lCid = 1046
	
	If NOT Session("mysAdmin") Then Response.Redirect "mysAdmSair.asp"
%>
<!--#include file="db.asp"-->
<html>
<head>
<title> &#9658;&#9658;&#9658; MySupport - Atendimento Online</title>
<style>
a:link	{font-size:8pt; font-family: Tahoma, Verdana; color:000000; TEXT-DECORATION: none;}
a:visited	{font-size:8pt; font-family: Tahoma, Verdana; color:000000; TEXT-DECORATION: none;}
a:hover	{font-size:8pt; font-family: Tahoma, Verdana; color:F4B511; TEXT-DECORATION: none;}
body	{ font-family: Tahoma, Verdana; font-size: 8pt; }

body 			{ scrollbar-face-color: #E2E5EA;}
body 			{ scrollbar-shadow-color: #687888;}
body 			{ scrollbar-highlight-color: #ffffff;}
body 			{ scrollbar-3dlight-color: #687888;}
body 			{ scrollbar-darkshadow-color: #E2E5EA;}
body 			{ scrollbar-track-color: #bcbfc0;}
body 			{ scrollbar-arrow-color: #6e7e88;}

td		{ font-family: Tahoma, Verdana; font-size: 8pt; }
.campo{				
		font-family: Arial, Verdana; 
		font-size: 11px; 
		background-color: #FFFFFF;	
		border-left: 1 solid #68A4C8; 
		border-right: 1 solid #B8D0D8; 
		border-top: 1 solid #68A4C8; 
		border-bottom: 1 solid #B8D0D8;
		}
				
.botao 	{
			background-color: #E8E8E8; 
			color: black; 
			border-color: #FFFFFF; 
			border-width: 1px; 
			font-family: Tahoma, Verdana; 
			font-size: 8pt;
		}
</style>
</head>
<body topmargin="0" leftmargin="0" bottommargin="0">
<table width="95%" cellpadding="1" align="center">
	<tr><td colspan="2" height="15"></td></tr>
	<td valign="top"><img src="img/t_historico.gif" alt="" border="0">
	</td><td align="right"></td></tr>
	<tr><td colspan="2" height="1" bgcolor="DDDDDD"></td></tr>
	<tr><td colspan="2" height="15"></td></tr>
</table>
<%
If Ok(Request.QueryString("msg_erro")) <> "" Then
	Response.Write "<center><FONT COLOR=""#FF0000"">" & Ok(Request.QueryString("msg_erro")) & "</FONT></center><br><br>"
End If
If Ok(Request.QueryString("msg_ok")) <> "" Then
	Response.Write "<center><FONT COLOR=""#009900"">" & Ok(Request.QueryString("msg_ok")) & "</FONT></center><br><br>"
End If
%>
<table width="95%" cellpadding="0" cellspacing="0" align="center">
<tr><td width="20%" valign="top">
	<table width="95%" cellpadding="1" align="center" style="border: 1 solid #F2F2F2">
	<tr>
		<td><select class="campo" name="operador" size="1" style="width: 180px" onchange="javascript:location.href='mysAdmHistorico.asp?operador='+this.value;">
		<option></option>
		<%
		Call AbreDB
		Dim intOperador
		intOperador = Request.QueryString("operador")
		If intOperador = "" OR isNumeric(intOperador) = False Then intOperador = "0"
		
		strSQL = "SELECT id, nome FROM operadores ORDER BY nome"
		Set rs = Conexao.Execute(strSQL)
		If NOT rs.EOF Then
			While Not rs.EOF
				Response.Write "<option value='"& rs("id") &"'"
				If rs("id") = cInt(intOperador) Then Response.Write "SELECTED"
				Response.Write ">&nbsp;" & rs("nome")
				rs.movenext
			Wend
		End If
		%>
		</select></td></tr>
	<tr><td><select class="campo" name="chamados" size="12" style="width: 180px" onchange="javascript:location.href='mysAdmHistorico.asp?operador='+operador.value+'&ID='+this.value;">
	<%
	Dim intConversa
	
	intConversa = OK(Request.QueryString("id"))
	
	strSQL = "SELECT TOP 30 inicio, id, usuario FROM conversas WHERE operador = " & intOperador & " ORDER BY ID DESC"
	Set rs = Conexao.Execute(strSQL)
	If NOT rs.EOF Then
		While Not rs.EOF
			Response.Write "<option value='"& rs("id") &"'"
			If CStr(rs("id")) = intConversa Then Response.Write "SELECTED"
			Response.Write ">&nbsp;" & FormataData(rs("inicio"),"dd.mm.yyyy") & " - " & rs("usuario")
			rs.movenext
		Wend
	Else
		Response.Write "Nenhum Chamado"
	End If
	%>
	</select></td></tr></table>
</td><td valign="top">
	<table width="95%" height="100%" cellpadding="1" align="center" style="border: 1 solid #F2F2F2">
	<%
	If intConversa <> "" AND isNumeric(intConversa) Then
		strSQL = "SELECT id, usuario, assunto, email, ip, inicio, texto FROM conversas WHERE id = " & intConversa
		Set rs = Conexao.Execute(strSQL)
		With Response
		If NOT rs.EOF Then
			.Write "<tr bgcolor=""F5F5F5""><td valign=""top"" width=""50%"">"
			.Write "<b>&nbsp;Chamado:</b>&nbsp; #" & rs("id") & "<br>"
			.Write "<b>&nbsp;Usuário:</b>&nbsp;" & rs("usuario") & "<br>"
			.Write "<b>&nbsp;Assunto:</b>&nbsp;" & rs("assunto") & "<br>"
			.Write "</td><td>"
			.Write "<b>&nbsp;Email:</b>&nbsp;" & rs("email") & "<br>"
			.Write "<b>&nbsp;IP:</b>&nbsp;" & rs("ip") & "<br>"
			.Write "<b>&nbsp;Inicio:</b>&nbsp;" & FormataData(rs("inicio"),"dd.mm.yyyy") & "<br>"
			.Write "</td></tr>"
			.Write "<tr><td colspan=""2"" height=""1"" bgcolor=""F5F5F5""></td></tr>"
			.Write "<tr><td colspan=""2"">"
			.Write rs("texto") & "</td></tr>"
		Else
			.Write "Nenhum Chamado"
		End If
		End With
	End If
	rs.close
	Call FechaDB
	%>
	<tr><td colspan="2">&nbsp;</td></tr>
	</table>
</td></tr>
</table>
</body>
</html>
