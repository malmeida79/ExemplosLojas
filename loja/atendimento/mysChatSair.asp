<%
'This is a human-readable summary of the Legal Code (the full license).
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
	
If Session("mysConversaID") <> "" Then
	
	intConversaID = Session("mysConversaID")
	
	Call AbreDB
	
	strUltMsg 		  = FormataData(now,"yyyy/mm/dd hh:nn:ss")
	
	strMensagem = Session("mysNome") & " deixou o atendimento."
	
	strMensagem = "<font id=""mysChatFinalizado"" color=""#0000FF""><b>&nbsp;</b>" & Server.HTMLEncode(strMensagem) & "</font><br><br>"
		strSQL = "UPDATE conversas SET ult_msg = #"& strUltMsg &"#, texto = texto & '" & strMensagem & "' WHERE id = " & Session("mysConversaID")
	
	Conexao.Execute(strSQL)
	
	Call FechaDB
	
	Session.Abandon
ElseIf OK(Request("ConversaID")) <> "" Then
	If Request.ServerVariables("HTTP_METHOD") = "POST" Then
	Call AbreDB
		intConversaID	=	OK(Request("ConversaID"))
		strIP			=	Request.ServerVariables("REMOTE_ADDR")
		intAvaliacao	=	OK(Request("avaliacao"))
		
		strSQL = "UPDATE conversas SET nota = " & intAvaliacao & " WHERE id = " & intConversaID & " AND ip = '" & strIP & "'"
		Conexao.Execute(strSQL)
	Call FechaDB
	End If
	Response.Write "<script>window.close();</script>"
	Response.End
Else
	Response.Write "<script>window.close();</script>"
	Response.End
End If
%>
<html>
<head>
<title><%=strConfigNome%></title>
<style>
a:link	{font-size:8pt; font-family: Tahoma, Verdana; color:000000; TEXT-DECORATION: none;}
a:visited	{font-size:8pt; font-family: Tahoma, Verdana; color:000000; TEXT-DECORATION: none;}
a:hover	{font-size:8pt; font-family: Tahoma, Verdana; color:F4B511; TEXT-DECORATION: none;}
body	{ font-family: Tahoma, Verdana; font-size: 8pt; color: <%=strConfigFonte%>; }
td		{ font-family: Tahoma, Verdana; font-size: 8pt; color: <%=strConfigFonte%>; }
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
<script language="JavaScript">
function validarForm(){
	if(form.avaliacao[0].checked == false && form.avaliacao[1].checked == false && form.avaliacao[2].checked == false && form.avaliacao[3].checked == false && form.avaliacao[4].checked == false) {
		alert('Favor escolher uma avalia��o para nosso atendimento.');
		return false;
	}
return true;
}
</script>
<body topmargin="0" leftmargin="0" bottommargin="0">
<table height="100%" width="100%" cellspacing="0">
	<tr bgcolor="<%=strConfigTopo%>">
		<td valign="top" width="50%" height="10%">
			<img src="<%=strConfigLogo%>" alt="" border="0">
		</td>
		<td valign="top">
		</td>
	</tr>
	<tr>
		<td colspan="2" height="75%" valign="top" style="border: 1 solid #DDDDDD" align="center"><br><br>
			<form name="form" action="mysChatSair.asp" method="post" onSubmit="return validarForm();">
			<table width="70%" align="center">
				<tr>
					<td colspan="2" align="center"><font size="2"><b>Por favor, avalie nosso atendimento.</b></font></td>
				</tr>
				<tr>
					<td colspan="2" height="15"><input type="hidden" name="ConversaID" value="<%=intConversaID%>"></td>
				</tr>
				<tr>
					<td width="50%" align="right"><input type="radio" name="avaliacao" value="5" checked>&nbsp;</td>
				<td>Excelente</td></tr>
				<tr>
				<td width="50%" align="right"><input type="radio" name="avaliacao" value="4">&nbsp;</td>
				<td>Muito Bom</td></tr>
				<tr>
				<td width="50%" align="right"><input type="radio" name="avaliacao" value="3">&nbsp;</td>
				<td>Bom</td></tr>
				<tr>
				<td width="50%" align="right"><input type="radio" name="avaliacao" value="2">&nbsp;</td>
				<td>Regular</td></tr>
				<tr>
				<td width="50%" align="right"><input type="radio" name="avaliacao" value="1">&nbsp;</td>
				<td>Ruim</td></tr>
				<tr>
					<td colspan="2" height="15"></td>
				</tr>
				<tr>
					<td colspan="2" align="center">
			<input type="submit" value="    Avaliar    " class="botao">
			</td></tr>
			</table>
			</form>
		</td>
	</tr>
	<tr bgcolor="F5F5F5">
		<td align="center" valign="middle" colspan="2">
			<br>
		</td>
	</tr>
	<tr bgcolor="<%=strConfigTopo%>">
		<td valign="middle" colspan="2" align="center">
			<font size="1">

			</font>
		</td>
	</tr>
</body>
</html>
