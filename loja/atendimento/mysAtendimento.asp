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

Call AbreDB
Dim vetDepartamentos, intIndice

Session("mysConversaID") = ""

strSQL = "SELECT count(id) AS Total FROM operadores WHERE  DateDiff(""s"", ping,now()) < 10"

Set rs = Conexao.execute(strSQL)

If rs("Total") = "0" OR isNull(rs("Total")) OR rs("Total") = 0 Then
	Response.Redirect "mysChatOff.asp"
End If


If Request.ServerVariables("HTTP_METHOD") = "POST" Then
	strNome 		= OK(Request.Form("nome"))
	strAssunto 		= OK(Request.Form("assunto"))
	strEmail 		= OK(Request.Form("email"))
	intDepartamento	= OK(Request.Form("departamento"))
	
	strSQL = "INSERT INTO conversas(departamentoID, usuario, assunto, email, ip, inicio, ult_msg, status) VALUES ("
	strSQL = strSQL & "" & intDepartamento &","
	strSQL = strSQL & "'"& strNome &"',"
	strSQL = strSQL & "'"& strAssunto &"',"
	strSQL = strSQL & "'"& strEmail &"',"
	strSQL = strSQL & "'"& Request.ServerVariables("REMOTE_ADDR") &"',"
	strSQL = strSQL & "#"& FormataData(now,"yyyy/mm/dd hh:nn:ss") &"#,"
	strSQL = strSQL & "#"& FormataData(now,"yyyy/mm/dd hh:nn:ss") &"#,"
	strSQL = strSQL & " True"
	strSQL = strSQL & ")"
	
	Conexao.Execute(strSQL)
	
	strSQL = "SELECT id, ult_msg FROM conversas WHERE ip = '"& Request.ServerVariables("REMOTE_ADDR") &"' ORDER BY id DESC"
	
	Set rs = Conexao.Execute(strSQL)
	
	Session("mysConversaID") = rs("id")
	Session("mysNome") = strNome
	Response.Cookies("MySupport").Expires=#Dec 31,2009#
	Response.Cookies("MySupport")("Nome") 	= strNome
	Response.Cookies("MySupport")("Email") 	= strEmail
	Response.Cookies("MySupport")("Assunto")= strAssunto
	
	rs.close
	Response.Redirect "mysPreChat.asp"
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
	if(form.nome.value == '') {
		alert('Voc� deve preencher o campo nome!');
		form.nome.focus();
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
			<form name="form" action="mysAtendimento.asp" method="post" onSubmit="return validarForm();">
			<table width="35%" align="center">
				<tr>
					<td colspan="2" align="center"></td>
				</tr>
				<tr>
					<td width="50%">Departamento:</td>
					<td><select class="campo" name="departamento" size="1" style="width: 180px">
					<%
				strSQL = "SELECT id, nome FROM departamentos ORDER BY nome"
				Set rs = Conexao.Execute(strSQL)
				vetDepartamentos = rs.getRows
				For intIndice = LBound(vetDepartamentos,2) To UBound(vetDepartamentos,2)
					strSQL = "SELECT count(id) AS Total FROM operadores WHERE DateDiff(""s"", ping,now()) < 10 AND departamentoID = " & vetDepartamentos(0,intIndice)
					Set rs = Conexao.Execute(strSQL)
					If rs("Total") = "0" OR isNull(rs("Total")) OR rs("Total") = 0 Then
						Response.Write "<option value='" & vetDepartamentos(0,intIndice) & "' style='color:CB0000;'>" & vetDepartamentos(1,intIndice) & " (Offline)</option>"
					Else
						Response.Write "<option value='" & vetDepartamentos(0,intIndice) & "' style='color:00A01E;'>" & vetDepartamentos(1,intIndice) & " (Online)</option>"
					End If
				Next
				%>
				</select></td></tr>
				<tr>
					<td colspan="2" height="15"></td>
				</tr>	
				<tr>
					<td width="50%">Nome:</td>
					<td><input type="text" name="nome" size="30" class="campo" value="<%=Request.Cookies("MySupport")("Nome")%>"></td></tr>
				<tr>
					<td width="50%">Email:</td>
					<td><input type="text" name="email" size="30" class="campo" value="<%=Request.Cookies("MySupport")("Email")%>"></td></tr>
				<tr>
					<td width="50%">Assunto:</td>
					<td><input type="text" name="assunto" size="30" class="campo"></td></tr>
				<tr>
					<td></td>
					<td>
			<input type="submit" value="    Entrar    " class="botao">
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
<%
rs.Close
Call FechaDB
%>