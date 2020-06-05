<div align="center">
<!-- #include file="df.asp" -->
  <!-- #include file="estats.asp"-->
  <!-- #include file="banner/include/admentor2.asp"-->
  <%
'INICIO DO HTML
'-----------------------------------------------------------------------------------------------------
response.Write	"<html>"&VBCRlf &_
				"<head>"&VBCRlf &_
				"<title>"&tituloloja&"</title>"&VBCRlf &_
				"<script language=""JavaScript"" src=""layer.js""></script>"&VBCRlf &_
				"<script language=""JavaScript"" src=""Browser_OS_Detection_css.js""></script>"&VBCRlf &_
				"<script language='JavaScript1.2'>"&VBCRlf &_
				"function click() {"&VBCRlf &_
				"if (event.button==2||event.button==3) {"&VBCRlf &_
				"oncontextmenu='return false';"&VBCRlf &_
				"}"&VBCRlf &_
				"}"&VBCRlf &_
				"document.onmousedown=click"&VBCRlf &_
				"document.oncontextmenu = new Function(""return false;"")"&VBCRlf &_
				pop & VBCRlf &_
  				"</script>"&VBCRlf &_
				"<style type=""text/css"">"&VBCRlf &_
				"<!--"&VBCRlf &_
				"a:link       { color: "&cor4&" }"&VBCRlf &_
				"a:visited    { color: "&cor4&" }"&VBCRlf &_
				"a:hover      { color: "&cor5&" }"&VBCRlf &_
				".menu:link { color:"&fontebranca&"; text-decoration: none;}"&VBCRlf &_
				".menu:visited { color:"&fontebranca&"; text-decoration: none;}"&VBCRlf &_
				".menu:hover { color:"&fontebranca&"; text-decoration: underline;}"&VBCRlf &_
				".baixo:link { color:"&cor4&"; text-decoration: none;}"&VBCRlf &_
				".baixo:visited { color:"&cor4&"; text-decoration: none;}"&VBCRlf &_
				".baixo:hover { color:"&cor4&"; text-decoration: underline;}"&VBCRlf &_
				"-->"&VBCRlf &_
				"</style>"&VBCRlf &_
				"</head>"&VBCRlf &_
				"<body bgcolor=""#00204"" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 text="&cor4&">"&VBCRlf

'Personaliza os links do  menu se o cliente estiver efetuando a compra
if session("usuario") = "" then
	link = "fechapedido.asp?compra=login"
else
	link = "fechapedido.asp?compra=ok"
end if
'-----------------------------------------------------------------------------------------------------
if session("ende1") = "" then
	link = link
else
	link = "formaspagamento.asp"
end if
'-----------------------------------------------------------------------------------------------------
response.Write  "<table width=100% height=80 cellpadding=""0"" cellspacing=""0"" align=""center"" >"&VBCRlf &_
				"<tr>"&VBCRlf &_
				"<td>"&VBCRlf &_
				"<basefont face="&fonte&">"&VBCRlf &_
				"<table valign=""top"" width=""100%"">"&VBCRlf
response.write	"		 <tr>"&VBCRlf &_
				"			"&VBCRlf

if Banner_AdMentor="Nao" then
response.Write	"<td valign=""middle"" align=""center"" width=234 height=60>"&VBCRlf &_
				AdMentor_GetAdASP("F=0&Z=0&N=1") &VBCRlf &_
				"</td>"&VBCRlf
elseif Banner_Fixo<>"" then
response.Write	"<td valign=""middle"" align=""center"" width=234 height=60><img src="" banners/"&Banner_Fixo&"""  border=""0""></td>"&VBCRlf
end if

response.Write	"<td valign=""bottom"" nowrap align=""right"">"&VBCRlf &_
				"<table>"&VBCRlf &_
				"<tr>"&VBCRlf &_
				"<td valign=center></td>"&VBCRlf &_
				"<td valign=top>"&VBCRlf 
'-----------------------------------------------------------------------------------------------------
'#### INICIO DO 1º IF 
'Cria o carrinho de compras no topo superior da loja
if cstr(Session("orderID")) = "" then
	'Chama o nome do cliente da tabela clientes
	Set dados = abredb.Execute("SELECT email,nome FROM clientes WHERE email='"&Cripto(session("usuario"),true)&"';")
	'Faz aparecer somento o primeiro nome do cliente
	if dados.eof then
		nomez = ""
	else
		nomeq = Cripto(dados("nome"),false)
		numeroz = Instr(1,nomeq," ")
		var5000 = Left(nomeq,numeroz)
			if var5000 = "" then
				var5000 = nomeq
			else
				var5000 =  Left(nomeq,numeroz)
			end if
	nomez = "&nbsp;"&var5000
	end if
	'Fecha tabela clientes
	dados.Close
	set dados = Nothing
	response.Write	"<table width=""100%"" border=""0"" onMouseOver=""window.status='"&strLg1&"';return true;"" onMouseOut=""window.status='';return true;"">"&VBCrlf &_
					" <tr><td >"&VBCrlf &_
					"<table>"&VBCrlf &_
					"<tr><td bgcolor=#ffffff>"&VBCrlf
%>
  <!-- #include file="carro_i.asp" -->
  <%
				response.Write " </td></tr></table>"&VBCrlf &_
				"</td>"&VBCrlf &_
				"</tr>"&VBCrlf &_
				"</table>"&VBCrlf 
end if
'#### TÉRMINO DO 1º IF 
Response.Write	"</td>"&VBCrlf &_
				"</tr>"&VBCrlf &_
				"</table>"&VBCrlf &_
				"</td>"&VBCrlf &_
				"</tr>"&VBCrlf &_
				"</table>"&VBCrlf
'-----------------------------------------------------------------------------------------------------
set rs = abredb.execute("SELECT nome FROM clientes where email='" & Cripto(Session("usuario"),true) & "'")
'-----------------------------------------------------------------------------------------------------
if rs.eof then
	strNome = strLg264
else
	nomeq = Cripto(rs("nome"),false)
	numeroz = Instr(1,nomeq," ")
	var5000 = Left(nomeq,numeroz)
		if var5000 = "" then
			var5000 = nomeq
		else
			var5000 =  Left(nomeq,numeroz)
		end if
strNome = Trim(var5000)
strNome2 = Trim(Cripto(rs("nome"),false))
end if
rs.close
set rs = nothing
'Personaliza o menu se o cliente estiver logado
%>
</div>

