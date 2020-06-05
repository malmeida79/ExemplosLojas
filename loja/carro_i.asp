<style type="text/css">
<!--
.style3 {color: #FFFFFF; font-weight: bold; }
.style4 {color: #FFFFFF}
-->
</style>
<%= strLgMoeda %>
<%
response.Write " </td></tr></table>"
else
	intOrderID = cstr(Session("orderID"))
	'Chama o nome do cliente da tabela clientes
	Set dados = abredb.Execute("SELECT email, nome FROM clientes WHERE email='"&Cripto(session("usuario"),true)&"';")
	'Faz aparecer somento o primeiro nome do cliente
	if dados.EOF then
		nomez = ""
	else
		nomeq = Cripto(dados("nome"),true)
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
'Chama os dados dos produtos comprados e monta o carrinho
	set pedidos = abredb.Execute("SELECT idprod, quantidade FROM pedidos WHERE idcompra='"&intOrderID&"';")
	while not pedidos.EOF
		idprod = pedidos("idprod")
		quantidade = pedidos("quantidade")
		set produtos = abredb.Execute("SELECT preco, nome FROM produtos WHERE idprod="&idprod&";")
		preco = produtos("preco")
		nome = produtos("nome")
	'Calcula os dados
		intProdID = idprod
		strProdName = nome
		intProdPrice = preco
		intQuant = cint(quantidade)
		intQuantx = cint(intQuantx) + cint(intQuant)	
		intTotal = intTotal + (intQuant * intProdPrice)
		inTotal = cint(inTotal)
		pedidos.MoveNext
	wend
'Fecha os dados dos produtos comprados
pedidos.Close
set pedidos = Nothing
'Valor do preço total
precito1 = formatNumber(intTotal,2)

Response.Write	" <table onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"&strLg1&"';return true;"">"&VBCrlf &_
				"<tr>"&VBCrlf &_
				"<td >"&VBCrlf &_
		   		"<table width=""10%"">"&VBCrlf &_
				"	   <tr>"&VBCrlf &_
				"<td>"&VBCrlf
%>