<%

%>
<!-- #include file="topo.asp" -->
<%
'Inserir a compra no banco de dados
dim vitenslista(10)

Sub adicionac(nOrderID, nProductID, nQuant)

textosql = "INSERT INTO pedidos (idcompra, idprod, quantidade, ic_tipo) values ("&nOrderID&", "&nProductID&", "&nQuant&", '*')"
response.Write textosql

abredb.Execute(textosql)
End Sub

Sub  contador(nOrderID,nProductID)
Set rs = abredb.execute("SELECT contador FROM produtos WHERE idprod=" & intProdID)
if VersaoDb = 1 Then
abredb.execute("UPDATE produtos SET contador='" & rs("contador") + 1 & "' WHERE idprod='" & intProdID & "'")
else
abredb.execute("UPDATE produtos SET contador=" & rs("contador") + 1 & " WHERE idprod=" & intProdID)
end if

End Sub 


vitenslista(0) = Request.Form("id_pm")
vitenslista(1) = Request.Form("id_pp")
vitenslista(2) = Request.Form("id_mm")
vitenslista(3) = Request.Form("id_hd")
vitenslista(4) = Request.Form("id_pv")
vitenslista(5) = Request.Form("id_pr")
vitenslista(6) = Request.Form("id_ps")
vitenslista(7) = Request.Form("id_pf")
vitenslista(8) = Request.Form("id_cd")
vitenslista(9) = Request.Form("id_dvd")



'Response.Write vitenslista(0)
'Response.Write vitenslista(1)
'Response.Write vitenslista(2)
'Response.Write vitenslista(3)
'Response.Write vitenslista(4)
'Response.Write vitenslista(5)
'Response.Write vitenslista(6)
'Response.Write vitenslista(7)
'Response.Write vitenslista(8)
'Response.Write vitenslista(9)






intQuant = 1
intOrderID = cstr(Session("orderID"))

'Incrementa o Contador do Produto. Este é usado para fazer os campeões de venda.


for i = 0 to 9
	if vitenslista(i)<> "" then 
	contador  nOrderID, vitenslista(i)
	adicionac intOrderID, vitenslista(i), intQuant
	end if 
	
next

items = Request.Form("id_dvd")
if items <>"" then 
items_dvd =split(items, "@esp@")
'Response.Write 		items_dvd(0)
'Response.Write 		items_dvd(1)
'Response.Write 		items_dvd(2)
'Response.end

for i = 0 to uBound(items_dvd) - 1
if items_dvd(i)<> "" then 
	contador  nOrderID, items_dvd(i)
	adicionac intOrderID, items_dvd(i), intQuant
	end if
	
Next 	
end if 				



'textosql = "SELECT * FROM pedidos WHERE idcompra ='" & intOrderID & "' AND idprod ='" & intProdID & "';"
'set adiciona = abredb.Execute(textosql)
'if adiciona.EOF then
'	txtInfo = "adicionado"  
	
'	estadozx = mid(request("frete"),2,3)
'	intOrderIDx = cstr(Session("orderID"))
'	set rsProdx = abredb.Execute("SELECT peso FROM produtos WHERE idprod="&intProdID&";")
'	peso = rsProdx("peso")
'	rsProdx.close

	'Calcula o frete
'	set rsprodx = nothing
'	quantidade = intQuant
'	valor = peso * quantidade
'	pesoz = pesoz + valor
'	if session("estado") = "" then
'	session("peso") = 0
'	else
'	session("peso") = pesoz
'	end if
'	fretexz = right(request("frete"),1)
'	numerox = left(request("frete"),1)
'	if numerox = "" then
'	on error resume next
'	end if
'	if fretexz = "c" then
'	sqlq = "SELECT * FROM fretes WHERE localidade='pesocapital';"
'	else
'	sqlq = "SELECT * FROM fretes WHERE localidade='pesointerior';"
'	end if
'	Set dadosz = abredb.Execute(sqlq)
'	if dadosz.EOF or dadosz.BOF then
'	regi = "0,00"
'	else
'	regi = dadosz("re"&numerox&"")
'	end if
'	dadosz.close
'	Set dadosz = nothing
'	fretez = right(request("frete"),1)
'	numero = left(request("frete"),1)
'	if fretez = "c" then
'	sql = "SELECT * FROM fretes WHERE localidade='capital';"
'	else
'	sql = "SELECT * FROM fretes WHERE localidade='interior';"
'	end if
'	Set dados = abredb.Execute(sql)
'	if dados.EOF or dados.BOF then
'	regiao = "0,00"
'	else
'	regiao = dados("re"&numero&"")
'	end if
'	if pesoz <= 1 then
'	session("frete") = regiao
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 1 AND pesoz <= 2 then
'	session("frete") = regiao + (regi * 1)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 2 AND pesoz <= 3 then
'	session("frete") = regiao + (regi * 2)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 3 AND pesoz <= 4 then
'	session("frete") = regiao + (regi * 3)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 4 AND pesoz <= 5 then
'	session("frete") = regiao + (regi * 4)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 5 AND pesoz <= 6 then
'	session("frete") = regiao + (regi * 5)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 6 AND pesoz <= 7 then
'	session("frete") = regiao + (regi * 6)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 7 AND pesoz <= 8 then
'	session("frete") = regiao + (regi * 7)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 8 AND pesoz <= 9 then
'	session("frete") = regiao + (regi * 8)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 9 AND pesoz <= 10 then
'	session("frete") = regiao + (regi * 9)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 10 AND pesoz <= 11 then
'	session("frete") = regiao + (regi * 10)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 11 AND pesoz <= 12 then
'	session("frete") = regiao + (regi * 11)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 12 AND pesoz <= 13 then
'	session("frete") = regiao + (regi * 12)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 13 AND pesoz <= 14 then
'	session("frete") = regiao + (regi * 13)
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	if pesoz > 14 AND pesoz => 15 then
'	session("frete") = regiao + (regi * int(pesoz))
'	session("estado") = estadozx
'	session("peso") = pesoz
'	end if
'	dados.close
'	Set dados = nothing
'	session("estado2") = request("frete")
'else
'session("frete") = session("frete")
'session("estado") = session("estado")
'session("peso") = session("peso")
'session("estado2") = session("estado2")
'txtInfo = "existente"
'end if

'Chama dados do produto
'Set rsProdInfo = abredb.Execute("SELECT * FROM produtos where idprod="&intProdID)
'idsessao = rsProdInfo("idsessao")
'nome = rsProdInfo("nome")
'Set nomes = abredb.Execute("SELECT * FROM sessoes where id="&idsessao)
'strNome = nomes("nome")
'nomes.Close
'set nomes = Nothing
'rsProdInfo.Close
'set rsProdInfo = Nothing

'Fecha banco de dados
abredb.Close
set abredb = Nothing
response.redirect "default.asp"
%>
