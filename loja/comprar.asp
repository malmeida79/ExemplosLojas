<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
'#
'#  C�DIGO: VirtuaStore Vers�o OPEN - Copyright 2001-2004 VirtuaStore
'#  URL: http://comunidade.virtuastore.com.br
'#  E-MAIL: comunidade@virtuastore.com.br
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
'#  // "bondhost - Hospedagem" ou "Software derivado de VirtuaStore 1.2" e 
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
'#     contato@virtuaworks.com.br ou pelo endere�o abaixo: 
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
%>
<!-- #include file="topo.asp" -->
<%
'Inserir a compra no banco de dados

Sub adicionac(nOrderID, nProductID, nQuant)
textosql = "INSERT INTO pedidos (idcompra, idprod, quantidade) values ("&nOrderID&", "&nProductID&", "&nQuant&")"
abredb.Execute(textosql)

'#########################################################################################
'#----------------------------------------------------------------------------------------
'# CONTROLE DE ESTOQUE POR ROGERIO SILVA
'#  O controle que estava originalmente nos arquivos comprar.asp e atualizapedido.asp foi removido destes 
'#  e readaptado para o arquivo pagamento.asp aqui para que a "baixa" do estoque seja realizado
'#  somente na Finalizacao da Compra  qdo houve a real sa�da do produto do estoque 
'#########################################################################################

Response.Redirect "carrinhodecompras.asp"
End Sub

intProdID = Request.form("intProdID")
intQuant = Request.form("intQuant")
intOrderID = cstr(Session("orderID"))

'Incrementa o Contador do Produto. Este � usado para fazer os campe�es de venda.
Set rs = abredb.execute("SELECT contador FROM produtos WHERE idprod=" & intProdID)
if VersaoDb = 1 Then
abredb.execute("UPDATE produtos SET contador='" & rs("contador") + 1 & "' WHERE idprod='" & intProdID & "'")
else
abredb.execute("UPDATE produtos SET contador=" & rs("contador") + 1 & " WHERE idprod=" & intProdID)
end if
textosql = "SELECT * FROM pedidos WHERE idcompra ='" & intOrderID & "' AND idprod ='" & intProdID & "';"
set adiciona = abredb.Execute(textosql)
if adiciona.EOF then
txtInfo = "adicionado"  

adicionac intOrderID, intProdID, intQuant
estadozx = mid(request("frete"),2,3)
intOrderIDx = cstr(Session("orderID"))
set rsProdx = abredb.Execute("SELECT peso FROM produtos WHERE idprod="&intProdID&";")
peso = rsProdx("peso")
rsProdx.close
set rsProdx = nothing

'Chama dados do produto
Set rsProdInfo = abredb.Execute("SELECT * FROM produtos where idprod="&intProdID)
idsessao = rsProdInfo("idsessao")
nome = rsProdInfo("nome")
Set nomes = abredb.Execute("SELECT * FROM sessoes where id="&idsessao)
strNome = nomes("nome")
nomes.Close
set nomes = Nothing
rsProdInfo.Close
set rsProdInfo = Nothing

end if


if Cstr(session("PesoTotalCep")) = Cstr("") then

response.redirect "compra.asp?prodid="&intProdID&"&item="&txtInfo

else

'##########################################################
'##########################################################
'##########################################################
'CALCULO DE FRETE USANDO A ROTINA DOS CORREIOS

'Chama os produtos comprados
intOrderID = Session("orderID")

set pedidos = abredb.Execute("SELECT idprod, quantidade FROM pedidos WHERE idcompra='" & intOrderID & "'")
if pedidos.eof then

response.redirect "compra.asp?prodid="&intProdID&"&item="&txtInfo

else

while not pedidos.EOF
idprod = pedidos("idprod")
quantidade = pedidos("quantidade")
set produtos = abredb.Execute("SELECT preco, nome, peso FROM produtos WHERE idprod="&idprod&";")
preco = produtos("preco")
peso = produtos("peso")
nome = produtos("nome")
intProdID = idprod
strProdNome = nome
pesoz = peso
intProdPrice = preco
intQuant = quantidade
if session("estado") = "" then
intFrete = 0
else
intFrete = valorfrete
end if
'Calcula o total do frete
intTotalFrete = intTotalFrete + (intQuant * intProdPrice)	
intTotal = intTotalFrete + intFrete	
subpreco = formatNumber(intProdPrice,2)
totpreco = formatNumber((intQuant * intProdPrice),2)
pesototal = 1 + FormatNumber(pesototal, 3) + FormatNumber((produtos("peso") * intQuant), 3) - 1
produtos.Close
set produtos = Nothing
pedidos.MoveNext
wend

end if
pedidos.Close
set pedidos = Nothing


Session("PesoTotalFrete") = FormatNumber(pesototal, 3)
suacompra = formatNumber(intTotal,2)


'--> Session("PesoTotalFrete") 'Peso Total
'--> application("CORREIOSseucep11") 'CEP da Loja
'--> application("CORREIOSmaop11") 'Mao Propria
'--> application("CORREIOSaviso11") 'Aviso de Recebimento
'--> session("PesoTotalCep") 'CEP do cliente
'--> suacompra 'Valor Declarado


UrlRetorno = "http://" & request.servervariables("Server_Name") & request.servervariables("Url")
UrlRetorno = Replace(UrlRetorno, "comprar.asp", "compra.asp?prodid="&intProdID&"&item="&txtInfo)

if Session("PesoTotalFrete") < 1 then
   Session("PesoTotalFrete")=1
   else
   if instr(Session("PesoTotalFrete"),",")<>0 then
      Session("PesoTotalFrete")=replace(Session("PesoTotalFrete"),",",".")
   end if
end if

Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")

objXMLHTTP.open "post", "http://www.correios.com.br/encomendas/precos/calculo.cfm?Servico=40010&CepDestino=" & session("PesoTotalCep") & "&CepOrigem=" & application("CORREIOSseucep11") & "&Peso=" & Session("PesoTotalFrete") & "&AvisoRecebimento=" & application("CORREIOSaviso11") & "&ValorDeclarado=" & Replace(suacompra, ".", "") & "&MaoPropria=" & application("CORREIOSmaop11"),false

objXMLHTTP.send
ValorBox = objXMLHTTP.responseText
set objXMLHTTP = nothing

primeira_tarifa = right(valorbox,len(valorbox)-instr(valorbox,"Tarifa"))
segunda_tarifa = right(primeira_tarifa,len(primeira_tarifa)-instr(primeira_tarifa,"Tarifa"))
final = replace(replace(left(right(segunda_tarifa,len(segunda_tarifa)-6),5),"&",""),".",",")

session("PesoTotalValor")=final

response.redirect UrlRetorno

'Fecha banco de dados
abredb.Close
set abredb = Nothing
%>
<table width=590 align=center><tr><td><font face=<%=fonte%> style=font-size:11px><b><br>Atualizando... por favor aguarde.</b></td></tr></table>
   </td></tr>
  </table></td></tr>
 </table>

<%end if%>


