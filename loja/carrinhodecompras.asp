<%
'IN�CIO DO C�DIGO
%>
<!-- #include file="topo.asp" -->

<table>
<table border="0" cellspacing="4" cellpadding="4" width="100%"><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b>
	<p align="left"><%=strLg4%></b></a> � <%=strLg1%></p>
<table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=20 bgcolor=f0f0f0>
<font face="<%=fonte%>" style=font-size:11px;color:000000> 
<p align="left">�<font face="<%=fonte%>" style=font-size:10px;color:#808080><font color="<%= Cor_texto_menu_fechamento_compras %>">1. <%=strLg282%></font><br>
</font>
�<font face="<%=fonte%>" style=font-size:10px;color:#808080>2. <%=strLg283%><br>
</font>
�<font face="<%=fonte%>" style=font-size:10px;color:#808080>3. <%=strLg284%><br>
</font>
�<font face="<%=fonte%>" style=font-size:10px;color:#808080>4. <%=strLg285%></td></tr><tr><td height=5></td></tr></table>
<table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table>
<%
'Chama os produtos comprados
set pedidos = abredb.Execute("SELECT idprod, quantidade FROM pedidos WHERE idcompra='" & intOrderID & "'")
if pedidos.eof then
Session("orderID") = "" 
%><div align=center><p><font face=<%=fonte%> style='font-size:15px'><b><%=strLg49%><br> <%=strLg50%><br><br></b><a href=default.asp><img src="<%=dirlingua%>/imagens/continuar.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg68%>';return true;" border=0></a></font></p></div><b><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
</table></td></tr>
</table>
<!-- #include file="baixo.asp" -->
<%
response.end
else
intOrderID = cstr(Session("orderID"))
%>  
<% 
'*** Se no login do cliente (comprareg.asp) foi detectado que o cliente informou CEP diferente do cadastro, for�a o cliente recalcular o frete com o seu CEP correto
if request.querystring("rec")="s" then
session("PesoTotalCep")=request.querystring("cep")
end if
%>
<font face="<%=fonte%>" style=font-size:13px><strong><%=strLg65%></strong></font></p><p><font face="<%=fonte%>" style=font-size:10px;color:000000><%=strLg56%><br><%=strLg57%></font></p><font face="<%=fonte%>" style=font-size:10px;color:red><%=request.querystring("erro")%></font><form action="atualizapedido.asp" method="post" name="registro1">
<table border="0" cellpadding="2" cellspacing="1" width="100%" align=center>
<tr bgcolor="<%=cor3%>"><td width="15%" align="left" nowrap bgcolor="#FF0000"><strong>
	<font color="#FFFFFF" face="<%=fonte%>" style=font-size:11px;><%=strLg58%></font></strong></td>
	<td width="36%" align="left" nowrap bgcolor="#FF0000"><strong>
	<font color="#FFFFFF" face="<%=fonte%>" style=font-size:11px;><%=strLg59%></font></strong></td>
	<td width="22%" align="left" bgcolor="#FF0000"><strong>
	<font color="#FFFFFF" face="<%=fonte%>" style=font-size:11px;><%=strLg60%></font></strong></td>
	<td width="20%" align="right" nowrap bgcolor="#FF0000"><strong>
	<font color="#FFFFFF" face="<%=fonte%>" style=font-size:11px;><%=strLg61%></font></strong></td>
	<td width="10%" align="right" nowrap bgcolor="#FF0000"><strong>
	<font color="#FFFFFF" face="<%=fonte%>" style=font-size:11px;><%=strLg62%></font></strong></td></tr>
<%
while not pedidos.EOF
idprod = pedidos("idprod")
quantidade = pedidos("quantidade")
set produtos = abredb.Execute("SELECT A.preco, A.nome, A.peso, B.QTDMAXIMA, B.ESTOQUE AS B FROM produtos AS A, ESTOQUE AS B WHERE  A.idprod = B.IDPRODUTO AND  A.idprod="&idprod&";")
if produtos.eof then
Session("orderID") = ""
response.redirect "carrinhodecompras.asp"
end if
preco = produtos("preco")
peso = produtos("peso")
nome = produtos("nome")
Maximo = produtos("QTDMAXIMA")
EsT =  produtos("B")
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
totpreco = formatNumber((intQuant * intProdPrice),2)%>
<script LANGUAGE="JScript">
		 	function AbortEntry(sMsg, eSrc){window.alert(sMsg);eSrc.focus();}
			function HandleError<%= intProdID %>(eSrc){var val = parseInt(eSrc.value);
				 if (isNaN(val)){
				 	document.registro1.quant<%= intProdID %>.value = '<%=intQuant%>';
					document.estado.quant<%= intProdID %>.value = '<%=intQuant%>';
				}
				// if (val > <%=Maximo%>){
				//   alert('<%=strLg270%> <%=Maximo%> <%=strLg271%>');
				// 	document.registro1.quant<%= intProdID %>.value = '<%=intQuant%>';
				// }

if  ((val > <%=produtos("B")%>) && (<%=MAXIMO%> > <%=produtos("B")%>))
			{
			/// Avisa qdo o cliente tenta comprar mais do que tem no estoque (e ve se nao passa do Maximo) 
			alert('<%=strLg277&" "&produtos("B")&" "& strLg278%>');
				if(confirm("<%=strLg279&" "&produtos("B")&" "& strLg280%>")==true)
				{		
				document.registro1.quant<%= intProdID %>.value = '<%= produtos("B") %>';
				document.registro1.submit();				
				}
				else
				{
				document.registro1.quant<%= intProdID %>.value = '<%=intQuant%>';
				document.registro1.quant<%= intProdID %>.select;
				}
			}
			else
			{
			/// Avisa q o cliente pode s� comprar um maximo X do produto
				if (val > <%=MAXIMO%>)
				{
				alert('<%=strLg270%> <%=MAXIMO%> <%=strLg271%>');
				document.registro1.quant<%= intProdID %>.value = '<%=MAXIMO%>';
				document.registro1.quant<%= intProdID %>.focus();
				}
			}
				if (val <= 0) {
				   document.registro1.quant<%= intProdID %>.value = '<%=intQuant%>';
				}
			}
		</script>
<SCRIPT LANGUAGE="JavaScript">
function jumpTo(URL_List){
   var URL = URL_List.options[URL_List.selectedIndex].value;
   window.location.href = URL;
}
</SCRIPT>
<tr><td width="15%" valign="center" align="left"><font face="<%=fonte%>" size="2"> <input name="quant<%= intProdID %>" size="2" value="<%=intQuant%>" <% If mostrar_limite_max_prod_compra="Sim" then%>  onChange="HandleError<%= intProdID %>(this)" <% End If %> style=font-size:11px;font-family:<%=fonte%>; maxlength=2></font></td>
	<td width="36%" align="left"><a href="produtos.asp?produto=<%= intProdID %>" ><font face="<%=fonte%>" style=font-size:11px;color:000000><b><%= strProdNome %></b></font></a></td>
	<td width="22%" align="right"><font face="<%=fonte%>" style=font-size:11px;color:000000><%= strLgMoeda & " " & subpreco %></font></td><td width="20%" align="right"><font face="<%=fonte%>" style=font-size:11px;color:000000><%= strLgMoeda & " " & totpreco %></font></td><td width="10%" align="right"><center>
<a onclick="return confirm('Voc� quer remover este produto da sua compra?')" href="atualizapedido.asp?acao=remover&produto=<%= intProdID %>&vvcep=<%=session("PesoTotalCep")%>&qtd=<%= intQuant %>" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg62%>';return true;" ><img src=<%=dirlingua%>/imagens/del.gif border=0></a></td></tr>
<%
pesototal = 1 + FormatNumber(pesototal, 3) + FormatNumber((produtos("peso") * intQuant), 3) - 1
pedidos.MoveNext
wend
pedidos.Close
set pedidos = Nothing
produtos.Close
set produtos = Nothing
Session("PesoTotalFrete") = FormatNumber(pesototal, 3)
'Valor total da compra
totaldasuacompra = formatNumber(intTotal,2)
suacompra = 1 + formatNumber(intTotal,2) + FormatNumber(session("PesoTotalValor"), 2) - 1
suacompra = FormatNumber(suacompra, 2)
%>    
<tr bgcolor=<%=cor3%>><td colspan=5 heigth=5></tr>
<tr><td valign="center" align="left" colspan=3><strong><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg63%></font></strong></td><td colspan="3" align="right" valign="center"><strong><font face="<%=fonte%>" style=font-size:11px;color:000000><b><%= strLgMoeda & " " & totaldasuacompra %></b></font></strong></td></tr>
<tr ><td colspan=5 heigth=5>&nbsp;<br></tr>
<tr bgcolor=<%=cor3%>><td colspan=5 heigth=5></tr>
<% '***Script Visivel/Invisivel %>
<script type=text/javascript>
var isNS = navigator.appName.indexOf("Netscape")  != -1
var isIE = navigator.appName.indexOf("Microsoft") != -1
function show1() {
  document.getElementById("d1").style.visibility = "visible";
}
function hide1() {
  document.getElementById("d1").style.visibility = "hidden";
}
function show2() {
  document.getElementById("d2").style.visibility = "visible";
}
function hide2() {
  document.getElementById("d2").style.visibility = "hidden";
}
function show3() {
  document.getElementById("d3").style.visibility = "visible";
}
function hide3() {
  document.getElementById("d3").style.visibility = "hidden";
}
function show4() {
  document.getElementById("d4").style.visibility = "visible";
}
function hide4() {
  document.getElementById("d4").style.visibility = "hidden";
}
</script>
<tr><td valign="center" align="left" colspan=6><font face="<%=fonte%>" style=font-size:11px;color:000000><strong>Forma de Entrega:<b><br>
<% If entrega_motoboy="Sim" then %>
<input type="radio" value="motoboy" name="opcao_entrega" onClick="document.location='atualizapedido.asp?modo_entrega=motoboy'" <%if session("modo_entrega")="motoboy" then response.write "checked"%> > <%= descricao_forma_entrega_motoboy %> <!-- Motoboy --> &nbsp;&nbsp;&nbsp;
<% End If %>
<br>
<% If entrega_encomenda="Sim" then %>
<input type="radio" value="encomenda" name="opcao_entrega" onClick="document.location='atualizapedido.asp?modo_entrega=encomenda'" <%if session("modo_entrega")="encomenda" then response.write "checked"%> > <%= descricao_forma_entrega_encomenda %> <!-- Encomenda Normal --> &nbsp;&nbsp;&nbsp;
<% End If %>
<br>
<% If entrega_download="Sim" then %>
<input type="radio" value="download" name="opcao_entrega" onClick="document.location='atualizapedido.asp?modo_entrega=download'" <%if session("modo_entrega")="download" then response.write "checked"%> > <%= descricao_forma_entrega_download %> <!-- Download --> &nbsp;&nbsp;&nbsp;
<% End If %>
<br>
<% If entrega_sedex="Sim" then %>
<input type="radio" value="sedex" name="opcao_entrega" onClick="show3();hide1();hide2()" <%if session("modo_entrega")="sedex" then response.write "checked"%>> Sedex&nbsp;&nbsp;
<% End If %>
</b></strong></font></td></tr>
<tr><td valign="center" align="left" colspan=3>
<!-- Area Encomenda Normal -->

<div id="d1" <% if session("modo_entrega")="encomenda" then %>
style="visibility:visible;position:absolute;"
<% else %>
style="visibility:hidden;position:absolute;"
<% End If %>>
<%
estaos = request("frete")
if estaos="acc" then e1 = "selected" end if
if estaos="acx" then e2 = "selected" end if
if estaos="alc" then e3 = "selected" end if
if estaos="alx" then e4 = "selected" end if
if estaos="apc" then e5 = "selected" end if
if estaos="apx" then e6 = "selected" end if
if estaos="amc" then e7 = "selected" end if
if estaos="amx" then e8 = "selected" end if
if estaos="bac" then e9 = "selected" end if
if estaos="bax" then e10 = "selected" end if
if estaos="cec" then e11 = "selected" end if
if estaos="cex" then e12 = "selected" end if
if estaos="dfc" then e13 = "selected" end if
if estaos="dfx" then e14 = "selected" end if
if estaos="esc" then e15 = "selected" end if
if estaos="esx" then e16 = "selected" end if
if estaos="goc" then e17 = "selected" end if
if estaos="gox" then e18 = "selected" end if
if estaos="mac" then e19 = "selected" end if
if estaos="max" then e20 = "selected" end if
if estaos="mtc" then e21 = "selected" end if
if estaos="mtx" then e22 = "selected" end if
if estaos="msc" then e23 = "selected" end if
if estaos="msx" then e24 = "selected" end if
if estaos="mgc" then e25 = "selected" end if
if estaos="mgx" then e26 = "selected" end if
if estaos="pac" then e27 = "selected" end if
if estaos="pax" then e28 = "selected" end if
if estaos="pbc" then e29 = "selected" end if
if estaos="pbx" then e30 = "selected" end if
if estaos="prc" then e31 = "selected" end if
if estaos="prx" then e32 = "selected" end if
if estaos="pec" then e33 = "selected" end if
if estaos="pex" then e34 = "selected" end if
if estaos="pic" then e35 = "selected" end if
if estaos="pix" then e36 = "selected" end if
if estaos="rjc" then e37 = "selected" end if
if estaos="rjx" then e38 = "selected" end if
if estaos="rnc" then e39 = "selected" end if
if estaos="rnx" then e40 = "selected" end if
if estaos="rsc" then e41 = "selected" end if
if estaos="rsx" then e42 = "selected" end if
if estaos="roc" then e43 = "selected" end if
if estaos="rox" then e44 = "selected" end if
if estaos="rrc" then e45 = "selected" end if
if estaos="rrx" then e46 = "selected" end if
if estaos="scc" then e47 = "selected" end if
if estaos="scx" then e48 = "selected" end if
if estaos="spc" then e49 = "selected" end if
if estaos="spx" then e50 = "selected" end if
if estaos="sec" then e51 = "selected" end if
if estaos="sex" then e52 = "selected" end if
if estaos="toc" then e53 = "selected" end if
if estaos="tox" then e54 = "selected" end if %>
<font face="<%=fonte%>" style=font-size:11px;color:000000><strong>Valor da entrega em </strong></font><form name="estado"><select name="estado" style=font-size:10px;font-family:<%=fonte%>; onChange="jumpTo(this);"><option value=""><%=strLg185%></option> <option value="">--------------------------------------</option>
<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ac%>acc" <%=e1%>>Acre - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ac%>acx" <%=e2%>>Acre - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_al%>alc" <%=e3%>>Alagoas - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_al%>alx" <%=e4%>>Alagoas - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ap%>apc" <%=e5%>>Amap� - Capital</option>	<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ap%>apx" <%=e6%>>Amap� - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_am%>amc" <%=e7%>>Amazonas - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_am%>amx" <%=e8%>>Amazonas - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ba%>bac" <%=e9%>>Bahia - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ba%>bax" <%=e10%>>Bahia - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ce%>cec" <%=e11%>>Cear� - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ce%>cex" <%=e12%>>Cear� - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_df%>dfc" <%=e13%>>Distrito Federal - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_df%>dfx" <%=e14%>>Distrito Federal - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_es%>esc" <%=e15%>>Espirito Santo - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_es%>esx" <%=e16%>>Espirito Santo - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_go%>goc" <%=e17%>>Goi�s - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_go%>gox" <%=e18%>>G�i�s - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ma%>mac" <%=e19%>>Maranh�o - Capital</option>
<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ma%>max" <%=e20%>>Maranh�o - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_mt%>mtc" <%=e21%>>Mato Grosso - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_mt%>mtx" <%=e22%>>Mato Grosso - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ms%>msc" <%=e23%>>Mato Grosso do Sul - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ms%>msx" <%=e24%>>Mato Grosso do Sul - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_mg%>mgc" <%=e25%>>Minas Gerais - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_mg%>mgx" <%=e26%>>Minas Gerais - Interior</option>	<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pa%>pac" <%=e27%>>Par� - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pa%>pax" <%=e28%>>Par� - Interior</option>	<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pb%>pbc" <%=e29%>>Paraiba - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pb%>pbx" <%=e30%>>Paraiba - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pr%>prc" <%=e31%>>Parana - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pr%>prx" <%=e32%>>Parana - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pe%>pec" <%=e33%>>Pernambuco - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pe%>pex" <%=e34%>>Pernambuco - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pi%>pic" <%=e35%>>Piau� - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_pi%>pix" <%=e36%>>Piau� - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rj%>rjc" <%=e37%>>Rio de Janeiro - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rj%>rjx" <%=e38%>>Rio de Janeiro - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rn%>rnc" <%=e39%>>Rio Grande do Norte - Capital</option>
<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rn%>rnx" <%=e40%>>Rio Grande do Norte - Interior</option>	<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rs%>rsc" <%=e41%>>Rio Grande do Sul - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rs%>rsx" <%=e42%>>Rio Grande do Sul - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ro%>roc" <%=e43%>>Rond�nia - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_ro%>rox" <%=e44%>>Rond�nia - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rr%>rrc" <%=e45%>>Roraima - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_rr%>rrx" <%=e46%>>Roraima - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_sc%>scc" <%=e47%>>Santa Catarina - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_sc%>scx" <%=e48%>>Santa Catarina - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_sp%>spc" <%=e49%>>S�o Paulo - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_sp%>spx" <%=e50%>>S�o Paulo - Interior</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_se%>sec" <%=e51%>>Sergipe - Capital</option>	<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_se%>sex" <%=e52%>>Sergipe - Interior</option>	<option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_to%>toc" <%=e53%>>Tocantins - Capital</option><option value="atualizapedido.asp?modo_entrega=encomenda&estado=<%=estado_to%>tox" <%=e54%>>Tocantins - Interior</option></select>
</div>
<!-- Area Motoboy -->
<div id="d4" <% if session("modo_entrega")="motoboy" then %>
style="visibility:visible;position:absolute;"
<% else %>
style="visibility:hidden;position:absolute;"
<% End If %>>
<font face="<%=fonte%>" style=font-size:11px;color:000000><strong><font color="#0000FF"><%= descricao_regiao_entrega_motoboy %></font></strong></font>
</div>
<!-- Area Download -->
<div id="d2" <% if session("modo_entrega")="download" then %>
style="visibility:visible;position:absolute;"
<% else %>
style="visibility:hidden;position:absolute;"
<% End If %>>
<font face="<%=fonte%>" style=font-size:11px;color:000000><strong><font color="#0000FF"><%= descricao_regiao_entrega_download %></font></strong></font>
</div>
<!-- Area Sedex -->
<div id="d3" <% if session("modo_entrega")="sedex" then %>
style="visibility:visible;position:relative;"
<% else %>
style="visibility:hidden;position:relative;"
<% End If %>>
<font face="<%=fonte%>" style=font-size:11px;color:000000><strong>Valor da entrega no CEP: 
<input type=text name=vvcep size=7 maxlength=8 style="font-size:11px;color:red;font-weight:bold;font-family:<%=fonte%>" value=<%=session("PesoTotalCep")%>>
<input type=hidden name=modo_entrega value=sedex>
</strong></strong></strong>&nbsp;&nbsp;&nbsp;
<br>
<!-- Link:  <a href=http://www.correios.com.br/servicos/cep/cep_default.cfm target=_novocep> -->
<a href="busca_cep.asp" target="novocep" onclick="window.open('','novocep','width=400,height=400,resizable=yes,scrollbars=yes')"<u>N�o sabe seu CEP?</u></a><br>
</strong></strong></strong>(Clique no bot�o "Atualizar" para calcular o valor da Entrega)<br>&nbsp;</div>
</td><td colspan="3" align="right" valign="top"><strong><font face="<%=fonte%>" style=font-size:11px;color:000000><b><%= strLgMoeda & " " & FormatNumber(session("PesoTotalValor"), 2)%></b></font></strong>
</td></tr>
<tr bgcolor=<%=cor3%>><td colspan=5 heigth=5></td></tr>
<tr><td valign="center" align="left" colspan=3><strong><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg65%></font></strong></td><td colspan=2 align="right" valign="center"><strong><font face="<%=fonte%>" style=font-size:11px;color:000000><b><%= strLgMoeda & " " & suacompra %></b></font></strong></td><% session("totalcompra1") = intTotal %></tr>
<tr><td colspan=5><br></td></tr><tr><td colspan=5 align=center><table><tr><td valign=top><center><input type="image" name="control" name="Atualizar" src=<%=dirlingua%>/imagens/atualiza.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg66%>';return true;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></form><td valign=top><%if request("atualiza") = "ok" then%><form action="<%=link%>" method="post" id="form1a" name="form1a"><input type="image" name="control" value="fpedido" src=<%=dirlingua%>/imagens/fecha.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg67%>';return true;"><%end if%></td></tr>
</table></td></tr>
</table><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><br><center></b><a href=default.asp><img src="<%=dirlingua%>/imagens/continuar_compras.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg68%>';return true;" border=0></a></center>
<br><br>
<font size="1" color="#ffffff"><%= "Session(PesoTotalFrete): "&Session("PesoTotalFrete") %></font>
<% end if %></td></tr></table></td></tr></table><!-- #include file="baixo.asp" -->