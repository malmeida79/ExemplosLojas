<% 
%>
<!-- #include file="topo.asp" -->
<%

'Requisita a variavel para saber a sessao
idsessao = Request.QueryString("item")
categoria = Request.QueryString("Categoria")
if idsessao="" then
idsessao ="000000000000000"
end if
if categoria="" then
categoria ="000000000000000"
end if
pag = request.querystring("itens")
pagdxx = request.querystring("pagina")
finalera = request.querystring("final")
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
' Pega a categoria com a sub-categoria do produto.
'----------------------------------------------------------------------------------------
'Abre os departamentos

'#########################################################################################
' Pega a categoria com a sub-categoria do produto.
'----------------------------------------------------------------------------------------
IF Categoria <> "000000000000000" THEN 
	SQL = "SELECT A.*, B.NOME AS SESSAO FROM CATEGORIA AS A, SESSOES AS B WHERE A.IDSESSAO = B.ID AND idcategoria="&categoria&" and idsessao="&idsessao
ELSE
	SQL = "SELECT top 1 A.*, B.NOME AS SESSAO FROM CATEGORIA AS A, SESSOES AS B WHERE A.IDSESSAO = B.ID AND idsessao="&idsessao	
END IF

Set nomes = abredb.execute(SQL)
if nomes.eof then%>
   			 <table><tr><td align="left" valign="top">
			 				<table border="0" cellspacing="4" cellpadding="4" width=100%><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strLg206%> <br><br><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><br><br><br><center><b><%=strLg207%>&nbsp;<%=nomeloja%>!</b></center><br><br><br><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><center><a HREF="default.asp" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="tahoma,arial,helvetica" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></center></center></td></tr>
							</table></td></tr>
			  </table>
			  <!-- #include file="baixo.asp" -->
<%
response.end
else

strSessao = nomes("Sessao")
strNome = nomes("nome")
strParam = nomes("idcategoria")

'Conta a quantidade de produtos
set rs2 = abredb.execute("SELECT count(nome) AS total FROM produtos WHERE idsessao = '"&strParam&"' AND estoque='s'")
totalreg = rs2("total")
rs2.close
set rs2 = nothing
%>
  		<table width="100%" align="center" valign="">
			   <table border="0" cellspacing="4" cellpadding="4" width="100%" align=center><tr><td> <font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strSessao%> � <%= strNome %><br><br><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></div><font color=#000000 style=font-size:11px;><%=strLg48%> <b><%=totalreg%></b><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></font><%if totalreg = "0" then%><br><br><center><font color=#000000 style=font-size:11px;><br><b> <%=strLg208%></b><br><br><br><br>
			   		  <table width="100%" cellspacing="0" cellpadding="0"><tr><td colspan=2><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%> <B>1</B> <%=strLg46%> <B>1</B></td><td align=right></td></tr><tr><td colspan=2 align=center><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><a HREF="javascript: history.back()" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
					  </table></td></tr>
			  </table></td></tr>
		</table>
		<!--#include file="baixo.asp"-->
<%
response.end
end if
nomes.Close
set nomes = Nothing
end if
%>
  	  	  <table align=center><tr align=center><td valign=top align=center><table align=center border=0 >

<%

if VersaoDb = 1 Then


'*** Inicio da Rotina MySQL

	
	if pag = "" then
	inicial = 0
	final = 4 
	inicialx = 4
	else
	inicial = pag 
	final = 4
	inicialx = 4 + pag
	end if
	
	'Abre os produtos do departamento na primeira coluna
	set rs = abredb.execute("SELECT * FROM produtos WHERE idsessao = '"&strParam&"' AND estoque='s' ORDER by nome LIMIT "&inicial&","&final&"")
	'Abre os produtos do departamento na segunda coluna
	set rs3 = abredb.execute("SELECT * FROM produtos WHERE idsessao = '"&strParam&"' AND estoque='s' ORDER by nome LIMIT "&inicialx&",4")
	if rs.eof or rs.bof then
	%>
	   		  	 <tr><td rowspan=4 valign=top></td></tr>
	<%
	else
	while not rs.EOF
	precito = formatNumber(rs("preco"),2)
	intProdID = rs("idprod")
	impequeno= rs("impeq")
	'Inicia a primeira coluna%>
	
	  <%  If len(cstr(impequeno))<3 then impeq="img_nao_disp.gif" else impeq=impequeno end if%>
	  		  		<tr><td>
							<table border="0" CELLSPACING="4" CELLPADDING="4"><tr><td bgcolor=#ffffff>
						    <table border="0" CELLSPACING="4" CELLPADDING="4"><tr><td bgcolor=#ffffff>
							<form action="comprar.asp" method="post" name=comprar<%=rs("idprod")%>>
							<table align=center width=100% cellspacing="1" cellpadding="1"><tr><td width=80 height=120><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src="produtos/<%=impeq%>"  border="0"></a></td><td valign=center width=120><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%= strlgMoeda & " " & precito%>

						<% '***  Verifica se tem Estoque do Produtos da Esquerda
						set rs_estoque1 = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs.fields("idprod")&" ;")
						if not rs_estoque1.eof then
						estoque_atual_1=rs_estoque1("estoque")
						end if
						rs_estoque1.close
						set rs_estoque1 = nothing
						 %>
						 
<br><b><%=strLg28%></b> <%if estoque_atual_1 > 0 then response.write strLg26 else response.write "<font color=red>" & strLg27 & "</font>" end if%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID %>"><input type="hidden" name="intQuant" value=1><%if estoque_atual_1 > 0 then response.write "<a href=""JavaScript: document.comprar"&rs("idprod")&".submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr>
							</table></td></tr>
						    </table></td></tr>
					        </table></td></tr>
	<%
	rs.MoveNext
	pagn = inicial + 8
	wend
	rs.close
	set rs = nothing
	end if
	%>
			 		</table>
							</td><td valign=top>
					<table>
	<%if rs3.eof or rs3.bof then%>
		 		 			<tr><td rowspan=4 valign=top></td></tr>
	<%else
	while not rs3.EOF
	
	'Formata o preco
	precito2 = formatNumber(rs3("preco"),2)
	'id do produto
	intProdID2 = rs3("idprod")
	impequeno2= rs3("impeq")
	'Inicia a segunda coluna
	%>
	
	 <%  If len(cstr(impequeno2))<3 then impeq="img_nao_disp.gif" else impeq=impequeno2 end if%>
	  		  		<tr><td>
							<table border="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#ffffff>
						    <table border="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
							<form action="comprar.asp" method="post" name=comprar<%=rs3("idprod")%>>
						    <table align=center width=100% cellspacing="1" cellpadding="1">
<tr><td width=80 height=120><a href="produtos.asp?produto=<%=rs3("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src="produtos/<%=impeq%>"  border="0"></a></td><td valign=center width=120><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs3("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%=strlgMoeda & " " & precito2%>

						<% '***  Verifica se tem Estoque do Produtos da Direita
						set rs_estoque2 = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs3.fields("idprod")&" ;")
						if not rs_estoque2.eof then
						estoque_atual_2=rs_estoque2("estoque")
						end if
						rs_estoque2.close
						set rs_estoque2 = nothing
						 %>
						 
						 <br><b><%=strLg28%></b> <%if estoque_atual_2 > 0 then response.write strLg26 else response.write "<font color=red>" & strLg27 & "</font>" end if%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID2 %>"><input type="hidden" name="intQuant" value=1><%if estoque_atual_2 > 0 then response.write "<a href=""JavaScript: document.comprar"&rs3("idprod")&".submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%><a href="produtos.asp?produto=<%=rs3("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></tr></form>
							</table></td></tr>
							
	<%
	rs3.MoveNext
	wend
	rs3.close
	set rs3 = nothing
	end if
	%>
	  	  	  	</table></td></tr>
			</table>
			<table width=100%>
				   <tr><td colspan=2>
	<%
	'Calcula o numero de pagina
	if totalreg Mod 8 <> 0 then
	valor = 1
	else
	valor = 0
	end if 
	pagina = fix(totalreg / 8) + valor
	pagina = replace(pagina,".","")
	if pagina = "0" then pagina = "1" end if
	%>
	   		  			 		<center>
	<%
	paga = pagn - 16
	paga = replace(paga,".","")
	pagn = replace(pagn,".","")
	pagdxx = pagdxx + 1
	pagdxx = replace(pagdxx,".","")
	pagdxx2 = pagdxx - 2
	pagdxx2 = replace(pagdxx2,".","")
	if pagdxx = "" then pagd = 1 end if
	if paga < 0 then
	paga = 0
	else
	%>
	  	 <a HREF="sessoes.asp?itens=<%=paga%>&pagina=<%=pagdxx2%>&item=<%=Server.URLEncode(strParam)%>" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg209%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg209%></b> ::</a></font>&nbsp;&nbsp;
	<%
	end if
	if totalreg < final OR pagdxx = pagina then
	totalreg = ""
	else
	variavel1 = CStr(pagdxx + 1)
	variavel2 = Cstr(pagina)
	variavel1 = replace(variavel1,".","")
	variavel2 = replace(variavel2,".","")
	%>
	  		  &nbsp;&nbsp;<a HREF="sessoes.asp?itens=<%=pagn%>&pagina=<%=pagdxx%>&item=<%= Server.URLEncode(strParam) %><%if variavel1  = variavel2 then response.write "&final=sim" end if%>" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg47%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg47%></b> ::</a></font>
	<%end if%>
		  	  <table width=100% cellspacing="0" cellpadding="0"><tr><td colspan=2>
			  		 			<table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%> <B><%=pagdxx%></B> <%=strLg46%> <B><%=pagina%></B></td><td align=right><font face="<%=fonte%>" style=font-size:11px;color:000000>
	<%
	if pagina = 1 then 
	finalera = "sim"
	end if
	if pagina <> 1 then
	For i=1 to pagina - 1
	i = replace(i,".","")
	i2 = i - 1
	i2 = replace(i2,".","")
	fator = 8
	totalfator = fator * i - 8
	totalfator = replace(totalfator,".","")
	paginadem = pagdxx
	paginadem = replace(paginadem,".","")
	%>
	  		  &nbsp;<a HREF="sessoes.asp?itens=<%=totalfator%>&pagina=<%=i2%>&item=<%=Server.URLEncode(strParam)%>" style=text-decoration:none;color:000000 onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg45%> <%=replace(i,",","")%>';return true;"><font face="<%=fonte%>" style=font-size:11px><%if paginadem = i  then response.write "<b><u>" end if%><%=replace(i,",","")%></u></b></a> &middot;</font>
	<%
	Next
	end if
	pagina2 = pagina * 8 - 8
	pagina2 = replace(pagina2,".","")
	pagina3 = pagina - 1
	pagina3 = replace(pagina3,".","")
	%>
			  &nbsp;<a HREF="sessoes.asp?itens=<%=pagina2%>&pagina=<%=pagina3%>&item=<%=Server.URLEncode(strParam)%>&final=sim" style=text-decoration:none;color:000000 onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg45%> <%=pagina%>';return true;"><font face="<%=fonte%>" style=font-size:11px><%if finalera = "sim"  then response.write "<b><u>" end if%><%=pagina%></u></b></a></td></tr>
			  		   <tr><td colspan=2 align=center><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><a HREF="javascript: history.back()" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
					 </table>
		   	 		   	  </td></tr>
			 		   	</table></td></tr>
				</table></td></tr>
		</table>

<%


'*** Fim da Rotina MySQL


Else


'*** Inicio da Rotina Access


regs = 8 'Aqui setamos quantos registros ser�o listados por p�gina
pag = request.querystring("pagina")

if pag = "" Then
   pag = 1
end if

set rs = createobject("adodb.recordset")
set rs.activeconnection = abredb

rs.cursortype = 3 'Definimos o cursor a ser utilizado
rs.CacheSize = regs
rs.pagesize = regs

sql = "SELECT * FROM produtos WHERE idsessao = '"&strParam&"' AND estoque='s' ORDER by impeq DESC ,nome"
rs.open sql
rs.movefirst
somando = 0
if rs.eof or rs.bof then%>

   		  <tr><td rowspan=4 valign=top></td></tr>
<%else
   rs.absolutepage = pag
   contador = 0
   do while not rs.eof and contador < rs.pagesize
   precito = formatNumber(rs("preco"),2)
   intProdID = rs("idprod")
   impequeno= rs("impeq")

   If len(cstr(impequeno))<3 then impeq="img_nao_disp.gif" else impeq=impequeno end if

if somando = 0 then

   '*** Inicia a primeira coluna
%>
<tr><td valign=top>


<table border="0" CELLSPACING="0" CELLPADDING="0"><tr><td bgcolor=#ffffff>
<table border="0" CELLSPACING="0" CELLPADDING="0"><tr><td bgcolor=#ffffff>
<form action="comprar.asp" method="post" name=comprar<%=rs("idprod")%>>
<table align=center width=100% cellspacing="1" cellpadding="1"><tr><td width=80 height=120><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src="produtos/<%=impeq%>"  border="0"></a></td>
<td valign=center width=150><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%=strLgMoeda & " " & precito%>
						
						<% '***  Verifica se tem Estoque do Produtos da Esquerda
						set rs_estoque1 = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs.fields("idprod")&" ;")
						if not rs_estoque1.eof then
						estoque_atual_1=rs_estoque1("estoque")
						end if
						rs_estoque1.close
						set rs_estoque1 = nothing
						 %>

<br><b><%=strLg28%></b> <%if estoque_atual_1 > 0 then response.write strLg26 else response.write "<font color=red>" & strLg27 & "</font>" end if%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID %>"><input type="hidden" name="intQuant" value=1><%if estoque_atual_1 > 0 then response.write "<a href=""JavaScript: document.comprar"&rs("idprod")&".submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>" end if%><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr></table>
</td></tr></table>
</td></tr></table>

</td>

<% contador = contador + 1
   somando = 1

elseif somando = 1 then

   '*** Inicia a segunda coluna
%>

<td valign=top>

<table border="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#ffffff>
<table border="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#ffffff>
<form action="comprar.asp" method="post" name=comprar<%=rs("idprod")%>>
<table align=center width=100% cellspacing="1" cellpadding="1"><tr><td width=80 height=120><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src="produtos/<%=impeq%>"  border="0"></a></td>
<td valign=center width=150><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%=strLgMoeda & " " & precito%>

						<% '***  Verifica se tem Estoque do Produtos da Direita
						set rs_estoque2 = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs.fields("idprod")&" ;")
						if not rs_estoque2.eof then
						estoque_atual_2=rs_estoque2("estoque")
						end if
						rs_estoque2.close
						set rs_estoque2 = nothing
						 %>
						 
<br><b><%=strLg28%></b> <%if estoque_atual_2 > 0 then response.write strLg26 else response.write "<font color=red>" & strLg27 & "</font>" end if%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID %>"><input type="hidden" name="intQuant" value=1><%if  estoque_atual_2 > 0 then response.write "<a href=""JavaScript: document.comprar"&rs("idprod")&".submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>" end if%><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg30%>';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr></table>
</td></tr></table>
</td></tr></table>

</td></tr>

<% contador = contador + 1
   somando = 0

end if

   rs.movenext
   loop
   end if
%>

  	  	  	</table></td></tr>
		</table>
		<table width=100%>
			   <tr><td colspan=2>

<%'Criando links para a navega��o%>

<table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
<tr><td width="100%"></td><td width="100%" align=right><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%>: &nbsp;


<%for i = 1 to rs.pagecount

if i = cint(pag) then
   response.write " <b>" & i & "</b> &nbsp;"
else
   response.write "<a href='" & request.servervariables("script_name") & "?item="&idsessao&"&categoria="&categoria&"&pagina=" & i & "'><font face="&fonte&" style=font-size:11px;color:000000><b>" & i & "</a> &nbsp;"
end if

next

rs.close
set rs = nothing
%>
</b></font></td></tr>
<tr><td colspan=2 align=center><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><a HREF="javascript: history.back()" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
<%'Fecha os produtos do departamento%>	

<td width="100%"></td></tr></table>
</td></tr></table>
</td></tr></table>
<%end if%>
	<!--#include file="baixo.asp"-->



