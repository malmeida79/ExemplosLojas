<!-- #include file="barrabotoes.asp" -->

<!--PRINCIPAL -->
<table border="0" bgcolor="#FFFFFF" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="161" valign="top" align="center">
			 
<table border="0" cellspacing="4" cellpadding="4" width="161">
<tr>
    <!-- #include file="topo_i.asp" -->
  
<td>
<table cellpadding="0" cellspacing="0" width="185" height="36">
<!-- MSTableType="layout" -->
<tr>
<td height="26" bgcolor="#ffffff" width="182">
<a href="carrinhodecompras.asp">

<img src="topo_bot_carrinho.gif" width="182" height="36" border="0" /></a></td>
</tr>
</table>
<TABLE width="21%" border=0 align="center" cellPadding=0 cellSpacing=0 bgColor=#000000>
<TR>
<TD align=left width="100%">
<table cellspacing="0" cellpadding="0" width="182" border="0" style="BACKGROUND: #ffffff">
<tr>
<td width="116"  nowrap="nowrap" colspan="2" ><font color="#000000">
Quantidade Itens:</font></td>
<td width="78" nowrap="nowrap"><div align="center"><span class="style3">
<font color="#000000"><%= intQuantx %></font></span></div></td>
</tr>
<tr>
<td align="left" nowrap="nowrap" width="76"><font color="#000000"><%= strLg3 %> </font></td>
<td align="left" nowrap="nowrap" width="20">
<p align="right"><font color="#000000"><%= strLgMoeda %></font></td>
<td align="left" nowrap="nowrap"><div align="center"><font color="#000000"><%= precito1 %></font></div></td>
</tr>
</table>
</TD>
</TR></TABLE>
<p align="left">
<!--Busca de produtos -->
<table width="184" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="184" height="23" background="linguagens/portuguesbr/imagens/menu_esq_top.gif">&nbsp;&nbsp;<font color="#FFFFFF"><strong> 
BUSCA DE PRODUTOS</strong></font> </td>
</tr>
<tr>
<td background="linguagens/portuguesbr/imagens/menu_esq_fundo.gif"> <BR>
<!--- start SEARCH --->
<TABLE width="90%" border=0 align="center" cellPadding=2 cellSpacing=0>
<FORM style="DISPLAY: inline" name=newsearch action=pesquisa.asp method=get>
<TR>
<td colspan=2> <div align="left">
<input type=text name=pesq value="<%= request.querystring("pesq") %>" size=23 style=font-size:11px;font-family:<%=fonte%>;>
</div></td></tr><tr><td>
<div align="right">
<select name=cat style=font-size:11px;font-family:<%=fonte%>>
<option value=todos><%=strLg15%> </option>
<option value=xxx>------------------</option>
<%Dim scat
'--------------------------------------------------------------------------------------------------
'Monta a select de pesquisa
'--------------------------------------------------------------------------------------------------
Set cat = abredb.Execute("SELECT * FROM sessoes WHERE ver = 's' ORDER by nome;")
While Not cat.EOF 
'Response.Write "<option value="& cat("id") &" style=color:#808080;>"&CHR(187)&"&nbsp;"&cat("nome") &"</option>"&VBCRlf 
Response.Write "<option value="""" style=color:#808080;>"&CHR(187)&"&nbsp;"&cat("nome") &"</option>"&VBCRlf 
		SQL = "SELECT C.idcategoria, C.nome FROM categoria as c ,sessoes as s  WHERE s.id = c.idsessao and c.idsessao = "&cat("id")&" AND C.ver = 's' ORDER by c.nome"
		Set scat = abredb.Execute(SQL)
			While Not scat.EOF
				Response.Write "<option value="&scat("idcategoria")&">&nbsp;&nbsp; " &CHR(149)&"&nbsp;" & scat("nome")&"</option>"&VBCRlf 
				scat.MoveNext
			Wend
cat.MoveNext
Wend
Response.Write "</select></div></font>"%>
									 
</TD>
<TD>
<div align="left">
	  
<INPUT class=searchforms type=image alt="Pesquisar" src="<%= dirlingua %>/imagens/botao_pesquisar.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg21%>';return true;" align="absmiddle"></div> </TD></TR></TBODY></FORM></TABLE>
<!--- end SEARCH --->
	
</td>
</tr>
<tr> 
<td><img src="linguagens/portuguesbr/imagens/menu_esq_fim.gif" width="184" height="6"></td>
</tr>
</table>
<!--Fim da busca de produtos-->
<IMG height=10 
            src="<%= dirlingua %>/imagens/spacer.gif" width=1>
<!-- Menu de seções e subseções -->
<table width="184" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="184" height="23" background="linguagens/portuguesbr/imagens/menu_esq_top.gif">&nbsp;&nbsp;<font color="#FFFFFF"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
DEPARTAMENTOS</strong></font></td>
</tr>
<tr>
<td background="linguagens/portuguesbr/imagens/menu_esq_fundo.gif">
<!-- Listando o menu e submenu -->
<TABLE width="96%" border=0 align="center" cellPadding=2 cellSpacing=0>
<TBODY>
<TR>
<TD>&nbsp; </TD>
</TR>
<%
Mostrar = Request.QueryString("item")
Dim Smenu
'###    MENU PRINCIPAL (Tabela Sessoes)
Set menu = abredb.Execute("SELECT * FROM sessoes WHERE ver = 's' ORDER by nome;")
DO While Not menu.EOF

'Verifica se existe alguma categoria sem sub-categoria, ou seja, se o produto estiver cadastrado em uma SubCategoria chamado 'TodoAs' , a Categoria será 'linkado' diretamente para os seus produtos

	SQL = "SELECT nome FROM categoria WHERE nome = 'Todos' AND idsessao = "&menu("id")&""
	Set sem_categ = abredb.Execute(SQL)
	if not sem_categ.EOF then %>
<TR>
<TD onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=menu("nome")%>';return true;">
<a href="sessoes.asp?item=<%=menu("id")%>&Categoria=" style="text-decoration:none;" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=menu("nome")%>';return true;">
&nbsp;<strong> <%=menu("nome")%></strong></a><BR><br>
<center> <img src="linguagens/portuguesbr/imagens/linhacinza.gif" width="90%" height="1"></center>
</TD>
</TR>
<%	 else %>
<TR>
<TD onMouseOut="window.status='';return true;" onMouseOver="window.status='<%= menu("nome") %>';return true;">
<strong>&nbsp;<%=menu("nome")%><br>
</strong>
</div>
<%IF cInt(Mostrar) = menu("id") then
                     Response.Write "<span class=""submenu"" id="""&menu("nome")&""" style=""display:block"">"
                     END IF
'### Mostra as Categorias das Sessoes da Loja

SQL = "SELECT C.idcategoria, C.nome FROM categoria as c ,sessoes as s  WHERE s.id = c.idsessao and c.idsessao = "&menu("id")&" AND C.ver = 's' AND C.nome <> 'Todos' ORDER by c.nome"
Set Smenu = abredb.Execute(SQL)
While Not Smenu.EOF%><a href="sessoes.asp?item=<%=menu("id")%>&Categoria=<%=Smenu("idcategoria")%>" style="text-decoration:none;"onMouseOut="window.status='';return true;" 
onMouseOver="window.status='<%= menu("nome")&" > "&Smenu("nome") %>';return true;"><img src='<%=dirlingua%>/imagens/menu_esq_seta.gif' border=0>&nbsp;<%=Smenu("nome")%></a><BR>
<%Smenu.MoveNext
Wend %>
<center>
<img src="linguagens/portuguesbr/imagens/linhacinza.gif" width="90%" height="1"> 
</center>		  
</span>
</TD>
</TR>
<%	end if	
	sem_categ.close
	set sem_categ=Nothing%>				

<%
menu.MoveNext
Loop
'Fecha o menu
menu.Close
Set menu = Nothing
'Response.Write "</table>"			%>
</TBODY>
</TABLE><BR>
</td>
</tr>
<tr> 
<td><img src="linguagens/portuguesbr/imagens/menu_esq_fim.gif" width="184" height="6"></td>
</tr>
</table>
<!--Fim de  Listando o menu e submenu -->
<IMG height=10 src="<%= dirlingua %>/imagens/spacer.gif" width=1>
</div>						
<br>
</tr>
<tr>
<td>
</td>
</tr>
</table>
<td width="1" ><img src="<%=dirlingua%>/imagens/spacer.gif" width="1" height="10" border="0">
<!-- <img src="<%=dirlingua%>/imagens/dot_gray_1x1.gif" width="1" height="100%" border="0"> --></td>
<td align="left" valign="top">
<!-- TABELA PRINCIPAL -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" valign=top>
<tr>
<td valign=top>	
<p align="center">	
<%
'-----------------------------------------------------------------------------------------------------
Sub BaixoC()
response.write "<a class=menu  href=""http://www.studio7seven.com.br"" target=_new>by STUDIOSEVEN</a>"
End Sub

Sub BaixoComunidade()
response.write application("link_comunidade") 
End Sub

Function formatazeros(dado, numero)
	if len(dado) > numero then
	dado = right(dado, numero)
	end if
	'--------------------------------------------------------------------------------------------------
	for i = 1 to len(dado)
		varn = (numero - 1) - i
		numezero = "0"
			if varn <> 0 then
				for i2 = 1 to varn
					numezero = numezero & "0"
				next
			end if
	next
formatazeros = right(numezero & dado, numero)
End Function
%>