<!-- #include file="topo.asp" -->
                        <%
function preparaPalavra(str)
  str = replace(str, "�" , "o")
  str = replace(str, "�" , "o")
  str = replace(str, "�" , "o")
  str = replace(str, "�" , "o")
  str = replace(str, "�" , "o")
  str = replace(str, "�" , "a")
  str = replace(str, "�" , "a")
  str = replace(str, "�" , "a")
  str = replace(str, "�" , "a")
  str = replace(str, "�" , "a")
  str = replace(str, "�" , "e")
  str = replace(str, "�" , "e")
  str = replace(str, "�" , "e")
  str = replace(str, "�" , "u")
  str = replace(str, "�" , "u")
  str = replace(str, "�" , "u")
  str = replace(str, "�" , "u")
  str = replace(str, "�" , "i")
  str = replace(str, "�" , "i")
  str = replace(str, "�" , "c")
  preparaPalavra = replace(LCASE(str),"a","[a,�,�,�,�,�]")
  preparaPalavra = replace(preparaPalavra,"e","[e,�,�,�]")
  preparaPalavra = replace(preparaPalavra,"i","[i,�,�]")
  preparaPalavra = replace(preparaPalavra,"o","[o,�,�,�,�,�]")
  preparaPalavra = replace(preparaPalavra,"u","[u,�,�,�,�]")
  preparaPalavra = replace(preparaPalavra,"c","[c,�]")
  preparaPalavra = replace(preparaPalavra,"'","['']")
  preparaPalavra = preparaPalavra
end function


'Requisita as variaveis
finalera = request.querystring("final")
pag = request.querystring("itens")
pesss = trim(request.querystring("pesq"))
palavra = Split(Trim(Request.QueryString("pesq")), " ")
tipo = request.querystring("cat")
xx = request.querystring("cat")
pagdxx = request.querystring("pagina")


If pesss = "" then%>
   		 <table width="100%"><tr><td align="left" valign="top" width="100%">
           <table border="0" cellspacing="4" cellpadding="4" width="100%"><tr>
             <td width="100%"><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strLg42%> <b> - </b><br><br>
             <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
               <td height=5 width="100%"></td></tr><tr>
                 <td height=1 bgcolor=<%=cor3%> width="100%"></td></tr><tr>
                 <td height=5 width="100%"></td></tr></table><%=strLg43%> <b>0</b> <%=strLg44%> 
             <table border=0 cellspacing=0 width="100%" cellpadding=0><tr>
               <td height=5 width="100%"></td></tr><tr>
                 <td height=1 bgcolor=<%=cor3%> width="100%"></td></tr><tr>
                 <td height=5 width="100%"></td></tr>
</table><br><br><br><center><b>
		 <%=strLg228%></b><br><br><br>
		 <%=strLg315%><br><br>
		 <%=strLg316%> <a href="ajuda_email.asp?duvida=2"><u>clique aqui</u></a> <%=strLg317%>
		 </center><br><br><br>
		 				<table width="100%" cellspacing="0" cellpadding="0"><tr>
                          <td colspan=2 width="100%">
                          <table border=0 cellspacing=0 width="100%" cellpadding=0><tr>
                            <td height=5 width="100%"></td></tr><tr>
                              <td height=1 bgcolor=<%=cor3%> width="100%"></td></tr><tr>
                              <td height=5 width="335"></td></tr></table></td></tr><tr>
                            <td width="100%"><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%> <B>1</B> <%=strLg46%> <B>1</B></td>
                            <td align=right width="1"></td></tr>
						<tr><td colspan=2 align=center width="100%"><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><a HREF="default.asp" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
						</table></td></tr>
		</table></td></tr>
	</table>
	<!--#include file="baixo.asp"-->
<%
response.end
end if


select case tipo
Case "todos"
tipos = " "
Case else
tipos = "AND idsessao='"&xx&"'"
end select


if VersaoDb = 1 then

'*** Pesquisa em MySQL


'Chama os requisitos e monta a SQL para pesquisa

if pag = "" then
inicial = 0
final = 5
else
inicial = pag
final = 5
end if

'Pesquisa no banco de dados %" & preparaPalavra(palavra(0)) & "%
set rs = abredb.Execute("SELECT * FROM produtos WHERE ( nome LIKE '%"&preparaPalavra(palavra(0))&"%' OR detalhe LIKE '%"&preparaPalavra(palavra(0))&"%' OR fabricante LIKE '%"&preparaPalavra(palavra(0))&"%') "&tipo&" ORDER by nome LIMIT " & inicial & "," & final & "")

if rs.eof or rs.bof then%>
   		  	 <table width="100%"><tr><td align="left" valign="top" width="381">
		<table border="0" cellspacing="4" cellpadding="4" width="100%"><tr>
          <td width="100%"><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strLg42%> <b> <%= pesss %> </b><br><br>
          <table border=0 cellspacing=0 width="100%" cellpadding=0><tr>
            <td height=5 width="100%"></td></tr><tr>
              <td height=1 bgcolor=<%=cor3%> width="100%"></td></tr><tr>
              <td height=5 width="100%"></td></tr></table><%=strLg43%> <b>0</b> <%=strLg44%> 
          <table border=0 cellspacing=0 width="100%" cellpadding=0><tr>
            <td height=5 width="100%"></td></tr><tr>
              <td height=1 bgcolor=<%=cor3%> width="100%"></td></tr><tr>
              <td height=5 width="100%"></td></tr>
		</table><br><br><br><center><b>
		 <%=strLg228%></b><br><br><br>
		 <%=strLg315%><br><br>
		 <%=strLg316%> <a href="ajuda_email.asp?duvida=2"><u>clique aqui</u></a> <%=strLg317%>
		 </center><br><br><br>
								   <table width=301 cellspacing="0" cellpadding="0"><tr>
                                     <td colspan=2 width="471">
                                     <table border=0 cellspacing=0 width=363 cellpadding=0><tr>
                                       <td height=5 width="363"></td></tr><tr>
                                         <td height=1 bgcolor=<%=cor3%> width="363"></td></tr><tr>
                                         <td height=5 width="363"></td></tr></table></td></tr><tr>
                                       <td width="414"><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%> <B>1</B> <%=strLg46%> <B>1</B></td>
                                       <td align=right width="58"></td></tr>
								   		  <tr>
                                            <td colspan=2 align=center width="471">
                                            <table border=0 cellspacing=0 width=389 cellpadding=0><tr>
                                              <td height=5 width="389"></td></tr><tr>
                                                <td height=1 bgcolor=<%=cor3%> width="389"></td></tr><tr>
                                                <td height=5 width="389"></td></tr></table><a HREF="default.asp" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
								   </table></td></tr>
							</table></td></tr>
	          </table>
			  <!--#include file="baixo.asp"-->
<%response.end
else
    'Conta o numero de registros encontrados
    set rs2 = abredb.Execute("SELECT count(nome) AS total FROM produtos WHERE ( nome LIKE '%"&preparaPalavra(palavra(0))&"%' OR detalhe LIKE '%"&preparaPalavra(palavra(0))&"%' OR fabricante LIKE '%"&preparaPalavra(palavra(0))&"%') "&tipo&"")
    totalreg = rs2("total")
    rs2.close
    set rs2 = nothing
    %>
      		  		 <table width="100%"><tr>
                       <td align="left" valign="top" width="100%">
    				 				<table border="0" cellspacing="4" cellpadding="4" width=100%><tr>
                                      <td width="100%"><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strLg42%> <b><%= pesss %></b><br><br>
                                      <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                                        <td height=5 width="374"></td></tr><tr>
                                          <td height=1 bgcolor=<%=cor3%> width="374"></td></tr><tr>
                                          <td height=5 width="374"></td></tr></table><%=strLg43%> <b><%=totalreg%></b> <%=strLg44%> 
                                      <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                                        <td height=5 width="100%"></td></tr><tr>
                                          <td height=1 bgcolor=<%=cor3%> width="100%"></td></tr><tr>
                                          <td height=5 width="100%"></td></tr></table><br><br><center>
    <%
    while not rs.EOF
    
    'Fomata o pre�o
    precito = formatNumber(rs("preco"),2)%>

						<% '***  Verifica se tem Estoque do Produto
						set rs_estoque = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs("idprod")&" ;")
						if not rs_estoque.eof then
						estoque_atual=rs_estoque("estoque")
						end if
						rs_estoque.close
						set rs_estoque = nothing
						 %>

    		  				     <table border="0" CELLSPACING="1" CELLPADDING="0" width="100%"><tr>
                                   <td bgcolor=#ffffff width="100%">
    							 		<table border="0" CELLSPACING="1" CELLPADDING="3" width="100%"><tr>
                                          
                      <td bgcolor=#ffffff width="100%"> <table align=center width=100% cellspacing="1" cellpadding="1">
                          <tr> 
                            <td valign=center width=100%><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs("nome")%></font></b><font color=000000><br>
                              <br>
                              <b>Pre�o:</b><%=strlgMoeda &" "& precito%><br>
                              <b>Estoque:</b> 
                              <% '*** Opcao de NAO Mostrar Estoque
    if estoque_atual > 0 then response.write strLg26  else response.write "<font color=red>" & strLg27 & "</font>" end if
%>
                              </font></font></td>
                            <td valign=center width="184"><div align=right><br>
                                <a href="produtos.asp?produto=<%=rs("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td>
                          </tr>
                        </table></table>
    						     </table>
    <%
    rs.MoveNext
    
    'Monta o numero de paginas
    pagn = inicial + 5
    wend
    paga = pagn - 10
    
    'Calcula  o numero de paginas
    if totalreg Mod 5 <> 0 then
    valor = 1
    else
    valor = 0
    end if 
    pagina = fix(totalreg/5) + valor
    pagina = replace(pagina,".","")
    pagdxx = pagdxx + 1
    pagdxx = replace(pagdxx,".","")
    pagdxx2 = pagdxx - 2
    pagdxx2 = replace(pagdxx2,",","")
    paga = replace(paga,".","")
    pagn = replace(pagn,".","")
    if pagdxx = "" then pagdxx = 1 end if
    if pagina = "0" then pagina = "1" end if%>
       		  	<table width=100%>
    				   <tr><td colspan=2 width="100%"><center><br>
  <%
    'Monta os links para navega��o
    if paga < 0 then
    paga = 0
    else%>
    	   	  			 		  <a HREF="pesquisa.asp?itens=<%=paga%>&pagina=<%=pagdxx2%>&pesq=<%=pesss%>&cat=<%=tipo%>" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg209%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg209%></b> ::</a></font>&nbsp;&nbsp;
  <%
    end if
    if totalreg < final OR totalreg = 5 OR pagdxx = pagina then
    totalreg = ""
    npc = 1
    totalp = 1
    else
    variavel1 = CStr(pagdxx + 1)
    variavel2 = Cstr(pagina)
    variavel1 = replace(variavel1,".","")
    variavel2 = replace(variavel2,".","")%>
    		  					&nbsp;&nbsp;<a HREF="pesquisa.asp?itens=<%=pagn%>&pagina=<%=pagdxx%>&pesq=<%=pesss%>&cat=<%=tipo%><%if variavel1  = variavel2 then response.write "&final=sim" end if%>" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg47%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg47%></b> ::</a></font>
<%end if%>
    	  	  					<table width=100% cellspacing="0" cellpadding="0">
    								   <tr><td colspan=2 width="100%">
                                         <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                                           <td height=5 width="100%"></td></tr><tr>
                                             <td height=1 bgcolor=<%=cor3%> width="382"></td></tr><tr>
                                             <td height=5 width="382"></td></tr></table></td></tr>
    								   <tr><td width="22"><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%> <B><%=pagdxx%></B> <%=strLg46%> <B><%=pagina%></B></td>
                                         <td align=right width="382"><font face="<%=fonte%>" style=font-size:11px;color:000000>
    <%
    if pagina = 1 then 
    finalera = "sim"
    end if
    
    'Monta o menu de navega��o interior
    if pagina <> 1 then
    For i=1 to pagina - 1
    i = replace(i,".","")
    i2 = i - 1
    i2 = replace(i2,".","")
    fator = 5
    totalfator = fator * i  - 5
    totalfator = replace(totalfator,".","")
    paginadem = pagdxx
    paginadem = replace(paginadem,".","")%>
    		  							   &nbsp;<a HREF="pesquisa.asp?itens=<%=totalfator%>&pagina=<%=i2%>&pesq=<%=pesss%>&cat=<%=tipo%>" style=text-decoration:none;color:000000 onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg45%> <%=replace(i,",","")%>';return true;"><font face="<%=fonte%>" style=font-size:11px><%if paginadem = i then response.write "<b><u>" end if%><%=replace(i,",","")%></u></b></a> &middot;</font>
    <%
    Next
    end if
    pagina2 = pagina * 5 - 5
    pagina2 = replace(pagina2,".","")
    pagina3 = pagina - 1
    pagina3 = replace(pagina3,".","")
    %>
    		  						      &nbsp;<a HREF="pesquisa.asp?itens=<%=pagina2%>&pagina=<%=pagina3%>&pesq=<%=pesss%>&cat=<%=tipo%>&final=sim" style=text-decoration:none;color:000000 onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg45%> <%=pagina%>';return true;"><font face="<%=fonte%>" style=font-size:11px><%if finalera = "sim"  then response.write "<b><u>" end if%><%=pagina%></u></b></a></td></tr>
    									  <tr>
                                            <td colspan=2 align=center width="404">
                                            <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                                              <td height=5 width="365"></td></tr><tr>
                                                <td height=1 bgcolor=<%=cor3%> width="365"></td></tr><tr>
                                                <td height=5 width="365"></td></tr></table><a HREF="javascript: history.back()" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
    								  </table>
    <%
    'Fecha a tabela
    rs.close
    set rs = nothing
    %>
      	   	 					</td></tr>
    						</table></td></tr>
    				</table></td></tr>
    		</table>
  <%end if%>
  
<%else

'*** Pesquisa em Access / SQL Server

    regs = mostrar_produtos_por_pagina_na_secao
    pag = request.querystring("pagina")
    
    if pag = "" Then
       pag = 1
    end if
    
    set rs = createobject("adodb.recordset")
    set rs.activeconnection = abredb
    
    rs.cursortype = 3
    rs.pagesize = regs
    
    'Pesquisa no banco de dados
    sql = "SELECT * FROM produtos WHERE ( nome LIKE '%"&preparaPalavra(palavra(0))&"%' OR detalhe LIKE '%"&preparaPalavra(palavra(0))&"%' OR fabricante LIKE '%"&preparaPalavra(palavra(0))&"%') "&tipos&" ORDER by nome"
    rs.open sql
    if rs.eof or rs.bof then%>
       		  	 <table width="100%"><tr><td align="left" valign="top">
    			<table border="0" cellspacing="4" cellpadding="4" width=100%><tr>
                  <td width="100%"><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strLg42%> <b> <%= pesss %> </b><br><br>
                  <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                    <td height=5 width="382"></td></tr><tr>
                      <td height=1 bgcolor=<%=cor3%> width="382"></td></tr><tr>
                      <td height=5 width="382"></td></tr></table><%=strLg43%> <b>0</b> <%=strLg44%> 
                  <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                    <td height=5 width="369"></td></tr><tr>
                      <td height=1 bgcolor=<%=cor3%> width="369"></td></tr><tr>
                      <td height=5 width="369"></td></tr>
				</table><br><br><br><center><b>
		 <%=strLg228%></b><br><br><br>
		 <%=strLg315%><br><br>
		 <%=strLg316%> <a href="ajuda_email.asp?duvida=2"><u>clique aqui</u></a> <%=strLg317%>
		 </center><br><br><br>
    								   <table width=100% cellspacing="0" cellpadding="0"><tr>
                                         <td colspan=2 width="415">
                                         <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                                           <td height=5 width="364"></td></tr><tr>
                                             <td height=1 bgcolor=<%=cor3%> width="364"></td></tr><tr>
                                             <td height=5 width="364"></td></tr></table></td></tr><tr>
                                           <td width="402"><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%> <B>1</B> <%=strLg46%> <B>1</B></td>
                                           <td align=right width="14"></td></tr>
    								   		  <tr>
                                                <td colspan=2 align=center width="415">
                                                <table border=0 cellspacing=0 width=100% cellpadding=0><tr>
                                                  <td height=5 width="371"></td></tr><tr>
                                                    <td height=1 bgcolor=<%=cor3%> width="371"></td></tr><tr>
                                                    <td height=5 width="371"></td></tr></table><a HREF="default.asp" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
    								   </table></td></tr>
    							</table></td></tr>
    	          </table>
    <%response.end
    else
    
    'Conta o numero de registros encontrados
    set rs2 = abredb.Execute("SELECT count(nome) AS total FROM produtos WHERE ( nome LIKE '%"&preparaPalavra(palavra(0))&"%' OR detalhe LIKE '%"&preparaPalavra(palavra(0))&"%' OR fabricante LIKE '%"&preparaPalavra(palavra(0))&"%') "&tipos&"")
    totalreg = rs2("total")
    rs2.close
    set rs2 = nothing%>
      		  		 <table width="100%"><tr>
                       <td align="left" valign="top" width="100%">
    				 				<table border="0" cellspacing="4" cellpadding="4" width=365><tr>
                                      <td width="437"><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strLg42%> <b><%= pesss %></b><br><br>
                                      <table border=0 cellspacing=0 width=389 cellpadding=0><tr>
                                        <td height=5 width="389"></td></tr><tr>
                                          <td height=1 bgcolor=<%=cor3%> width="389"></td></tr><tr>
                                          <td height=5 width="389"></td></tr></table><%=strLg43%> <b><%=totalreg%></b> <%=strLg44%> 
                                      <table border=0 cellspacing=0 width=390 cellpadding=0><tr>
                                        <td height=5 width="390"></td></tr><tr>
                                          <td height=1 bgcolor=<%=cor3%> width="390"></td></tr><tr>
                                          <td height=5 width="390"></td></tr></table><br><br><center>
    <% rs.absolutepage = pag
       contador = 0
       do while not rs.eof and contador < rs.pagesize
    
    'Fomata o pre�o
    precito = formatNumber(rs("preco"),2)%>
	
						<% '***  Verifica se tem Estoque do Produto
						set rs_estoque = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs("idprod")&" ;")
						if not rs_estoque.eof then
						estoque_atual=rs_estoque("estoque")
						end if
						rs_estoque.close
						set rs_estoque = nothing
						 %>
	
    		  				     <table border="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#cccccc>
    							 		<table border="0" CELLSPACING="1" CELLPADDING="3" width="141"><tr>
                                          <td bgcolor=#ffffff width="503">
    										   <table align=center width=371 cellspacing="1" cellpadding="1">
    										   		  <tr>
                                                        <td valign=center width=252><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs("nome")%></b><br><br><b>Pre�o:</b><%=strLgMoeda & " " & precito%><br><b>Estoque:</b> 


<% '*** Opcao de NAO Mostrar Estoque
    if estoque_atual > 0 then response.write strLg26  else response.write "<font color=red>" & strLg27 & "</font>" end if
%>
</td><td valign=center width="112"><div align=right><br><a href="produtos.asp?produto=<%=rs("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></tr>
    										   </table>
    								    </table>
    						     </table>
    <%
    contador = contador + 1
    rs.MoveNext
    loop%>
    <table width=394><tr><td colspan=2 width="410"><center><br>
    
    <table width=362 cellspacing="0" cellpadding="0"><tr>
      <td colspan=2 width="379">
    
    <%'Criando links para a navega��o%>
    
    <table border=0 cellspacing=0 width=314 cellpadding=0><tr>
      <td height=5 width="314"></td></tr><tr>
        <td height=1 bgcolor=<%=cor3%> width="314"></td></tr><tr>
        <td height=5 width="314"></td></tr></table></td></tr>
    <tr><td width="63"></td><td align=right width="316"><font face="<%=fonte%>" style=font-size:11px;color:000000><%=strLg45%>: &nbsp;
    
    
    <%for i = 1 to rs.pagecount
    
    if i = cint(pag) then
       response.write " <b>" & i & "</b> &nbsp;"
    else
       response.write "<a href='" & request.servervariables("script_name") & "?pesq="&trim(request.querystring("pesq"))&"&cat="&request.querystring("cat")&"&pagina=" & i & "'><font face="&fonte&" style=font-size:11px;color:000000><b>" & i & "</a> &nbsp;"
    end if
    
    next
    
    rs.close
    set rs = nothing
    end if%>
    
    </b></font></td></tr>
    <tr><td colspan=2 align=center width="379"><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><a HREF="javascript: history.back()" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></td></tr>
    
    </td></tr></table>
    </td></tr></table>
    </td></tr></table>
    </td></tr></table>
<%end if%><!-- #include file="baixo.asp" -->