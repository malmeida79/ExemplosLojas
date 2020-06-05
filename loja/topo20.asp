<!-- #include file="topo_i.asp" -->
<!-- #include file="barrabotoes.asp" -->
<%
'-----------------------------------------------------------------------------------------------------
%>


		<!-- LAYER PRINCIPAL -->
			
	 <table border="0" bgcolor="#FFFFFF" cellpadding="0" cellspacing="0" width="750">
		 <tr>
		 <td width="161" valign="top" align="center">
			 
		 <table border="0" cellspacing="4" cellpadding="4" width="161">
 <tr>
 <td>
 
<!--Busca de produtos -->
<table width="184" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="184" height="23" background="linguagens/portuguesbr/imagens/menu_esq_top.gif">&nbsp;&nbsp; 
      <font color="#FFFFFF"><strong>BUSCA DE PRODUTOS</strong></font> </td>
  </tr>
  <tr>
    <td background="linguagens/portuguesbr/imagens/menu_esq_fundo.gif"> <BR>
	
		<!--- start SEARCH --->

      
        <TABLE width="90%" border=0 align="center" cellPadding=2 cellSpacing=0>
		<FORM style="DISPLAY: inline" name=newsearch action=pesquisa.asp method=get>
          <TR>
            <td colspan=2> <div align="center">
                <input type=text name=pesq value="<%= request.querystring("pesq") %>" size=26 style=font-size:11px;font-family:<%=fonte%>;>
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
	'#########################################################################################
	'#----------------------------------------------------------------------------------------
	'#########################################################################################
	'### Mostra categorias e sub-categorias
	'### Rogério Silva - WBSolutions - http://www.libihost.net/wbsolutions.
	'###    SUB MENU / CATEGORIA
	'#########################################################################################
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
    <td width="184" height="23" background="linguagens/portuguesbr/imagens/menu_esq_top.gif">&nbsp;&nbsp; 
     <font color="#FFFFFF"> <strong><%=(strLg12)%></strong></font> </td>
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

'-----------------------------------------------------------------------------------------------------
'Monta o menu de departamentos (Sessoes e Categorias)
'-----------------------------------------------------------------------------------------------------

		'#########################################################################################
		'### Mostra as Sessoes da Loja
		'### Rogério Silva - WBSolutions - http://www.libihost.net/wbsolutions.
		'###    MENU PRINCIPAL (Tabela Sessoes)
		'#########################################################################################

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
                        <strong>&nbsp;<%=menu("nome")%></strong>
                        </div>
                        <%IF cInt(Mostrar) = menu("id") then
                        Response.Write "<span class=""submenu"" id="""&menu("nome")&""" style=""display:block"">"
                        END IF

      
      '############################################################

            '### Mostra as Categorias das Sessoes da Loja
            '### Rogério Silva - WBSolutions - 
            '###    SUB MENU (Tabela Categoria)
      
      '############################################################

            SQL = "SELECT C.idcategoria, C.nome FROM categoria as c ,sessoes as s  WHERE s.id = c.idsessao and c.idsessao = "&menu("id")&" AND C.ver = 's' AND C.nome <> 'Todos' ORDER by c.nome"
            Set Smenu = abredb.Execute(SQL)
            While Not Smenu.EOF%>
            <a href="sessoes.asp?item=<%=menu("id")%>&Categoria=<%=Smenu("idcategoria")%>" style="text-decoration:none;"onMouseOut="window.status='';return true;" 
onMouseOver="window.status='<%= menu("nome")&" > "&Smenu("nome") %>';return true;">
              &nbsp;&nbsp;<img src='<%=dirlingua%>/imagens/menu_esq_seta.gif' border=0>&nbsp;<%=Smenu("nome")%></a><BR>

		
		
		
		
		
		
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
'-----------------------------------------------------------------------------------------------------
'Fecha o menu
'-----------------------------------------------------------------------------------------------------
menu.Close
Set menu = Nothing


'-------------------------------------------------------------------------------------------------??????KE???E?e?¹??----
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


<!-- /////     Quadro Atendimento - Lateral Esquerdo  

			
            <TABLE cellSpacing=0 cellPadding=0 width=160 border=0 >
                   <TR><TD style="BACKGROUND: <%= Cor_principal %>; HEIGHT: 22px" colSpan=3> &nbsp;&nbsp;<B><font color="#FFFFFF"><%=strLg14%></font></B></TD></TR>
                   </TBODY></TABLE>
				   
					<TABLE cellSpacing=0 cellPadding=2 width=160 border=0 bgcolor="#ffffff" style="BACKGROUND: #ffffff; BORDER-LEFT: #bbbaba 1px solid; BORDER-RIGHT: #bbbaba 1px solid; BORDER-BOTTOM: #bbbaba 1px solid">

                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="como.asp"><%=strLg16%></A></TD></TR>

                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="comopagar.asp"><%=strLg287%></A></TD></TR>
                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="como_reimprimir_boleto.asp"><%=strLg288%></A></TD></TR>						
                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="como_conf_pagamento.asp"><%=strLg313%></A></TD></TR>						
                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="ajuda_email.asp"><%=strLg17%></A></TD></TR>
                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="quemsomos.asp"><%=strLg20%></A></TD></TR>
                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="seguranca.asp"><%=strLg22%></A></TD></TR>
						<% If mostrar_politica_de_trocas="Sim" then%>
                    <TR>
                      <TD class=ti><IMG alt="Blue Arrow" 
                        src="<%= dirlingua %>/imagens/arrow_blue.gif" 
                        border=0>&nbsp;<A class=bbh 
                        href="politica_de_trocas.asp"><%=strLg296%></A></TD></TR>
						<% End If %>

</TBODY></TABLE>
			  
			  
			  <BR><IMG height=10 src="<%= dirlingua %>/imagens/spacer.gif" width=1><BR>
			
-->
			  
<!-- /////    Fim do Quadro Atendimento - Lateral Esquerdo  -->





<!-- /////     Quadro Linguagens - Lateral Esquerdo  -->
<%
If mostrar_quadro_linguagens="Nao" then 

	'*** Colcoado este filtro, pois se for usado a traducao nesta pagina, ocorrerá um erro, pois é ENCERRADO a sessão no pagamento.asp
	 If InStr(Request.ServerVariables("SCRIPT_NAME"),"pagamento.asp") = 0 then %>
			
            <TABLE cellSpacing=0 cellPadding=0 width=160 border=0 >
                   <TR><TD style="BACKGROUND: <%= Cor_principal %>; HEIGHT: 22px" colSpan=3> &nbsp;&nbsp;<B><font color="#FFFFFF"><%=strLg260%></font></B></TD></TR>
                   </TBODY></TABLE>
				   
					<TABLE cellSpacing=0 cellPadding=2 width=160 border=0 bgcolor="#ffffff" style="BACKGROUND: #ffffff; BORDER-LEFT: #bbbaba 1px solid; BORDER-RIGHT: #bbbaba 1px solid; BORDER-BOTTOM: #bbbaba 1px solid">

                    <TR>
                      <TD class=ti align="center"><%= OpcaoLingua %></TD></TR>

</TBODY></TABLE>
			  
			  
<BR><IMG height=10 src="<%= dirlingua %>/imagens/spacer.gif" width=1><BR>
			

<%
	end If
end if %>			  
<!-- /////    Fim do Quadro Linguagens - Lateral Esquerdo  -->


<!-- /////    Quadro Atendimento Online - Lateral Esquerdo  -->

<table width="184" height="87"  border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td background="linguagens/portuguesbr/imagens/atendimentoonline_fundo.gif"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td><div align="right">
		  
		  
		  <%
Set Conn= Server.CreateObject("adodb.connection")
Conn.open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("chat/LiveSupport.mdb") 
'--------------------------------------------------------------------------------------------------      
SQl = "select * from tblSettings where Online = '0'"
set Rs = Conn.execute(SQL)
'--------------------------------------------------------------------------------------------------
	if Rs.eof then

		'*** Para aparecer a imagem de Atendimento "Online"

		Response.Write " <img src="&dirlingua&"/imagens/atendimentoonline_online.gif onclick=""javascript:generico('chat/default.asp','aol',400,300,20,20,'no','no')""  border=0 style=""cursor:hand"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"&strLg258&"';return true;"">&nbsp;&nbsp;<BR><br><br>"
	else

		'*** Para aparecer a imagem de Atendimento "Offline"  >>> Ative se vc quiser que avise que nao em ninguem "Online"

		Response.Write " <img src="&dirlingua&"/imagens/atendimentoonline_offline.gif border=0 onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"&strLg259&"';return true;"" alt="&strLg259&">&nbsp;&nbsp;<BR><br><br>"
	end if
'--------------------------------------------------------------------------------------------------	
set rs = nothing
set atend = nothing%>
		  
		  
		  </div></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!-- /////    Fim do Quadro Atendimento Online - Lateral Esquerdo  -->



<!-- /////    Inicio do Quadro Contador - Lateral Esquerdo  -->		
			   
<!--#include file="contador.asp"-->

<% If mostrar_contador="Sim" then %>
             <TABLE cellSpacing=0 cellPadding=0 width=160 border=0 >
                   <TR><TD style="BACKGROUND: <%= Cor_principal %>; HEIGHT: 22px" colSpan=3> &nbsp;&nbsp;<font color="#FFFFFF" size="1"><%=strLg263&date%></font></TD></TR>
                   </TBODY></TABLE>
				   
					<TABLE cellSpacing=0 cellPadding=2 width=160 border=0 bgcolor="#ffffff" style="BACKGROUND: #ffffff; BORDER-LEFT: #bbbaba 1px solid; BORDER-RIGHT: #bbbaba 1px solid; BORDER-BOTTOM: #bbbaba 1px solid">

                    <TR>
                      <TD class=ti align="center"><strong><font face="Tahoma" color="#555555" style=font-size:12px><%=formatazeros(UltimoNumero, 6)%></strong></TD></TR>

</TBODY></TABLE>
<% End If %>

<br><br><br>
<!-- /////    Fim do Quadro Contador - Lateral Esquerdo  -->			   

</div>						

</td>
</tr>
</table>
		</td>
		<td width="1" ><img src="<%=dirlingua%>/imagens/spacer.gif" width="1" height="10" border="0">
		<!-- <img src="<%=dirlingua%>/imagens/dot_gray_1x1.gif" width="1" height="100%" border="0"> --></td>
		<td align="left" valign="top">
		
		<!-- TABELA PRINCIPAL -->
		<table width="100%" border="0" cellspacing="0" cellpadding="0" valign=top>
		<tr>
		<td valign=top>	
		
<%
'-----------------------------------------------------------------------------------------------------
Sub BaixoC()
response.write "<a class=menu  href=""http://www.bondhost.com.br"" target=_new>bondhost - Hospedagem</a>"
End Sub

Sub BaixoComunidade()
response.write application("link_comunidade") 
End Sub
'-----------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------
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
'-----------------------------------------------------------------------------------------------------
%>