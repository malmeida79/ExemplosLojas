<%

%><body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<TABLE width="100%" border=0 align="center" cellPadding=0 cellSpacing=0 bordercolor="#FF6600">
  <TBODY>
    <TR> 
      <TD colspan="3"> <% 
if session("usuario") = "" or InStr(Request.ServerVariables("SCRIPT_NAME"),"/pagamento.asp") > 0 then
'*** MENU
%> <TABLE width="100%" height="26" border=0 cellPadding=0 cellSpacing=0 background="imagens_loja/bgbg_menu.gif" bgcolor="#F07642">
          <TBODY>
            <TR bgcolor="#003366"> 
              <td width="86%" bgcolor="#FF9900"> <p align="left"><font face="Verdana" size="2">&nbsp;</font><font color="#FFFFFF" size="2" face="Verdana"><a href="default.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Home</font></span></a> 
                  | <span style="text-decoration: none"> <font color="#FFFFFF"><a href="quemsomos.asp"> 
                  <span style="text-decoration: none"><b><font color="#FFFFFF">Nossa 
                  Empresa</font></b><font color="#FFFFFF"> </font> </span></a></font></span>| 
                  <a href="carrinhodecompras.asp"> <span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Finalizar 
                  Compras</font></span></a> | <a href="registro.asp"> <span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Registro 
                  de Clientes</font></span></a> | <a href="fechapedido.asp?compra=login&menu=sim"> 
                  <span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Login</font></span></a> 
                  | <a href="ajuda_email.asp"> <span style="text-decoration: none; font-weight:700"> 
              <font color="#FFFFFF">Contato&nbsp;&nbsp;</font></span></a></font></td>
              <td width="14%" align="right" bgcolor="#FF9933"><b><span class="style1 style3"><%=strNome%></span>&nbsp;</b></td>
            </TR>
          </TBODY>
        </TABLE>
        <% Else  %> <TABLE width="100%" border=0 cellPadding=0 cellSpacing=0 background="imagens_loja/bgbg_menu.gif" height="26">
          <TBODY>
            <TR bgcolor="#006699"> 
              <TD width="86%"> <p align="left"><font face="Verdana" size="2" color="#FFFFFF">&nbsp;<a href="default.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Home</font></span></a> 
                  |<a href="quemsomos.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Empresa</font></span></a> 
                  |<a href="carrinhodecompras.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Produtos</font></span></a> 
                  |<a href="dados.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Meus 
                  Dados</font></span></a> |<a href="minhascompras.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Historico 
                  Compras</font></span></a> |<a href="conf_pagamento.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Pagamento</font></span></a> 
                  |<a href="logout.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Sair</font></span></a> 
                  |<a href="ajuda_email.asp"><span style="text-decoration: none; font-weight:700"><font color="#FFFFFF">Contato</font></span></a></font></TD>
              <TD width="14%" align="right"><span class="style1 style3"><b><%=strNome%></b></span><span class="style1"><b>&nbsp;&nbsp;</b></span></TD>
            </TR>
          </TBODY>
        </TABLE>
        <% End If %> </TD>
    </TR>
    <TR> 
       <TD background="linguagens/portuguesbr/imagens/bck.jpg" bgcolor="#A99EC8"><img src="linguagens/portuguesbr/imagens/barra1.jpg" width="443" height="152"><img src="linguagens/portuguesbr/imagens/barra2.jpg"><img src="linguagens/portuguesbr/imagens/barra3.jpg" width="343" height="152"></TD>
    </TR>
  </TBODY>
</TABLE>
  
