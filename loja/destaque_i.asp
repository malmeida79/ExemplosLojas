 <script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<div align="center"><br>
  <table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr> 
      <td bgcolor="#FFFFFF"> <% If mostrar_produto_destaque_fachada="Sim" and intProdID9<>"" then %> 
    <tr> 
      <td width="100%" height="0" bgcolor="#cccccc" > <%'############################   INÍCIO DO PRODUTO DESTAQUE   ##############################

' Esta área mostra um produto em destaque acima dos produtos da vitrine.
' Adaptacao do codigo feito por jusivansjc@yahoo.com.br
%> <table width="100%" BORDER="0" CELLPADDING="0" CELLSPACING="1" style="BORDER: whithe 1px dotted; cellSpacing=1 cellPadding=1 width=213 border=0">
          <tr> 
            <td bgcolor="<%=fontebranca%>"> 
              <form action="comprar.asp" method="post" name="comprar9"><table BORDER="0" CELLSPACING="1" CELLPADDING="3"> 
                <tr> 
                  <td bgcolor="<%=fontebranca%>"> 
                    <% If len(cstr(rs9.fields("imgra")))<3 then imgra="img_nao_disp.gif" else imgra=rs9.fields("imgra")  end if%><table align=center width="100%" cellspacing="1" cellpadding="1"> 
                    <tr> 
                      <td width="435" align="center" height=180><a href="produtos.asp?produto=<%=rs9.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:"<%=fontebranca%>";text-decoration:none;"><img src="produtos/<%=imgra%>" width="180" border="0"></a></td>
                      <td align="left" valign=center width=643><font style=font-size:13px;font-family:<%=fonte%>><b><font color=<%=fontepreta%>><%=rs9.fields("nome")%></font></b><font color=<%=fontepreta%>><br>
                        <br>
                        <b><%=strLg29%></b> <FONT face="arial, helvetica" color=#0000ff size=2><B>&nbsp;<%= strlgmoeda & " " & precito9%></b></font><br>
                        <br>
                        <b><%=strLg28%></b> 
                        <% '***  Verifica se tem Estoque do Produto 9 (Produto Destaque)
						set rs_estoque9 = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs9.fields("idprod")&" ;")
						if not rs_estoque9.eof then
						estoque_atual_9=rs_estoque9("estoque")
						end if
						rs_estoque9.close
						set rs_estoque9 = nothing
						 %>
                        <%if estoque_atual_9 > 0 then response.write " " & strLg26 else response.write " " & strLg27 end if%>
                        <br>
                        <b>Entrega:</b> &nbsp; <%=application("diasentrega")%> 
                        dias(s) + transporte 
                        <div align=right><br>
                          <input type="hidden" name="intProdID" value="<%= intProdID9 %>">
                          <input type="hidden" name="intQuant" value=1>
                          <%if estoque_atual_9 > 0 then response.write "<a href=""JavaScript: document.comprar9.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt='"& strLg276 & "' onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"& strLg276 & "';return true;""></a>&nbsp;&nbsp;" end if%>
                          <a href="produtos.asp?produto=<%=rs9.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:"<%=fontebranca%>";text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div>
                        </font></font></td>
              </form></tr>
        </table> 
    <tr> 
      <td height="0" bordercolor="#00CC99" bgcolor="#336699" ><img src="linguagens/portuguesbr/imagens/destaques.gif" width="163" height="28"> 
    </table>
  
  </table>
  <%'############################   FIM DO PRODUTO DESTAQUE   ##############################%></td></tr> </td> </tr> </table> 
</div>
