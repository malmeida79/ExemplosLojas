<% 'Produto 9 %>
<style type="text/css">
<!--
.STYLE222 {
	font-size: 10px;
	font-family: Tahoma;
	font-weight: bold;
	color: #000000;
}
.style4 {
	font-family: Tahoma;
	font-size: 12px;
	color: #FF0000;
	font-weight: bold;
}
.style6 {font-family: Tahoma; color: #000000; font-size: 10px;}
-->
</style>

 <form action="comprar.asp" method="post" name="comprar99">	
   <div align="center"><span class="style6">Fabricante: </span><span class="STYLE222"><%=rs99.fields("fabricante")%>
    </span><br>
     <a href="produtos.asp?produto=<%=rs99.fields("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="text-decoration:none;"><font color=#000000><%=rs99.fields("nome")%></font></a><br>
     <BR>
			              <% If len(cstr(rs99.fields("impeq")))<3 then impeq="img_nao_disp.gif" else impeq=rs99.fields("impeq")  end if%>
        
                        <a href="produtos.asp?produto=<%=rs99.fields("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style <%=fontebranca%>";text-decoration:none;"><IMG src="produtos/<%=impeq%>" alt='<%=rs99.fields("nome")%>' border=0 ></A>
     <BR>
                          <span class="style4"><%= strlgmoeda & " " & precito99%></span><strong><FONT 
                        face="verdana, arial, helvetica" color=#666666 
                        size=1></FONT></strong><br>
						        
			            <% '***  Verifica se tem Estoque do Produto 8
						set rs_estoque99 = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs99.fields("idprod")&" ;")
						if not rs_estoque99.eof then
						estoque_atual_99=rs_estoque99("estoque")
						end if
						rs_estoque99.close
						set rs_estoque99 = nothing
						 %>
						 						
						<font size="1" face="arial"><%=strLg28%> <%if estoque_atual_99 > 0 then response.write " " & strLg26 else response.write " <font color=red>" & strLg27 & "</font> " end if%>
						<br>
						<br>
						</font>
					    <%if estoque_atual_99 > 0 then 
				response.write "<a href=""JavaScript: document.comprar99.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt='"& strLg276 & "' onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"& strLg276 & "';return true;""></a>&nbsp;&nbsp;"
				end if%>                                  
                  <input type="hidden" name="intProdID" value="<%= intProdID99 %>">
                  <input type="hidden" name="intQuant" value="1">
                  <a href="produtos.asp?produto=<%=rs99.fields("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style <%=fontebranca%>";text-decoration:none;"><img src="<%=dirlingua%>/imagens/detalhes.gif" border="0">
				  				  
   <BR></a>   </div>
 </form>
