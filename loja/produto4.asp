<% 'Produto 4 %>
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

 <form action="comprar.asp" method="post" name="comprar4">	
   <div align="center"><span class="style6">Fabricante: </span><span class="STYLE222"><%=rs4.fields("fabricante")%>
    </span><br>
     <a href="produtos.asp?produto=<%=rs4.fields("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="text-decoration:none;"><font color=#000000><%=rs4.fields("nome")%></font></a><br>
     <BR>
			              <% If len(cstr(rs4.fields("impeq")))<3 then impeq="img_nao_disp.gif" else impeq=rs4.fields("impeq")  end if%>
        
                        <a href="produtos.asp?produto=<%=rs4.fields("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style <%=fontebranca%>";text-decoration:none;"><IMG src="produtos/<%=impeq%>" alt='<%=rs4.fields("nome")%>' border=0 ></A>
     <BR>
                          <span class="style4"><%= strlgmoeda & " " & precito4%></span><strong><FONT 
                        face="verdana, arial, helvetica" color=#666666 
                        size=1></FONT></strong><br>
						        
			            <% '***  Verifica se tem Estoque do Produto 4
						set rs_estoque4 = abredb.execute("SELECT estoque FROM estoque WHERE idproduto="&rs4.fields("idprod")&" ;")
						if not rs_estoque4.eof then
						estoque_atual_4=rs_estoque4("estoque")
						end if
						rs_estoque4.close
						set rs_estoque4 = nothing
						 %>
						 						
						<font size="1" face="arial"><%=strLg28%> <%if estoque_atual_4 > 0 then response.write " " & strLg26 else response.write " <font color=red>" & strLg27 & "</font> " end if%>
						<br>
						<br>
						</font>
					    <%if estoque_atual_4 > 0 then 
				response.write "<a href=""JavaScript: document.comprar4.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt='"& strLg276 & "' onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"& strLg276 & "';return true;""></a>&nbsp;&nbsp;"
				end if%>                                  
                  <input type="hidden" name="intProdID" value="<%= intProdID4 %>">
                  <input type="hidden" name="intQuant" value="1">
                  <a href="produtos.asp?produto=<%=rs4.fields("idprod")%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style <%=fontebranca%>";text-decoration:none;"><img src="<%=dirlingua%>/imagens/detalhes.gif" border="0">
				  				  
   <BR></a>   </div>
 </form>
