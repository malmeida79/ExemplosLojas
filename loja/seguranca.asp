<%

%>
  	  <!-- #include file="topo.asp" -->
	 
	 	  <table><tr><td align="left" valign="top">
		  				 <table border="0" cellspacing="4" cellpadding="4" width="100%"><tr>
                      <td> <font face="<%=fonte%>" style=font-size:11px;color:000000> 
                        <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='Home';return true;"><b>Home</b></a> 
                        » <%=strLg19%><br><br><div align=left> 
						
						<table border=0 cellspacing=0 width="100%" cellpadding=0>
						  <tr>
						    <td height=5></td>
						  </tr>
						  <tr>
						     <td height=1 bgcolor=<%=cor3%>></td>
						  </tr>
						  <tr>
						     <td height=5></td>
						  </tr>
						</table>
						
						<img src="adm_imagens/senha_admin.gif" width="48" height="48" border="0" align="right"></div><font face="<%=fonte%>" style=font-size:13px><strong><br><%=strLg19%></strong></font></p><div align=justify><font style=font-size:11px face=<%=fonte%>><b><%=strLg239%></b><br><%=strLgcomprasegura%><br><br><b><%=strLg240%></b><br><%=strLgcriptografia%>

<% If mostrar_pgtos_com_cartoescredito="Sim" then %>
<br><br><b><%=strLg241%></b><br><%=strLgcomprascc%>
<% End If %>

<br><br><b><%=strLg242%></b><br><%=strLginformacoes%><br><br><b><%=strLg243%></b><br><%=strLgduvsug%><br>
<img src="linguagens\portuguesbr\imagens/site_seguro.gif" alt="A imagem &ldquo;http://www.superasp.com.br/totalecommerce/imagens/catalogo/site_seguro.gif&rdquo; cont&eacute;m erros e n&atilde;o pode ser exibida." width="145" height="90" style="width: 145px; height: 90px;" /><br>
<table border=0 cellspacing=0 width="100%" cellpadding=0>
<tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr>
<tr><td height=5></td></tr></table>
<center><a HREF="ajuda.asp" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></center></td></tr>
						 </table></td></tr>
		 </table>
		 <!-- #include file="baixo.asp" -->