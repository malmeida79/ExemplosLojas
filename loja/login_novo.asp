  		   <table><tr><td align="left" valign="top">
		   				  <table border="0" cellspacing="4" cellpadding="4" width=590>
						    <tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> » <%=strLg6%><br><br><div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table>  </div><font face="<%=fonte%>" style=font-size:13px><strong><br><%=strLg179%></strong></font></p><p align="left"><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg180%> <a href="registro.asp" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg71%>';return true;"><b> <%=strLg71%></b></a><br><br><%=strLg181%></p><form method="post" action="comprareg.asp?login=sim">
<%
'Chama o cookie gravado no pc do usuario
if request.cookies(""&nomeloja&"")("usuario") = "" or request.querystring("user") <> "" or request.querystring("user") = "x" then
 user = request.querystring("user")
 if user = "x" then
 user = ""
end if
else
user = request.cookies(""&nomeloja&"")("usuario")
end if 
%>
  	  <table>
	   		  <tr><td><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg108%></td><td><input name="email" size="35" value="<%= user %>" style=font-size:11px;font-family:<%=fonte%>> <font face=<%=fonte%> style=font-size:11px; color=red> <%=request("erro")%></td></tr>
	   		  <tr><td><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg123%></td><td><div id="senha" style="postion:absolute;z-index:80"><input name="senha" type="password" size="25" value="" style=font-size:11px;font-family:<%=fonte%>></div> <font face=<%=fonte%> style=font-size:11px; color=red> <%=request("erro2")%></td></tr>
			  <tr><td></td><td><input type="image" src="<%=dirlingua%>/imagens/login.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg179%>';return true;"></td></tr>
	   </table></form><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg182%> <a href="senha.asp" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg182%>';return true;"><b> <%=strLg71%></b></a><br><br><br><div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
	</table></td></tr>
</table>  		   
		   
		   <%
'Chama o cookie gravado no pc do usuario
if request.cookies(""&nomeloja&"")("usuario") = "" or request.querystring("user") <> "" or request.querystring("user") = "x" then
 user = request.querystring("user")
 if user = "x" then
 user = ""
end if
else
user = request.cookies(""&nomeloja&"")("usuario")
end if 
%>
  	  <table>
	   		  <tr><td><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg108%></td><td><input name="email" size="35" value="<%= user %>" style=font-size:11px;font-family:<%=fonte%>> <font face=<%=fonte%> style=font-size:11px; color=red> <%=request("erro")%></td></tr>
	   		  <tr><td><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg123%></td><td><div id="senha" style="postion:absolute;z-index:80"><input name="senha" type="password" size="25" value="" style=font-size:11px;font-family:<%=fonte%>></div> <font face=<%=fonte%> style=font-size:11px; color=red> <%=request("erro2")%></td></tr>
			  <tr><td></td><td><input type="image" src="<%=dirlingua%>/imagens/login.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg179%>';return true;"></td></tr>
	   </table></form><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg182%> <a href="senha.asp" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg182%>';return true;"><b> <%=strLg71%></b></a><br><br><br><div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
	</table></td></tr>
</table>
'Verifica se a compra está vazia
if cstr(Session("orderID")) = "" then%>
   <table><tr><td align="left" valign="top">
			   <table border="0" cellspacing="4" cellpadding="4" width=590><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> » 


<br><br><div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></div><br><br><div align='center'><p><font face=<%=fonte%> style='font-size:17px'><b><%=strLg49%><br><%=strLg50%><br><br></b><a href=default.asp><img src="<%=dirlingua%>/imagens/continuar.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg68%>';return true;" border=0></a><b></font></p></div><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
			   </table></td></tr>
	</table>
<%
response.end
else

'Login na loja      
intOrderID = cstr(Session("orderID"))%>
		   <table><tr><td align="left" valign="top">
		   				  <table border="0" cellspacing="4" cellpadding="4" width=590>
						   <tr><td><font face=<%=fonte%> style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> » 
<% if Cstr(Request("menu")) = Cstr("sim") then %>
<%=strLg6%>
<% Else  %>
<%=strLg8%>
<%end if%>
<br><br><div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table>  </div><font face="<%=fonte%>" style=font-size:13px><strong><br><%=strLg55%></strong></font></p><p align="left"><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg180%> <a href="registro.asp" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg71%>';return true;"><b> <%=strLg71%></b></a><br><br><%=strLg181%></p><form method="post" action="comprareg.asp">
<%
'Chama os cookies gravados
if request.cookies(""&nomeloja&"")("usuario") = "" or request.querystring("user") <> "" then
user = request.querystring("user")
if user = "x" then
user = ""
end if
else
user = request.cookies(""&nomeloja&"")("usuario")
end if 
%>
  	   <table>
	   		  <tr><td><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg108%></td><td><input name="email" size="35" value="<%= user %>" style=font-size:11px;font-family:<%=fonte%>> <font face=<%=fonte%> style=font-size:11px; color=red> <%=request("erro")%></td></tr>
	   		  <tr><td><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg123%></td><td><div id="senha" style="postion:absolute;z-index:80"><input name="senha" type="password" size="25" value="" style=font-size:11px;font-family:<%=fonte%>></div> <font face=<%=fonte%> style=font-size:11px; color=red> <%=request("erro2")%></td></tr>
			  <tr><td></td><td><input type="image" src="<%=dirlingua%>/imagens/login.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg179%>';return true;"></td></tr>

<%
if Cstr(Request("menu")) = Cstr("sim") then
%>
<input type=hidden name=login value=sim>
<%
end if
%>
	   </table></form><font face="<%=fonte%>"  style=font-size:11px color=000000><%=strLg182%> <a href="senha.asp" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg182%>';return true;"><b> <%=strLg71%></b></a><br><br><br><div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
	</table></td></tr>
</table>
 