<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Documento sem t&iacute;tulo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<p>&nbsp;</p>
<table width="157"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="157" height="24" background="linguagens/portuguesbr/imagens/menu_dir_top.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td background="linguagens/portuguesbr/imagens/menu_dir_fundo.gif">
	
	<%
				if VersaoDb = 1 then
				set rs = abredb.execute("SELECT idprod, nome, preco FROM produtos ORDER BY contador DESC LIMIT 0,3")
				else
				set rs = abredb.execute("SELECT TOP 5 idprod, nome, preco FROM produtos ORDER BY contador DESC")
				end if
				cont = 1
				if not rs.eof then
				  for x = 1 to 5%>
				  					  
      <%=cont%>) - <a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="text-decoration: none;"><%=rs("nome")%></A><BR>
        <b><%= strlgMoeda & " " & formatnumber(rs("preco"),2)%></b></font><BR><BR>
							  <%cont = cont + 1%>
				              <%rs.movenext%>
				  <%next%>
				<%end if%>
	</td>
  </tr>
  <tr> 
    <td><img src="linguagens/portuguesbr/imagens/menu_dir_fim.gif" width="157" height="8"></td>
  </tr>
</table>
<br>
<table width="157"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="157" height="24" background="linguagens/portuguesbr/imagens/menu_dir_top.gif">&nbsp;&nbsp;  <%= strLg159 %>
    </td>
  </tr>
  <tr> 
    <td background="linguagens/portuguesbr/imagens/menu_dir_fundo.gif">
	
					
<form method=post action="javascript: cadmail()" name=emailx>
         
          <%= strLg159 %> 
            <input type=text size=15 value="<%=strLg160%>" style="font-size:11px;font-family:tahoma,verdana,arial,helvetica;color: Blue;" name=email onfocus="limpa();">
            <br>
             <select name='tipo' size='1' style="font-size:11px;font-family:tahoma,verdana,arial,helvetica;">
              <option value='1' selected ><%=strLg267%></option>
              <option value='0' ><%=strLg266%></option>
            </select>
            <br><br>
            <input type=image src="<%= dirlingua %>/imagens/cadastro1.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=srtLg161%>';return true;"></div>
 
				   
				   </form>
	
	
	 </td>
  </tr>
  <tr> 
    
  <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
