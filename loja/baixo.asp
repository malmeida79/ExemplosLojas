<% 
'INÍCIO DO CÓDIGO
%>    
</td>
<td width="185" valign="top">
<!--#include file="cantodir.asp" -->
</td></tr></table>
</table> 
<!-- ////   Sombra do Rodape -->
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD bgColor=#808080><IMG height=1 
      src="<%= dirlingua %>/imagens/spacer.gif" width=1></TD></TR></TBODY></TABLE>

<!-- ////   Fim da Sombra do Rodape -->


<!--   /////   Rodape    -->


<TABLE cellSpacing=0 cellPadding=0 width="100%" bgColor=#FBFBFB border=0>
<TBODY>
<TR>
<TD align=middle><IMG height=3 
      src="<%= dirlingua %>/imagens/spacer_1x1.gif" width=1 border=0><BR>
<TABLE cellSpacing=0 cellPadding=2 width="50%" border=0>
<TBODY>
<TR>
              <TD align="center"><a class=bb href="default.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><font 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%=strLg4%></font></a><a class=bb href="default.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"></a><SPAN 
            > | </SPAN> 
                <%
Set menui = abredb.Execute("SELECT * FROM sessoes WHERE ver='s' ORDER by nome;")
While Not menui.EOF%>
                <a class=bb href="sessoes.asp?item=<%= menui("id") %>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%= menui("nome") %>';return true;"><FONT 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%= menui("nome") %></font></a> <SPAN> | </SPAN>
<%menui.MoveNext
Wend
menui.Close
Set menui = Nothing%> 
</TD></TR><TR>
<TD noWrap align="center">
<br>
<a class=bb href="quemsomos.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg20%>';return true;"><FONT 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%=strLg20%></font></a><SPAN > | </SPAN>
<a class=bb href="seguranca.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg19%>';return true;"><FONT 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%=strLg19%></font></a><SPAN 
            > | </SPAN>
<a class=bb href="como.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg16%>';return true;"><FONT 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%=strLg16%></font></a><SPAN 
            > | </SPAN>
<a class=bb href="carrinhodecompras.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg51%>';return true;"><FONT 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%=strLg51%></font></a><SPAN 
            > | </SPAN>
<a class=bb href="<%=link%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg8%>';return true;"><FONT 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%=strLg8%></font></a><SPAN 
            > | </SPAN><a class=bb href="registro.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg5%>';return true;"><FONT 
                        face="tahoma, arial, helvetica" color=#000000 
                        size=1><%=strLg5%></font></a>
</TD></TR></TBODY></TABLE>
      </TD>
    </TR></TBODY></TABLE>



<TABLE cellSpacing=0 cellPadding=0 width="100%" bgColor= "#FF9933" border=0>
<TBODY>
<TR>
      <TD align=middle><BR>
<TABLE cellSpacing=0 cellPadding=2 width="50%" border=0>
<TBODY>
<TR>
<TD noWrap align="center"> 
  <p><font size="1" color="#ffffff"><%= nomeloja %> - <%= Slogan_da_sua_loja %><br>
  <font size="1" color="#ffffff">Av. Onze de Agosto, 79 Jardim Silvestre - São Bernardo do Campo Fone/Fax: (011) 4367-1415
    <br>
  © Copyright 2007-<%= year(now) %> &nbsp;&nbsp;Todos os direitos Reservados a Compuware Technologies&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <%call BaixoC()%>
  &nbsp; &nbsp;
    </font><br>
    <br>
  </p></TD></TR></TBODY></TABLE>
</TD></TR></TBODY></TABLE>
			
			


<!--   /////   Fim do Rodape    -->
	

<%
'Fecha o banco de bados
abredb.close
set abredb = nothing
%>