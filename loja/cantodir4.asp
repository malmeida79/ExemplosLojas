<link href="stylesheet.css" rel="stylesheet" type="text/css">
<table width="11%"  border="0" cellpadding="0" cellspacing="0" bordercolor="#99CC00">
  <tr> 
    <td><p><img src="images/receba_produtos.gif" width="167" height="42" /></p>
    <table width="145"  border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="36"><div align="center"><font color="#999999"><strong><img src="comprados.gif" width="145" height="36"></strong></font></div></td>
        </tr>
      </table>
      <table cellspacing=0 cellpadding=5 width="145" border=0>
        <tbody>
          <%
				if VersaoDb = 1 then
				set rs = abredb.execute("SELECT idprod, nome, preco FROM produtos ORDER BY contador DESC LIMIT 0,3")
				else
				set rs = abredb.execute("SELECT TOP 5 idprod, nome, preco FROM produtos ORDER BY contador DESC")
				end if
				cont = 1
				if not rs.eof then
				  for x = 1 to 5%>
          <tr bgcolor="#FFFFFF"> 
            <td align=right valign=center><img 
                        src="linguagens/portuguesbr/imagens/seta.gif" width="6" height="5" 
                        border=0></td>
            <td class="inp"><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="text-decoration: none;"><%=rs("nome")%></a><br> 
              <b><%= strlgMoeda & " " & formatnumber(rs("preco"),2)%></b></td>
          </tr>
          <%cont = cont + 1%>
          <%rs.movenext%>
          <%next%>
          <%end if%>
        </tbody>
      </table>
      <table width="136"  border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="136"><div align="center"> <font color="#999999"><img src="news.gif" width="145" height="36"></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> <table width="131" height="164" border=0 cellpadding=0 cellspacing=1 bgcolor="#F7F7F7">
        <tbody>
          <tr> 
            <td width="48%" valign=top bgcolor="#FFFFFF"><form method=post action="javascript: cadmail()" name=emailx>
                <table cellspacing=0 cellpadding=5 width="100%" border=0>
                  <tbody>
                    <tr> </tr>
                    <tr>
                      <td align=left valign=center bgcolor="#FFFFFF" class="inp"><div align="left">
                          <h6><font face="Verdana, Arial, Helvetica, sans-serif">Caso 
                            queira receber ofertas especiais e novidades , preencha 
                            o campo abaixo(e-mail):</font></h6>
                        </div></td>
                    </tr>
                    <tr> 
                      <td align=left valign=center bgcolor="#FFFFFF" class="inp"><input name=email type=text class="inp"  value=""  size=15></td>
                    </tr>
                    <tr> 
                      <td align=left valign=center bgcolor="#FFFFFF"><select name='tipo' size='1' style="font-size:11px;font-family:tahoma,tahoma,arial,helvetica;">
                          <option value='1' selected ><%=strLg267%></option>
                          <option value='0' ><%=strLg266%></option>
                        </select></td>
                    </tr>
                    <tr> 
                      <td align=left valign=top bgcolor="#FFFFFF"><input name="image" type=image onMouseOver="window.status='<%=srtLg161%>';return true;" onMouseOut="window.status='';return true;" src="<%= dirlingua %>/imagens/cadastro.gif"></td>
                    </tr>
                    <tr> 
                      <td align=left valign=center bgcolor="#FFFFFF"><a href="ajuda_email.asp"><img src="images/televendas.gif" width="136" height="42" border="0"></a></td>
                    </tr>
                    <tr> 
                      <td align=left valign=center bgcolor="#FFFFFF"><!-- /////    Fim do Quadro Linguagens - Lateral Esquerdo  -->
                        <!-- /////    Quadro Atendimento Online - Lateral Esquerdo  -->
                        <%
Set Conn= Server.CreateObject("adodb.connection")
Conn.open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("chat/LiveSupport.mdb") 
'--------------------------------------------------------------------------------------------------      
SQl = "select * from tblSettings where Online = '0'"
set Rs = Conn.execute(SQL)
'--------------------------------------------------------------------------------------------------
	if Rs.eof then

		'*** Para aparecer a imagem de Atendimento "Online"

		Response.Write " <img src="&dirlingua&"/images/atend_on.gif onclick=""javascript:generico('chat/default.asp','aol',400,300,20,20,'no','no')""  border=0 style=""cursor:hand"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"&strLg258&"';return true;""><BR><br><br>"
	else

		'*** Para aparecer a imagem de Atendimento "Offline"  >>> Ative se vc quiser que avise que nao em ninguem "Online"

		Response.Write " <img src="&dirlingua&"/images/atend_on.gif border=0 onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='"&strLg259&"';return true;"" alt="&strLg259&"><BR><br><br>"
	end if
'--------------------------------------------------------------------------------------------------	
set rs = nothing
set atend = nothing%></td>
                    </tr>
                    <tr> 
                      <td align=left valign=center bgcolor="#FFFFFF"><p class="ti"><img src="images/atend_on.gif" width="145" height="72" /></p>
                        <p class="ti"><img src="images/site_seguro2.gif" width="147" height="77" /></p>
                        <p class="ti"><img src="images/sedex.gif" width="150" height="46" /></p></td>
                    </tr>
                  </tbody>
                </table>
              </form></td>
          </tr>
        </tbody>
      </table>
      <table width="131"  border="0" cellspacing="0" cellpadding="0">
        <tr> </tr>
      </table></td>
  </tr>
</table>
<p>&nbsp;</p>
