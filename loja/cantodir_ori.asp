<link href="stylesheet.css" rel="stylesheet" type="text/css">
<table width="11%"  border="0" cellpadding="0" cellspacing="0" bordercolor="#99CC00">
  <tr> 
    <td><table width="151" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="36"><div align="center">
          <iframe style="width: 151px; height: 82px" name="I1" frameBorder="0" scrolling="no" src="http://www.studio7seven.com/midia/atendimento/mysStatus.asp">
			</iframe>
          </div></td>
        </tr>
      </table>
      <table width="145"  border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="36"><div align="center"><font color="#999999"><strong><img src="comprados.gif" width="145" height="36"></strong></font></div></td>
        </tr>
      </table>
      <table cellspacing=0 cellpadding=5 width="145" border=0>
        <tbody>
          <%
				if VersaoDb = 1 then
				set rs = abredb.execute("SELECT idprod, nome, preco FROM produtos ORDER BY contador DESC LIMIT 100")
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
                      <td align=left valign=center bgcolor="#FFFFFF"><img src="linguagens/portuguesbr/imagens/sedex.gif" width="150" height="46"></td>
                    </tr>
                    <tr> 
                      <td align=left valign=center bgcolor="#FFFFFF"><a href="seguranca.asp"><img src="linguagens/portuguesbr/imagens/site_seguro2.gif" width="147" height="77" border="0" /></a></td>
                    </tr>
                    <tr> 
                      <td align=left valign=center bgcolor="#FFFFFF">&nbsp;</td>
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
