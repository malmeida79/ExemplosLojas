<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
'#
'#  C�DIGO: VirtuaStore Vers�o OPEN - Copyright 2001-2004 VirtuaStore
'#  URL: http://comunidade.virtuastore.com.br
'#  E-MAIL: comunidade@virtuastore.com.br
'#  AUTORES: Ot�vio Dias(Desenvolvedor)
'#  GRUPO DE DISCUSS�O: Virtua Developer (http://br.groups.yahoo.com/group/virtuadeveloper)
'# 
'#     Este programa � um software livre; voc� pode redistribu�-lo e/ou 
'#     modific�-lo sob os termos do GNU General Public License como 
'#     publicado pela Free Software Foundation.
'#     � importante lembrar que qualquer altera��o feita no programa 
'#     deve ser informada e enviada para os criadores, atrav�s de e-mail 
'#     ou da p�gina de onde foi baixado o c�digo.
'#
'#  //-------------------------------------------------------------------------------------
'#  // 
'#  // LEIA COM ATEN��O: O software VirtuaStore OPEN deve conter as frases
'#  // "bondhost - Hospedagem" ou "Software derivado de VirtuaStore 3.0" e 
'#  // o link para o site http://comunidade.virtuastore.com.br no cr�ditos da 
'#  // sua loja virtual para ser utilizado gratuitamente.
'#  // Se o link e/ou uma das frases n�o estiver presentes ou visiveis este
'#  // software deixar� de serconsiderado Open-source(gratuito) e o uso sem autoriza��o 
'#  // ter� penalidades judiciais de acordo com as leis de pirataria de software.
'#  // 
'#  // Se voc� concordar, pedimos que visando divulgar o nosso grupo de discuss�o, pedimos
'#  // que seja mantido o link para o Grupo Virtua Developer
'#  // (http://br.groups.yahoo.com/group/virtuadeveloper)
'#  // 
'#  //--------------------------------------------------------------------------------------
'#      
'#     Este programa � distribu�do com a esperan�a de que lhe seja �til,
'#     por�m SEM NENHUMA GARANTIA. Veja a GNU General Public License
'#     abaixo (GNU Licen�a P�blica Geral) para mais detalhes.
'# 
'#     Voc� deve receber a c�pia da Licen�a GNU com este programa, 
'#     caso contr�rio escreva para
'#     Free Software Foundation, Inc., 59 Temple Place, Suite 330, 
'#     Boston, MA  02111-1307  USA
'# 
'#     Para ajuda e suporte acesse um dos sites abaixo:
'#     http://comunidade.virtuastore.com.br
'#     http://br.groups.yahoo.com/group/virtuadeveloper
'#     
'#     ********************************
'#     Cr�ditos do C�digo Original
'#     ********************************
'#     Os arquivos originais da vers�o 3.0 foram baixados do Grupo VirtuaStore
'#     (http://br.groups.yahoo.com/group/virtuastore) e revisados e adaptados pela
'#     equipe do Grupo Virtua Developer (http://br.groups.yahoo.com/group/virtuadeveloper)
'#
'#     BOM PROVEITO!          
'#     Equipe VirtuaStore & Grupo Virtua Developer
'#     []'s
'#
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################

'Este codigo foi disponibilizado para group/virtuadeveloper por Dubairro comercial@bondhost.com.br

'IN�CIO DO C�DIGO

'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
  	  <!-- #include file="topo.asp" -->
	 <div id="layer1" style="position:absolute; left:600px; top:60px; z-index:1"></div>
	 	  <table><tr><td align="left" valign="top">
		  				 <table border="0" cellspacing="4" cellpadding="4" width=590><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> 
                           <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='Home';return true;"><b>Home</b></a> 
                           � <%=strLg319%><br><br><div align=left> <table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></div><font face="<%=fonte%>" style=font-size:13px><strong><br>
                           Rastreamento<br>
&nbsp;</strong></font><font face="tahoma,verdana,arial,helvetica" style=font-size:11px;color:000000><font face="tahoma,verdana,arial,helvetica" style=font-size:13px><strong><body><table>
                          <tr> 
                            <td nowrap> <div align="center">
                              <a href="http://www.correios.com.br" style="color: #000000"> 
                                <img src="linguagens/portuguesbr/imagens/logo_correios.gif" border="0"> 
                                </a> </div></td>
                            <td> <div align="center"><font face="Arial"> <b>&nbsp;&nbsp;</b><font size="-1" face="Verdana, Arial, Helvetica, sans-serif">Rastreamento 
                                de Objetos</font> </font> </div></td>
                          </tr>
                        </table>
                        </strong> 
                        <form method="post" action="http://www.correios.com.br/servicos/rastreamento/remoto.cfm" name="rastreamento">
                          <p>
                          <font face=arial size=2><font size="-2" face="Verdana, Arial, Helvetica, sans-serif">Objeto(s):</font><strong> 
                          &nbsp; 
                          <input name="p_codigo" type="text" value="" size="40">
                          <br>
                          <br>
                          <!-- Objetos Nacionais -->
                          <!----------------------->
                          <input type="radio" name="tipo_pesquisa" value="1" checked onclick="lingua(1,2,3,1,0);">
                          </strong></font><font face=Verdana, Arial, Helvetica, sans-serif size=-2><strong>Objetos 
                          Nacionais</strong> <br>
                          <br>
                          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Exemplos de consulta 
                          de:</font><font face=arial size=2> </p>
                          <ul>
                            <li><font size="-2" face="Verdana, Arial, Helvetica, sans-serif"> 
                              Exemplos &quot;Um&quot; objeto: SSXXXXXXXXXBR</font></li>
                            <li><font size="-2" face="Verdana, Arial, Helvetica, sans-serif">
                            Exemplo &quot;Dois&quot;Intervalo 
                              de objetos: SSXXXXXXXXXBR-SSXXXXXXXXXBR</font></li>
                            <li><font size="-2" face="Verdana, Arial, Helvetica, sans-serif">
                            Exemplos &quot;Lista 
                              de objetos&quot;: SSXXXXXXXXXBR; SSXXXXXXXXXBR; SSXXXXXXXXXBR</font> 
                            </li>
                          </ul>
                          <p>
                          <font size="-2" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <!-- Objetos Internacionais -->
                          <!---------------------------->
                          <input type="radio" name="tipo_pesquisa" value="2" onclick="lingua(1,2,3,0,1);">
                          <strong> Objetos internacionais</strong> <br>
                          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Informe o c&oacute;digo 
                          de no m&aacute;ximo 10 objetos,<br>
                          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; separando-os por 
                          ponto-e-v&iacute;rgula 
                          <!----------------------------------------->
                          <!--   C�digo para selecionar a l�ngua   -->
                          <br>
                          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Resposta em: &nbsp;&nbsp; 
                          <input type="radio" name="LINGUA" value="P" onclick="lingua(0,2,3,0,1);" checked>
                          Portugu�s &nbsp;&nbsp;&nbsp; 
                          <input type="radio" name="LINGUA" value="E" onclick="lingua(0,2,3,0,1);">
                          English 
                          <!-- Fim C�digo para selecionar a l�ngua -->
                          <!----------------------------------------->
                          </font> <strong> <br>
                          <br>
                          <br>
                          <input type="submit" name="Submit" value="Pesquisar" >
                          &nbsp; 
                          <input type="reset" name="Submit2" value="Limpar">
                          &nbsp; 
                          <input type="button" name="Submit3" value="Ajuda" onClick="return window.open('http://www.correios.com.br/sedexonline/rastreamento_ajuda.cfm'
	  ,'ajuda','width=460,height=500, scrollbars=yes' );" >
                          </strong></font> 
                          </p>
                        </form>
                        <strong> 
                        <script language="JavaScript">
<!--

//-->
                           </script>
                        </strong></font> 
                           </font><div align=justify><font style=font-size:11px face=<%=fonte%>><%=strLgquemsomos%><br><br><table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><center>
                             <a HREF="default.asp" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></center></td></tr>
						 </table></td></tr>
		 </table>
		 <!-- #include file="baixo.asp" -->