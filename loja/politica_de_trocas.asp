<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
'#
'#  C�DIGO: VirtuaStore Vers�o OPEN - Copyright 2001-2004 VirtuaStore
'#  URL: http://comunidade.virtuastore.com.br
'#  E-MAIL: comunidade@virtuastore.com.br
'#  AUTORES: Ot�vio Dias(Desenvolvedor)
'# 
'#     Este programa � um software livre; voc� pode redistribu�-lo e/ou 
'#     modific�-lo sob os termos do GNU General Public License como 
'#     publicado pela Free Software Foundation.
'#     � importante lembrar que qualquer altera��o feita no programa 
'#     deve ser informada e enviada para os criadores, atrav�s de e-mail 
'#     ou da p�gina de onde foi baixado o c�digo.
'#
'#  //-------------------------------------------------------------------------------------
'#  // LEIA COM ATEN��O: O software VirtuaStore OPEN deve conter as frases
'#  // "bondhost - Hospedagem" ou "Software derivado de VirtuaStore 1.2" e 
'#  // o link para o site http://comunidade.virtuastore.com.br no cr�ditos da 
'#  // sua loja virtual para ser utilizado gratuitamente. Se o link e/ou uma das 
'#  // frases n�o estiver presentes ou visiveis este software deixar� de ser
'#  // considerado Open-source(gratuito) e o uso sem autoriza��o ter� 
'#  // penalidades judiciais de acordo com as leis de pirataria de software.
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
'#     Para enviar suas d�vidas, sugest�es e/ou contratar a VirtuaWorks 
'#     Internet Design entre em contato atrav�s do e-mail 
'#     contato@virtuaworks.com.br ou pelo endere�o abaixo: 
'#     Rua Ven�ncio Aires, 1001 - Niter�i - Canoas - RS - Brasil. Cep 92110-000.
'#
'#     Para ajuda e suporte acesse um dos sites abaixo:
'#     http://comunidade.virtuastore.com.br
'#     http://br.groups.yahoo.com/group/virtuastore
'#
'#     BOM PROVEITO!          
'#     Equipe VirtuaStore
'#     []'s
'#
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################

'IN�CIO DO C�DIGO
'----------------------------------------------------------------------------------------------------------------
'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<script language="JavaScript">
var NomeAtual;
var quantItem = 3;
	function mostraDiv(NomeDiv){
		for (cont=0;cont<=quantItem;cont++){
			eval('document.all.menu'+cont+'.style.display = "none"');
		}
		if (NomeAtual!=NomeDiv){
			eval('document.all.menu'+NomeDiv+'.style.display = ""');
			NomeAtual = NomeDiv;
		}else{
			NomeAtual = "";
		}
	}
</script>
  	  <!-- #include file="topo.asp" -->
	 <div id="layer1" style="position:absolute; left:600px; top:60px; z-index:1"></div>
	 	  <table><tr><td align="left" valign="top">
		  				 <table border="0" cellspacing="4" cellpadding="4" width=410><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='Home';return true;"><b>Home</b></a> � Pol�tica de Trocas<br><br><div align=left> <table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></div><font face="<%=fonte%>" style=font-size:13px><strong><br>Pol�tica de trocas</strong></font></p><div align=justify><font style=font-size:11px face=<%=fonte%>> Veja as pol�ticas de troca de produtos comprados atrav�s da  <%=  nomeloja %>. Aqui, voc� conta com a nossa total assist�ncia ap�s a compra de seu produto. Fique tranq�ilo e conte conosco!<br><br>
						 
						 <table width="100%" cellpadding="0" cellspacing="0" border="0" align="center">
		<tr>
		  <td><br><br>				  
				  1.&nbsp;<a href="javascript:onclick=mostraDiv('0');" class="link">Insatisfa��o do Cliente</a>
		  </td>
		</tr>
		<tr>
		  <td><div id="menu0" style="display:none">
		  <br><br>
		    <table width="100%" cellpadding="0" cellspacing="0" align="center" border="0">
		<tr>
		  <td height="1" bgcolor="#009933"></td>
		</tr>
		<tr>		  
		  <td>	  <font style=font-size:11px face="<%=fonte%>">
				  <br><b>Desist�ncia de Compra:</b><br><br>
				  Se voc� efetuou a compra atrav�s da <%= nomeloja %>, recebeu o produto em perfeitas condi��es, e ainda sim n�o se sentiu feliz com a compra, voc� poder� solicitar atrav�s de nossa central de atendimento o cancelamento da compra. Mas fique atento para as regras abaixo.<br><br>
				  Prazo para desist�ncia de compra:<br><br>
				  Conforme ditam as normas do CDC (C�digo de Defesa do Consumidor), o cliente que realiza compras atrav�s de lojas virtuais, gozam de at� 7 (sete) dias ap�s o recebimento do produto para registrar a desist�ncia da compra.<br><br>
				  O desejo de cancelamento deve ser imediatamente comunicado ao fornecedor do servi�o, seguindo as seguintes regras:<br><br>
				  <b>Devolu��o da Mercadoria:</b>
                  <ul>
				    <li> A mercadoria dever� ser devolvida atrav�s dos correios com porte pago pelo cliente, para o endere�o constante na Nota Fiscal de compra.</li>
				    <li> O produto deve ser devolvido em sua embalagem original, acompanhado de todos os acess�rios e manuais.</li>
				    <li> O produto deve ser devolvido acompanhado da 2� via da Nota Fiscal de Compra.</li>
				    <li> O produto n�o poder� apresentar qualquer ind�cio de uso.</li>
				  </ul>
				  <font color="#CC0000">ATEN��O:</font> O produto que n�o atender as condi��es exigidas acima, n�o ser� aceito como devolu��o, e, automaticamente ser� remetido de volta ao endere�o de origem. Nessas condi��es, a <%= nomeloja %> se reservar� no direito de fazer nova cobran�a de frete.<br><br>

				  <b>Restitui��o dos Valores:</b>
				  <ul>
				  	<% If mostrar_pgtos_com_cartoescredito="Sim" then %>
				    <li> <b>Cart�o de cr�dito:</b> O estorno poder� ocorrer em at� 2 faturas subsequentes.</li>
					<% End If %>
				    <li> <b>Boleto Banc�rio:</b> Cr�dito em conta corrente em at� 10 dias �teis.</li>
				    <li> N�o haver� restitui��o do valor do frete.</li>
				  </ul>
				  <font color = "#CC0000">ATEN��O:</font> A restitui��o dos valores ser�o processadas somente ap�s o recebimento e an�lise das condi��es do(s) produto(s) em nosso Estoque. (O produto n�o poder� trazer qualquer ind�cio de uso).<br><br>

				  <font color = "#CC0000">NOTA:</font> <strong>Toda inten��o de cancelamento dever� obrigatoriamente ser comunicada � <%= nomeloja %> antes do envio do(s) produto(s). Caso contr�rio, os pedidos n�o ser�o aceitos.</strong><br><br>
				</font>
  		  </td>
		</tr>
		<tr>
		  <td height="1" bgcolor="#009933"></td>
		</tr>
		    </table>
		  </div></td>
		</tr>
		<tr>
		  <td><br><br>
		    	 2.&nbsp;<a href="javascript:onclick=mostraDiv('1');" class="link">Produtos sem defeito</a>
		  </td>
		</tr>
		<tr>
		  <td><div id="menu1" style="display:none">
		  <br><br>
		    <table width="100%" cellpadding="0" cellspacing="0" align="center" border="0">
		<tr>
		  <td height="1" bgcolor="#009933"></td>
		</tr>
		<tr>
		  <td>
				  <font style=font-size:11px face="<%=fonte%>">
				  <br><b>Produto Novo:</b><br><br>
				  O produto novo goza de at� 30 dias para ser substitu�do. Desde que este seja apresentado nas mesmas condi��es em que foi recebido/comprado (na embalagem original e sem uso).<br><br> 
				  <b>"Lembre-se que para efetuarmos a troca, o produto dever� estar na embalagem original e sem uso."</b><br><br>
				  Se voc� escolheu seu produto em nossa loja virtual, mas ao receb�-lo n�o se agradou e deseja fazer a substitui��o por outro item. Observe as regras:<br><br>
				  <b>Devolu��o da Mercadoria:</b>
				  <ul>
				    <li> O produto n�o poder� apresentar qualquer ind�cio de uso.</li>
				    <li> O produto deve ser devolvido acompanhado da 2� via da Nota Fiscal de Compra.</li>
				    <li> O produto deve ser devolvido em sua embalagem original, acompanhado de todos os acess�rios e manuais.</li>
				    <li> A mercadoria dever� ser devolvida atrav�s dos correios com porte pago e remetido ao endere�o constante na Nota Fiscal de compra.</li>
				  </ul>
				  <b>Substitui��o da Mercadoria:</b>
				  <ul>
				    <li> Voc� poder� fazer a escolha de outro item conforme estoque dispon�vel em nossa loja virtual.</li>
				    <li> A escolha dever� se limitar ao valor limite do produto. Se houver diferen�a de pre�o para maior, dever� ser providenciado o pagamento da diferen�a, via op��es existentes no site.</li>
				    <li> O cliente dever� indicar no verso da 2� via da Nota Fiscal de compra, o item escolhido para troca.</li>
				    <li> O produto ser� despachado para a resid�ncia do cliente, mediante pagamento de novo frete.</li>
				  </ul>
				  <br>
				  </font>
		  </td>
		</tr>
		<tr>
		  <td height="1" bgcolor="#009933"></td>
		</tr>
		    </table>
		  </div></td>
		</tr>
		<tr>
		  <td><br><br>
		    	  3.&nbsp;<a href="javascript:onclick=mostraDiv('2');" class="link">Produtos com defeito</a>
		  </td>
		</tr>
		<tr>
		  <td><div id="menu2" style="display:none">
		  <br><br>
		    <table width="100%" cellpadding="0" cellspacing="0" align="center" border="0">
		<tr>
		  <td height="1" bgcolor="#009933"></td>
		</tr>
		<tr>
		  <td>
		    	<font style=font-size:11px face="<%=fonte%>">
				<br><b>Produto com suposta Falha de Fabrica��o:</b><br><br>
				Os produtos comercializados pela <%= nomeloja %>, gozam de 90 dias de garantia para falhas naturais de fabrica��o. Salvo a determinados produtos, onde o pr�prio fabricante indica em seu certificado de garantia o prazo de garantia de f�brica e presta atendimento direto ao cliente atrav�s dos postos de Assist�ncia T�cnica.<br><br>
				Veja como solicitar o pedido de avalia��o de supostas falhas de fabrica��o:<br><br>
	
				<b>Devolu��o da Mercadoria:</b>
				<ul>
				  <li>A mercadoria dever� ser encaminhada atrav�s dos correios para o endere�o constante na Nota Fiscal de compra.</li>
				  <li>Devolver o produto na embalagem original.</li>
				  <li>Obrigatoriamente o cliente dever� encaminhar a 2� via da Nota Fiscal de Compra juntamente com o produto.</li>
				  <li>O cliente dever� enviar por escrito, um breve relato sobre o suposto defeito a qual se refere a reclama��o, esclarecendo sobre o problema ocorrido, etc.</li>
				</ul>
				<font color = "#CC0000">ATEN��O:</font> Produtos encaminhados fora das especifica��es acima, n�o ser� aceito para an�lise de defeitos e automaticamente ser� devolvido ao remetente.<br><br>

				<b>An�lise de defeitos dos produtos:</b><br><br>
				A avalia��o de defeitos ser� realizada pelos nossos fornecedores, de onde receberemos o laudo final do pedido de troca.<br><br>
				Prazo m�dio de conclus�o: 15 dias �teis ap�s o recebimento do produto.<br><br>

				<b>Laudo Favor�vel a troca:</b>
				<ul>	
				  <li> O cliente receber� no endere�o de origem, sem custos adicionais, a substitui��o pelo mesmo produto.</li>
				  <li> Na aus�ncia do mesmo produto em estoque, o cliente ser� comunicado e  poder� escolher um outro produto para troca, entre as op��es existentes no site, respeitando o valor limite do cr�dito.</li>
				  <li> Se houver diferen�a de pre�o entre o produto escolhido e o produto reclamado, dever� ser providenciado o pagamento da diferen�a.</li>
				</ul>
				
				<b>Laudo Contr�rio a troca:</b><br><br>
				O produto ser� devolvido ao cliente com a carta/laudo da reprova��o, sem direito de substitui��o.<br><br>
				Itens de reprova��o:
				<ul>
				  <li>Aus�ncia de defeito (n�o constata��o do dano apontado pelo cliente).</li>
				  <li>Ind�cios de uso inadequado do produto.</li>
				  <li>Ind�cios de dano acidental.</li>
				  <li>Desgaste natural em decorr�ncia do uso.</li>
				  </font>
			</ul><br>
		  </td>
		</tr>
		<tr>
		  <td height="1" bgcolor="#009933"></td>
		</tr>
		    </table>
		  </div></td>
		</tr>
		<tr>
		  <td><br><br>
		    	  4.&nbsp;<a href="javascript:onclick=mostraDiv('3');" class="link">Por parte da <%= nomeloja %></a>
		  </td>
		</tr>
		<tr>
		  <td><div id="menu3" style="display:none">
		  <br><br>
		    <table width="100%" cellpadding="0" cellspacing="0" align="center" border="0">
		<tr>
		  <td height="1" bgcolor="#009933"></td>
		</tr>
		<tr>
		  <td>
		    	<font style=font-size:11px face="<%=fonte%>">
				<br><b>Cancelamento Autom�tico de compra:</b><br><br>
				O Cancelamento do pedido e libera��o dos produtos adquiridos por iniciativa da <%= nomeloja %> ser� autom�tico nas seguintes situa��es:
				<ul>
				  <li>Impossibilidade de execu��o do d�bito correspondente � compra no cart�o de cr�dito. </li>
				  <li>Inconsist�ncia de dados preenchidos no pedido.</li>
				  <li>N�o pagamento do boleto banc�rio.</li>
				  <li>Aus�ncia de estoque.</li>
				</ul>
				A Restitui��o de valores por iniciativa da <%= nomeloja %> ocorrer� nas seguintes situa��es:
				<ul>
				  <li>Impossibilidade de entrega da mercadoria adquirida, tendo em vista a inexist�ncia do endere�o de entrega como indicado pelo comprador, ou a sua n�o acessibilidade. Na impossibilidade da entrega por essa raz�o, o produto retornar� para o nosso Centro de Distribui��o, gerando devolu��o por parte da <%= nomeloja %> dos valores correspondentes aos pre�os dos produtos, excluindo-se os custos com frete.</li><br><br>
				  <li>Caso ocorra a falta da mercadoria adquirida pelo comprador no site da <%= nomeloja %>, em fun��o de ocorr�ncia de vendas do produto, haver� a devolu��o dos valores pagos ao comprador, atrav�s do mesmo meio de pagamento utilizado na compra. Nessas situa��es, o comprador ser� comunicado do ocorrido.</li>
				</ul><br>
				</font>
		  </td>
		</tr>
		    </table>
		  </div></td>
		</tr>
		
</table>
						<br> 
						 <table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><center><a HREF="default.asp" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg41%>';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b><%=strLg41%></b> ::</a></font></center></td></tr>
						 </table></td></tr>
		 </table>
		 <!-- #include file="baixo.asp" -->