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
'#     http://www.bondhost.com.br
'#
'#     BOM PROVEITO!          
'#     Equipe VirtuaStore
'#     []'s
'#
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################

'IN�CIO DO C�DIGO

'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP

'Desenvolvido por: Cirilo Jos� Veloso  (cjveloso@ig.com.br - www.veloso.kit.net)

%>
<!-- #include file="topo.asp" -->
<table>
 <tr><td align="left" valign="top">
   <table border="0" cellspacing="4" cellpadding="4" width="100%">
     <tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='Home';return true;"><b>Home</b></a> � Como Pagar?<br><br><div align=left> <table border=0 cellspacing=0 width="100%" cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table>
       </div>
    <font face="<%=fonte%>" style=font-size:13px><strong><br>Como pagar na <%=nomeloja%></strong></font><br><br><br>
    <div align=center>
<!--CONTEUDO-->
<!--cart�o de cr�dito-->
    <table width="100%" align="center" border="0">
	
	<% If mostrar_pgtos_com_cartoescredito="Sim" then %>
      <tr bgColor="#006666">
        <td width="100%" height="20" bgcolor="#f0f0f0"><font face="<%=fonte%>"  style=font-size:11px color=000000><strong><font style=font-size:15px><b>�</b></font> Cart�o de cr�dito</strong></font></td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td> 
      <table>
      <tr vAlign="center">
          <td align="center"><font style=font-size:11px face="<%=fonte%>"><img src=<%=dirlingua%>/imagens/visa.gif border=0></td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>"><img src=<%=dirlingua%>/imagens/mastercard.gif border=0></td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>"> <img src=<%=dirlingua%>/imagens/amex.gif border=0></td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>"> <img src=<%=dirlingua%>/imagens/diners.jpg border=1></td>
        </tr>
        <tr>
          <td align="center"><font style=font-size:11px face="<%=fonte%>">Visa</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>">MasterCard</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>">Amex</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>">Diners</td>
        </tr>
       </table><br><br>
	   
       </td></tr>
       <tr><td width=50% valign=bottom>&nbsp;</td></tr>
	<% End If %>
		
       <tr bgColor="#006666">
        <td width="100%" height="20" bgcolor="#f0f0f0"><font face="<%=fonte%>"  style=font-size:11px color=000000><strong><font style=font-size:15px><b>�</b></font> Boleto Banc�rio</strong></font></td>	   
      </tr>

      <tr vAlign="center">
       <td><br>
        <p><font face="<%=fonte%>" style=font-size:11px;color:000000>
        <img height="44" src="./linguagens/portuguesbr/imagens/pg_boleto.gif" width="87" align="baseline">
        <br><br>
         O pagamento atrav�s de <strong>Boleto banc�rio</strong>
         � a forma mais pr�tica e segura para compras realizadas na Internet,
         pois os clientes n�o necessitam inserir dados como numero de cart�o de
         cr�dito ou senha.</font></p>
         <div align="justify">
           <ul>
           <li><font face="<%=fonte%>" style=font-size:11px;color:000000>O boleto dever� ser impresso ao
            final de sua compra e pode ser pago em qualquer ag�ncia banc�ria do
            pa�s ou via <strong>Home-Bank.</strong></font>
           <li><font face="<%=fonte%>" style=font-size:11px;color:000000>Os pedidos ser�o
            enviados ap�s confirma��o do pagamento pela institui��o banc�ria, conforme o prazo de entrega.</font></li>
          </ul>
         </div>
          </tr> </td>

<script LANGUAGE="JavaScript">
  function Boleto() {
  remote = window.open("","remotewin","'toolbar=no,location=no,directories=no,status=yes,menubar=yes,scrollbars=yes,resizable=no,copyhistory=no,width=720,height=500'");
  remote.location.href = "boleto_hsbc.asp?sacador=RAZ�O SOCIAL OU NOME DO COMPRADOR&Cpf=11111111111&endereco=ENDERE�O DO COMPRADOR, 1000&bairro=BAIRRO&cidade=CIDADE&estado=UF&cep=64999000&nossonumero=1234567890&datadocumento=01/01/2003&datavencimento=06/01/2003&valordocumento=<%=formatnumber(10000,2)%>&numerodoc=999999";
  remote.opener = window;
  remote.opener.name = "imagem1";
 }
</script>
   <tr><td valign=center><font face="<%=fonte%>" style=font-size:11px;color:000000>Veja um teste de impress&atilde;o do boleto<br><br></font><a href="javascript:Boleto()"><img border="0" src="<%=dirlingua%>/imagens/ver_boleto.gif"></a></td></tr>
</table>

  <tr><td width=50% valign=bottom>&nbsp;</td></tr>

<!--Dep�sito Banc�rio-->
<!--INICIO DESABILITAR DEPOSITO BANC�RIO-->
    <table width="100%" align="center" border="0">
     <tbody>
      <tr bgColor="#006666">
        <td width="100%" height="20" bgcolor="#f0f0f0"><font face="<%=fonte%>"  style=font-size:11px color=000000><strong><font style=font-size:15px><b>�</b></font> Dep&oacute;sito Banc&aacute;rio ou Transfer&ecirc;ncia</strong></font></td>	   
	  
      </tr>
      <tr vAlign="center">
       <td>
        <p><font face="<%=fonte%>" style=font-size:11px;color:000000><br>
        <img border="0" src="<%=dirlingua%>/imagens/pg_deposito.gif" width="87" height="44">
		&nbsp;&nbsp;&nbsp;
        <img src="<%=dirlingua%>/imagens/pg_transf.gif" border=0>
        <br><br>
        O pagamento atrav�s de <strong>Dep�sito Banc�rio,
        Tranfer�ncia Eletr�nica (entre contas) ou DOC</strong> s�o tamb�m uma forma pr�tica e segura para
        compras realizadas na Internet, pois os clientes n�o necessitam inserir
        dados como numero de cart�o de cr�dito ou senha.
		<br><br>
		A forma de pagamento por Transfer�ncia Eletr�nica consiste na simples trasnfer�ncia entre contas, seja entre contas do mesmo banco ou entre contas de bancos diferentes (chamado de DOC) , que voc� pode fazer pelo seu banco na Internet ou em qualquer ag�ncia banc�ria do seu banco.</font></p>
        <div align="justify">
          <ul>
            <li><font face="<%=fonte%>" style=font-size:11px;color:000000>As contas banc�rias aparecer�o na
              finaliza��o da compra, podendo ser feito o pagamento de uma das 3 formas acima, da maneira que for melhor para voc� no(s) banco(s) abaixo:</font><br><br>
			  <font face="<%=fonte%>" style=font-size:11px;color:blue>
			  - <%= Seu_banco %><br><br>
			  <% If Seu_banco2<>"" then %>
			  - <%= Seu_banco2 %>
			  <br><br>
			  <% End If %></font>
			  
			  
            <li><font face="<%=fonte%>" style=font-size:11px;color:000000>Os pedidos ser�o
              enviados ap�s confirma��o do pagamento, que pode ser feita via
              e-mail <a href=mailto:<%=emailloja%>><%=emailloja%></a>
              ou via fone/fax <%=fone11%>.*</font></li>
          </ul>

        </div>
       </td>
      </tr>

   <tr><td width=50% valign=bottom>&nbsp;</td></tr>

<!--FIM DESABILITAR DEPOSITO BANC�RIO-->
<!--Sedex a Cobrar-->
<!--INICIO DESABILITAR SEDEX-->

<% If entrega_sedex_a_cobrar="Sim" then %>
      <tr bgColor="#006666">
	  
        <td width="100%" height="20" bgcolor="#f0f0f0"><font face="<%=fonte%>"  style=font-size:11px color=000000><strong><font style=font-size:15px><b>�</b></font> Sedex a Cobrar</strong></font></td>		  

      </tr>
      <tr>
       <td vAlign="top">
        <p align="left"><br>
        <img border="0" src="./linguagens/portuguesbr/imagens/pg_sedexacobrar.gif" width="130" height="50"></p><br>
        <ul>
          <li><font align="left" face="<%=fonte%>" style=font-size:11px;color:000000>O pagamento � feito na
            retirada do (s) produto (s) nas ag�ncias dos <strong>Correios</strong>,
            com uma taxa diferenciada, que � cobrada pelos Correios.</font>
          <li><font face="<%=fonte%>" style=font-size:11px;color:000000>Voc� receber� um aviso quando seu
            (s) produto (s) estiver (em) dispon�vel para retirada na ag�ncia
            dos Correios mais pr�xima do endere�o informado no seu pedido.</font></li>
        </ul>
        <p><font face="<%=fonte%>" style=font-size:11px;color:000000>Uma&nbsp; forma&nbsp; segura de comprar
        na internet, haja visto que o pagamento do valor total � feito somente
        no ato do recebimento da mercadoria na ag�ncia dos correios.</font></p>
        <div align="justify">
          <ul>
            <li><font style=font-size:11px;color:000000><strong><font face="<%=fonte%>">Obs.</font></strong><font face="<%=fonte%>">
              O pagamento s� podera ser feito em <strong>Dinheiro</strong>, pois os
              correios n�o aceitam cheques</font></font></li>
            <li><font style=font-size:11px;color:000000><font face="<%=fonte%>">
              Tempo de entrega: 2 a 3 dia ap�s  libera��o</font></font></li>
          </ul>

          <p><font face="<%=fonte%>" style=font-size:11px;color:000000>&nbsp;</font></p>
        </div>
       </td>
      </tr>
	  <% End If %>
	  
     </tbody>
    </table>
<!--FIM DESABILITAR SEDEX-->
<!--FIM DO CONTEUDO-->
     <table border=0 cellspacing=0 width="100%" cellpadding=0>
         <tr><td height=5></td></tr>
         <tr><td height=1 bgcolor=<%=cor3%>></td></tr>
         <tr><td height=5></td></tr>
       </table>
	   <div align="center">
       <A HREF="javascript:history.back();" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='Voltar';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b>Voltar</b> ::</font></A></div></ul></td></tr>
   </table>

<!-- #include file="baixo.asp" -->
