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
     <tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='Home';return true;"><b>Home</b></a> � Como Reimprimir o Boleto Banc�rio?<br><br><div align=left>
       <table border=0 cellspacing=0 width="100%" cellpadding=0>
         <tr><td height=5></td></tr>
         <tr><td height=1 bgcolor=<%=cor3%>></td></tr>
         <tr><td height=5></td></tr>
       </table>
       </div>
    <br><br><font face="<%=fonte%>" style=font-size:13px><strong>Como reimprimir o Boleto na <%=nomeloja%></strong></font><br><br>
    <div align=center>
<!--CONTEUDO-->
<!--Personaliza o menu se o cliente estiver logado-->
<% if session("usuario") = "" then %>

    <table width="100%" align="center" border="0">
      <tr bgColor="#006666">
        <td width="100%" height="20" bgcolor="#f0f0f0"><font face="<%=fonte%>"  style=font-size:11px color=000000><strong>Roteiro</strong></font></td>        
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
        <td><font face="<%=fonte%>" style=font-size:11px>
           <ul>1. Clicar no bot&atilde;o <A href="fechapedido.asp?compra=login&menu=sim"><IMG src="<%= dirlingua %>/imagens/botao_superior_login.gif" border=0 target="home"></A></ul>
           <ul>2. Em seguida aparecer&aacute; o bot&atilde;o <A href="minhascompras.asp"><IMG src="<%= dirlingua %>/imagens/botao_superior_historico_compras.gif" border=0 target="home"></A> clique nele !</ul>
           <ul>3. Na p&aacute;gina de <%=strLg119%> voc&ecirc;, clicar no bot&atilde;o <img border="0" src="<%=dirlingua%>/imagens/detalhes.gif"></ul>
           <ul>4. Em detalhes da Compra, clicar no bot&atilde;o <img border="0" src="<%=dirlingua%>/imagens/2via_boleto.gif"></ul>
        </td></font>
      </tr>
    </table>

<%else%>

    <table width="100%" align="center" border="0">
      <tr bgColor="#006666">
        <td width="100%" height="20" bgcolor="#f0f0f0"><font face="<%=fonte%>"  style=font-size:11px color=000000><strong>Roteiro</strong></font></td>        
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
        <td><font face="<%=fonte%>" style=font-size:11px>
           <ul>1. Clicar no bot&atilde;o <A href="minhascompras.asp"><IMG src="<%= dirlingua %>/imagens/botao_superior_historico_compras.gif" border=0 target="home"></A></ul>
           <ul>2. Na p&aacute;gina de <%=strLg119%>&nbsp;&nbsp;<%=strNome%>, clicar no bot&atilde;o <img border="0" src="<%=dirlingua%>/imagens/detalhes.gif"></ul>
           <ul>3. Em detalhes da Compra, clicar no bot&atilde;o <img border="0" src="<%=dirlingua%>/imagens/2via_boleto.gif"></ul>
        </td></font>
      </tr>
    </table>

<%end if%>
<br>
<br>
<br>
<!--FIM DO CONTEUDO-->
       <table border=0 cellspacing=0 width="100%" cellpadding=0>
         <tr><td height=5></td></tr>
         <tr><td height=1 bgcolor=<%=cor3%>></td></tr>
         <tr><td height=5></td></tr>
       </table>
       <A HREF="javascript:history.back();" style=text-decoration:none onMouseOut="window.status='';return true;" onMouseOver="window.status='Voltar';return true;"><font face="<%=fonte%>" style=font-size:11px>:: <b>Voltar</b> ::</font></A></div></ul></td></tr>
   </table></td></tr>
</table>
<!-- #include file="baixo.asp" -->