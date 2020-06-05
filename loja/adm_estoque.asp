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
'#  // "Powered by VirtuaStore OPEN" ou "Software derivado de VirtuaStore 1.2" e 
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


'#########################################################################################
'# Consulta posi��o de estoque dos produtos da loja										 #
'# produtos que tenham estoque maior ou menor que o informado pelo usuario               #
'# CRIADO/MELHORADO por: F�bio V.Coelho 												 #
'# AJUDA: Rogerio 																		 #
'#########################################################################################

%><!--#include file="adm_restrito.asp"--><%

total_vendido=0
subtotal_vendido=0

set conn = Server.CreateObject("ADODB.Connection")
conn.Open(StringdeConexao)

 strRegistro = " SELECT produtos.nome, produtos.preco, produtos.vendas, estoque.idproduto, estoque.estoque FROM ESTOQUE, PRODUTOS where produtos.idprod=estoque.idproduto order by estoque.estoque desc" 
 set rs= conn.execute(strRegistro) 

   SQL = "pr.nome, es.idproduto, es.estoque "
   SQL = SQL & " produtos pr, estoque es" 
   SQL = SQL & " where es.estoque like '0' pr.idprod=es.idproduto" 

				strTextoHtml = strTextoHtml & "<FONT face=tahoma color=#003366 style=font-size:17px><img src=""adm_imagens/prod_g.gif"" width=""30"" height=""33"" hspace=""1"" vspace=""1"" border=""0"" align=""left""><B>&nbsp;Estoque Atual e Vendas Efetuadas dos produtos na loja</B></FONT><hr size=1 color=#cccccc><br>"

IF NOT RS.EOF THEN 
 strTextoHtml = strTextoHtml & "<br><table width=100% cellspacing=""0"" cellpadding=""0"">"
DO WHILE NOT RS.EOF 
  strTextoHtml = strTextoHtml & "<tr><td colspan=3 align=left><hr size=1 color=#cccccc width=""101%""></td></tr>"
  strTextoHtml = strTextoHtml & "<tr><td colspan=1 align=left><font face=""tahoma"" style=font-size:11px>" & RS("nome") 

 	if RS("ESTOQUE")<4 then
  strTextoHtml = strTextoHtml & "  </font></td><td colspan=1 nowrap><font face=""tahoma"" style=font-size:11px > Vendas: " &  RS("vendas") & " &nbsp;</font>"
  strTextoHtml = strTextoHtml & "  </font></td><td colspan=1 nowrap><font face=""tahoma"" style=font-size:11px color=red><strong> Estoque: " &  RS("ESTOQUE") & " &nbsp;</strong></font>"
  strTextoHtml = strTextoHtml & "  </font></td><td colspan=1 nowrap><font face=""tahoma"" style=font-size:11px > <a href=?link=produtos&acao=edita&prod=" & rs("idproduto") & "><strong>Colocar Estoque</strong></a></font><br></td></tr>"
	else
  strTextoHtml = strTextoHtml & "  </font></td><td colspan=1 nowrap><font face=""tahoma"" style=font-size:11px > Vendas: " &  RS("vendas") & " </font>"
  strTextoHtml = strTextoHtml & "  </font></td><td colspan=1 nowrap><font face=""tahoma"" style=font-size:11px> Estoque: " &  RS("ESTOQUE") & " </font>"
  strTextoHtml = strTextoHtml & "  </font></td><td colspan=1 nowrap><font face=""tahoma"" style=font-size:11px> <a href=?link=produtos&acao=edita&prod=" & rs("idproduto") & "><strong>Alterar Estoque</strong></a></font><br></td></tr>"
	end if

  subtotal_vendido=rs("vendas")*rs("preco") 
  total_vendido=total_vendido+subtotal_vendido

RS.MOVENEXT 
LOOP 
  strTextoHtml = strTextoHtml & "<tr><td colspan=3 align=left><hr size=1 color=#cccccc width=""101%""></td></tr>"
  strTextoHtml = strTextoHtml & "</table>"
  

ELSE 
strTextoHtml = strTextoHtml & "<br><font face=""tahoma"" style=font-size:11px color=red><strong>N�o existem produtos na sua loja! </font><br>" 
END IF

Conn.close
set Conn = nothing

strTextoHtml = strTextoHtml & "<br><hr size=1 color=dddddd><center><FONT face=tahoma  style=font-size:11px><img src=""adm_imagens/mao.gif"" width=""26"" height=""28"" hspace=""1"" vspace=""1"" border=""0"" align=""absmiddle"" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>A loja j� efetuou "& strLgMoeda & " " & formatnumber(total_vendido,2)&" nas vendas acima</b></center>" 

strTextoHtml = strTextoHtml & "<br><hr size=1 color=dddddd><center><FONT face=tahoma  style=font-size:11px><img align=absmiddle border=0 src=""adm_imagens/xis.gif"" > &nbsp;&nbsp;&nbsp;<b><A onclick=""return confirm('Voc� quer zerar a contagem de vendas?')"" href=""?link=util&acao=vendas&acaba=sim"">Voc� quer Zerar a Contagem das Vendas acima?</A></b></center><br>" 

%>

