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
'#
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
%>
<!--#include file="config.asp"-->
<!--#include file="adm_restrito.asp"-->
<%
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Aten��o: Este arquivo n�o pode ser distrubu�do ou separado desta VirtuaStore OPEN
'Colabora��o: Elizeu Alcantara - elizeu@onda.com.br / elizeu@cristaosite.com.br
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

response.buffer = true
response.expires = 10

if request.form("acao") = "ExecSQL" then
	Dim Conn
	set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open StringdeConexao 'Ponha aqui sua string de Conex�o
	executar_sql=request.form("SQL")
	set qrSQL = Conn.Execute(executar_sql)
	
	strArea_transferencia = Request.form("area_transferencia")
	'grava a area de transferencia em cookies
	if strArea_transferencia<>"" then
	response.cookies("AreaTransferenciaSql")= strArea_transferencia
	response.cookies("AreaTransferenciaSql").Expires= #January 01, 2010#
	end if  
end if
%>
<html>
<head>
<title><%= Nome_da_sua_loja %> - Administra��o On-line</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<% area_transferencia = Request.Cookies("AreaTransferenciaSql") %>

<body bgcolor="#FFFFFF">

<form name="formSQL" method="post" action="">
  <input type="hidden" name="acao" value="ExecSQL">
  <table width="80%" border="0" cellspacing="0" cellpadding="0">
  
      <tr align="center"> 
      <td > 
        <p><font face=tahoma style=font-size:11px>Execu&ccedil;&atilde;o de Comandos SQL</font></p>
      </td>
      <td > 
        <p><font face=tahoma style=font-size:11px>�rea de Transfer�ncia</font></p>
      </td>	  
    </tr>
  
    <tr align="center"> 
      <td > 
        <textarea name="SQL" rows="8" cols="70" style="font-family:tahoma; font-size: 11px; color:003366"><%= executar_sql %>
</textarea>
      </td>
      <td > 
        <textarea name="area_transferencia" rows="8" cols="60" style="font-family:tahoma; font-size: 11px; color:003366"><%= area_transferencia %>
</textarea>
      </td>	  
    </tr>
	<tr align="center"> 
      <td > 
        <input type=submit value="      Executar      " style="cursor:hand;font-family:tahoma; font-size: 11px;" onclick="Javascript: document.forms['formSQL'].submit()">
      </td>
      <td > &nbsp;  </td>	  
    </tr>
	

  </table>
  <br><br>
<%if UCase(Mid(request.form("SQL"),1,6)) = "SELECT" then%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#cccccc" bordercolordark="#cccccc">
  <tr bgcolor="#f5f5f5">
 <%for i = 0 to qrSQL.Fields.Count - 1%> 
    <td><font face=tahoma style="font-size:11px; color:003366"><%=qrSQL.Fields(i).name%>&nbsp;</font></td>
 <%next%>
    </tr>
 <%while not qrSQL.EOF%>
     <tr>
  <%for i = 0 to qrSQL.Fields.Count - 1%> 
     <td><font face=tahoma style=font-size:11px><%=qrSQL.Fields(i)%>&nbsp;</font></td>
  <%next%>
     </tr>
 <%qrSQL.MoveNext%>
 <%wend%>
  </table>
<%
Conn.close
set qrSQL = nothing
set Conn = nothing
%>
<%end if%>

</body>
</html>
