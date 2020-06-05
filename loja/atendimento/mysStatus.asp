<!--#include file="db.asp"-->
<!--#include file="mysConfiguracoes.asp"-->
<%
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.addHeader "pragma","no-cache"
	Response.addHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	Session.LCID = 1033
%>
<%
Call AbreDB
Dim strImagem

%>
<html>
<SCRIPT LANGUAGE="JavaScript">
<!--
function abre_popup(width, height, nome) {
  var top; var left;
  top = ( (screen.height/2) - (height/2) )
  left = ( (screen.width/2) - (width/2) )
  window.open('',nome,'width='+width+',height='+height+',scrollbars=no,toolbar=no,location=no,status=no,menubar=no,resizable=no,left='+left+',top='+top);
}
function atualiza_pagina(){
  document.location.href=document.location.href
}
atualiza = setTimeout("atualiza_pagina()",10000);
//-->
</SCRIPT>

<%
'Verifica se existe algum online
strSQL = "SELECT count(id) AS Total FROM operadores WHERE  DateDiff(""s"", ping, now()) < 10"

Set rs = Conexao.execute(strSQL)

If rs("Total") = "0" OR isNull(rs("Total")) OR rs("Total") = 0 Then
	strImagem = strConfigOffline
Else
	strImagem = strConfigOnline
End If

rs.Close

Call FechaDB
%>
</head>
<script language="JavaScript">
function abrirChat(){
	var width = 450;
	var height = 350;
	var esquerda = (screen.width/2) - (width/2);
	var topo = (screen.height/2) - (height/2);
	window.open("<%=strConfigEndereco%>/mysAtendimento.asp","MySupport","top="+topo+",left="+esquerda+",width="+width+",height="+height+",scrollbars=yes, menu=0,status=yes");
}
document.write("<a href='javascript:abrirChat();' title='Atendimento'>");
document.write("<img border='0' src='<%=strImagem%>'>");
document.write("</a>");
</script>