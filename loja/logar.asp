<% Option Explicit %>
<!--#include file="_conexao.asp" -->
<%
'Criando vari�veis
Dim localbd, bd, usuario

'Abriremos a conex�o criada do include _conexao.asp
call abrir_conexao

'Criaremos um Recordset para selecionar os usu�rios cadastrados
set usuario=Server.CreateObject("ADODB.Recordset")

'Selecionar de acordo com o Login digitado no campo
usuario.Open "SELECT * from usuario where usuario='"& request.form("usuario") &"'", bd

'Se o usu�rio n�o for encontrado, iremos fechar a conex�o, dar um alert e voltar 
if usuario.EOF then
	call fechar_conexao
	response.write "<script>history.back(1);alert('Login incorreto. Tente novamente.')</script>"
else
	'Caso tenha achado o usu�rio, o sistema ir� verificar a senha
	if usuario("senha")=request.form("senha") then
		'Se a senha for v�lida, ele ir� gravar um cookie com o codigo do usuario
		response.cookies("access")("usuario")=usuario("cod")
		'Fecharemos a conex�o
		call fechar_conexao	
		'Redirecionaremos para a p�gina principal
		response.redirect("tabela.asp")
	else
	'Caso n�o valide a senha, ser� dada uma mensagem de senha incorreta e voltar�
		call fechar_conexao
		response.write "<script>history.back(1);alert('Senha incorreta. Tente novamente.')</script>"
	end if
end if
%>
