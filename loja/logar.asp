<% Option Explicit %>
<!--#include file="_conexao.asp" -->
<%
'Criando variáveis
Dim localbd, bd, usuario

'Abriremos a conexão criada do include _conexao.asp
call abrir_conexao

'Criaremos um Recordset para selecionar os usuários cadastrados
set usuario=Server.CreateObject("ADODB.Recordset")

'Selecionar de acordo com o Login digitado no campo
usuario.Open "SELECT * from usuario where usuario='"& request.form("usuario") &"'", bd

'Se o usuário não for encontrado, iremos fechar a conexão, dar um alert e voltar 
if usuario.EOF then
	call fechar_conexao
	response.write "<script>history.back(1);alert('Login incorreto. Tente novamente.')</script>"
else
	'Caso tenha achado o usuário, o sistema irá verificar a senha
	if usuario("senha")=request.form("senha") then
		'Se a senha for válida, ele irá gravar um cookie com o codigo do usuario
		response.cookies("access")("usuario")=usuario("cod")
		'Fecharemos a conexão
		call fechar_conexao	
		'Redirecionaremos para a página principal
		response.redirect("tabela.asp")
	else
	'Caso não valide a senha, será dada uma mensagem de senha incorreta e voltará
		call fechar_conexao
		response.write "<script>history.back(1);alert('Senha incorreta. Tente novamente.')</script>"
	end if
end if
%>
