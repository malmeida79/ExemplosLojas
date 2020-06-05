<%
'Criando variáveis
Dim localbd, bd, verificar_usuario, varCookie

'Se o cookie for vazio iremos dar um valor ZERO para não dar erro
'Se não irá setar o valor da variável com o valor do cookie
if request.cookies("access")("usuario")="" then
	varCookie=0
else
	varCookie=request.cookies("access")("usuario")
end if

'Abrir conexão
call abrir_conexao

'Criaremos um Recordset para verificar se o Codigo do Cookie existe no banco de dados
set verificar_usuario=Server.CreateObject("ADODB.Recordset")

'Selecionar o usuário
verificar_usuario.Open "SELECT cod from usuario where cod="&varCookie&"", bd

'Se o usuário não existir, fecharemos a conexão e redirecionaremos para a página de logar
if verificar_usuario.EOF then
	call fechar_conexao
	response.redirect "restrito.asp"
end if

call fechar_conexao
%>