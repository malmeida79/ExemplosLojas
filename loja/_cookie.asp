<%
'Criando vari�veis
Dim localbd, bd, verificar_usuario, varCookie

'Se o cookie for vazio iremos dar um valor ZERO para n�o dar erro
'Se n�o ir� setar o valor da vari�vel com o valor do cookie
if request.cookies("access")("usuario")="" then
	varCookie=0
else
	varCookie=request.cookies("access")("usuario")
end if

'Abrir conex�o
call abrir_conexao

'Criaremos um Recordset para verificar se o Codigo do Cookie existe no banco de dados
set verificar_usuario=Server.CreateObject("ADODB.Recordset")

'Selecionar o usu�rio
verificar_usuario.Open "SELECT cod from usuario where cod="&varCookie&"", bd

'Se o usu�rio n�o existir, fecharemos a conex�o e redirecionaremos para a p�gina de logar
if verificar_usuario.EOF then
	call fechar_conexao
	response.redirect "restrito.asp"
end if

call fechar_conexao
%>