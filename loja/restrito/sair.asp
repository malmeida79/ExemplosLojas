<%
'Daremos um valor vazio ao cookie
response.Cookies("access")("usuario")=""
'Retornaremos para a página de logar
response.redirect "painel.asp"
%>