<%
'Daremos um valor vazio ao cookie
response.Cookies("access")("usuario")=""
'Retornaremos para a p�gina de logar
response.redirect "painel.asp"
%>