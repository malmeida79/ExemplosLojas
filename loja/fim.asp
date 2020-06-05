<%



'Desloga o usuario
response.buffer = true
session("usuario") = request.form("usuarioz")
response.redirect "default.asp"
%>