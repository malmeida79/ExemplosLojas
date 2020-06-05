<%



'Desloga o usuario
response.buffer = true
session.abandon
response.redirect "default.asp"%>