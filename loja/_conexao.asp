<%
'Criaremos uma SUB para a conex�o para conectarmos com o banco de dados do AccessAdmin
'Detalhe: N�o � o banco que ser� administrado
sub abrir_conexao	
	localbd = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("db/configaccessadmin.mdb")
	set bd=Server.CreateObject("ADODB.Connection")
	bd.open localbd
end sub

'SUB que fechar� a conex�o
sub fechar_conexao
	bd.close
	Set bd = nothing
end sub
%>