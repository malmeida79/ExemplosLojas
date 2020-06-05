<%
'Criaremos uma SUB para a conexão para conectarmos com o banco de dados do AccessAdmin
'Detalhe: Não é o banco que será administrado
sub abrir_conexao	
	localbd = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("db/configaccessadmin.mdb")
	set bd=Server.CreateObject("ADODB.Connection")
	bd.open localbd
end sub

'SUB que fechará a conexão
sub fechar_conexao
	bd.close
	Set bd = nothing
end sub
%>