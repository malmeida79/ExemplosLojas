<%
'#################################################################
'INICIO - Rotina do Contador de acessos 
'Este contador inicia uma nova visita a cada sess�o de compras iniciada na loja
'Criado em 5/11/2003 por Ot�vio Dias
'#################################################################
'--------------------------------------------------------------------------------
path_do_arquivo = Replace(Server.MapPath("topo.asp"), "topo.asp", "")
Const forReading = 1, forWriting = 2, forAppending = 8
Const TriDef = -2, TriTrue = -1, TriFalse = 0
arquivo = path_do_arquivo & "contador.log"
'--------------------------------------------------------------------------------
'a linha abaixo abre a inst�ncia com o objeto Scripting. FileSystemObject
'--------------------------------------------------------------------------------
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
'--------------------------------------------------------------------------------
'abaixo, � feita a verifica��o da exist�ncia do arquivo procurado. Caso ele ainda n�o exista (o que ocorre 1 vez por dia, no primeiro acesso), ele � criado
'--------------------------------------------------------------------------------
If ObjFSO.FileExists(arquivo) = False then
	'--------------------------------------------------------------------------------
	'abaixo, a linha que cria o arquivo LOG especificado
	'--------------------------------------------------------------------------------
	objFSO.CreateTextFile(arquivo)
	Set objStream = objFSO.OpenTextFile(arquivo, 2, True, False)
	ObjStream.WriteLine 0
	ObjStream.close 
	Set ObjStream = nothing
End If
'--------------------------------------------------------------------------------
'a linha abaixo abre o arquivo desejado. Lembre-se, ou ele j� existe ou ele foi criado na rotina acima.
'--------------------------------------------------------------------------------
Set ObjFile = objFSO.GetFile(arquivo)
'--------------------------------------------------------------------------------
'a linha abaixo diz o tipo de manipula��o que ser� utilizada no arquivo LOG, no caso � para adicionar dados.
'--------------------------------------------------------------------------------
Set objStream = objFSO.OpenTextFile(arquivo, 8, True, False)
Set objRead = objFSO.OpenTextFile(arquivo, forReading, TriTrue) 
	Do While Not objRead.AtEndOfStream
	UltimoNumero = Clng(objRead.readline) + 1
	Loop
'--------------------------------------------------------------------------------
objRead.close 
Set objRead = nothing
'--------------------------------------------------------------------------------
If Session("RotinaAccess") <> "Sim" then
	'--------------------------------------------------------------------------------
	'o comando WriteLine, abaixo, grava os dados no arquivo LOG especificado.
	'--------------------------------------------------------------------------------
	ObjStream.WriteLine UltimoNumero
	Session("RotinaAccess") = "Sim" 
end if
'--------------------------------------------------------------------------------
'abaixo, o objeto ObjStream � fechado
'--------------------------------------------------------------------------------
ObjStream.close 
Set ObjStream = nothing
'--------------------------------------------------------------------------------
If Session("RotinaAccess2") <> "Sim" then
Set objStream = objFSO.OpenTextFile(arquivo, 2, True, False)
ObjStream.WriteLine UltimoNumero
ObjStream.close 
	Set ObjStream = nothing
		Session("RotinaAccess2") = "Sim" 
end if
'#################################################################
'FIM - Rotina do Contador de acessos
'#################################################################
%>