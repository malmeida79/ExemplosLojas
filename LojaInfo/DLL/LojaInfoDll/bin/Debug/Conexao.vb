Imports System.Data.SqlClient

Public Class Conexao

#Region " Variaveis "

    ' Variaveis padrao para conexao
    Dim cn As SqlConnection
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim cmd As SqlCommand

    'metodos para manipulação de dados
    Dim cnstr As String
    Dim sqlCn As String
    Dim sqlDa As String
    Dim sqlCmd As String
    Dim tabela As String
    Dim nomeProc As String

#End Region

#Region " Constructors "
    Public Sub New()
        'constructor vazio para garantir a instancia do metodo puro
    End Sub

    Public Sub New(ByVal server As String, ByVal banco As String)
        cnstr = "Server=" & server & ";Integrated Security=SSPI;DataBase=" & banco
    End Sub

    Public Sub New(ByVal cnx As String)
        cnstr = cnx
    End Sub

#End Region

#Region " Manipulador de  Conexao "
    'conecta ao banco
    Public Sub conecta()
        cn = New SqlConnection(cnstr)
        cn.Open()
    End Sub

    'conecta ao banco
    Public Sub conecta(ByVal server As String, ByVal banco As String)
        cn = New SqlConnection("Server=" & server & ";Integrated Security=SSPI;DataBase=" & banco)
        cn.Open()
    End Sub

    Public Sub encerra()
        'mata conexao
        cn.Close()
    End Sub

#End Region

#Region " Metodos de Entrada "

    'recebe tabela
    Public Sub setTabela(ByVal tbl As String)
        tabela = tbl
    End Sub

    'recebe comando para o adapter
    Public Sub setComandoAdapter(ByVal cmd As String)
        sqlDa = cmd
    End Sub

    'utilizado para comands e datareaders
    Public Sub setComando(ByVal cmd As String)
        sqlCmd = cmd
    End Sub

#End Region

#Region " Manipulador de Dados "
    'prepara comando para ser executado como TEXT
    Public Sub iniciaCmd()
        cmd = cn.CreateCommand
        cmd.CommandText = sqlCmd
        cmd.CommandType = CommandType.Text
    End Sub

    'prepara comando para ser executado como PROCEDURE
    Public Sub iniciaCmd(ByVal spc As String)
        cmd = cn.CreateCommand
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = spc
    End Sub

    'executando comandos no banco de dados 
    Public Sub executaComando()
        cmd.ExecuteNonQuery()
    End Sub

    'monta dataset com dados da consulta passada no adapter
    Public Function buscaDs() As DataSet
        buscaDs = New DataSet()
        da = New SqlDataAdapter(sqlDa, cn)
        da.Fill(buscaDs, tabela)
    End Function

    'executa data reader
    Public Function buscaDr() As SqlDataReader
        buscaDr = cmd.ExecuteReader
    End Function

#End Region

#Region " Entrada de parametros "

    ' passagem de parametros string para o comand
    Public Sub addParametro(ByVal nome As String, ByVal par As String, ByVal direcao As String)

        'comando que leva o parametro ao command
        cmd.Parameters.Add(New SqlParameter(nome, par))

        'controle de direção de dados dos parametros
        If direcao = "i" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Input
        ElseIf direcao = "io" Then
            cmd.Parameters(nome).Direction = ParameterDirection.InputOutput
        ElseIf direcao = "v" Then
            cmd.Parameters(nome).Direction = ParameterDirection.ReturnValue
        ElseIf direcao = "o" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Output
        End If
    End Sub

    ' passagem de parametros integer para o comand
    Public Sub addParametro(ByVal nome As String, ByVal par As Integer, ByVal direcao As String)
        cmd.Parameters.Add(New SqlParameter(nome, par))
        If direcao = "i" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Input
        ElseIf direcao = "io" Then
            cmd.Parameters(nome).Direction = ParameterDirection.InputOutput
        ElseIf direcao = "v" Then
            cmd.Parameters(nome).Direction = ParameterDirection.ReturnValue
        ElseIf direcao = "o" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Output
        End If
    End Sub

    ' passagem de parametros decimal para o comand
    Public Sub addParametro(ByVal nome As String, ByVal par As Decimal, ByVal direcao As String)
        cmd.Parameters.Add(New SqlParameter(nome, par))
        If direcao = "i" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Input
        ElseIf direcao = "io" Then
            cmd.Parameters(nome).Direction = ParameterDirection.InputOutput
        ElseIf direcao = "v" Then
            cmd.Parameters(nome).Direction = ParameterDirection.ReturnValue
        ElseIf direcao = "o" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Output
        End If
    End Sub

    ' passagem de parametros boolean para o comand
    Public Sub addParametro(ByVal nome As String, ByVal par As Boolean, ByVal direcao As String)
        cmd.Parameters.Add(New SqlParameter(nome, par))
        If direcao = "i" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Input
        ElseIf direcao = "io" Then
            cmd.Parameters(nome).Direction = ParameterDirection.InputOutput
        ElseIf direcao = "v" Then
            cmd.Parameters(nome).Direction = ParameterDirection.ReturnValue
        ElseIf direcao = "o" Then
            cmd.Parameters(nome).Direction = ParameterDirection.Output
        End If
    End Sub

#End Region

#Region " Parametros de retorno "

    'retorna parametro Integer
    Public Function retParInt(ByVal nome As String) As Integer
        retParInt = CType(cmd.Parameters(nome).Value, Integer)
    End Function

    'retorna parametro String
    Public Function retParStr(ByVal nome As String) As String
        retParStr = CType(cmd.Parameters(nome).Value, String)
    End Function

    'retorna parametro Decimal
    Public Function retParDec(ByVal nome As String) As Decimal
        retParDec = CType(cmd.Parameters(nome).Value, Decimal)
    End Function

    'retorna parametro Boolean
    Public Function retParBool(ByVal nome As String) As Boolean
        retParBool = CType(cmd.Parameters(nome).Value, Boolean)
    End Function

#End Region

End Class
