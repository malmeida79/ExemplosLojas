Imports System.Data
Imports System.Data.SqlClient

Public Class Endereco

#Region " Declarações "

    Dim codEnd As Integer
    Dim cep As String
    Dim sigla As Integer
    Dim endereco As String
    Dim bairro As Integer
    Dim cidade As Integer
    Dim estado As Integer
    Dim listaEnd As New DataSet()
    Dim totalLinhasEnd As Integer

    ' instancias 
    Dim cnn As New Conexao(".\sqlexpress", "dbLojaInfo")

#End Region

#Region " Metodos Input / Output "


    Public Sub setCep(ByVal cp As String)
        cep = cp
    End Sub

    Public Function getCep() As String
        getCep = cep
    End Function

    Public Sub setEndereco(ByVal ed As String)
        endereco = ed
    End Sub

    Public Function getEndereco() As String
        getEndereco = endereco
    End Function

    Public Sub setSigla(ByVal sg As Integer)
        sigla = sg
    End Sub

    Public Function getSigla() As Integer
        getSigla = sigla
    End Function

    Public Sub setBairro(ByVal br As Integer)
        bairro = br
    End Sub

    Public Function getBairro() As Integer
        getBairro = bairro
    End Function

    Public Sub setCidade(ByVal cid As Integer)
        cidade = cid
    End Sub

    Public Function getCidade() As Integer
        getCidade = cidade
    End Function

    Public Sub setEstado(ByVal es As Integer)
        estado = es
    End Sub

    Public Function getEstado() As Integer
        getEstado = estado
    End Function

    Public Function getTotLinhasEnd() As Integer
        getTotLinhasEnd = totalLinhasEnd
    End Function

    Public Function getCodEnd() As Integer
        getCodEnd = codEnd
    End Function

#End Region

#Region " Ações "

    Public Sub consultaCep()

        'connecta o banco
        cnn.conecta()

        'alimenta o dataAdapter
        cnn.setComandoAdapter("select * from cadastro_endereco where endCep='" & cep & "'")

        'nomeia a tabela
        cnn.setTabela("endereco")

        'executa busca
        listaEnd = cnn.buscaDs()

        'retorna total de linhas 
        totalLinhasEnd = listaEnd.Tables("Endereco").Rows.Count

    End Sub

    Public Sub consultaId(ByVal idend As Integer)

        'connecta o banco
        cnn.conecta()

        'alimenta o dataAdapter
        cnn.setComandoAdapter("select * from cadastro_endereco where idEndereco=" & idend)

        'nomeia a tabela
        cnn.setTabela("endereco")

        'executa busca
        listaEnd = cnn.buscaDs()

        'retorna total de linhas 
        totalLinhasEnd = listaEnd.Tables("Endereco").Rows.Count

    End Sub

    Public Sub carregaEndereco()

        If totalLinhasEnd > 0 Then
            codEnd = listaEnd.Tables("endereco").Rows(0)(0).ToString()
            cep = listaEnd.Tables("endereco").Rows(0)(1).ToString()
            endereco = listaEnd.Tables("endereco").Rows(0)(2).ToString()
            sigla = listaEnd.Tables("endereco").Rows(0)(3).ToString()
            bairro = listaEnd.Tables("endereco").Rows(0)(4).ToString()
            cidade = listaEnd.Tables("endereco").Rows(0)(5).ToString()
            estado = listaEnd.Tables("endereco").Rows(0)(6).ToString()
        End If

    End Sub

    Public Function carregaCidade() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from cadastro_cidade")
        cnn.iniciaCmd()
        carregaCidade = cnn.buscaDr
    End Function

    Public Function carregaEstado() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from cadastro_estado")
        cnn.iniciaCmd()
        carregaEstado = cnn.buscaDr
    End Function

    Public Function carregaBairro() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from cadastro_bairro")
        cnn.iniciaCmd()
        carregaBairro = cnn.buscaDr
    End Function

    Public Function carregaSigla() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from cadastro_sigla")
        cnn.iniciaCmd()
        carregaSigla = cnn.buscaDr
    End Function

    Public Sub cadastraendereco()

        cnn.conecta()
        cnn.iniciaCmd("spc_cad_endereco")
        cnn.addParametro("@endCep", cep, "i")
        cnn.addParametro("@idSigla", sigla, "i")
        cnn.addParametro("@endEndereco", endereco, "i")
        cnn.addParametro("@idBairro", bairro, "i")
        cnn.addParametro("@idCidade", cidade, "i")
        cnn.addParametro("@IdEstado", estado, "i")
        cnn.addParametro("@codigo", codEnd, "o")

        'cadastra novo resgistro
        cnn.executaComando()

        'capturando retorno de cadastro (capturando antes de matar o campo)
        codEnd = cnn.retParInt("@codigo")
        cnn.encerra()

    End Sub

#End Region

End Class
