Imports LojaInfoDll

Public Class clsCliente

    Inherits Endereco

#Region " Variaveis e instancias "

    Dim cnn As New Conexao(".\sqlexpress", "dbLojaInfo")

    Dim codigo As Integer
    Dim nome As String
    Dim fantasia As String
    Dim cnpj As String
    Dim numero As Integer
    Dim parametro As String
    Dim totalLinhasCli As Integer
    Dim listaCli As DataSet

#End Region

#Region " Metodos de Input / output "

    Public Sub setParametro(ByVal par As String)
        parametro = par
    End Sub

    Public Sub setNome(ByVal nm As String)
        nome = nm
    End Sub

    Public Function getNome() As String
        getNome = nome
    End Function

    Public Sub setFantasia(ByVal ft As String)
        fantasia = ft
    End Sub

    Public Function getFantasia() As String
        getFantasia = fantasia
    End Function

    Public Sub setCnpj(ByVal pj As String)
        cnpj = pj
    End Sub

    Public Function getCnpj() As String
        getCnpj = cnpj
    End Function

    Public Sub setNumero(ByVal num As String)
        numero = num
    End Sub

    Public Function getNumero() As String
        getNumero = numero
    End Function

    Public Function getTotalLinhasCli() As Integer
        getTotalLinhasCli = totalLinhasCli
    End Function

#End Region

#Region " Ações "

    Public Sub cadastrar()

        cnn.conecta()
        cnn.iniciaCmd("spc_cad_cliente")
        cnn.addParametro("@cliNome", nome, "i")
        cnn.addParametro("@cliFantasia", fantasia, "i")
        cnn.addParametro("@cliCnpj", cnpj, "i")
        cnn.addParametro("@idEndereco", getCodEnd(), "i")
        cnn.addParametro("@cliNumero", numero, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub salvar()

        cnn.conecta()
        cnn.iniciaCmd("spc_alt_cliente")
        cnn.addParametro("@idCliente", codigo, "i")
        cnn.addParametro("@cliNome", nome, "i")
        cnn.addParametro("@cliFantasia", fantasia, "i")
        cnn.addParametro("@cliCnpj", cnpj, "i")
        cnn.addParametro("@idEndereco", getCodEnd(), "i")
        cnn.addParametro("@cliNumero", numero, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub apagar()

        cnn.conecta()
        cnn.iniciaCmd("spc_del_cliente")
        cnn.addParametro("@idCliente", codigo, "i")
        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub buscar()


        'connecta o banco
        cnn.conecta()

        'alimenta o dataAdapter
        cnn.setComandoAdapter("select * from cadastro_cliente where cliNome like '%" & parametro & "%'")

        'nomeia a tabela
        cnn.setTabela("cliente")

        'executa busca
        listaCli = cnn.buscaDs()

        'retorna total de linhas 
        totalLinhasCli = listaCli.Tables("Cliente").Rows.Count

    End Sub

    Public Sub navega(ByVal pos As Integer)

        If totalLinhasCli > 0 Then
            codigo = listaCli.Tables("cliente").Rows(pos)(0).ToString()
            nome = listaCli.Tables("cliente").Rows(pos)(1).ToString()
            fantasia = listaCli.Tables("cliente").Rows(pos)(2).ToString()
            cnpj = listaCli.Tables("cliente").Rows(pos)(3).ToString()
            numero = Convert.ToInt32(listaCli.Tables("cliente").Rows(pos)(5).ToString())
            consultaId(Convert.ToInt32(listaCli.Tables("cliente").Rows(pos)(4).ToString()))
            carregaEndereco()
        End If

    End Sub

#End Region

End Class
