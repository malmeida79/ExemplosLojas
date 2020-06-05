Imports LojaInfodll

Public Class clsEstoque

#Region " Variaveis e instancias "

    Dim cnn As New Conexao(".\sqlexpress", "dblojaInfo")

    Dim codigo As Integer
    Dim Estoque As String
    Dim lista As DataSet
    Dim parametro As String
    Dim totalLinhas As Integer

#End Region

#Region " Metodos de Input / output "

    Public Sub setParametro(ByVal par As String)
        parametro = par
    End Sub

    Public Sub setEstoque(ByVal nmEstq As String)
        Estoque = nmEstq
    End Sub

    Public Function getEstoque() As String
        getEstoque = Estoque
    End Function


    Public Function getTotalLinhas() As Integer
        getTotalLinhas = totalLinhas
    End Function

#End Region

#Region " Ações "

    Public Sub cadastrar()

        cnn.conecta()
        cnn.iniciaCmd("spc_cad_estq")
        cnn.addParametro("@estqDesc", Estoque, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub salvar()

        cnn.conecta()
        cnn.iniciaCmd("spc_alt_estq")
        cnn.addParametro("@estqId", codigo, "i")
        cnn.addParametro("@estqDesc", Estoque, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub apagar()

        cnn.conecta()
        cnn.iniciaCmd("spc_del_estq")
        cnn.addParametro("@estqId", codigo, "i")
		
        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub buscar()


        'connecta o banco
        cnn.conecta()

        'alimenta o dataAdapter
        cnn.setComandoAdapter("select * from tblEstoque where estqDesc like '%" & parametro & "%'")

        'nomeia a tabela
        cnn.setTabela("estq")

        'executa busca
        lista = cnn.buscaDs()

        'retorna total de linhas 
        totalLinhas = lista.Tables("estq").Rows.Count

    End Sub

    Public Sub carregar(ByVal pos As Integer)

        If totalLinhas > 0 Then
            codigo = lista.Tables("estq").Rows(pos)(0).ToString()
            Estoque = lista.Tables("estq").Rows(pos)(1).ToString()
        End If

    End Sub

#End Region

End Class
