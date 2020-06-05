Imports LojaInfodll

Public Class clsGrpProd

#Region " Variaveis e instancias "

    Dim cnn As New Conexao(".\sqlexpress", "dblojaInfo")

    Dim codigo As Integer
    Dim descricao As String
    Dim lista As DataSet
    Dim parametro As String
    Dim totalLinhas As Integer

#End Region

#Region " Metodos de Input / output "

    Public Sub setParametro(ByVal par As String)
        parametro = par
    End Sub

    Public Sub setDescricao(ByVal dsc As String)
        descricao = dsc
    End Sub

    Public Function getDescricao() As String
        getDescricao = descricao
    End Function


    Public Function getTotalLinhas() As Integer
        getTotalLinhas = totalLinhas
    End Function

#End Region

#Region " Ações "

    Public Sub cadastrar()

        cnn.conecta()
        cnn.iniciaCmd("spc_cad_grp")
        cnn.addParametro("@grpDesc", descricao, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub salvar()

        cnn.conecta()
        cnn.iniciaCmd("spc_alt_grp")
        cnn.addParametro("@grpId", codigo, "i")
        cnn.addParametro("@grpDesc", descricao, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub apagar()

        cnn.conecta()
        cnn.iniciaCmd("spc_del_grp")
        cnn.addParametro("@grpId", codigo, "i")
        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub buscar()


        'connecta o banco
        cnn.conecta()

        'alimenta o dataAdapter
        cnn.setComandoAdapter("select * from tblGrpProd where grpDesc like '%" & parametro & "%'")

        'nomeia a tabela
        cnn.setTabela("grp")

        'executa busca
        lista = cnn.buscaDs()

        'retorna total de linhas 
        totalLinhas = lista.Tables("grp").Rows.Count

    End Sub

    Public Sub carregar(ByVal pos As Integer)

        If totalLinhas > 0 Then
            codigo = lista.Tables("grp").Rows(pos)(0).ToString()
            descricao = lista.Tables("grp").Rows(pos)(1).ToString()
        End If

    End Sub

#End Region

End Class
