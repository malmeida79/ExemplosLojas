Imports LojaInfodll
Imports System.Data.SqlClient

Public Class clsEntrada

#Region " Variaveis e instancias "

    Dim cnn As New Conexao(".\sqlexpress", "dblojaInfo")

    Dim quantidade As Short
    Dim valorUnitario As Decimal
    Dim subGrupo As Short
    Dim estoque As Short
    Dim lista As DataSet
    Dim parametro As String
    Dim totalLinhas As Integer

#End Region

#Region " Metodos de Input / output "

    Public Sub setParametro(ByVal par As String)
        parametro = par
    End Sub

    Public Sub setQuantidade(ByVal dsc As Short)
        quantidade = dsc
    End Sub

    Public Function getQuantidade() As Short
        getQuantidade = quantidade
    End Function

    Public Sub setValUnitario(ByVal valUnit As Decimal)
        quantidade = valUnit
    End Sub

    Public Function getValUnitario() As Decimal
        getValUnitario = quantidade
    End Function

    Public Sub setsubGrupo(ByVal grp As Short)
        subGrupo = grp
    End Sub

    Public Function getsubGrupo() As Short
        getsubGrupo = subGrupo
    End Function

    Public Sub setEstoque(ByVal grp As Short)
        estoque = grp
    End Sub

    Public Function getEstoque() As Short
        getEstoque = estoque
    End Function

    Public Function getTotalLinhas() As Integer
        getTotalLinhas = totalLinhas
    End Function

#End Region

#Region " Ações "

    Public Sub cadastrar()

        cnn.conecta()
        cnn.iniciaCmd("spc_cad_it_estq")
        cnn.addParametro("@estqItemId", estoque, "i")
        cnn.addParametro("@estqValUnit", valorUnitario, "i")
        cnn.addParametro("@idprod", subGrupo, "i")
        cnn.addParametro("@estqQtde", quantidade, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub salvar()

        cnn.conecta()
        cnn.iniciaCmd("spc_alt_it_estq")
        cnn.addParametro("@estqItemId", estoque, "i")
        cnn.addParametro("@estqValUnit", valorUnitario, "i")
        cnn.addParametro("@estqQtde", quantidade, "i")
        cnn.addParametro("@idprod", subGrupo, "i")


        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub apagar()

        cnn.conecta()
        cnn.iniciaCmd("spc_del_it_estq")
        cnn.addParametro("@estqItemId", estoque, "i")
        cnn.addParametro("@idProd", subGrupo, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub buscar()


        'connecta o banco
        cnn.conecta()

        'alimenta o dataAdapter
        cnn.setComandoAdapter("select * from tblItemEstoque inner join tblSGrpProd on tblItemEstoque.idProd=tblSGrpProd.grpId where sgrpDesc like '%" & parametro & "%'")

        'nomeia a tabela
        cnn.setTabela("entrada")

        'executa busca
        lista = cnn.buscaDs()

        'retorna total de linhas 
        totalLinhas = lista.Tables("entrada").Rows.Count

    End Sub

    Public Sub carregar(ByVal pos As Integer)

        If totalLinhas > 0 Then
            estoque = lista.Tables("entrada").Rows(pos)(0).ToString()
            subGrupo = lista.Tables("entrada").Rows(pos)(1).ToString()
            quantidade = lista.Tables("entrada").Rows(pos)(2).ToString()
            valorUnitario = lista.Tables("entrada").Rows(pos)(3).ToString()
        End If

    End Sub

    Public Function carregaGrp() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from tblsGrpProd")
        cnn.iniciaCmd()
        carregaGrp = cnn.buscaDr
    End Function

    Public Function carregaEstq() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from tblEstoque")
        cnn.iniciaCmd()
        carregaEstq = cnn.buscaDr
    End Function

#End Region

End Class
