Imports LojaInfodll
Imports System.Data.SqlClient

Public Class clsProduto

#Region " Variaveis e instancias "

    Dim cnn As New Conexao(".\sqlexpress", "dblojaInfo")

    Dim codigo As Integer
    Dim descricao As String
    Dim grupo As Short
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

    Public Sub setGrupo(ByVal grp As Short)
        grupo = grp
    End Sub

    Public Function getGrupo() As Short
        getGrupo = grupo
    End Function


    Public Function getTotalLinhas() As Integer
        getTotalLinhas = totalLinhas
    End Function

#End Region

#Region " Ações "

    Public Sub cadastrar()

        cnn.conecta()
        cnn.iniciaCmd("spc_cad_sgrp")
        cnn.addParametro("@sgrpDesc", descricao, "i")
        cnn.addParametro("@grpid", grupo, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub salvar()

        cnn.conecta()
        cnn.iniciaCmd("spc_alt_sgrp")
        cnn.addParametro("@grpId", codigo, "i")
        cnn.addParametro("@grpDesc", descricao, "i")
        cnn.addParametro("@grpid", grupo, "i")

        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub apagar()

        cnn.conecta()
        cnn.iniciaCmd("spc_del_sgrp")
        cnn.addParametro("@grpId", codigo, "i")
        'cadastra novo resgistro
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub buscar()


        'connecta o banco
        cnn.conecta()

        'alimenta o dataAdapter
        cnn.setComandoAdapter("select * from tblsGrpProd where sGrpDesc like '%" & parametro & "%'")

        'nomeia a tabela
        cnn.setTabela("sgrp")

        'executa busca
        lista = cnn.buscaDs()

        'retorna total de linhas 
        totalLinhas = lista.Tables("sgrp").Rows.Count

    End Sub

    Public Sub carregar(ByVal pos As Integer)

        If totalLinhas > 0 Then
            codigo = lista.Tables("sgrp").Rows(pos)(0).ToString()
            descricao = lista.Tables("sgrp").Rows(pos)(1).ToString()
            grupo = lista.Tables("sgrp").Rows(pos)(2).ToString()
        End If

    End Sub

    Public Function carregaSGrp() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from tblSGrpProd")
        cnn.iniciaCmd()
        carregaSGrp = cnn.buscaDr
    End Function

    Public Function carregaValor() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select idprod,avg(estqValUnit) from tblItemEstoque group by idprod")
        cnn.iniciaCmd()
        carregaValor = cnn.buscaDr
    End Function

    Public Function carregaEstoque() As SqlDataReader
        cnn.conecta()
        cnn.setComando("select * from tblEstoque")
        cnn.iniciaCmd()
        carregaEstoque = cnn.buscaDr
    End Function

#End Region

End Class
