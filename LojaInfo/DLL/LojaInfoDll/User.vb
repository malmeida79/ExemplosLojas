Imports system.data

Public Class User

#Region " Declarações "

    Dim codigo As Short
    Dim login As String
    Dim senha As String 
    Dim bloqueio As Boolean
    Dim ativo As Boolean 
    Dim permissao As Boolean
    Dim usuarios As 
    Dim parametro As String
    Dim totalLinhas As Short
    Dim cnn As New Conexao(".\sqlexpress", "dbLojaInfo")

#End Region

#Region " Entrada de valores "

    Public Sub setLogin(ByVal lgn As String)
        login = lgn
    End Sub

    Public Sub setSenha(ByVal pwd As String)
        senha = pwd
    End Sub

    Public Sub setBloqueio(ByVal blk As Boolean)
        bloqueio = blk
    End Sub

    Public Sub setAtivo(ByVal atv As Boolean)
        ativo = atv
    End Sub

    Public Sub setPermissao(ByVal pms As Boolean)
        permissao = pms
    End Sub

    Public Sub setParametro(ByVal par As String)
        parametro = par
    End Sub

#End Region

#Region "Saida de Valores"

    Public Function getLogin() As String
        getLogin = login
    End Function

    Public Function getBloqueio() As Boolean
        getBloqueio = bloqueio
    End Function

    Public Function getPermissao() As Boolean
        getPermissao = permissao
    End Function

    Public Function getSenha() As String
        getSenha = senha
    End Function

    Public Function getAtivo() As Boolean
        getAtivo = ativo
    End Function

    Public Function getTotalLinhas() As Short
        getTotalLinhas = totalLinhas
    End Function

#End Region

#Region " Ações "

    Public Sub busca()

        cnn.conecta()
        cnn.setComandoAdapter("Select * from tbluser  where logUser like '%" & parametro & "%'")
        cnn.setTabela("user")
        usuarios = cnn.buscaDs()

        'preenchendo o total de linhas
        totalLinhas = usuarios.Tables(0).Rows.Count

    End Sub

    Public Sub navega(ByVal pos As Short)

        ' se linhas foram encontradas ... ccarregar as variaveis
        If totalLinhas > 0 Then
            codigo = Short.Parse(usuarios.Tables(0).Rows(pos)(0).ToString)
            login = usuarios.Tables(0).Rows(pos)(1).ToString
            senha = usuarios.Tables(0).Rows(pos)(2).ToString
            bloqueio = usuarios.Tables(0).Rows(pos)(3).ToString
            ativo = usuarios.Tables(0).Rows(pos)(4).ToString
            permissao = usuarios.Tables(0).Rows(pos)(5).ToString
        End If

    End Sub

    Public Sub cadastrar()

        cnn.conecta()
        cnn.iniciaCmd("spc_cad_user")
        cnn.addParametro("@logUser", login, "i")
        cnn.addParametro("@pwdUser", senha, "i")
        cnn.addParametro("@blkUser", bloqueio, "i")
        cnn.addParametro("@atvUser", ativo, "i")
        cnn.addParametro("@pmsUser", permissao, "i")
        cnn.executaComando()
        cnn.encerra()

    End Sub

    Public Sub salvar()

        cnn.conecta()
        cnn.iniciaCmd("spc_alt_user")
        cnn.addParametro("@codigo", codigo, "i")
        cnn.addParametro("@logUser", login, "i")
        cnn.addParametro("@pwdUser", senha, "i")
        cnn.addParametro("@blkUser", bloqueio, "i")
        cnn.addParametro("@atvUser", ativo, "i")
        cnn.addParametro("@pmsUser", permissao, "i")
        cnn.executaComando()
        cnn.encerra()

    End Sub

#End Region

End Class