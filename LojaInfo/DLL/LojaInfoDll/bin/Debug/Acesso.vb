Imports system.data

Public Class Acesso

#Region " Declarações "

    Dim login As String
    Dim senha As String
    Dim resultado As DataSet ' Dataset com Acessos buscados
    Dim totalLinhas As Integer ' controlar o total de linhas retornados pela busca
    Dim cnn As New Conexao() ' Instancia do objeto conexao

#End Region

#Region " Entrada de valores "

    Public Sub setLogin(ByVal lgn As String)
        login = lgn
    End Sub

    Public Sub setSenha(ByVal pwd As String)
        senha = pwd
    End Sub

#End Region

#Region "Saida de Valores"

    Public Function getTotalLinhas() As Integer
        getTotalLinhas = totalLinhas
    End Function

#End Region

#Region " Ações "

    Public Sub executaLogin()

        cnn.conecta()
        cnn.setComandoAdapter("select * from tblUser where logUser='" & login & "' and pwdUser='" & senha & "'")
        cnn.setTabela("users")
        resultado = cnn.buscaDs()

        'preenchendo o total de linhas
        totalLinhas = resultado.Tables(0).Rows.Count

    End Sub

#End Region

End Class