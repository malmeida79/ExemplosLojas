Public Class valida

    Public Function validaQtChar(ByVal valor As String, ByVal tamanho As Integer) As Boolean

        If valor.Length < tamanho Then
            validaQtChar = False
        Else
            validaQtChar = True
        End If

    End Function

    Public Function validaNaoVazio(ByVal valor As String) As Boolean

        If valor = "" Or valor = String.Empty Then
            validaNaoVazio = False
        Else
            validaNaoVazio = True
        End If

    End Function

    Public Function validaNumerico(ByVal valor As String) As Boolean

        If IsNumeric(valor) = False Then
            validaNumerico = False
        Else
            validaNumerico = True
        End If

    End Function

    Public Function validaTexto(ByVal valor As String) As Boolean

        If IsNumeric(valor) = True Then
            validaTexto = False
        Else
            validaTexto = True
        End If

    End Function




End Class
