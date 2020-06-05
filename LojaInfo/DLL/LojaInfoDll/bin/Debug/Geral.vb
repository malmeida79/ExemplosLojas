Imports system
Imports System.Windows.Forms

Public Class Geral

    Public Sub alerta(ByVal frase As String)
        MessageBox.Show(frase, " :: Alerta :: ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Public Sub erro(ByVal frase As String)
        MessageBox.Show(frase, " :: Erro :: ", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Public Function duvida(ByVal frase As String) As Integer

        duvida = MessageBox.Show(frase, ":: Confirmação ::", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

    End Function

End Class
