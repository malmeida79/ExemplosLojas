Public Class Principal

    Dim ger As New Geral()

    Private Sub VeiculosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VeiculosToolStripMenuItem.Click

        Dim veiculos As New frmVeiculo
        veiculos.MdiParent = Me
        veiculos.Show()

    End Sub

    Private Sub HorizonotalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizonotalToolStripMenuItem.Click

        Me.LayoutMdi(MdiLayout.TileHorizontal)

    End Sub

    Private Sub VerticalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerticalToolStripMenuItem.Click

        Me.LayoutMdi(MdiLayout.TileVertical)

    End Sub

    Private Sub CascataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CascataToolStripMenuItem.Click

        Me.LayoutMdi(MdiLayout.Cascade)

    End Sub



    Private Sub atalho(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        ' necessário ativar propriedade keypreview do form principal
        ' ainda em propriedades clicar no relampago e ativar
        ' keydown

        Select Case e.KeyCode
            Case Keys.F5
                VeiculosToolStripMenuItem_Click(sender, e)
            Case Keys.F6
                VeiculosToolStripMenuItem_Click(sender, e)
        End Select



    End Sub


    Private Sub VeiculosF5ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        VeiculosToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub SairToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SairToolStripMenuItem.Click
        If ger.duvida("Deseja realmente encerrar a aplicação?") = vbYes Then End
    End Sub

    Private Sub SobreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SobreToolStripMenuItem.Click

        Dim sobre As New frmAbout
        sobre.mdiparent = Me
        sobre.show()

    End Sub

    'ação resize do form (caixa de propriedades, clicar em relampago)
    Private Sub minimizar(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize

        If WindowState = FormWindowState.Minimized Then
            Me.Hide()
        End If

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

        ' duplo clique em notify icon
        Me.Show()
        WindowState = FormWindowState.Maximized

    End Sub

    Private Sub AbrirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AbrirToolStripMenuItem.Click
        Me.Show()
        WindowState = FormWindowState.Maximized
    End Sub

    Private Sub EncerrarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EncerrarToolStripMenuItem.Click
        SairToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub MinimizarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MinimizarToolStripMenuItem.Click
        WindowState = FormWindowState.Minimized
    End Sub

    Private Sub VeículosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VeículosToolStripMenuItem.Click

        Dim rlt As New relatorio01()
        rlt.MdiParent = Me
        rlt.Show()

    End Sub

    Private Sub VeículosToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VeículosToolStripMenuItem1.Click
        VeiculosToolStripMenuItem_Click(sender, e)
    End Sub

End Class