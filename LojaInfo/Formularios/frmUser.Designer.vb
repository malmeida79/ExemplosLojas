<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUser
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdBusca = New System.Windows.Forms.Button
        Me.txtPesquisar = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.grpCampos = New System.Windows.Forms.GroupBox
        Me.chkAtivo = New System.Windows.Forms.CheckBox
        Me.txtConfirmacao = New System.Windows.Forms.TextBox
        Me.txtSenha = New System.Windows.Forms.TextBox
        Me.txtLogin = New System.Windows.Forms.TextBox
        Me.chkBloqueado = New System.Windows.Forms.CheckBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdPrimeiro = New System.Windows.Forms.Button
        Me.cmdAnterior = New System.Windows.Forms.Button
        Me.cmdProximo = New System.Windows.Forms.Button
        Me.cmdUltimo = New System.Windows.Forms.Button
        Me.lbldisplay = New System.Windows.Forms.Label
        Me.cmdCria = New System.Windows.Forms.Button
        Me.cmdSalva = New System.Windows.Forms.Button
        Me.cmdSair = New System.Windows.Forms.Button
        Me.cmdNovo = New System.Windows.Forms.Button
        Me.cmdCancela = New System.Windows.Forms.Button
        Me.chkPermissao = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.grpCampos.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdBusca)
        Me.GroupBox1.Controls.Add(Me.txtPesquisar)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(3, -1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(353, 41)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'cmdBusca
        '
        Me.cmdBusca.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdBusca.Location = New System.Drawing.Point(272, 12)
        Me.cmdBusca.Name = "cmdBusca"
        Me.cmdBusca.Size = New System.Drawing.Size(75, 23)
        Me.cmdBusca.TabIndex = 2
        Me.cmdBusca.Text = "GO !"
        Me.cmdBusca.UseVisualStyleBackColor = False
        '
        'txtPesquisar
        '
        Me.txtPesquisar.Location = New System.Drawing.Point(70, 13)
        Me.txtPesquisar.Name = "txtPesquisar"
        Me.txtPesquisar.Size = New System.Drawing.Size(194, 20)
        Me.txtPesquisar.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Pesquisar:"
        '
        'grpCampos
        '
        Me.grpCampos.Controls.Add(Me.chkPermissao)
        Me.grpCampos.Controls.Add(Me.chkAtivo)
        Me.grpCampos.Controls.Add(Me.txtConfirmacao)
        Me.grpCampos.Controls.Add(Me.txtSenha)
        Me.grpCampos.Controls.Add(Me.txtLogin)
        Me.grpCampos.Controls.Add(Me.chkBloqueado)
        Me.grpCampos.Controls.Add(Me.Label4)
        Me.grpCampos.Controls.Add(Me.Label3)
        Me.grpCampos.Controls.Add(Me.Label2)
        Me.grpCampos.Location = New System.Drawing.Point(4, 54)
        Me.grpCampos.Name = "grpCampos"
        Me.grpCampos.Size = New System.Drawing.Size(353, 134)
        Me.grpCampos.TabIndex = 1
        Me.grpCampos.TabStop = False
        '
        'chkAtivo
        '
        Me.chkAtivo.AutoSize = True
        Me.chkAtivo.Location = New System.Drawing.Point(172, 105)
        Me.chkAtivo.Name = "chkAtivo"
        Me.chkAtivo.Size = New System.Drawing.Size(50, 17)
        Me.chkAtivo.TabIndex = 7
        Me.chkAtivo.Text = "Ativo"
        Me.chkAtivo.UseVisualStyleBackColor = True
        '
        'txtConfirmacao
        '
        Me.txtConfirmacao.Location = New System.Drawing.Point(82, 72)
        Me.txtConfirmacao.Name = "txtConfirmacao"
        Me.txtConfirmacao.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtConfirmacao.Size = New System.Drawing.Size(250, 20)
        Me.txtConfirmacao.TabIndex = 6
        '
        'txtSenha
        '
        Me.txtSenha.Location = New System.Drawing.Point(82, 46)
        Me.txtSenha.Name = "txtSenha"
        Me.txtSenha.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSenha.Size = New System.Drawing.Size(250, 20)
        Me.txtSenha.TabIndex = 5
        '
        'txtLogin
        '
        Me.txtLogin.Location = New System.Drawing.Point(82, 20)
        Me.txtLogin.Name = "txtLogin"
        Me.txtLogin.Size = New System.Drawing.Size(250, 20)
        Me.txtLogin.TabIndex = 4
        '
        'chkBloqueado
        '
        Me.chkBloqueado.AutoSize = True
        Me.chkBloqueado.Location = New System.Drawing.Point(80, 105)
        Me.chkBloqueado.Name = "chkBloqueado"
        Me.chkBloqueado.Size = New System.Drawing.Size(83, 17)
        Me.chkBloqueado.TabIndex = 3
        Me.chkBloqueado.Text = "Bloquaeado"
        Me.chkBloqueado.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 73)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Confirmação:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(38, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Senha:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(43, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Login:"
        '
        'cmdPrimeiro
        '
        Me.cmdPrimeiro.Location = New System.Drawing.Point(7, 194)
        Me.cmdPrimeiro.Name = "cmdPrimeiro"
        Me.cmdPrimeiro.Size = New System.Drawing.Size(48, 23)
        Me.cmdPrimeiro.TabIndex = 2
        Me.cmdPrimeiro.Text = "< |"
        Me.cmdPrimeiro.UseVisualStyleBackColor = True
        '
        'cmdAnterior
        '
        Me.cmdAnterior.Location = New System.Drawing.Point(62, 194)
        Me.cmdAnterior.Name = "cmdAnterior"
        Me.cmdAnterior.Size = New System.Drawing.Size(49, 23)
        Me.cmdAnterior.TabIndex = 3
        Me.cmdAnterior.Text = "<<"
        Me.cmdAnterior.UseVisualStyleBackColor = True
        '
        'cmdProximo
        '
        Me.cmdProximo.Location = New System.Drawing.Point(257, 194)
        Me.cmdProximo.Name = "cmdProximo"
        Me.cmdProximo.Size = New System.Drawing.Size(46, 23)
        Me.cmdProximo.TabIndex = 4
        Me.cmdProximo.Text = ">>"
        Me.cmdProximo.UseVisualStyleBackColor = True
        '
        'cmdUltimo
        '
        Me.cmdUltimo.Location = New System.Drawing.Point(308, 194)
        Me.cmdUltimo.Name = "cmdUltimo"
        Me.cmdUltimo.Size = New System.Drawing.Size(48, 23)
        Me.cmdUltimo.TabIndex = 5
        Me.cmdUltimo.Text = "| >"
        Me.cmdUltimo.UseVisualStyleBackColor = True
        '
        'lbldisplay
        '
        Me.lbldisplay.AutoSize = True
        Me.lbldisplay.Location = New System.Drawing.Point(159, 199)
        Me.lbldisplay.Name = "lbldisplay"
        Me.lbldisplay.Size = New System.Drawing.Size(0, 13)
        Me.lbldisplay.TabIndex = 6
        '
        'cmdCria
        '
        Me.cmdCria.Location = New System.Drawing.Point(133, 235)
        Me.cmdCria.Name = "cmdCria"
        Me.cmdCria.Size = New System.Drawing.Size(60, 23)
        Me.cmdCria.TabIndex = 7
        Me.cmdCria.Text = "Ca&dastrar"
        Me.cmdCria.UseVisualStyleBackColor = True
        '
        'cmdSalva
        '
        Me.cmdSalva.Location = New System.Drawing.Point(199, 235)
        Me.cmdSalva.Name = "cmdSalva"
        Me.cmdSalva.Size = New System.Drawing.Size(60, 23)
        Me.cmdSalva.TabIndex = 8
        Me.cmdSalva.Text = "&Modificar"
        Me.cmdSalva.UseVisualStyleBackColor = True
        '
        'cmdSair
        '
        Me.cmdSair.Location = New System.Drawing.Point(308, 235)
        Me.cmdSair.Name = "cmdSair"
        Me.cmdSair.Size = New System.Drawing.Size(48, 23)
        Me.cmdSair.TabIndex = 9
        Me.cmdSair.Text = "&Sair"
        Me.cmdSair.UseVisualStyleBackColor = True
        '
        'cmdNovo
        '
        Me.cmdNovo.Location = New System.Drawing.Point(5, 235)
        Me.cmdNovo.Name = "cmdNovo"
        Me.cmdNovo.Size = New System.Drawing.Size(60, 23)
        Me.cmdNovo.TabIndex = 10
        Me.cmdNovo.Text = "&Novo"
        Me.cmdNovo.UseVisualStyleBackColor = True
        '
        'cmdCancela
        '
        Me.cmdCancela.Location = New System.Drawing.Point(70, 235)
        Me.cmdCancela.Name = "cmdCancela"
        Me.cmdCancela.Size = New System.Drawing.Size(60, 23)
        Me.cmdCancela.TabIndex = 11
        Me.cmdCancela.Text = "&Cancelar"
        Me.cmdCancela.UseVisualStyleBackColor = True
        '
        'chkPermissao
        '
        Me.chkPermissao.AutoSize = True
        Me.chkPermissao.Location = New System.Drawing.Point(228, 105)
        Me.chkPermissao.Name = "chkPermissao"
        Me.chkPermissao.Size = New System.Drawing.Size(105, 17)
        Me.chkPermissao.TabIndex = 8
        Me.chkPermissao.Text = "Acesso Usuários"
        Me.chkPermissao.UseVisualStyleBackColor = True
        '
        'frmUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(360, 263)
        Me.Controls.Add(Me.cmdCancela)
        Me.Controls.Add(Me.cmdNovo)
        Me.Controls.Add(Me.cmdSair)
        Me.Controls.Add(Me.cmdSalva)
        Me.Controls.Add(Me.cmdCria)
        Me.Controls.Add(Me.lbldisplay)
        Me.Controls.Add(Me.cmdUltimo)
        Me.Controls.Add(Me.cmdProximo)
        Me.Controls.Add(Me.cmdAnterior)
        Me.Controls.Add(Me.cmdPrimeiro)
        Me.Controls.Add(Me.grpCampos)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = ":: Controle de usuários ::"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.grpCampos.ResumeLayout(False)
        Me.grpCampos.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdBusca As System.Windows.Forms.Button
    Friend WithEvents txtPesquisar As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpCampos As System.Windows.Forms.GroupBox
    Friend WithEvents txtConfirmacao As System.Windows.Forms.TextBox
    Friend WithEvents txtSenha As System.Windows.Forms.TextBox
    Friend WithEvents txtLogin As System.Windows.Forms.TextBox
    Friend WithEvents chkBloqueado As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkAtivo As System.Windows.Forms.CheckBox
    Friend WithEvents cmdPrimeiro As System.Windows.Forms.Button
    Friend WithEvents cmdAnterior As System.Windows.Forms.Button
    Friend WithEvents cmdProximo As System.Windows.Forms.Button
    Friend WithEvents cmdUltimo As System.Windows.Forms.Button
    Friend WithEvents lbldisplay As System.Windows.Forms.Label
    Friend WithEvents cmdCria As System.Windows.Forms.Button
    Friend WithEvents cmdSalva As System.Windows.Forms.Button
    Friend WithEvents cmdSair As System.Windows.Forms.Button
    Friend WithEvents cmdNovo As System.Windows.Forms.Button
    Friend WithEvents cmdCancela As System.Windows.Forms.Button
    Friend WithEvents chkPermissao As System.Windows.Forms.CheckBox
End Class
