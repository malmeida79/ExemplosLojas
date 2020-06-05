<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGrpProd
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
        Me.grpDescricao = New System.Windows.Forms.GroupBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.grpNavega = New System.Windows.Forms.GroupBox
        Me.btnUltimo = New System.Windows.Forms.Button
        Me.btnProximo = New System.Windows.Forms.Button
        Me.btnAnterior = New System.Windows.Forms.Button
        Me.btnPrimeiro = New System.Windows.Forms.Button
        Me.lblDisplay = New System.Windows.Forms.Label
        Me.grpBotoes1 = New System.Windows.Forms.GroupBox
        Me.btnCadastrar = New System.Windows.Forms.Button
        Me.btnCancelar = New System.Windows.Forms.Button
        Me.btnNovo = New System.Windows.Forms.Button
        Me.grpBotoes2 = New System.Windows.Forms.GroupBox
        Me.btnLimpa = New System.Windows.Forms.Button
        Me.btnExcluir = New System.Windows.Forms.Button
        Me.btnSalvar = New System.Windows.Forms.Button
        Me.btnBusca = New System.Windows.Forms.Button
        Me.btnFechar = New System.Windows.Forms.Button
        Me.grpDescricao.SuspendLayout()
        Me.grpNavega.SuspendLayout()
        Me.grpBotoes1.SuspendLayout()
        Me.grpBotoes2.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpDescricao
        '
        Me.grpDescricao.Controls.Add(Me.txtDesc)
        Me.grpDescricao.Controls.Add(Me.Label1)
        Me.grpDescricao.Location = New System.Drawing.Point(2, 0)
        Me.grpDescricao.Name = "grpDescricao"
        Me.grpDescricao.Size = New System.Drawing.Size(278, 72)
        Me.grpDescricao.TabIndex = 0
        Me.grpDescricao.TabStop = False
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(78, 29)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(175, 20)
        Me.txtDesc.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Descrição:"
        '
        'grpNavega
        '
        Me.grpNavega.Controls.Add(Me.btnUltimo)
        Me.grpNavega.Controls.Add(Me.btnProximo)
        Me.grpNavega.Controls.Add(Me.btnAnterior)
        Me.grpNavega.Controls.Add(Me.btnPrimeiro)
        Me.grpNavega.Controls.Add(Me.lblDisplay)
        Me.grpNavega.Location = New System.Drawing.Point(3, 74)
        Me.grpNavega.Name = "grpNavega"
        Me.grpNavega.Size = New System.Drawing.Size(277, 66)
        Me.grpNavega.TabIndex = 1
        Me.grpNavega.TabStop = False
        '
        'btnUltimo
        '
        Me.btnUltimo.Location = New System.Drawing.Point(222, 28)
        Me.btnUltimo.Name = "btnUltimo"
        Me.btnUltimo.Size = New System.Drawing.Size(39, 23)
        Me.btnUltimo.TabIndex = 4
        Me.btnUltimo.Text = ">>"
        Me.btnUltimo.UseVisualStyleBackColor = True
        '
        'btnProximo
        '
        Me.btnProximo.Location = New System.Drawing.Point(174, 28)
        Me.btnProximo.Name = "btnProximo"
        Me.btnProximo.Size = New System.Drawing.Size(39, 23)
        Me.btnProximo.TabIndex = 3
        Me.btnProximo.Text = ">"
        Me.btnProximo.UseVisualStyleBackColor = True
        '
        'btnAnterior
        '
        Me.btnAnterior.Location = New System.Drawing.Point(57, 28)
        Me.btnAnterior.Name = "btnAnterior"
        Me.btnAnterior.Size = New System.Drawing.Size(39, 23)
        Me.btnAnterior.TabIndex = 2
        Me.btnAnterior.Text = "<"
        Me.btnAnterior.UseVisualStyleBackColor = True
        '
        'btnPrimeiro
        '
        Me.btnPrimeiro.Location = New System.Drawing.Point(11, 28)
        Me.btnPrimeiro.Name = "btnPrimeiro"
        Me.btnPrimeiro.Size = New System.Drawing.Size(39, 23)
        Me.btnPrimeiro.TabIndex = 1
        Me.btnPrimeiro.Text = "<<"
        Me.btnPrimeiro.UseVisualStyleBackColor = True
        '
        'lblDisplay
        '
        Me.lblDisplay.AutoSize = True
        Me.lblDisplay.Location = New System.Drawing.Point(117, 34)
        Me.lblDisplay.Name = "lblDisplay"
        Me.lblDisplay.Size = New System.Drawing.Size(0, 13)
        Me.lblDisplay.TabIndex = 0
        '
        'grpBotoes1
        '
        Me.grpBotoes1.Controls.Add(Me.btnCadastrar)
        Me.grpBotoes1.Controls.Add(Me.btnCancelar)
        Me.grpBotoes1.Controls.Add(Me.btnNovo)
        Me.grpBotoes1.Location = New System.Drawing.Point(3, 141)
        Me.grpBotoes1.Name = "grpBotoes1"
        Me.grpBotoes1.Size = New System.Drawing.Size(277, 48)
        Me.grpBotoes1.TabIndex = 2
        Me.grpBotoes1.TabStop = False
        '
        'btnCadastrar
        '
        Me.btnCadastrar.Location = New System.Drawing.Point(172, 16)
        Me.btnCadastrar.Name = "btnCadastrar"
        Me.btnCadastrar.Size = New System.Drawing.Size(60, 23)
        Me.btnCadastrar.TabIndex = 11
        Me.btnCadastrar.Text = "Ca&dastrar"
        Me.btnCadastrar.UseVisualStyleBackColor = True
        '
        'btnCancelar
        '
        Me.btnCancelar.Location = New System.Drawing.Point(103, 16)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(60, 23)
        Me.btnCancelar.TabIndex = 10
        Me.btnCancelar.Text = "&Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = True
        '
        'btnNovo
        '
        Me.btnNovo.Location = New System.Drawing.Point(34, 16)
        Me.btnNovo.Name = "btnNovo"
        Me.btnNovo.Size = New System.Drawing.Size(60, 23)
        Me.btnNovo.TabIndex = 9
        Me.btnNovo.Text = "N&ovo"
        Me.btnNovo.UseVisualStyleBackColor = True
        '
        'grpBotoes2
        '
        Me.grpBotoes2.Controls.Add(Me.btnLimpa)
        Me.grpBotoes2.Controls.Add(Me.btnExcluir)
        Me.grpBotoes2.Controls.Add(Me.btnSalvar)
        Me.grpBotoes2.Controls.Add(Me.btnBusca)
        Me.grpBotoes2.Location = New System.Drawing.Point(287, 0)
        Me.grpBotoes2.Name = "grpBotoes2"
        Me.grpBotoes2.Size = New System.Drawing.Size(112, 140)
        Me.grpBotoes2.TabIndex = 3
        Me.grpBotoes2.TabStop = False
        '
        'btnLimpa
        '
        Me.btnLimpa.Location = New System.Drawing.Point(14, 105)
        Me.btnLimpa.Name = "btnLimpa"
        Me.btnLimpa.Size = New System.Drawing.Size(86, 23)
        Me.btnLimpa.TabIndex = 8
        Me.btnLimpa.Text = "Li&mpar"
        Me.btnLimpa.UseVisualStyleBackColor = True
        '
        'btnExcluir
        '
        Me.btnExcluir.Location = New System.Drawing.Point(14, 76)
        Me.btnExcluir.Name = "btnExcluir"
        Me.btnExcluir.Size = New System.Drawing.Size(86, 23)
        Me.btnExcluir.TabIndex = 7
        Me.btnExcluir.Text = "E&xcluir"
        Me.btnExcluir.UseVisualStyleBackColor = True
        '
        'btnSalvar
        '
        Me.btnSalvar.Location = New System.Drawing.Point(14, 47)
        Me.btnSalvar.Name = "btnSalvar"
        Me.btnSalvar.Size = New System.Drawing.Size(86, 23)
        Me.btnSalvar.TabIndex = 6
        Me.btnSalvar.Text = "Sa&lvar"
        Me.btnSalvar.UseVisualStyleBackColor = True
        '
        'btnBusca
        '
        Me.btnBusca.Location = New System.Drawing.Point(14, 18)
        Me.btnBusca.Name = "btnBusca"
        Me.btnBusca.Size = New System.Drawing.Size(86, 23)
        Me.btnBusca.TabIndex = 5
        Me.btnBusca.Text = "B&uscar"
        Me.btnBusca.UseVisualStyleBackColor = True
        '
        'btnFechar
        '
        Me.btnFechar.Location = New System.Drawing.Point(314, 156)
        Me.btnFechar.Name = "btnFechar"
        Me.btnFechar.Size = New System.Drawing.Size(60, 23)
        Me.btnFechar.TabIndex = 12
        Me.btnFechar.Text = "&Fechar"
        Me.btnFechar.UseVisualStyleBackColor = True
        '
        'frmGrpProd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(403, 194)
        Me.Controls.Add(Me.btnFechar)
        Me.Controls.Add(Me.grpBotoes2)
        Me.Controls.Add(Me.grpBotoes1)
        Me.Controls.Add(Me.grpNavega)
        Me.Controls.Add(Me.grpDescricao)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGrpProd"
        Me.Text = ":: Grupos de Produto ::"
        Me.grpDescricao.ResumeLayout(False)
        Me.grpDescricao.PerformLayout()
        Me.grpNavega.ResumeLayout(False)
        Me.grpNavega.PerformLayout()
        Me.grpBotoes1.ResumeLayout(False)
        Me.grpBotoes2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpDescricao As System.Windows.Forms.GroupBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpNavega As System.Windows.Forms.GroupBox
    Friend WithEvents grpBotoes1 As System.Windows.Forms.GroupBox
    Friend WithEvents grpBotoes2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnUltimo As System.Windows.Forms.Button
    Friend WithEvents btnProximo As System.Windows.Forms.Button
    Friend WithEvents btnAnterior As System.Windows.Forms.Button
    Friend WithEvents btnPrimeiro As System.Windows.Forms.Button
    Friend WithEvents lblDisplay As System.Windows.Forms.Label
    Friend WithEvents btnSalvar As System.Windows.Forms.Button
    Friend WithEvents btnBusca As System.Windows.Forms.Button
    Friend WithEvents btnLimpa As System.Windows.Forms.Button
    Friend WithEvents btnExcluir As System.Windows.Forms.Button
    Friend WithEvents btnCadastrar As System.Windows.Forms.Button
    Friend WithEvents btnCancelar As System.Windows.Forms.Button
    Friend WithEvents btnNovo As System.Windows.Forms.Button
    Friend WithEvents btnFechar As System.Windows.Forms.Button

End Class
