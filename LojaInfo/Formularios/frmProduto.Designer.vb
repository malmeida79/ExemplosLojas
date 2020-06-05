<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProduto
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
        Me.cmbValor = New System.Windows.Forms.ComboBox
        Me.btnExclui = New System.Windows.Forms.Button
        Me.btnbAdiciona = New System.Windows.Forms.Button
        Me.txtQtd = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmbSGrp = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dtgItemProd = New System.Windows.Forms.DataGridView
        Me.btnFechar = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.lbltotal = New System.Windows.Forms.Label
        Me.btnSalvar = New System.Windows.Forms.Button
        Me.cmbProduto = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmbLocalEstoque = New System.Windows.Forms.ComboBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dtgItemProd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.cmbLocalEstoque)
        Me.GroupBox1.Controls.Add(Me.cmbProduto)
        Me.GroupBox1.Controls.Add(Me.cmbValor)
        Me.GroupBox1.Controls.Add(Me.btnExclui)
        Me.GroupBox1.Controls.Add(Me.btnbAdiciona)
        Me.GroupBox1.Controls.Add(Me.txtQtd)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cmbSGrp)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(357, 110)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Produto:"
        '
        'cmbValor
        '
        Me.cmbValor.FormattingEnabled = True
        Me.cmbValor.Location = New System.Drawing.Point(131, 120)
        Me.cmbValor.Name = "cmbValor"
        Me.cmbValor.Size = New System.Drawing.Size(73, 21)
        Me.cmbValor.TabIndex = 8
        '
        'btnExclui
        '
        Me.btnExclui.Location = New System.Drawing.Point(272, 75)
        Me.btnExclui.Name = "btnExclui"
        Me.btnExclui.Size = New System.Drawing.Size(66, 23)
        Me.btnExclui.TabIndex = 7
        Me.btnExclui.Text = "&Excluir"
        Me.btnExclui.UseVisualStyleBackColor = True
        '
        'btnbAdiciona
        '
        Me.btnbAdiciona.Location = New System.Drawing.Point(193, 75)
        Me.btnbAdiciona.Name = "btnbAdiciona"
        Me.btnbAdiciona.Size = New System.Drawing.Size(75, 23)
        Me.btnbAdiciona.TabIndex = 6
        Me.btnbAdiciona.Text = "&Adicionar"
        Me.btnbAdiciona.UseVisualStyleBackColor = True
        '
        'txtQtd
        '
        Me.txtQtd.Location = New System.Drawing.Point(89, 76)
        Me.txtQtd.Name = "txtQtd"
        Me.txtQtd.Size = New System.Drawing.Size(91, 20)
        Me.txtQtd.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(19, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Item:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Quantidade:"
        '
        'cmbSGrp
        '
        Me.cmbSGrp.FormattingEnabled = True
        Me.cmbSGrp.Location = New System.Drawing.Point(59, 49)
        Me.cmbSGrp.Name = "cmbSGrp"
        Me.cmbSGrp.Size = New System.Drawing.Size(121, 21)
        Me.cmbSGrp.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nome:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dtgItemProd)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 123)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(356, 184)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Itens do produto:"
        '
        'dtgItemProd
        '
        Me.dtgItemProd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtgItemProd.Location = New System.Drawing.Point(12, 20)
        Me.dtgItemProd.Name = "dtgItemProd"
        Me.dtgItemProd.Size = New System.Drawing.Size(333, 150)
        Me.dtgItemProd.TabIndex = 0
        '
        'btnFechar
        '
        Me.btnFechar.Location = New System.Drawing.Point(295, 318)
        Me.btnFechar.Name = "btnFechar"
        Me.btnFechar.Size = New System.Drawing.Size(53, 23)
        Me.btnFechar.TabIndex = 8
        Me.btnFechar.Text = "&Fechar"
        Me.btnFechar.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 322)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Valor total do Produto:"
        '
        'lbltotal
        '
        Me.lbltotal.AutoSize = True
        Me.lbltotal.Location = New System.Drawing.Point(136, 324)
        Me.lbltotal.Name = "lbltotal"
        Me.lbltotal.Size = New System.Drawing.Size(0, 13)
        Me.lbltotal.TabIndex = 10
        '
        'btnSalvar
        '
        Me.btnSalvar.Location = New System.Drawing.Point(236, 318)
        Me.btnSalvar.Name = "btnSalvar"
        Me.btnSalvar.Size = New System.Drawing.Size(53, 23)
        Me.btnSalvar.TabIndex = 11
        Me.btnSalvar.Text = "&Salvar"
        Me.btnSalvar.UseVisualStyleBackColor = True
        '
        'cmbProduto
        '
        Me.cmbProduto.FormattingEnabled = True
        Me.cmbProduto.Location = New System.Drawing.Point(59, 19)
        Me.cmbProduto.Name = "cmbProduto"
        Me.cmbProduto.Size = New System.Drawing.Size(280, 21)
        Me.cmbProduto.TabIndex = 9
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(186, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(31, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Estq:"
        '
        'cmbLocalEstoque
        '
        Me.cmbLocalEstoque.FormattingEnabled = True
        Me.cmbLocalEstoque.Location = New System.Drawing.Point(218, 49)
        Me.cmbLocalEstoque.Name = "cmbLocalEstoque"
        Me.cmbLocalEstoque.Size = New System.Drawing.Size(121, 21)
        Me.cmbLocalEstoque.TabIndex = 11
        '
        'frmProduto
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 349)
        Me.Controls.Add(Me.btnSalvar)
        Me.Controls.Add(Me.lbltotal)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnFechar)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProduto"
        Me.Text = ":: Produtos ::"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dtgItemProd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtQtd As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmbSGrp As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExclui As System.Windows.Forms.Button
    Friend WithEvents btnbAdiciona As System.Windows.Forms.Button
    Friend WithEvents dtgItemProd As System.Windows.Forms.DataGridView
    Friend WithEvents btnFechar As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lbltotal As System.Windows.Forms.Label
    Friend WithEvents cmbValor As System.Windows.Forms.ComboBox
    Friend WithEvents btnSalvar As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbLocalEstoque As System.Windows.Forms.ComboBox
    Friend WithEvents cmbProduto As System.Windows.Forms.ComboBox
End Class
