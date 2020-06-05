<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Principal
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Principal))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ArquivoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CadastrosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.VeiculosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.SairToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.JanelaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.HorizonotalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.VerticalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CascataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AjudaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SobreToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AbrirToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EncerrarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MinimizarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RelatóriosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.VeículosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CadastrosToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.VeiculosF5ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RelatóriosToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.VeículosToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ArquivoToolStripMenuItem, Me.RelatóriosToolStripMenuItem, Me.JanelaToolStripMenuItem, Me.AjudaToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(477, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ArquivoToolStripMenuItem
        '
        Me.ArquivoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CadastrosToolStripMenuItem, Me.ToolStripSeparator1, Me.SairToolStripMenuItem})
        Me.ArquivoToolStripMenuItem.Name = "ArquivoToolStripMenuItem"
        Me.ArquivoToolStripMenuItem.Size = New System.Drawing.Size(56, 20)
        Me.ArquivoToolStripMenuItem.Text = "&Arquivo"
        '
        'CadastrosToolStripMenuItem
        '
        Me.CadastrosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VeiculosToolStripMenuItem})
        Me.CadastrosToolStripMenuItem.Name = "CadastrosToolStripMenuItem"
        Me.CadastrosToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.CadastrosToolStripMenuItem.Text = "&Cadastros"
        '
        'VeiculosToolStripMenuItem
        '
        Me.VeiculosToolStripMenuItem.Name = "VeiculosToolStripMenuItem"
        Me.VeiculosToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
        Me.VeiculosToolStripMenuItem.Text = "&Veiculos      F5"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(131, 6)
        '
        'SairToolStripMenuItem
        '
        Me.SairToolStripMenuItem.Name = "SairToolStripMenuItem"
        Me.SairToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.SairToolStripMenuItem.Text = "&Sair"
        '
        'JanelaToolStripMenuItem
        '
        Me.JanelaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HorizonotalToolStripMenuItem, Me.VerticalToolStripMenuItem, Me.CascataToolStripMenuItem})
        Me.JanelaToolStripMenuItem.Name = "JanelaToolStripMenuItem"
        Me.JanelaToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.JanelaToolStripMenuItem.Text = "&Janela"
        '
        'HorizonotalToolStripMenuItem
        '
        Me.HorizonotalToolStripMenuItem.Name = "HorizonotalToolStripMenuItem"
        Me.HorizonotalToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.HorizonotalToolStripMenuItem.Text = "&Horizonotal"
        '
        'VerticalToolStripMenuItem
        '
        Me.VerticalToolStripMenuItem.Name = "VerticalToolStripMenuItem"
        Me.VerticalToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.VerticalToolStripMenuItem.Text = "&Vertical"
        '
        'CascataToolStripMenuItem
        '
        Me.CascataToolStripMenuItem.Name = "CascataToolStripMenuItem"
        Me.CascataToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.CascataToolStripMenuItem.Text = "&Cascata"
        '
        'AjudaToolStripMenuItem
        '
        Me.AjudaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SobreToolStripMenuItem})
        Me.AjudaToolStripMenuItem.Name = "AjudaToolStripMenuItem"
        Me.AjudaToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.AjudaToolStripMenuItem.Text = "A&juda"
        '
        'SobreToolStripMenuItem
        '
        Me.SobreToolStripMenuItem.Name = "SobreToolStripMenuItem"
        Me.SobreToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SobreToolStripMenuItem.Text = "Sobre ..."
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CadastrosToolStripMenuItem1, Me.RelatóriosToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(153, 70)
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon1.BalloonTipText = "Aplicação para controle de agencias de veiculos"
        Me.NotifyIcon1.BalloonTipTitle = "Informativo"
        Me.NotifyIcon1.ContextMenuStrip = Me.ContextMenuStrip2
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = ":: Sistema de Controle de Agências ::"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AbrirToolStripMenuItem, Me.EncerrarToolStripMenuItem, Me.MinimizarToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(127, 70)
        '
        'AbrirToolStripMenuItem
        '
        Me.AbrirToolStripMenuItem.Name = "AbrirToolStripMenuItem"
        Me.AbrirToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.AbrirToolStripMenuItem.Text = "Abrir"
        '
        'EncerrarToolStripMenuItem
        '
        Me.EncerrarToolStripMenuItem.Name = "EncerrarToolStripMenuItem"
        Me.EncerrarToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.EncerrarToolStripMenuItem.Text = "Encerrar"
        '
        'MinimizarToolStripMenuItem
        '
        Me.MinimizarToolStripMenuItem.Name = "MinimizarToolStripMenuItem"
        Me.MinimizarToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.MinimizarToolStripMenuItem.Text = "Ocultar"
        '
        'RelatóriosToolStripMenuItem
        '
        Me.RelatóriosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VeículosToolStripMenuItem})
        Me.RelatóriosToolStripMenuItem.Name = "RelatóriosToolStripMenuItem"
        Me.RelatóriosToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.RelatóriosToolStripMenuItem.Text = "R&elatórios"
        '
        'VeículosToolStripMenuItem
        '
        Me.VeículosToolStripMenuItem.Name = "VeículosToolStripMenuItem"
        Me.VeículosToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.VeículosToolStripMenuItem.Text = "&Veículos"
        '
        'CadastrosToolStripMenuItem1
        '
        Me.CadastrosToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VeiculosF5ToolStripMenuItem})
        Me.CadastrosToolStripMenuItem1.Name = "CadastrosToolStripMenuItem1"
        Me.CadastrosToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.CadastrosToolStripMenuItem1.Text = "Cadastros"
        '
        'VeiculosF5ToolStripMenuItem
        '
        Me.VeiculosF5ToolStripMenuItem.Name = "VeiculosF5ToolStripMenuItem"
        Me.VeiculosF5ToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.VeiculosF5ToolStripMenuItem.Text = "&Veiculos     F5"
        '
        'RelatóriosToolStripMenuItem1
        '
        Me.RelatóriosToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VeículosToolStripMenuItem1})
        Me.RelatóriosToolStripMenuItem1.Name = "RelatóriosToolStripMenuItem1"
        Me.RelatóriosToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.RelatóriosToolStripMenuItem1.Text = "Relatórios"
        '
        'VeículosToolStripMenuItem1
        '
        Me.VeículosToolStripMenuItem1.Name = "VeículosToolStripMenuItem1"
        Me.VeículosToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.VeículosToolStripMenuItem1.Text = "Veículos"
        '
        'Principal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(477, 349)
        Me.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Controls.Add(Me.MenuStrip1)
        Me.IsMdiContainer = True
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Principal"
        Me.Text = ":: Sistema de Controle de Agências ::"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ArquivoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CadastrosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VeiculosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SairToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents JanelaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HorizonotalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VerticalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CascataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AjudaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SobreToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AbrirToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EncerrarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MinimizarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RelatóriosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VeículosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CadastrosToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VeiculosF5ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RelatóriosToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VeículosToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
End Class
