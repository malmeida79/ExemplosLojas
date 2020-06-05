Imports LojaInfoDll
Imports System.Data.SqlClient

Public Class frmCliente

    Dim cli As New clsCliente()
    Dim ger As New geral()
    Dim posicao As Integer

    Private Sub frmCliente_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        loadCombos()
        painel(True, False, False, False, True, False, False, False, False)

    End Sub

    Private Sub testaCep(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCep.LostFocus

        cli.setCep(txtCep.Text)
        cli.consultaCep()

        'verifica se achou alguma coisa ...
        If cli.getTotLinhasEnd() > 0 Then

            mostraend()

        Else

            txtEnd.Text = String.Empty

            cboCidade.SelectedValue = 0
            cboBairro.SelectedValue = 0
            cboSigla.SelectedValue = 0
            cboEstado.SelectedValue = 0

            txtEnd.Enabled = True
            cboCidade.Enabled = True
            cboBairro.Enabled = True
            cboEstado.Enabled = True
            cboSigla.Enabled = True

        End If

    End Sub

    Private Sub cmdCadastrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCadastrar.Click

        With cli

            Try
                ' fornece dados para cadastrar o cliente
                .setNome(txtNome.Text)
                .setFantasia(txtFantasia.Text)
                .setCnpj(txtCnpj.Text)
                .setNumero(txtNumero.Text)

                If .getTotLinhasEnd() > 0 And txtCep.Text.Trim() <> "" Then
                    'cadastra o cliente caso endereco ja conhecido
                    .cadastrar()
                Else
                    'fornece os dados para o cep
                    .setCep(txtCep.Text)
                    .setSigla(cboSigla.SelectedValue)
                    .setEndereco(txtEnd.Text)
                    .setCidade(cboCidade.SelectedValue)
                    .setBairro(cboBairro.SelectedValue)
                    .setEstado(cboEstado.SelectedValue)

                    'cadastra endereco
                    .cadastraendereco()
                    'cadastra o cliente
                    .cadastrar()
                End If

                ger.alerta("Cliente cadastrado com sucesso.")

            Catch ex As Exception
                ger.erro("Ocorreu um erro, nada foi realizado.")
            End Try

        End With

    End Sub

    Private Sub cmdBusca_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBusca.Click


        Dim parametro As String

        parametro = InputBox("Digite o nome da empresa", " :: Pesquisa de empresas ::")

        cli.setParametro(parametro)
        cli.buscar()

        If cli.getTotalLinhasCli() > 0 Then
            cli.navega(0)
            display(0)
            painel(True, False, False, False, True, True, True, True, True)
        Else
            ger.erro("Nenhum Cliente foi localizado.")
        End If

    End Sub

    Public Sub mostraEnd()

        'copiar os dados do dataset deste endereco para as vcaraiveis da classe
        cli.carregaEndereco()

        ' utilizando as variaveis da classe
        txtEnd.Text = cli.getEndereco()

        cboCidade.SelectedValue = cli.getCidade()
        cboBairro.SelectedValue = cli.getBairro()
        cboSigla.SelectedValue = cli.getSigla()
        cboEstado.SelectedValue = cli.getEstado()

        txtEnd.Enabled = False
        cboCidade.Enabled = False
        cboBairro.Enabled = False
        cboEstado.Enabled = False
        cboSigla.Enabled = False

        txtNumero.Focus()

    End Sub

    Public Sub mostraCli()

        txtNome.Text = cli.getNome()
        txtFantasia.Text = cli.getFantasia()
        txtCnpj.Text = cli.getCnpj()
        txtNumero.Text = cli.getNumero()
        txtCep.Text = cli.getCep()

    End Sub

    Private Sub cmdSair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSair.Click
        If ger.duvida("Deseja realmente sair ?") = vbYes Then Me.Close()
    End Sub

    Private Sub cmdPrimeiro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrimeiro.Click

        posicao = 0
        cli.navega(posicao)
        display(posicao)

    End Sub

    Private Sub cmdAnterior_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAnterior.Click

        If posicao > 0 Then
            posicao = posicao - 1
            cli.navega(posicao)
            display(posicao)
        End If

    End Sub

    Private Sub cmdProximo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProximo.Click
        If posicao < (cli.getTotalLinhasCli - 1) Then
            posicao = posicao + 1
            cli.navega(posicao)
            display(posicao)
        End If
    End Sub

    Private Sub cmdUltimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUltimo.Click
        posicao = cli.getTotalLinhasCli - 1
        cli.navega(posicao)
        display(posicao)
    End Sub
    Public Sub display(ByVal inicio As Integer)
        lblDisplay.Text = (inicio + 1) & " de " & cli.getTotalLinhasCli
        mostraCli()
        mostraEnd()
    End Sub

    Private Sub CmdSalvar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSalvar.Click
        With cli

            Try
                ' fornece dados para cadastrar o cliente
                .setNome(txtNome.Text)
                .setFantasia(txtFantasia.Text)
                .setCnpj(txtCnpj.Text)
                .setNumero(txtNumero.Text)

                If .getTotLinhasEnd() > 0 And txtCep.Text.Trim() <> "" Then
                    'cadastra o cliente caso endereco ja conhecido
                    .salvar()
                Else
                    'fornece os dados para o cep
                    .setCep(txtCep.Text)
                    .setSigla(cboSigla.SelectedValue)
                    .setEndereco(txtEnd.Text)
                    .setCidade(cboCidade.SelectedValue)
                    .setBairro(cboBairro.SelectedValue)
                    .setEstado(cboEstado.SelectedValue)

                    'cadastra endereco
                    .cadastraendereco()
                    'cadastra o cliente
                    .salvar()
                End If
                cli.buscar()
                ger.alerta("Cliente Alterado com sucesso.")

            Catch ex As Exception
                ger.erro("Ocorreu um erro, nada foi realizado.")
            End Try

        End With
    End Sub

    Private Sub cmdExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcluir.Click
        ' teste se usuario deseja realmente excluir
        If ger.duvida("Deseja realmente exluir o cliente:") = vbYes Then
            With cli
                Try
                    'apaga o cliente
                    .apagar()
                    'refaz o dataset atualizando o mesmo
                    .buscar()
                    'avisa o cliente
                    ger.alerta("Cliente excluido com sucesso.")
                    'verifica se tem linha para onde ir
                    If .getTotalLinhasCli > 0 Then
                        'caso tenha verificado se nao estamos com a ultima linha 
                        ' do banco, ou seja registro numero 1
                        If posicao - 1 > 0 Then
                            posicao = posicao - 1
                        Else
                            'se ultimo registro, entao para zero.
                            posicao = 0
                        End If
                        'carrega dados do cliente novo na tela,
                        'ou seja um antes do excluido
                        cli.navega(posicao)
                        'atualiza display.
                        display(posicao)
                    Else
                        limparCampos()
                        painel(True, False, False, False, True, False, False, False, False)
                    End If
                Catch ex As Exception
                    ger.erro("Ocorreu um erro, nada foi realizado.")
                End Try
            End With
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLimpar.Click

        limparcampos()

    End Sub
    Public Sub loadCombos()

        Dim arrCidade As New ArrayList
        Dim arrEstado As New ArrayList
        Dim arrSigla As New ArrayList
        Dim arrBairro As New ArrayList

        Dim drCid As SqlDataReader
        Dim drEst As SqlDataReader
        Dim drBrr As SqlDataReader
        Dim drSg As SqlDataReader

        drCid = cli.carregaCidade()
        drEst = cli.carregaEstado()
        drBrr = cli.carregaBairro()
        drSg = cli.carregaSigla()

        While drCid.Read
            arrCidade.Add(New Drop(drCid(1).ToString(), Convert.ToInt32(drCid(0).ToString())))
        End While

        While drEst.Read
            arrEstado.Add(New Drop(drEst(1).ToString(), Convert.ToInt32(drEst(0).ToString())))
        End While

        While drBrr.Read
            arrBairro.Add(New Drop(drBrr(1).ToString(), Convert.ToInt32(drBrr(0).ToString())))
        End While

        While drSg.Read
            arrSigla.Add(New Drop(drSg(1).ToString(), Convert.ToInt32(drSg(0).ToString())))
        End While

        cboCidade.DataSource = arrCidade
        cboCidade.ValueMember = "valor"
        cboCidade.DisplayMember = "nome"

        cboEstado.DataSource = arrEstado
        cboEstado.ValueMember = "valor"
        cboEstado.DisplayMember = "nome"

        cboBairro.DataSource = arrBairro
        cboBairro.ValueMember = "valor"
        cboBairro.DisplayMember = "nome"

        cboSigla.DataSource = arrSigla
        cboSigla.ValueMember = "valor"
        cboSigla.DisplayMember = "nome"

        cboCidade.SelectedValue = 0
        cboBairro.SelectedValue = 0
        cboSigla.SelectedValue = 0
        cboEstado.SelectedValue = 0

    End Sub

    Public Sub painel(ByVal novo As Boolean, ByVal cancela As Boolean, ByVal cadastra As Boolean, ByVal limpa As Boolean, ByVal busca As Boolean, ByVal salva As Boolean, ByVal exclui As Boolean, ByVal navega As Boolean, ByVal campos As Boolean)

        cmdNovo.Enabled = novo
        cmdCancelar.Enabled = cancela
        cmdCadastrar.Enabled = cadastra
        cmdLimpar.Enabled = limpa
        cmdBusca.Enabled = busca
        CmdSalvar.Enabled = salva
        cmdExcluir.Enabled = exclui

        cmdPrimeiro.Enabled = navega
        cmdProximo.Enabled = navega
        cmdAnterior.Enabled = navega
        cmdUltimo.Enabled = navega

        grpCampos.Enabled = campos

    End Sub


    Private Sub cmdNovo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNovo.Click
        painel(False, True, True, True, False, False, False, False, True)
        limparCampos()
    End Sub

    Private Sub cmdCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelar.Click
        painel(True, False, False, False, True, False, False, False, False)
    End Sub

    Public Sub limparCampos()

        Dim campo As Object

        For Each campo In Me.grpCampos.Controls

            If TypeOf campo Is TextBox Then
                campo.text = String.Empty
            ElseIf TypeOf campo Is RadioButton Then
                campo.checked = False
            ElseIf TypeOf campo Is ComboBox Then
                campo.SelectedValue = 0
            End If

            campo.enabled = True

        Next

        lblDisplay.Text = String.Empty

    End Sub
End Class
