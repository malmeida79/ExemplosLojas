Imports LojaInfoDll

Public Class frmEstoque

    Dim estq As New clsEstoque
    Dim ger As New Geral
    Dim posicao As Integer

    Private Sub frmEstoque_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        painel(True, False, False, False, True, False, False, False, False)

    End Sub

    Private Sub btnCadastrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCadastrar.Click

        With estq

            Try
                ' fornece dados para cadastrar o Estoque
                .setEstoque(txtDesc.Text)
                .cadastrar()

                ger.alerta("Estoque cadastrado com sucesso.")

            Catch ex As Exception
                ger.erro("Ocorreu um erro, nada foi realizado.")
            End Try

        End With

    End Sub

    Private Sub btnBusca_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBusca.Click

        Dim parametro As String

        parametro = InputBox("Digite o nome da Estoque", " :: Pesquisa de empresas ::")

        estq.setParametro(parametro)
        estq.buscar()

        If estq.getTotalLinhas() > 0 Then
            estq.carregar(0)
            display(0)
            painel(True, False, False, False, True, True, True, True, True)
        Else
            ger.erro("Nenhum Estoque foi localizado.")
        End If

    End Sub

    Public Sub mostraDesc()

        txtDesc.Text = estq.getEstoque()

    End Sub

    Private Sub btnSair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFechar.Click
        If ger.duvida("Deseja realmente sair ?") = vbYes Then Me.Close()
    End Sub

    Private Sub btnPrimeiro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrimeiro.Click

        posicao = 0
        estq.carregar(posicao)
        display(posicao)

    End Sub

    Private Sub btnAnterior_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAnterior.Click

        If posicao > 0 Then
            posicao = posicao - 1
            estq.carregar(posicao)
            display(posicao)
        End If

    End Sub

    Private Sub btnProximo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProximo.Click
        If posicao < (estq.getTotalLinhas - 1) Then
            posicao = posicao + 1
            estq.carregar(posicao)
            display(posicao)
        End If
    End Sub

    Private Sub btnUltimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUltimo.Click
        posicao = estq.getTotalLinhas - 1
        estq.carregar(posicao)
        display(posicao)
    End Sub
    Public Sub display(ByVal inicio As Integer)
        lblDisplay.Text = (inicio + 1) & " de " & estq.getTotalLinhas
        mostraDesc()
    End Sub

    Private Sub btnSalvar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalvar.Click
        With estq

            Try
                ' fornece dados para cadastrar o Estoque
                .setEstoque(txtDesc.Text)
                .salvar()
                .buscar()
                ger.alerta("Grupo alterado com sucesso.")

            Catch ex As Exception
                ger.erro("Ocorreu um erro, nada foi realizado.")
            End Try

        End With
    End Sub

    Private Sub btnExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcluir.Click
        ' teste se usuario deseja realmente excluir
        If ger.duvida("Deseja realmente exluir o grupo:") = vbYes Then
            With estq
                Try
                    'apaga o Estoque
                    .apagar()
                    'refaz o dataset atualizando o mesmo
                    .buscar()
                    'avisa o Estoque
                    ger.alerta("Grupo excluido com sucesso.")
                    'verifica se tem linha para onde ir
                    If .getTotalLinhas > 0 Then
                        'caso tenha verificado se nao estamos com a ultima linha 
                        ' do banco, ou seja registro numero 1
                        If posicao - 1 > 0 Then
                            posicao = posicao - 1
                        Else
                            'se ultimo registro, entao para zero.
                            posicao = 0
                        End If
                        'carrega dados do Estoque novo na tela,
                        'ou seja um antes do excluido
                        estq.carregar(posicao)
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpa.Click

        limparCampos()

    End Sub
    
    Public Sub painel(ByVal novo As Boolean, ByVal cancela As Boolean, ByVal cadastra As Boolean, ByVal limpa As Boolean, ByVal busca As Boolean, ByVal salva As Boolean, ByVal exclui As Boolean, ByVal navega As Boolean, ByVal campos As Boolean)

        btnNovo.Enabled = novo
        btnCancelar.Enabled = cancela
        btnCadastrar.Enabled = cadastra
        btnLimpa.Enabled = limpa
        btnBusca.Enabled = busca
        btnSalvar.Enabled = salva
        btnExcluir.Enabled = exclui

        btnPrimeiro.Enabled = navega
        btnProximo.Enabled = navega
        btnAnterior.Enabled = navega
        btnUltimo.Enabled = navega

        grpDescricao.Enabled = campos

    End Sub


    Private Sub btnNovo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNovo.Click
        painel(False, True, True, True, False, False, False, False, True)
        limparCampos()
    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        painel(True, False, False, False, True, False, False, False, False)
    End Sub

    Public Sub limparCampos()

        Dim campo As Object

        For Each campo In Me.grpDescricao.Controls

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
