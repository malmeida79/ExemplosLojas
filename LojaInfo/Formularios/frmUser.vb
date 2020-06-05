Public Class frmUser

    Dim usr As New User()
    Dim vld As New valida()
    Dim ger As New Geral()
    Dim posicao As Short

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        painel(True, False, False, True, False, False, False)

    End Sub

    ' Recebe um objeto campo, esvazia e coloca foco no mesmo.
    Public Sub foca(ByVal campo As Object)
        campo.text = String.Empty
        campo.focus()
    End Sub

    ' Atalho para validação de campos TXT
    Public Function vlTxt(ByVal campo As Object, ByVal nome As String) As Boolean

        If vld.validaNaoVazio(campo.Text) = False Then
            ger.erro("Preencha o campo " & nome & ".")
            foca(campo)
            vlTxt = False
        Else
            vlTxt = True
        End If

    End Function

    Private Sub cmdSair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSair.Click
        'função para sair do formulário
        If ger.duvida("Deseja realmente Fechar?") = vbYes Then Me.Close()
    End Sub

    Public Function aplicaValidacao() As Boolean

        aplicaValidacao = True

        ' validação de fabricante
        If vlTxt(txtLogin, "Login") = False Then
            aplicaValidacao = False
            Exit Function
        End If

        ' Validação de veiculo
        If vlTxt(txtSenha, "Senha") = False Then
            aplicaValidacao = False
            Exit Function
        End If

        ' Validação de valor
        If vlTxt(txtConfirmacao, "Confirmação de senha") = False Then
            aplicaValidacao = False
            Exit Function
        End If

    End Function


    Public Sub painel(ByVal novo As Boolean, ByVal cancela As Boolean, ByVal cria As Boolean, ByVal busca As Boolean, ByVal salva As Boolean, ByVal navega As Boolean, ByVal campos As Boolean)

        'botoes de comando
        cmdNovo.Enabled = novo
        cmdCancela.Enabled = cancela
        cmdCria.Enabled = cria
        cmdBusca.Enabled = busca
        cmdSalva.Enabled = salva


        'botoes de navegacao
        cmdPrimeiro.Enabled = navega
        cmdAnterior.Enabled = navega
        cmdProximo.Enabled = navega
        cmdUltimo.Enabled = navega

        'controle dos campos
        grpCampos.Enabled = campos

    End Sub

    Private Sub cmdNovo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNovo.Click

        painel(False, True, True, False, False, False, True)
        limpaTela()

    End Sub

    Private Sub cmdCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancela.Click

        painel(True, False, False, True, False, False, False)
        limpaTela()

    End Sub

    Public Sub limpaTela()

        Dim obj As Object

        ' limpar a tela analisando qual tipo de controle e agindo para tal
        For Each obj In Me.grpCampos.Controls
            If TypeOf obj Is TextBox Then
                obj.text = String.Empty
            ElseIf TypeOf obj Is RadioButton Then
                obj.checked = False
            ElseIf TypeOf obj Is ComboBox Then
                obj.selecteditem = Nothing
                obj.text = String.Empty
            ElseIf TypeOf obj Is CheckBox Then
                obj.checked = False
            End If
        Next

        lblDisplay.Text = String.Empty

    End Sub

    Private Sub cmdCria_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCria.Click

        If aplicaValidacao() = False Then Exit Sub

        If txtSenha.Text <> txtConfirmacao.Text Then
            ger.erro("Sua senha não confere com a confirmação, por favor digite novamente.")
            txtSenha.Text = String.Empty
            txtConfirmacao.Text = String.Empty
            txtSenha.Focus()
            Exit Sub
        End If

        'carrega a classe e exibe a mensagem
        With usr
            .setLogin(txtLogin.Text)
            .setSenha(txtSenha.Text)
            .setBloqueio(clicaCheck("blk"))
            .setAtivo(clicaCheck("atv"))
            .setPermissao(clicaCheck("pms"))
        End With

        Try 'tente execuar as linhas que se seguem ...
            usr.cadastrar() ' cadastrar
            ger.alerta("Dados cadastrados com sucesso.")
            limpaTela()
        Catch ex As Exception ' caso ocorra qualquer excessão ... 
            'ger.erro("Houve o erro:" & ex.Message.ToString() & " no processo. Nada foi realizado.")
            ger.erro("Houve um erro no processo, nada foi realizado.")
        End Try

    End Sub


    Private Sub cmdSalva_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSalva.Click

        If aplicaValidacao() = False Then Exit Sub

        If txtSenha.Text <> txtConfirmacao.Text Then
            ger.erro("Sua senha não confere com a confirmação, por favor digite novamente.")
            txtSenha.Text = String.Empty
            txtConfirmacao.Text = String.Empty
            txtSenha.Focus()
            Exit Sub
        End If

        'carrega a classe e exibe a mensagem
        With usr
            .setLogin(txtLogin.Text)
            .setSenha(txtSenha.Text)
            .setBloqueio(clicaCheck("blk"))
            .setAtivo(clicaCheck("atv"))
            .setPermissao(clicaCheck("pms"))
        End With

        Try 'tente execuar as linhas que se seguem ...
            usr.salvar() ' cadastrar
            ger.alerta("Dados modificados com sucesso.")
            usr.busca()
        Catch ex As Exception ' caso ocorra qualquer excessão ... 
            'ger.erro("Houve o erro:" & ex.Message.ToString() & " no processo. Nada foi realizado.")
            ger.erro("Houve um erro no processo, nada foi realizado.")
        End Try

    End Sub

    Private Sub cmdBusca_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBusca.Click

        If txtPesquisar.Text <> "" Then

            'passagem do parametro de busca para classe
            usr.setParametro(txtPesquisar.Text)

            'realizando buwsca
            usr.busca()

            ' se algo foi localizado exibe
            If usr.getTotalLinhas > 0 Then
                'posiciona o cursor no primeiro registro encontrado
                usr.navega(0)
                painel(True, False, False, True, True, True, True)
                display(0, usr.getTotalLinhas())
            Else
                'senão, avisa que nada foi encontrado
                ger.erro("Nenhum registro encotrado.")
            End If

        Else
            ger.erro("Informe algo para Busca ...")
        End If


    End Sub

    Public Sub mostraDados()

        With usr
            txtLogin.Text = .getLogin()
            txtSenha.Text = .getSenha()
            mostraCheck("blk")
            mostraCheck("atv")
            mostraCheck("pms")
        End With

    End Sub

    Public Sub mostraCheck(ByVal opc As String)

        'verifica opçao do usuário se block ou aitvo
        If opc = "blk" Then
            'verifica se bloqueio esta marcado e se sim, marca na tela
            If usr.getBloqueio = True Then
                chkBloqueado.Checked = True
            Else
                chkBloqueado.Checked = False
            End If

        ElseIf opc = "atv" Then
            'verifica se ativo esta marcado e se sim, marca na tela
            If usr.getAtivo = True Then
                chkAtivo.Checked = True
            Else
                chkAtivo.Checked = False
            End If
        ElseIf opc = "pms" Then
            'verifica se ativo esta marcado e se sim, marca na tela
            If usr.getAtivo = True Then
                chkPermissao.Checked = True
            Else
                chkPermissao.Checked = False
            End If
        End If

    End Sub

    Public Function clicaCheck(ByVal opc As String) As Boolean

        'verifica opçao do usuário se block ou aitvo
        If opc = "blk" Then
            'verifica se bloqueio esta marcado e se sim, fornece true
            If chkBloqueado.Checked = True Then
                clicaCheck = True
            Else
                clicaCheck = False
            End If

        ElseIf opc = "atv" Then
            'verifica se ativo esta marcado e se sim, fornece false
            If chkAtivo.Checked = True Then
                clicaCheck = True
            Else
                clicaCheck = False
            End If
        ElseIf opc = "pms" Then
            'verifica se ativo esta marcado e se sim, fornece false
            If chkAtivo.Checked = True Then
                clicaCheck = True
            Else
                clicaCheck = False
            End If
        End If

    End Function

    Public Sub display(ByVal inicio As Short, ByVal fim As Short)
        mostraDados()
        lblDisplay.Text = (inicio + 1) & " de " & fim
    End Sub


    Private Sub cmdPrimeiro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrimeiro.Click

        posicao = 0
        usr.navega(0)
        display(0, usr.getTotalLinhas)

    End Sub

    Private Sub cmdAnterior_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAnterior.Click

        If posicao > 0 Then
            posicao = posicao - 1
            usr.navega(posicao)
            display(posicao, usr.getTotalLinhas)
        End If

    End Sub

    Private Sub cmdProximo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProximo.Click

        If posicao < (usr.getTotalLinhas - 1) Then
            posicao = posicao + 1
            usr.navega(posicao)
            display(posicao, usr.getTotalLinhas)
        End If

    End Sub

    Private Sub cmdUltimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUltimo.Click

        posicao = (usr.getTotalLinhas - 1)
        usr.navega(posicao)
        display(posicao, usr.getTotalLinhas)

    End Sub

End Class