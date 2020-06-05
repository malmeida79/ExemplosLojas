Public Class Login

    Dim log As New Acesso()
    Dim vld As New valida()
    Dim ger As New Geral()
    Dim contagem As Short

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'verifica se login foi preenchido
        If vld.validaNaoVazio(txtlog.Text) = True Then
            'verifica se senha foi preenchida
            If vld.validaNaoVazio(txtsenha.Text) = True Then

                ' alimenta a classe login e realiza cosulta
                log.setLogin(txtlog.Text)
                log.setSenha(txtsenha.Text)

                ' passagem de parametros para classe
                log.testaLogin()
                log.testaSenha()
                log.testaBlock()
                log.testaAtivo()
                log.testaPermissao()

                'se login esta correto
                If log.getTotalLinhasLogin > 0 Then
                    'se senha � correta
                    If log.getTotalLinhasSenha > 0 Then
                        'se usu�rio ativo
                        If log.getTotalLinhasAtivo > 0 Then
                            'se usuario bloqueado
                            If log.getTotalLinhasBlock > 0 Then
                                'avisar que esta bloqueado
                                ger.erro("Seu usu�rio se encontra bloqueado, por favor contate o administrador do sistema.")
                                txtlog.Text = String.Empty
                                txtsenha.Text = String.Empty
                                txtlog.Focus()
                                'se  n�o, logar o formulario principal.
                            Else
                                'esconde a tela de login

                                ' Verifica op��o do usuario, se o mesmo quer trabalhar com usu�rios ou no
                                ' sistema (ambiente operacional)
                                If CboAmbiente.SelectedIndex = 0 Then
                                    Me.Hide()
                                    Dim prc As New Principal
                                    prc.Show()
                                ElseIf CboAmbiente.SelectedIndex = 1 Then
                                    'valida a permissao de acesso para area de usuarios
                                    If log.getTotalLinhasPermissao > 0 Then
                                        Me.Hide()
                                        Dim usr As New frmUser
                                        usr.Show()
                                    Else
                                        'notifica��o de senha incorreta
                                        ger.erro("Permiss�o Negada.")
                                        txtlog.Text = String.Empty
                                        txtsenha.Text = String.Empty
                                        txtlog.Focus()
                                    End If
                                End If
                            End If
                        Else
                            'notifica��o de senha incorreta
                            ger.erro("Usu�rio Inativo.")
                            txtlog.Text = String.Empty
                            txtsenha.Text = String.Empty
                            txtlog.Focus()
                        End If
                    Else
                        'notifica��o de senha incorreta
                        ger.erro("Senha Inv�lida.")
                        contagem = contagem + 1
                        boqueiaUsuario()
                        txtlog.Text = String.Empty
                        txtsenha.Text = String.Empty
                        txtlog.Focus()
                    End If

                Else
                    'notifica��o de usu�rio incorreto.
                    ger.erro("usu�rio Invalido.")
                    contagem = contagem + 1
                    boqueiaUsuario()
                    txtlog.Text = String.Empty
                    txtsenha.Text = String.Empty
                    txtlog.Focus()
                End If

            Else
                ' notifica��o de campo vazio
                ger.erro("Preencha o campo Senha ...")
                txtsenha.Focus()
            End If
        Else
            ' notifica��o de login vazio
            ger.erro("Preencha o campo Login ...")
            txtlog.Focus()
        End If


    End Sub

    Public Sub boqueiaUsuario()

        If contagem = 3 Then
            log.blockUser()
            ger.alerta("Seu usuario foi bloqueado.Contacte o Administrador do sistema.")
        End If

    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CboAmbiente.Items.Add("Sistema")
        CboAmbiente.Items.Add("Usu�rios")
        CboAmbiente.SelectedIndex = 0
    End Sub

End Class