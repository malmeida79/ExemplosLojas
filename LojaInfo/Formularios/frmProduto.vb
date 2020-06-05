Imports LojaInfoDll
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataGridView

Public Class frmProduto
    Dim total As Decimal
    Dim ger As New Geral()
    Dim prd As New clsProduto()

    Private Sub btnbAdiciona_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbAdiciona.Click

        dtgItemProd.AllowUserToAddRows = False
        dtgItemProd.RowHeadersVisible = False
        dtgItemProd.AllowUserToDeleteRows = False
        dtgItemProd.AllowUserToResizeColumns = False

        If dtgItemProd.Columns.Count <= 0 Then

            dtgItemProd.Columns.Add("Teste0", "Codigo")
            dtgItemProd.Columns.Add("Teste1", "ITEM")
            dtgItemProd.Columns.Add("Teste2", "QTDE")
            dtgItemProd.Columns.Add("Teste3", "VALOR")
            dtgItemProd.Columns.Add("Teste4", "Local")

            dtgItemProd.Columns(1).Width = 170
            dtgItemProd.Columns(2).Width = 50
            dtgItemProd.Columns(3).Width = 110

            dtgItemProd.Columns(0).ReadOnly = True
            dtgItemProd.Columns(0).Visible = False
            dtgItemProd.Columns(4).Visible = False

            dtgItemProd.Columns(1).ReadOnly = True
            dtgItemProd.Columns(3).ReadOnly = True
            dtgItemProd.Columns(4).ReadOnly = True

        End If

        cmbValor.SelectedValue = cmbSGrp.SelectedValue

        somaTotal(txtQtd.Text, Decimal.Parse(cmbValor.Text))
        dtgItemProd.Rows.Add(cmbSGrp.SelectedValue, cmbSGrp.Text, txtQtd.Text, FormatCurrency(cmbValor.Text))

        txtQtd.Text = String.Empty
        txtQtd.Focus()

    End Sub

    Private Sub btnExclui_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExclui.Click

        If dtgItemProd.Rows.Count > 0 Then
            If dtgItemProd.CurrentRow.Index > -1 And dtgItemProd.CurrentRow.Index < dtgItemProd.Rows.Count Then
                If Not String.IsNullOrEmpty(dtgItemProd.CurrentRow.Cells(2).Value) And Not String.IsNullOrEmpty(dtgItemProd.CurrentRow.Cells(3).Value) Then
                    tiraTotal(Short.Parse(dtgItemProd.CurrentRow.Cells(2).Value), Decimal.Parse(dtgItemProd.CurrentRow.Cells(3).Value.ToString.Replace("R$", "")))
                Else
                    lbltotal.Text = "R$ 0,00"
                End If
                dtgItemProd.Rows.Remove(dtgItemProd.CurrentRow)
            End If
        End If

    End Sub


    Public Sub loadCombos()

        Dim arrProdutos As New ArrayList
        Dim arrValores As New ArrayList
        Dim arrLocais As New ArrayList

        Dim drPrd As SqlDataReader
        drPrd = prd.carregaSGrp()

        Dim drVlPrd As SqlDataReader
        drVlPrd = prd.carregaValor()

        Dim drLocais As SqlDataReader
        drLocais = prd.carregaEstoque()

        While drPrd.Read
            arrProdutos.Add(New Drop(drPrd(1).ToString(), Convert.ToInt32(drPrd(0).ToString())))
        End While

        While drVlPrd.Read
            arrValores.Add(New Drop(drVlPrd(1).ToString(), Convert.ToInt32(drVlPrd(0).ToString())))
        End While

        While drLocais.Read
            arrLocais.Add(New Drop(drLocais(1).ToString(), Convert.ToInt32(drLocais(0).ToString())))
        End While

        cmbSGrp.DataSource = arrProdutos
        cmbSGrp.ValueMember = "valor"
        cmbSGrp.DisplayMember = "nome"

        cmbValor.DataSource = arrValores
        cmbValor.ValueMember = "valor"
        cmbValor.DisplayMember = "nome"

        cmbLocalEstoque.DataSource = arrLocais
        cmbLocalEstoque.ValueMember = "valor"
        cmbLocalEstoque.DisplayMember = "nome"

    End Sub


    Private Sub frmProduto_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCombos()
        lbltotal.Text = "R$ 0,00"
        dtgItemProd.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
    End Sub

    Public Sub somaTotal(ByVal unidade As Short, ByVal valor As Decimal)
        total = total + (valor * unidade)
        lbltotal.Text = FormatCurrency(total)
    End Sub

    Public Sub tiraTotal(ByVal unidade As Short, ByVal valor As Decimal)
        total = total - (valor * unidade)
        lbltotal.Text = FormatCurrency(total)
    End Sub

    Private Sub btnFechar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFechar.Click
        If ger.duvida("Deseja realmente sair ?") = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub remover(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dtgItemProd.CellBeginEdit
        tiraTotal(Short.Parse(dtgItemProd.CurrentCell.Value), Decimal.Parse(dtgItemProd.Rows(e.RowIndex).Cells(3).Value.ToString.Replace("R$", "")))
    End Sub

    Private Sub alterar(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dtgItemProd.CellEndEdit
        somaTotal(Short.Parse(dtgItemProd.CurrentCell.Value), Decimal.Parse(dtgItemProd.Rows(e.RowIndex).Cells(3).Value.ToString.Replace("R$", "")))
    End Sub


    Private Sub btnSalvar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalvar.Click

        For Each item As DataGridViewRow In dtgItemProd.Rows
            If Not String.IsNullOrEmpty(item.Cells(1).Value) Then
                MsgBox(item.Cells(1).Value)
            End If
        Next

    End Sub

End Class