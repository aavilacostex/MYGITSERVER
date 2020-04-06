Imports System.Globalization

Public Class frmproductsdevelopmentmanu

    Dim gnr As Gn1 = New Gn1()

    Private Sub frmproductsdevelopmentmanu_Load(sender As Object, e As EventArgs)
        Form_Load()
    End Sub

    Private Sub Form_Load()
        Dim exMessage As String = " "
        Try
            If gnr.ConnSql.State = 1 Then
            Else
                gnr.ConnSql.ConnectionString = gnr.strconnSQL
                gnr.ConnSql.Open()
            End If

            Dim codeproject = frmProductsDevelopment.txtCode.Text
            'lblproject.Text = frmProductsDevelopment.txtCode.Text & " - " & Trim(frmProductsDevelopment.txtname.Text)

            'check delete temp


            Dim dsInvPrdoDet = gnr.GetInvProdDetailByProject(codeproject)
            'fillcell2(dsInvPrdoDet)

            'If dsInvPrdoDet IsNot Nothing Then
            '    If dsInvPrdoDet.Tables(0).Rows.Count > 0 Then
            '        For Each ttt As DataRow In dsInvPrdoDet.Tables(0).Rows


            '        Next
            '    End If
            'End If


        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView1.DoubleClick
        Dim Index As Integer
        Dim ds As New DataSet()
        Dim ds1 As New DataSet()
        Dim RowDs As DataRow
        ds.Locale = CultureInfo.InvariantCulture
        ds1.Locale = CultureInfo.InvariantCulture
        Dim exMessage As String = " "


        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                Index = DataGridView1.CurrentCell.RowIndex
                If DataGridView1.Rows(Index).Selected = True Then
                    Me.DataGridView1.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight
                End If
            Next
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub frmproductsdevelopmentmanu_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class