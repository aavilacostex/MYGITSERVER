Public Class frmproductsdevelopmentvendor

    Dim gnr As Gn1 = New Gn1()

    Private Sub frmproductsdevelopmentvendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdexit1_Click(sender As Object, e As EventArgs) Handles cmdexit1.Click
        Me.Close()
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
            lblproject.Text = frmProductsDevelopment.txtCode.Text & " - " & Trim(frmProductsDevelopment.txtname.Text)

            'check delete temp


            Dim dsInvPrdoDet = gnr.GetInvProdDetailByProject(codeproject)
            If dsInvPrdoDet IsNot Nothing Then
                If dsInvPrdoDet.Tables(0).Rows.Count > 0 Then
                    For Each ttt As DataRow In dsInvPrdoDet.Tables(0).Rows


                    Next
                End If
            End If


        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

End Class