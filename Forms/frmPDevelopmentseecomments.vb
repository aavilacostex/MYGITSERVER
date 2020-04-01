Imports System.Globalization

Public Class frmPDevelopmentseecomments

    Dim gnr As Gn1 = New Gn1()
    Dim sql As String
    Public userid As String
    Public flagallow As Integer
    Public cod_detcomment As Integer

    Private Sub frmPDevelopmentseecomments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim exMessage As String = " "
        cmddelete.Enabled = False
        Try
            userid = frmLogin.txtUserName.Text
            If UCase(userid) = "AALZATE" Then
                flagallow = 1
            End If

            Dim rsDeletionSql = gnr.DeleteDataSqlByUser("tbtemppdevelopmentseecomment", userid)
            Dim dsSelection = gnr.GetDataSqlByUser("tbtemppdevelopmentseecomment", userid)

            gnr.seeaddprocomments = lblNotVisible.Text
            If gnr.seeaddprocomments = 5 Then
                txtCode.Text = frmProductsDevelopment.txtCode.Text
                txtpartno.Text = Trim(UCase(frmProductsDevelopment.txtpartno.Text))
            End If

            'test purpose
            Dim code = Trim(txtCode.Text)
            Dim partNO = Trim(txtpartno.Text)
            Dim dsProdHeaderMess = gnr.GetDataByCodAndPartProdAndComm1(code, partNO)

            'Dim dsProdHeaderMess = gnr.GetDataByCodAndPartProdAndComm1(Trim(txtCode.Text), Trim(txtpartno.Text))
            fillDgvProjectMessages(dsProdHeaderMess)

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub


    Public Sub fillDgvProjectMessages(dsData As DataSet)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            If dsData.Tables(0).Rows.Count > 0 Then
                dgvProjectMessages.DataSource = Nothing
                dgvProjectMessages.Refresh()
                dgvProjectMessages.AutoGenerateColumns = False

                'Add Columns
                dgvProjectMessages.Columns(0).Name = "clSubject"
                dgvProjectMessages.Columns(0).HeaderText = "Subject"
                dgvProjectMessages.Columns(0).DataPropertyName = "PRDCSU"

                dgvProjectMessages.Columns(0).Name = "clDateEntered"
                dgvProjectMessages.Columns(0).HeaderText = "Date Entered"
                dgvProjectMessages.Columns(0).DataPropertyName = "PRDCDA"

                dgvProjectMessages.Columns(0).Name = "clTimeEntered"
                dgvProjectMessages.Columns(0).HeaderText = "Time Entered"
                dgvProjectMessages.Columns(0).DataPropertyName = "PRDCTI"

                dgvProjectMessages.Columns(0).Name = "clUser"
                dgvProjectMessages.Columns(0).HeaderText = "User"
                dgvProjectMessages.Columns(0).DataPropertyName = "USUSER"

                dgvProjectMessages.DataSource = dsData.Tables(0)
            End If

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub
End Class