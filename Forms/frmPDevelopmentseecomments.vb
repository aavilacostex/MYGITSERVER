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
            'Dim code = Trim(3245)
            'Dim partNO = Trim("1260588")
            'Dim dsProdHeaderMess = gnr.GetDataByCodAndPartProdAndComm1(code, partNO)

            Dim dsProdHeaderMess = gnr.GetDataByCodAndPartProdAndComm1(Trim(txtCode.Text), Trim(txtpartno.Text))
            fillDgvProjectMessages(dsProdHeaderMess)
            'TAB()

            TabPage1.Text = ""
            'SSTab1.TabCaption(0) = ""
            'tabtab.TabCaption(1) = ""

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub SSTab1_Selected(ByVal sender As Object, ByVal e As TabControlEventArgs) _
    Handles SSTab1.Selected
        If SSTab1.SelectedTab.Name = "TabPage1" Then
            cmddelete.Enabled = False
        Else
            cmddelete.Enabled = True
        End If
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

                dgvProjectMessages.Columns(1).Name = "clDateEntered"
                dgvProjectMessages.Columns(1).HeaderText = "Date Entered"
                dgvProjectMessages.Columns(1).DataPropertyName = "PRDCDA"

                dgvProjectMessages.Columns(2).Name = "clTimeEntered"
                dgvProjectMessages.Columns(2).HeaderText = "Time Entered"
                dgvProjectMessages.Columns(2).DataPropertyName = "PRDCTI"

                dgvProjectMessages.Columns(3).Name = "clUser"
                dgvProjectMessages.Columns(3).HeaderText = "User"
                dgvProjectMessages.Columns(3).DataPropertyName = "USUSER"

                dgvProjectMessages.Columns(4).Name = "clTableCode"
                dgvProjectMessages.Columns(4).HeaderText = "Table Code"
                dgvProjectMessages.Columns(4).DataPropertyName = "PRDCCO"

                dgvProjectMessages.DataSource = dsData.Tables(0)
            End If

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub

    Public Sub fillDgvProjectMessage2(dsData As DataSet)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            If dsData.Tables(0).Rows.Count > 0 Then
                dgvProjectMessage2.DataSource = Nothing
                dgvProjectMessage2.Refresh()
                dgvProjectMessage2.AutoGenerateColumns = False

                'Add Columns
                dgvProjectMessage2.Columns(1).Name = "clTableCode"
                dgvProjectMessage2.Columns(1).HeaderText = "Table Code"
                dgvProjectMessage2.Columns(1).DataPropertyName = "PRDCCO"

                dgvProjectMessage2.Columns(2).Name = "clCommentNo"
                dgvProjectMessage2.Columns(2).HeaderText = "Comment No"
                dgvProjectMessage2.Columns(2).DataPropertyName = "PRDCDC"

                dgvProjectMessage2.Columns(0).Name = "clComments"
                dgvProjectMessage2.Columns(0).HeaderText = "Comments"
                dgvProjectMessage2.Columns(0).DataPropertyName = "PRDCTX"

                dgvProjectMessage2.DataSource = dsData.Tables(0)


                'Dim headerCellLocation As Point = Me.dgvProjectMessage2.GetCellDisplayRectangle(0, -1, True).Location
                ''Place the Header CheckBox in the Location of the Header Cell.
                'headerCheckBox.Location = New Point(headerCellLocation.X + 8, headerCellLocation.Y + 2)
                'headerCheckBox.BackColor = Color.White
                'headerCheckBox.Size = New Size(18, 18)

                Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                checkBoxColumn.HeaderText = ""
                checkBoxColumn.Width = 30
                checkBoxColumn.Name = "checkBoxColumn"
                dgvProjectMessage2.Columns.Insert(0, checkBoxColumn)
                'AddHandler dgvProjectMessage2.CellContentClick, AddressOf dgvProjectMessage2_CellClick

            End If

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub

    Private Sub dgvProjectMessages_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles dgvProjectMessages.DoubleClick
        Dim exMessage As String = " "
        Dim Index As Integer
        Dim dsResult As DataSet
        Try
            Dim rsDeletionSql = gnr.DeleteDataSqlByUser("tbtemppdevelopmentseecomment", userid)
            Dim dsSelection = gnr.GetDataSqlByUser("tbtemppdevelopmentseecomment", userid)



            For Each row As DataGridViewRow In dgvProjectMessages.SelectedRows
                Index = dgvProjectMessages.CurrentCell.RowIndex

                If dgvProjectMessages.Rows(Index).Selected = True Then
                    Dim tableCode As String = row.Cells(4).Value.ToString()
                    dsResult = gnr.GetDataByCodAndPartProdAndComm2(tableCode)
                    If dsResult IsNot Nothing Then
                        If dsResult.Tables(0).Rows.Count > 0 Then
                            fillDgvProjectMessage2(dsResult)
                            SSTab1.SelectedIndex = 1
                        End If
                    End If
                End If
            Next


        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

End Class