Public Class frmproductsdevelopmentstatus

    Dim gnr As Gn1 = New Gn1()
    Public userid As String

    Private Sub frmproductsdevelopmentstatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

            FillDDLStatus()

            Dim codeproject = frmProductsDevelopment.txtCode.Text
            lblproject.Text = frmProductsDevelopment.txtCode.Text & " - " & Trim(frmProductsDevelopment.txtname.Text)

            'check delete temp

            Dim dsInvPrdoDet = gnr.GetInvProdDetailByProject(codeproject)
            fillcell2(dsInvPrdoDet)

            userid = frmLogin.txtUserName.Text
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub FillDDLStatus()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsStatuses = gnr.GetAllStatuses()

            dsStatuses.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsStatuses.Tables(0).Rows.Count - 1
                If dsStatuses.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsStatuses.Tables(0).Rows(i).Item(2).ToString() + " -- " + dsStatuses.Tables(0).Rows(i).Item(3).ToString()
                    'CleanUser = Trim(dsStatuses.Tables(0).Rows(i).Item(0).ToString())
                    dsStatuses.Tables(0).Rows(i).Item(5) = fllValueName
                    'dsStatuses.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next

            cmbstatus.DataSource = dsStatuses.Tables(0)
            cmbstatus.DisplayMember = "FullValue"
            cmbstatus.ValueMember = "CNT03"

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) _
    Handles DataGridView1.CellFormatting
        Dim CurrentState As String = ""
        If e.ColumnIndex = 5 Then
            If e.Value IsNot Nothing Then
                CurrentState = e.Value.ToString
                If CurrentState = "A" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Approved"
                ElseIf CurrentState = "R " Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Rejected"
                ElseIf CurrentState = "NS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Negotiation with Supplier"
                ElseIf CurrentState = "RP" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Receiving of First Production"
                ElseIf CurrentState = "CS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed Successfully"
                ElseIf CurrentState = "CN" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed-Approved w/o Negotiation"
                ElseIf CurrentState = "CD" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed-Rejected"
                ElseIf CurrentState = "CL" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed w/o negotiation"
                ElseIf CurrentState = "AA" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Approved with advice"
                ElseIf CurrentState = "Q" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Quoting"
                ElseIf CurrentState = "TD" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Technical Documentation"
                ElseIf CurrentState = "DP" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Documentation in Process"
                ElseIf CurrentState = "DF" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Documentation Finalized"
                ElseIf CurrentState = "SS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Sample already Sent"
                ElseIf CurrentState = "PS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Pending from Supplier"
                ElseIf CurrentState = "AS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Analysis of Samples"
                    'ElseIf CurrentState = "NS" Then
                    '    DataGridView1.Rows(e.RowIndex).Cells("Status").Value = "Negotiation with Supplier"
                End If
            End If
        End If
    End Sub

    Private Sub fillcell2(ds As DataSet)
        Dim exMessage As String = " "
        Try

            If Not ds Is Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    DataGridView1.Columns.Clear()
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 5

                    'Add Columns
                    DataGridView1.Columns(0).Name = "clPartNo"
                    DataGridView1.Columns(0).HeaderText = "Part No."
                    DataGridView1.Columns(0).DataPropertyName = "PRDPTN"

                    DataGridView1.Columns(1).Name = "clCtpNo"
                    DataGridView1.Columns(1).HeaderText = "CTP No."
                    DataGridView1.Columns(1).DataPropertyName = "PRDCTP"

                    DataGridView1.Columns(2).Name = "clMfrNo"
                    DataGridView1.Columns(2).HeaderText = "MFR No."
                    DataGridView1.Columns(2).DataPropertyName = "PRDMFR#"

                    DataGridView1.Columns(3).Name = "clDescription"
                    DataGridView1.Columns(3).HeaderText = "Descripcion"
                    DataGridView1.Columns(3).DataPropertyName = "IMDSC"

                    DataGridView1.Columns(4).Name = "clStatus"
                    DataGridView1.Columns(4).HeaderText = "Status"
                    DataGridView1.Columns(4).DataPropertyName = "PRDSTS"

                    'FILL GRID
                    DataGridView1.DataSource = ds.Tables(0)

                    Dim headerCellLocation As Point = Me.DataGridView1.GetCellDisplayRectangle(0, -1, True).Location

                    'Place the Header CheckBox in the Location of the Header Cell.
                    Dim headerCheckBox As New CheckBox
                    headerCheckBox.Location = New Point(headerCellLocation.X + 8, headerCellLocation.Y + 2)
                    headerCheckBox.BackColor = Color.White
                    headerCheckBox.Size = New Size(18, 18)

                    'Assign Click event to the Header CheckBox.
                    AddHandler headerCheckBox.Click, AddressOf HeaderCheckBox_Clicked
                    DataGridView1.Controls.Add(headerCheckBox)

                    'Add a CheckBox Column to the DataGridView at the first position.
                    Dim checkBoxColumn As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                    checkBoxColumn.HeaderText = "All"
                    checkBoxColumn.Width = 30
                    checkBoxColumn.Name = "checkBoxColumn"
                    DataGridView1.Columns.Insert(0, checkBoxColumn)

                    DataGridView1.Columns("clPartNo").ReadOnly = True
                    DataGridView1.Columns("clDescription").ReadOnly = True
                    DataGridView1.Columns("clMfrNo").ReadOnly = True
                    DataGridView1.Columns("clCtpNo").ReadOnly = True

                Else
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If
        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub HeaderCheckBox_Clicked(ByVal sender As Object, ByVal e As EventArgs)
        'Necessary to end the edit mode of the Cell.
        DataGridView1.EndEdit()

        'Loop and check and uncheck all row CheckBoxes based on Header Cell CheckBox.
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim checkBox As DataGridViewCheckBoxCell = (TryCast(row.Cells("checkBoxColumn"), DataGridViewCheckBoxCell))

            Dim myItem As CheckBox = CType(sender, CheckBox)
            checkBox.Value = myItem.Checked
        Next
    End Sub

    Private Sub cmdexit1_Click(sender As Object, e As EventArgs) Handles cmdexit1.Click
        Me.Close()
    End Sub

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click

    End Sub

    Private Sub cmdSelectAll_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class