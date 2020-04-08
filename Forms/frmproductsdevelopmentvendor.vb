Imports System.Globalization

Public Class frmproductsdevelopmentvendor

    Dim gnr As Gn1 = New Gn1()
    Public userid As String

    Private Sub frmproductsdevelopmentvendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form_Load()
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
            fillcell2(dsInvPrdoDet)

            userid = frmLogin.txtUserName.Text
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
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
                    DataGridView1.ColumnCount = 3

                    'Add Columns
                    DataGridView1.Columns(0).Name = "clPartNo"
                    DataGridView1.Columns(0).HeaderText = "Part No."
                    DataGridView1.Columns(0).DataPropertyName = "PRDPTN"

                    DataGridView1.Columns(1).Name = "clDescription"
                    DataGridView1.Columns(1).HeaderText = "Descripcion"
                    DataGridView1.Columns(1).DataPropertyName = "IMDSC"

                    DataGridView1.Columns(2).Name = "clVendorNo"
                    DataGridView1.Columns(2).HeaderText = "Vendor No."
                    DataGridView1.Columns(2).DataPropertyName = "VMVNUM"

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

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click
        Dim exMessage As String = " "
        Dim ds As DataSet
        Dim updatedRecords As Integer = 0
        Try

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("checkBoxColumn").Value = True Then
                    ds = gnr.GetDataByCodeAndPartNoProdDesc(frmProductsDevelopment.txtCode.Text, row.Cells("clPartNo").Value.ToString())
                    Dim oldVendorNo = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("VMVNUM").Ordinal)
                    If Trim(UCase(oldVendorNo)) <> Trim(UCase(row.Cells("clVendorNo").Value.ToString())) Then
                        PoQotaFunction(oldVendorNo, row.Cells("clPartNo").Value.ToString(), row.Cells("clVendorNo").Value.ToString())
                        gnr.UpdateChangedVendor(userid, row.Cells("clVendorNo").Value.ToString(), row.Cells("clPartNo").Value.ToString(), frmProductsDevelopment.txtCode.Text)
                        'update validation
                        updatedRecords += 1
                    End If
                End If
            Next

            If updatedRecords > 0 Then
                MessageBox.Show("Records Updated.", "CTP System", MessageBoxButtons.OK)
                Form_Load()
                frmProductsDevelopment.fillcell2(frmProductsDevelopment.txtCode.Text)
            Else
                MessageBox.Show("No records to update.", "CTP System", MessageBoxButtons.OK)
            End If

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub PoQotaFunction(oldVendorNo As String, partNo As String, newVendorNo As String)
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = ""
        Dim dsUpdatedData As Integer
        Dim strQueryAdd As String = "WHERE PQVND = " & Trim(newVendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "'"
        Try
            statusquote = "D-" & Status2
            Dim spacepoqota As String = String.Empty
            'Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
            Dim dsPoQota = gnr.GetPOQotaData(oldVendorNo, partNo)
            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    Dim poqSeq = dsPoQota.Tables(0).Rows(0).ItemArray(dsPoQota.Tables(0).Columns("PQSEQ").Ordinal)
                    Dim rsResult = PoQotaFunctionDuplex(newVendorNo, partNo, poqSeq)
                    If rsResult = 0 Then
                        Dim updatedSeq = CInt(poqSeq) + 1
                        dsUpdatedData = gnr.UpdatePoQoraRowVendor(oldVendorNo, newVendorNo, partNo, updatedSeq)
                        'validation result
                    Else
                        dsUpdatedData = gnr.UpdatePoQoraRowVendor(oldVendorNo, newVendorNo, partNo, poqSeq)
                        'validation result
                    End If
                Else
                    'error message
                End If
            Else
                Dim maxValue As Integer = 1
                Dim maxQotaVal = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)
                If Not String.IsNullOrEmpty(maxQotaVal) Then
                    maxValue = CInt(Trim(maxQotaVal)) + 1
                End If
                spacepoqota = "                               DEV"
                Dim rsInsertion = gnr.InsertNewPOQotaLess(partNo, newVendorNo, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), "", DateTime.Now.Day.ToString(), "", spacepoqota, 0)
                'insertion validation
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function PoQotaFunctionDuplex(newVendorNo As String, partNo As String, seqNo As String) As Integer
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = ""
        Dim strQueryAdd As String = " WHERE PQVND = " & Trim(newVendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND PQSEQ = '" & Trim(seqNo) & "'"
        Try
            statusquote = "D-" & Status2
            Dim spacepoqota As String = String.Empty
            'Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
            Dim dsPoQota = gnr.GetPOQotaDataDuplex(strQueryAdd)
            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    Return 0
                    'validation result
                Else
                    Return -1
                    'error message
                End If
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
    End Function

End Class