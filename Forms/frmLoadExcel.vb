Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports Microsoft.Office.Interop

Public Class frmLoadExcel

    Private Excel03ConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'"
    Private Excel07ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'"
    Dim gnr As Gn1 = New Gn1()
    Public userid As String
    Public flagallow As Integer

    Private Const totalRecords As Integer = 43
    Private Const pageSize As Integer = 10


#Region "Page Load"

    Private Sub frmLoadExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim exMessage As String = " "
        Try
            userid = frmLogin.txtUserName.Text
            If UCase(userid) = "AALZATE" Then
                flagallow = 1
            End If

            btnSuccess.Enabled = False
            btnInsert.Enabled = False
            btnCheck.Enabled = False
            btnSelect.Enabled = False
            dtProjectDate.Value = Now

            cmbStatus.Items.Add("-- Select Status --")
            cmbStatus.Items.Add("I - In Process")
            cmbStatus.Items.Add("F - Finished")
            cmbStatus.SelectedIndex = 1
            FillDDlUser1()

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Gridview,  dropdowns and textboxes methods"

    Private Sub fillData(dt As DataTable)
        Dim exMessage As String = " "
        Dim mandatoryMissed As String = String.Empty
        Dim dsResult As DataSet = New DataSet()
        Dim dsError As DataSet = New DataSet()
        Dim dtError As DataTable = New DataTable()
        Dim dtResult As DataTable = New DataTable()
        Try
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then

                    Dim dictionary As New Dictionary(Of String, String)
                    'dictionary.Add("PRNAME", "Project Name")
                    dictionary.Add("PRDPTN", "Part Number")
                    dictionary.Add("VMVNUM", "Vendor Number")
                    'Dim lstRequiredColumns As New List(Of String)({"PRNAME", "PRDPTN", "VMVNUM"})
                    For Each pair As KeyValuePair(Of String, String) In dictionary
                        If dt.Columns(pair.Key) Is Nothing Then
                            mandatoryMissed += pair.Value + " is missed. ,"
                        End If
                    Next

                    If Not String.IsNullOrEmpty(mandatoryMissed) Then
                        mandatoryMissed.Insert(0, "The mandatory fields must be filled. ")
                        mandatoryMissed.Remove(mandatoryMissed.LastIndexOf(","), 1)
                        MessageBox.Show(mandatoryMissed, "CTP System", MessageBoxButtons.OK)
                    Else
                        dtError = dt.Clone()
                        dtResult = dt.Clone()
                        dsError.Tables.Add(dtError)
                        dsResult.Tables.Add(dtResult)

                        For Each item As DataRow In dt.Rows
                            If String.IsNullOrEmpty(item.ItemArray(dt.Columns("PRPECH").Ordinal).ToString()) Then
                                item.Item(dt.Columns("PRPECH").Ordinal) = userid
                            End If
                            If String.IsNullOrEmpty(item.ItemArray(dt.Columns("PRDPTN").Ordinal).ToString()) Or
                                String.IsNullOrEmpty(item.ItemArray(dt.Columns("VMVNUM").Ordinal).ToString()) Then
                                dsError.Tables(0).ImportRow(item)
                            Else
                                If checkIfPartAndVdrExist(item.ItemArray(dt.Columns("PRDPTN").Ordinal).ToString(), item.ItemArray(dt.Columns("VMVNUM").Ordinal).ToString()) Then
                                    dsError.Tables(0).ImportRow(item)
                                Else
                                    dsResult.Tables(0).ImportRow(item)
                                End If
                            End If
                        Next

                        LikeSession.dsErrorSession = dsError
                        LikeSession.dsResultsSession = dsResult

                        fillcell1(dsResult.Tables(0), 0)
                    End If
                Else
                    MessageBox.Show("Error reading excel data.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("Error reading excel data.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtProjectName_TextChanged(sender As Object, e As EventArgs) Handles txtProjectName.TextChanged
        If Not String.IsNullOrEmpty(txtProjectName.Text) Then
            btnSelect.Enabled = True
        Else
            btnSelect.Enabled = False
        End If
    End Sub

    Private Sub fillcell1(dt As DataTable, flag As Integer)
        Dim exMessage As String = " "
        Try
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            DataGridView1.AutoGenerateColumns = False
            DataGridView1.ColumnCount = 6

            'Add Columns
            DataGridView1.Columns(0).Name = "clPRHCOD"
            DataGridView1.Columns(0).HeaderText = "Project No."
            DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

            DataGridView1.Columns(1).Name = "clPRDPTN"
            DataGridView1.Columns(1).HeaderText = "Part No."
            DataGridView1.Columns(1).DataPropertyName = "PRDPTN"

            DataGridView1.Columns(2).Name = "clPRDCTP"
            DataGridView1.Columns(2).HeaderText = "CTP No."
            DataGridView1.Columns(2).DataPropertyName = "PRDCTP"

            DataGridView1.Columns(3).Name = "clPRDMFR"
            DataGridView1.Columns(3).HeaderText = "Manufacturer No."
            DataGridView1.Columns(3).DataPropertyName = "PRDMFR#"

            DataGridView1.Columns(4).Name = "clVMVNUM"
            DataGridView1.Columns(4).HeaderText = "Vendor No."
            DataGridView1.Columns(4).DataPropertyName = "VMVNUM"

            DataGridView1.Columns(5).Name = "clPRDSTS"
            DataGridView1.Columns(5).HeaderText = "Status"
            DataGridView1.Columns(5).DataPropertyName = "PRDSTS"

            'FILL GRID
            DataGridView1.DataSource = dt

            If flag = 0 Then
                btnInsert.Enabled = True
                btnCheck.Enabled = True
            End If

        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) _
    Handles DataGridView1.CellFormatting
        Dim CurrentState As String = ""
        Dim NewState As String = ""
        If e.ColumnIndex = 5 Then
            If Not IsDBNull(Trim(e.Value.ToString())) Then
                CurrentState = e.Value.ToString
                NewState = buildStatusString(CurrentState)
                DataGridView1.Rows(e.RowIndex).Cells("clPRDSTS").Value = NewState
            Else
                CurrentState = "E"
                NewState = buildStatusString(CurrentState)
                DataGridView1.Rows(e.RowIndex).Cells("clPRDSTS").Value = NewState
            End If
        End If
    End Sub

    Private Sub FillDDlUser1()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsUser = gnr.FillDDLUser()

            dsUser.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsUser.Tables(0).Rows.Count - 1
                If dsUser.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsUser.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsUser.Tables(0).Rows(i).Item(1).ToString()
                    CleanUser = Trim(dsUser.Tables(0).Rows(i).Item(0).ToString())
                    dsUser.Tables(0).Rows(i).Item(2) = fllValueName
                    dsUser.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next


            Dim newRow As DataRow = dsUser.Tables(0).NewRow
            newRow("USUSER") = "N/A"
            newRow("USNAME") = "NO NAME"
            newRow("FullValue") = "N/A -- NO NAME"
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow, 0)

            cmbPerCharge.DataSource = dsUser.Tables(0)
            cmbPerCharge.DisplayMember = "FullValue"
            cmbPerCharge.ValueMember = "USUSER"


        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

#End Region

#Region "Excel process"

    Private Sub openFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim exMessage As String = " "
        Try
            Dim filePath As String = OpenFileDialog1.FileName
            Dim extension As String = Path.GetExtension(filePath)
            'Dim header As String = If(rbHeaderYes.Checked, "YES", "NO")
            Dim conStr As String, sheetName As String

            conStr = String.Empty
            Select Case extension

                Case ".xls"
                    'Excel 97-03
                    conStr = String.Format(Excel03ConString, filePath, "YES")
                    Exit Select

                Case ".xlsx"
                    'Excel 07
                    conStr = String.Format(Excel07ConString, filePath, "YES")
                    Exit Select
            End Select

            'Get the name of the First Sheet.
            Using con As New OleDbConnection(conStr)
                Using cmd As New OleDbCommand()
                    cmd.Connection = con
                    con.Open()
                    Dim dtExcelSchema As DataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                    sheetName = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
                    con.Close()
                End Using
            End Using

            'Read Data from the First Sheet.
            Using con As New OleDbConnection(conStr)
                Using cmd As New OleDbCommand()
                    Using oda As New OleDbDataAdapter()
                        Dim dt As New DataTable()
                        cmd.CommandText = (Convert.ToString("SELECT * From [") & sheetName) + "]"
                        cmd.Connection = con
                        con.Open()
                        oda.SelectCommand = cmd
                        oda.TableMappings.Add("Table", "Net-informations.com")
                        oda.Fill(dt)
                        LikeSession.dsData = dt
                        fillData(dt)
                        con.Close()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub


#End Region

#Region "button methods"

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Dim exMessage As String = " "
        Dim countErrors As Integer = 0
        Dim Qry As New DataTable
        Dim arraySuccess As New List(Of Integer)
        Dim arrayError As New List(Of Integer)
        Try
            If String.IsNullOrEmpty(txtProjectName.Text) Then
                MessageBox.Show("The Project Name is a required field.", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If

            'Dim dt As New DataTable
            'dt = (DirectCast(DataGridView1.DataSource, DataTable))

            Dim dsResult = LikeSession.dsResultsSession
            If dsResult IsNot Nothing Then
                If dsResult.Tables(0).Rows.Count <= 0 Then
                    MessageBox.Show("There is an error in the data.", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                MessageBox.Show("There is an error in the data.", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If

            Dim maxProjectNo = gnr.getmax("PRDVLH", "PRHCOD")
            Dim ProjectNoCurrent = CInt(maxProjectNo) + 1
            Dim projectPerCharge = If(cmbPerCharge.SelectedIndex = 0, userid, cmbPerCharge.SelectedValue)

            Dim dsExistsProject = gnr.GetExistByPRNAME(txtProjectName.Text)
            If dsExistsProject IsNot Nothing Then
                'warning message project name exists
            Else
                Dim queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, txtDesc.Text, txtProjectName.Text, cmbStatus, projectPerCharge)
                If queryResult < 0 Then
                    'error message insertion
                Else
                    For Each tt As DataRow In dsResult.Tables(0).Rows
#Region "not in use validate"

                        'If dsExistsProject.Tables(0).Rows.Count > 0 Then
                        '    'update

                        'Else
                        '    'insert
                        '    Dim maxProjectNo = gnr.getmax("PRDVLH", "PRHCOD")
                        '    Dim ProjectNoCurrent = CInt(maxProjectNo) + 1



                        '    Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
                        '                 .Where(Function(x) Trim(UCase(x.Field(Of String)("PRNAME")).ToString()) = Trim(UCase(txtProjectName.Text)) And
                        '                 Trim(UCase(x.Field(Of Double)("PRDPTN"))) = Trim(UCase(partNo)))

                        '    If Qry1.Count > 0 Then
                        '        Qry = Qry1.CopyToDataTable

                        '        Dim projectNameValue = txtProjectName.Text
                        '        Dim personInChargeValue = Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString()
                        '        Dim detailsValue = txtDesc.Text

                        '        Dim queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, detailsValue, projectNameValue, cmbStatus, personInChargeValue)
                        '        If queryResult < 0 Then
                        '            'error message insertion
                        '        Else
                        '            Dim rsInsert = InsertProductDetails(Qry, ProjectNoCurrent)
                        '            If rsInsert > 0 Then
                        '                'delete project no
                        '                Dim rsDelete = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
                        '                If rsDelete < 0 Then
                        '                    'error
                        '                End If
                        '                countErrors += rsInsert
                        '                arrayError.Add(ProjectNoCurrent)
                        '            Else
                        '                If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                        '                    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                        '                End If

                        '                tt("PRHCOD") = ProjectNoCurrent
                        '                dsResult.AcceptChanges()
                        '                arraySuccess.Add(ProjectNoCurrent)
                        '            End If
                        '            'countErrors += InsertProductDetails(Qry)
                        '        End If
                        '    Else
                        '        MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                        '    End If


                        '    'If Qry IsNot Nothing Then
                        '    '    If Qry.Rows.Count > 0 Then

                        '    '    Else
                        '    '        MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                        '    '    End If
                        '    'Else
                        '    '    MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                        '    'End If
                        'End If

#End Region

                        'insert
                        Dim partNo = tt.Item(dsResult.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                        Dim vendorNo = tt.Item(dsResult.Tables(0).Columns("VMVNUM").Ordinal).ToString()

                        Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
                                             .Where(Function(x) Trim(UCase(x.Field(Of Double)("VMVNUM")).ToString()) = Trim(UCase(vendorNo)) And
                                             Trim(UCase(x.Field(Of Double)("PRDPTN"))) = Trim(UCase(partNo)))

                        If Qry1.Count > 0 Then
                            Qry = Qry1.CopyToDataTable
                            Dim personInChargeValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString()), userid, Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString())

                            Dim rsInsert = InsertProductDetails(Qry, ProjectNoCurrent)
                            If rsInsert > 0 Then
                                'delete project no
                                Dim rsDelete = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
                                If rsDelete < 0 Then
                                    'error borrando
                                End If
                                countErrors += rsInsert
                                arrayError.Add(ProjectNoCurrent)
                            Else
                                'right insertion
                                If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                                    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                                End If

                                tt("PRHCOD") = ProjectNoCurrent
                                dsResult.AcceptChanges()

                                txtProjectNo.Text = ProjectNoCurrent
                                If cmbPerCharge.FindStringExact(Trim(projectPerCharge)) Then
                                    cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(projectPerCharge))
                                End If


                                arraySuccess.Add(ProjectNoCurrent)
                            End If
                            'countErrors += InsertProductDetails(Qry)

                        Else
                            MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                        End If
                    Next
                End If
            End If

            If countErrors > 0 Then
                MessageBox.Show("The insertion process finished with some fails inserting data.", "CTP System", MessageBoxButtons.OK)
            Else
                MessageBox.Show("The insertion process finished successfully.", "CTP System", MessageBoxButtons.OK)
                disableAfterInsert()
            End If
            'cleanFormValues()

            'LikeSession.dsData = dsProcess
            'Dim dsRestore = LikeSession.dsData
            'Dim dtTemp = New DataTable()
            'dtTemp = dsRestore.Clone()
            'For Each item As Integer In arraySuccess
            '    Dim Qry1 = dsRestore.AsEnumerable() _
            '                         .Where(Function(x) Trim(UCase(x.Field(Of Integer)("PRHCOD")).ToString()) = Trim(UCase(item).ToString()))
            '    If Qry1.Count > 0 Then

            '        dtTemp.Rows.Add(Qry1)
            '    End If
            'Next
            'DataGridView1.DataSource = dtTemp
            'DataGridView1.Refresh()

            'lblMessage.Text = arraySuccess.Count & ": Records Inserted Successfully."
            'lblMessage.Visible = True
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub btnCheck_Click(sender As Object, e As EventArgs) Handles btnCheck.Click
        Dim exMessage As String = " "
        Try
            Dim dsValue = LikeSession.dsErrorSession
            fillcell1(dsValue.Tables(0), 1)
            btnSuccess.Enabled = True
            btnCheck.Enabled = False
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub btnSuccess_Click(sender As Object, e As EventArgs) Handles btnSuccess.Click
        Dim exMessage As String = " "
        Try
            Dim dsValue = LikeSession.dsResultsSession
            fillcell1(dsValue.Tables(0), 0)
            btnSuccess.Enabled = False
            btnCheck.Enabled = True
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Utils"

    Private Function checkIfPartAndVdrExist(partNo As String, vendorNo As String) As Boolean
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Dim rsReturn As Boolean = False
        Try
            ds = gnr.GetDataByVendorAndPartNoProdDesc(partNo, vendorNo)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    Return rsReturn = True
                End If
                Return False
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return rsReturn
        End Try
    End Function

    Private Function InsertProductDetails(Qry As DataTable, code As String) As Integer
        Dim dtTime As DateTimePicker = New DateTimePicker()
        Dim dtTime1 As DateTimePicker = New DateTimePicker()
        Dim dtTime2 As DateTimePicker = New DateTimePicker()
        Dim dtTime3 As DateTimePicker = New DateTimePicker()
        Dim dtTime4 As DateTimePicker = New DateTimePicker()
        Dim dtTime5 As DateTimePicker = New DateTimePicker()
        Dim dtTime6 As DateTimePicker = New DateTimePicker()
        Dim QueryDetailResult As Integer = -1
        Dim partstoshow As String
        Dim exMessage As String = " "
        Try

            'Dim strCheck = gnr.checkFields(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, txtqty.Text,
            '                                                    txtmfr.Text, txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
            '                                                    cmbuser.SelectedValue, chknew, dtTime3, txtsample.Text, txttcost.Text, txtvendorno.Text, partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, dtTime4,
            '                                                    dtTime5.Value.ToShortDateString(), txtsampleqty.Text)
            Dim strCheck = Nothing
            If String.IsNullOrEmpty(strCheck) Then

#Region "Variable assign"

                Dim projectNoValue = code
                Dim PartNoValue = Qry.Rows(0).ItemArray(Qry.Columns("PRDPTN").Ordinal).ToString()
                Dim CTPNoValue = Qry.Rows(0).ItemArray(Qry.Columns("PRDCTP").Ordinal).ToString()
                Dim qtyValue = Qry.Rows(0).ItemArray(Qry.Columns("PRDQTY").Ordinal).ToString()
                Dim MFRValue = ""
                Dim MFRNoValue = Qry.Rows(0).ItemArray(Qry.Columns("PRDMFR#").Ordinal).ToString()
                Dim unitcostValue = Qry.Rows(0).ItemArray(Qry.Columns("PRDCOS").Ordinal).ToString()
                Dim unitcostVValue = Qry.Rows(0).ItemArray(Qry.Columns("PRDCON").Ordinal).ToString()
                Dim vendorNoValue = Qry.Rows(0).ItemArray(Qry.Columns("VMVNUM").Ordinal).ToString()
                Dim statusValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDSTS").Ordinal).ToString()),
                    "E", Qry.Rows(0).ItemArray(Qry.Columns("PRDSTS").Ordinal).ToString())
                Dim chkControl = New CheckBox()
                Dim chkValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDNEW").Ordinal).ToString()),
                    "0", Qry.Rows(0).ItemArray(Qry.Columns("PRDNEW").Ordinal).ToString())
                Dim chkSelection = If(chkValue = "0", Not chkControl.Checked, chkControl.Checked)
                chkControl.Checked = chkSelection

                'Dim poNoValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDPO#").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDPO#").Ordinal).ToString())
                'Dim benefitsValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDBEN").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDBEN").Ordinal).ToString())
                'Dim DetailsValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDINF").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDINF").Ordinal).ToString())
                'Dim personChValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDUSR").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDUSR").Ordinal).ToString())
                'Dim samplecostValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDSCO").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDSCO").Ordinal).ToString())
                'Dim misccostValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDTTC").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDTTC").Ordinal).ToString())
                'Dim minorcodeValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDMPC").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDMPC").Ordinal).ToString())
                'Dim toolingcostValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDTCO").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDTCO").Ordinal).ToString())
                'Dim sampleQtyValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDSQTY").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDSQTY").Ordinal).ToString())
                'Dim optSelected = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDPTS").Ordinal).ToString()),
                '    "", Qry.Rows(0).ItemArray(Qry.Columns("PRDPTS").Ordinal).ToString())
                'partstoshow = displayPart(optSelected)

#End Region

#Region "Set Date values"

                dtTime.Value = Now 'PRDDAT
                dtTime.CustomFormat = "yyyy/MM/dd/"
                dtTime1.Value = Now 'CRDATE
                dtTime1.CustomFormat = "yyyy/MM/dd/"
                dtTime2.Value = Now 'MODATE
                dtTime2.CustomFormat = "yyyy/MM/dd/"
                dtTime3.Value = Now 'PODATE
                dtTime3.CustomFormat = "yyyy/MM/dd/"
                dtTime4.Value = Now 'PODATE
                dtTime4.CustomFormat = "yyyy/MM/dd/"
                dtTime5.Value = Now 'PODATE
                dtTime5.CustomFormat = "yyyy/MM/dd/"
                dtTime6.Value = Now 'PODATE
                dtTime6.CustomFormat = "yyyy/MM/dd/"

                'dtTime5.Value = New DateTime(1900, 1, 1)

#End Region

                'PRHCOD,PRDPTN,PRDDAT,CRUSER,CRDATE,MOUSER,MODATE,PRDCTP,PRDQTY,PRDMFR,PRDMFR#,PRDCOS,PRDCON,PRDPO#,PODATE,PRDSTS,PRDBEN,PRDINF,PRDUSR,PRDNEW,
                'PRDEDD, PRDSCO, PRDTTC, VMVNUM, PRDPTS, PRDMPC, PRDTCO, PRDERD, PRDPDA, PRDSQTY

                'QueryDetailResult = gnr.InsertProductDetail(projectNoValue, PartNoValue, dtTime,
                '                    userid, dtTime1, userid, dtTime2, CTPNoValue, qtyValue,
                '                    MFRValue, MFRNoValue, unitcostValue, unitcostVValue,
                '                    poNoValue, dtTime3, statusValue, benefitsValue,
                '                    DetailsValue, personChValue, chkControl, dtTime4, samplecostValue,
                '                    misccostValue, vendorNoValue, partstoshow, minorcodeValue, toolingcostValue, dtTime5,
                '                    dtTime6, If(Not String.IsNullOrEmpty(sampleQtyValue), CInt(sampleQtyValue), 0))

                QueryDetailResult = gnr.InsertProductDetail(projectNoValue, PartNoValue, dtTime,
                                    userid, dtTime1, userid, dtTime2, CTPNoValue, qtyValue,
                                    MFRValue, MFRNoValue, unitcostValue, unitcostVValue,
                                    "", dtTime3, statusValue, "",
                                    "", userid, chkControl, dtTime4, "0",
                                    "0", vendorNoValue, "", "", "0", dtTime5,
                                    dtTime6, If(Not String.IsNullOrEmpty(""), CInt(""), "0"))

                If QueryDetailResult < 0 Then
                    MessageBox.Show("An error ocurred in the process.", "CTP System", MessageBoxButtons.OK)
                    Return 1
                Else
                    Return 0
                End If
                'Else
                '    Dim arrayCheck As New List(Of String)
                'arrayCheck = strCheck.Split(",").ToList()
                'For Each item As String In arrayCheck
                '    If item = "Project Number" Then
                '        'show error message must have data
                '        Exit For
                '    ElseIf item = "Quantity" Then
                '        txtqty.Text = "0"
                '    ElseIf item = "Unit Cost" Then
                '        txtunitcost.Text = "0"
                '    ElseIf item = "Unit Cost New" Then
                '        txtunitcostnew.Text = "0"
                '    ElseIf item = "Sample Cost" Then
                '        txtsample.Text = "0"
                '    ElseIf item = "Misc. Cost" Then
                '        txttcost.Text = "0"
                '    ElseIf item = "Vendor Number" Then
                '        Exit For
                '        'txtvendorno.Text = "0"  must have data
                '    ElseIf item = "Tooling Cost" Then
                '        txttoocost.Text = "0"
                '    ElseIf item = "Sample Quantity" Then
                '        txtsampleqty.Text = "0"
                '    End If
                'Next

                'If txtvendorno.Text <> "" And projectNo <> 0 Then
                '    QueryDetailResult = gnr.InsertProductDetail(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, CInt(txtqty.Text),
                '                    txtmfr.Text, txtmfrno.Text, CInt(txtunitcost.Text), CInt(txtunitcostnew.Text), txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                '                    cmbuser.SelectedValue, chknew, dtTime3, CInt(txtsample.Text), CInt(txttcost.Text), CInt(txtvendorno.Text), partstoshow, cmbminorcode.SelectedValue, CInt(txttoocost.Text), dtTime4,
                '                    dtTime5, CInt(txtsampleqty.Text))
                'Else
                '    QueryDetailResult = -1
                '    MessageBox.Show("The project number an d vendor number must have value.", "CTP System", MessageBoxButtons.OK)
                'End If

                'If QueryDetailResult < 0 Then
                '    MessageBox.Show("Ann error ocurred inserting data in database.", "CTP System", MessageBoxButtons.OK)
                'End If
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
            Return 1
        End Try
    End Function

    Private Shared Function IsWorkbookAlreadyOpen(app1 As Excel.Application, workbookName As String) As Boolean
        Dim isAlreadyOpen As Boolean = True

        Try
            'app.Workbooks(workbookName)
        Catch theException As Exception
            isAlreadyOpen = False
        End Try

        Return isAlreadyOpen
    End Function

    'part to show column display de option selected. Ex: CTP, Vendor or Both
    Private Function displayPart(opt As String) As String
        Dim result As String = "-1"
        If opt = "CTP" Then
            result = "1"
        ElseIf opt = "Vendor" Then
            result = "2"
        ElseIf opt = "Both" Then
            result = ""
        End If
        Return result
    End Function

    Private Sub cleanFormValues()
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            Dim lstLayouts As New List(Of TableLayoutPanel)

            myTableLayout = Me.TableLayoutPanel2
            lstLayouts.Add(myTableLayout)

            For Each ttt In lstLayouts
                For Each tt In ttt.Controls
                    If TypeOf tt Is Windows.Forms.TextBox Then
                        tt.Text = ""
                    ElseIf TypeOf tt Is Windows.Forms.ComboBox Then
                        tt.selectedIndex = 0
                    ElseIf TypeOf tt Is Windows.Forms.DateTimePicker Then
                        tt.Value = DateTime.Now
                    End If
                Next
            Next

            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Function buildStatusString(status As String) As String
        Dim exMessage As String = " "
        Try
            Dim dsStatuses = gnr.GetAllStatuses()

            dsStatuses.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsStatuses.Tables(0).Rows.Count - 1
                If dsStatuses.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsStatuses.Tables(0).Rows(i).Item(2).ToString() + " -- " + dsStatuses.Tables(0).Rows(i).Item(3).ToString()
                    dsStatuses.Tables(0).Rows(i).Item(5) = fllValueName
                End If
            Next

            Dim dwResult = dsStatuses.Tables(0).AsEnumerable() _
                          .Where(Function(x) Trim(UCase(x.Field(Of String)("CNT03"))) = Trim(UCase(status)))
            Dim newValue = Trim(dwResult(0).ItemArray(3).ToString())
            Return newValue

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Sub disableAfterInsert()
        Dim exMessage As String = " "
        Dim myTableLayout As TableLayoutPanel
        Try
            myTableLayout = Me.TableLayoutPanel2
            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Then
                    tt.Enabled = False
                ElseIf TypeOf tt Is Windows.Forms.ComboBox Then
                    tt.Enabled = False
                ElseIf TypeOf tt Is Windows.Forms.DateTimePicker Then
                    tt.Enabled = False
                ElseIf TypeOf tt Is Windows.Forms.Button Then
                    If tt.Name = "btnSuccess" Or tt.Name = "btnInsert" Then
                        tt.Enabled = False
                    Else
                        tt.Enabled = True
                    End If
                End If
            Next
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub


#End Region

#Region "Grid other Process"

    'Dim customGrid As New Supergrid()

    'customGrid.PageSize = 5
    'customGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    'customGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    'Dim dsData = gnr.getCell2("3221")

    ''DataTable dt = DataProvider.ExecuteDt("select * from test order by col");
    'customGrid.SetPagedDataSource(dsData.Tables(0), BindingNavigator1)

    'Controls.Add(customGrid)
    'customGrid.BringToFront()

#End Region

#Region "Not Used Now"

    'Public Sub frmLoadExcel()
    'InitializeComponent()
    'DataGridView1.Columns.Add(New DataGridViewTextBoxColumn())
    ''DataGridView1.Columns.Add(New DataGridViewTextBoxColumn(DataPropertyName = "Index"))
    'BindingNavigator1.BindingSource = BindingSource1
    'AddHandler BindingSource1.CurrentChanged, AddressOf bindingSource1_CurrentChanged
    'BindingSource1.CurrentChanged += New System.EventHandler(bindingSource1_CurrentChanged);

    'AddHandler vScrollBar1.Scroll, AddressOf vScrollBar1_Scroll
    'BindingSource1.CurrentChanged += New System.EventHandler(bindingSource1_CurrentChanged);
    'BindingSource1.DataSource = New PageOffsetList()
    'End Sub

    'Private Sub fillcell1Other(dw As DataGridViewRow)
    '    Dim exMessage As String = " "
    '    Try
    '        Dim dt As New DataTable
    '        dt = (DirectCast(DataGridView1.DataSource, DataTable))
    '        'Dim projectNo = dw.Cells("clPRHCOD").Value.ToString()
    '        Dim partNo = dw.Cells("clPRDPTN").Value.ToString()
    '        Dim vendorNo = dw.Cells("clVMVNUM").Value.ToString()
    '        'Dim partNo = dw.Cells("clPRDPTN").Value.ToString()

    '        'Dim Qry = dt.AsEnumerable() _
    '        '              .Where(Function(x) Trim(UCase(x.Field(Of Double)("PRHCOD")).ToString()) = Trim(UCase(projectNo)) And
    '        '              Trim(UCase(x.Field(Of Double)("PRDPTN"))) = Trim(UCase(partNo))) _
    '        '              .CopyToDataTable


    '        'txtProjectNo.Text = Qry.Rows(0).ItemArray(0).ToString()
    '        'txtProjectName.Text = Qry.Rows(0).ItemArray(0).ToString()
    '        'dtProjectDate.Text = Qry.Rows(0).ItemArray(1).ToString()
    '        'txtPerCharge.Text = Qry.Rows(0).ItemArray(3).ToString()
    '        'txtStatus.Text = Qry.Rows(0).ItemArray(2).ToString()
    '        'txtDesc.Text = dt.Rows(0).ItemArray(4).ToString()

    '    Catch ex As Exception
    '        exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
    '        MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
    '    End Try
    'End Sub

    'Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
    '    Dim tempView = DirectCast(sender, DataGridView)
    '    Dim Index As Integer

    '    For Each row As DataGridViewRow In DataGridView1.SelectedRows
    '        Index = DataGridView1.CurrentCell.RowIndex
    '        If DataGridView1.Rows(Index).Selected = True Then
    '            fillcell1Other(DataGridView1.Rows(Index))
    '            'Dim code As String = row.Cells(0).Value.ToString()
    '        End If
    '    Next
    'End Sub

#End Region

End Class