Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Schema
Imports ClosedXML.Excel
Imports Microsoft.Win32
Imports Excel = Microsoft.Office.Interop.Excel
'Dim ac As New Autocomplete__module()

Public Class frmLoadExcel

    Private Excel03ConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX={2}'"
    Private Excel07ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX={2}'"
    Dim gnr As Gn1 = New Gn1()
    Dim prd As Product = New Product()
    Dim prdMt As ProductMetadata = New ProductMetadata()
    Dim prdHd As ProductHeader = New ProductHeader()
    Dim xmlConvertClass As ConvertXml = New ConvertXml()
    Public userid As String
    Public flagallow As Integer
    Dim errors As Boolean = False
    Dim schemaErrorDesc As String = Nothing


    Private Const totalRecords As Integer = 43
    Private Const pageSize As Integer = 10

    Dim bs As BindingSource = New BindingSource()
    Dim bs1 As BindingSource = New BindingSource()
    Dim Tables = New BindingList(Of DataTable)()
    Dim Tables1 = New BindingList(Of DataTable)()
    Dim form As frmProductsDevelopment = New frmProductsDevelopment()

#Region "Page Load"

    Private Sub frmLoadExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadCombos(sender, e)
        frmLoadExcel_Load()

    End Sub

    Private Sub frmLoadExcel_Load()
        Dim exMessage As String = " "
        Try
            If CInt(gnr.FlagProductionMethod).Equals(1) Then
                userid = LikeSession.retrieveUser
            Else
                userid = Trim(UCase(frmLogin.txtUserName.Text))
                lblUsrLog.Text += userid
                userid = "LREDONDO"
            End If

            If gnr.getFlagAllow(userid) = 1 Then
                flagallow = 1
            End If

            'test

            'LikeSession.objToCast = prd.ToString().Split(".")(1)
            ' Dim rsFlag = prd.IsValid(prdHd)

            'test

            setValues()

            'Then Set ComboBox AutoComplete properties
            Dim ds = gnr.getVendorNoAndNameByNameDS()
            'Dim ds1 = gnr.getVendorsAccepted(ds)
            Dim bs = New BindingSource()
            bs.DataSource = ds.Tables(0)
            Dim dataview = New DataView(ds.Tables(0))
            Dim myTable As DataTable = dataview.ToTable(False, "VMNAME", "VMVNUM")

            Dim newRow As DataRow = myTable.NewRow
            newRow("VMNAME") = " "
            newRow("VMVNUM") = 0
            'dsUser.Tables(0).Rows.Add(newRow)
            myTable.Rows.InsertAt(newRow, 0)

            ComboBox1.DisplayMember = "VMNAME"
            ComboBox1.ValueMember = "VMVNUM"
            ComboBox1.DataSource = myTable


            Dim myList As String() = New String(myTable.Rows.Count) {}
            Dim i As Integer = 0
            For Each item As DataRow In myTable.Rows
                If Not item("VMNAME").ToString().Equals("") Then
                    If item("VMNAME").ToString() IsNot Nothing Then
                        myList(i) = item("VMNAME").ToString()
                        i += 1
                    End If
                End If
            Next

            ac1.Values = myList



            'Dim newRow As DataRow = myTable.NewRow
            'newRow("VMNAME") = ""
            'newRow("VMVNUM") = -1
            ''dsUser.Tables(0).Rows.Add(newRow)
            'myTable.Rows.InsertAt(newRow, 0)

            'With ComboBox1
            '    .DisplayMember = "VMNAME"
            '    .ValueMember = "VMVNUM"
            '    .DataSource = myTable
            '    .DropDownStyle = ComboBoxStyle.DropDown
            '    .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            '    .AutoCompleteSource = AutoCompleteSource.ListItems
            'End With


        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Threads"

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) _
        Handles BackgroundWorker2.RunWorkerCompleted
        Loading.Close()
    End Sub

    Private Sub backgroundWorker2_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) _
        Handles BackgroundWorker2.DoWork
        ExecuteCombos(sender, e)
    End Sub

    Private Sub backgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) _
        Handles BackgroundWorker2.ProgressChanged
        'txtMfrNoSearch.Text = e.ProgressPercentage.ToString()
    End Sub

#End Region

#Region "DropDowns"

    Private Sub FillDDLStatus1()
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

            Dim newRow As DataRow = dsStatuses.Tables(0).NewRow
            newRow("CNT01") = ""
            newRow("CNT02") = ""
            newRow("CNT03") = ""
            newRow("CNTDE1") = ""
            newRow("CNTDE2") = ""
            newRow("FullValue") = ""
            'dsUser.Tables(0).Rows.Add(newRow)
            dsStatuses.Tables(0).Rows.InsertAt(newRow, 0)

            cmbStatusMore.DataSource = dsStatuses.Tables(0)
            cmbStatusMore.DisplayMember = "FullValue"
            cmbStatusMore.ValueMember = "CNT03"

            'cmbstatus1.DataSource = dsStatuses.Tables(0)
            'cmbstatus1.DisplayMember = "FullValue"
            'cmbstatus1.ValueMember = "CNT03"

            'cmbstatus1.SelectedIndex = -1
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
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
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox1.TextChanged
        If ComboBox1.SelectedValue IsNot Nothing Then
            txtVendorNo.Text = ComboBox1.SelectedValue.ToString()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim exMessage As String = Nothing
        Try
            If ComboBox2.SelectedValue IsNot Nothing And ComboBox2.SelectedIndex <> 0 Then
                txtVendorNo.Text = ComboBox2.SelectedValue.ToString()
                lblVendorDesc.Text = ComboBox2.GetItemText(ComboBox2.SelectedItem)
                'ac1.Text = lblVendorDesc.Text
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "TextBox"

    Private Sub txtProjectNo_TextChanged(sender As Object, e As EventArgs) Handles txtProjectNo.TextChanged
        If Not String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtProjectNo.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        ElseIf Not String.IsNullOrEmpty(txtProjectNo.Text) And String.IsNullOrEmpty(txtProjectName.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
            Dim ds = gnr.GetDataByPRHCOD(txtProjectNo.Text)
            Dim message = If(ds IsNot Nothing, "", "This project number is invalid.")
            If (Not String.IsNullOrEmpty(message)) Then
                MessageBox.Show(message, "CTP System", MessageBoxButtons.OK)
                txtProjectNo.Text = Nothing
            End If
        Else
            btnSelect.Enabled = False
        End If
    End Sub

    Private Sub txtProjectName_TextChanged(sender As Object, e As EventArgs) Handles txtProjectName.TextChanged
        If Not String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtProjectNo.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        ElseIf Not String.IsNullOrEmpty(txtProjectNo.Text) And String.IsNullOrEmpty(txtProjectName.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        Else
            btnSelect.Enabled = False
        End If
    End Sub

    Private Sub txtVendorName_TextChanged(sender As Object, e As EventArgs)

        'Dim result = gnr.getVendorNoAndNameByNameLike(txtVendorName.Text)
        'Dim strValue = txtVendorName.Text
        'Dim DataCollection As New AutoCompleteStringCollection()
        'Dim collection = gnr.getVendorNoAndNameByName()
        'txtVendorName.AutoCompleteCustomSource = collection

    End Sub

    Private Sub txtVendorNo_TextChanged_1(sender As Object, e As EventArgs) Handles txtVendorNo.TextChanged
        If Not String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtProjectNo.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        ElseIf Not String.IsNullOrEmpty(txtProjectNo.Text) And String.IsNullOrEmpty(txtProjectName.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        Else
            btnSelect.Enabled = False
        End If
        btnValidVendor.Enabled = True

        Dim txtValue As String = txtVendorNo.Text

        If True Then

        End If
        'txtVendorNo.Text = If(txtVendorNo.Text IsNot Nothing Or txtVendorNo.Text <> "", txtVendorNo.Text.Replace(Environment.NewLine, ""), " ")
        txtVendorNo.Text = If(txtValue = "0" Or txtValue = Environment.NewLine, txtVendorNo.Text.Replace(txtValue, ""), txtValue)
        ''txtVendorNo.Text = txtVendorNo.Text.Replace(Environment.NewLine, "")
        'If (Regex.IsMatch(txtVendorNo.Text, "^[0-9]{1,6}$") And gnr.isVendorAccepted(txtVendorNo.Text)) Then
        'ComboBox1.SelectedIndex = ComboBox1.FindString(Trim(lblVendorDesc.Text))
        'If ComboBox1.SelectedIndex > 0 Then
        '    ac1.Text = lblVendorDesc.Text
        'End If
        'End If

        If txtVendorNo.Text = "-1" Then
            Dim selIndex = ComboBox1.FindString(Trim(lblVendorDesc.Text))
            Dim curSel As DataRowView = ComboBox1.Items(selIndex)
            txtVendorNo.Text = curSel.Row.ItemArray(1).ToString()
            lblVendorDesc.Text = curSel.Row.ItemArray(0).ToString()
        End If
    End Sub

#End Region

#Region "Gridview and Pagination methods"

    Private Function fromListboxToDatatable(lst As ListBox, Optional dtBase As DataTable = Nothing) As DataTable
        Dim exMessage As String = Nothing
        Try
            Dim dt As DataTable = New DataTable()
            If lst.Items.Count > 0 Then
                dt = dtBase.Clone()
                Dim x As Integer = 0
                Dim vndCode As String = Nothing
                For Each item As String In lst.Items
                    vndCode = If(Integer.TryParse(utilDT(item, dtBase), x), CInt(utilDT(item, dtBase)), 0)
                    If vndCode IsNot Nothing Then
                        Dim row = dt.NewRow()
                        row("VMNAME") = item
                        row("VMVNUM") = vndCode
                        dt.Rows.Add(row)
                    End If
                Next
                dt.AcceptChanges()
            End If
            Return dt
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Sub fillData(dt As DataTable)
        Dim exMessage As String = " "
        Dim mandatoryMissed As String = String.Empty
        Dim dsResult As DataSet = New DataSet()
        Dim dsError As DataSet = New DataSet()
        Dim dtError As DataTable = New DataTable()
        Dim dtResult As DataTable = New DataTable()
        Dim errorMessagee As String
        Dim message3 As String = "This project reference for this part number and vendor already exist."
        Dim message4 As String = "This Part is not in existence in our inventary."
        Dim message5 As String = "The Price and Minimun QT must have a number."
        Dim aditionMessage As String = ""
        Try
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then

                    Dim dictionary As New Dictionary(Of String, String)
                    'preparar logica para que lea automaticamente del xsd file las columnas en el diccionario
                    'dictionary.Add("PRNAME", "Project Name")
                    dictionary.Add("PartNo", "Part No")
                    dictionary.Add("UnitCost", "Unit Cost")
                    dictionary.Add("MOQ", "Minimun Cost")
                    dictionary.Add("MFRNo", "MFR No")
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

                        If Not dt.Columns.Contains("ErrorDesc") Then
                            dtError.Columns.Add("ErrorDesc", GetType(String))
                        End If

                        dtResult = dt.Clone()

                        dsError.Tables.Add(dtError)
                        dsResult.Tables.Add(dtResult)

                        dsError.Namespace = "dsError"
                        dsResult.Namespace = "dsResult"
                        Dim i = 0
                        Dim j = 0
                        For Each item As DataRow In dt.Rows
                            'If String.IsNullOrEmpty(item.ItemArray(dt.Columns("PRPECH").Ordinal).ToString()) Then
                            '    item.Item(dt.Columns("PRPECH").Ordinal) = userid
                            'End If
                            If Not String.IsNullOrEmpty(item.ItemArray(0).ToString()) Then
                                If checkIfPartAndVdrExist(item.ItemArray(dt.Columns("PartNo").Ordinal).ToString(), txtVendorNo.Text) Then
                                    dsError.Tables(0).ImportRow(item)
                                    errorMessagee = message3
                                    dsError.Tables(0).Rows(j).Item("ErrorDesc") = errorMessagee
                                    j += 1
                                Else
                                    If gnr.isPartInExistence(item.ItemArray(dt.Columns("PartNo").Ordinal).ToString()) Then
                                        Dim checkDuplicates = From data In dsResult.Tables(0).AsEnumerable()
                                                              Where Trim(UCase(data.Item("PartNo").ToString())) = Trim(UCase(item.ItemArray(dt.Columns("PartNo").Ordinal).ToString()))

                                        If checkDuplicates IsNot Nothing Then
                                            If Not checkDuplicates.Any() Then
                                                Dim canImportUC As Boolean = False
                                                Dim canImportMC As Boolean = False
                                                Dim uc = item.ItemArray(dt.Columns("UnitCost").Ordinal).ToString()
                                                If Not String.IsNullOrEmpty(uc) Then

                                                    Dim CanConvertUC As Boolean = IsNumeric(uc)
                                                    If CanConvertUC Then
                                                        'dsResult.Tables(0).ImportRow(item)
                                                        'i += 1
                                                        canImportUC = True
                                                    Else
                                                        'dsError.Tables(0).ImportRow(item)
                                                        'errorMessagee = message5
                                                        'dsError.Tables(0).Rows(j).Item("ErrorDesc") = errorMessagee
                                                        'j += 1
                                                        canImportUC = False
                                                    End If
                                                Else
                                                    canImportUC = True
                                                End If

                                                Dim mc = item.ItemArray(dt.Columns("MOQ").Ordinal).ToString()
                                                If Not String.IsNullOrEmpty(mc) Then

                                                    Dim CanConvertMC As Boolean = IsNumeric(mc)
                                                    If CanConvertMC Then
                                                        'dsResult.Tables(0).ImportRow(item)
                                                        'i += 1
                                                        canImportMC = True
                                                    Else
                                                        'dsError.Tables(0).ImportRow(item)
                                                        'errorMessagee = message5
                                                        'dsError.Tables(0).Rows(j).Item("ErrorDesc") = errorMessagee
                                                        'j += 1
                                                        canImportMC = False
                                                    End If
                                                Else
                                                    canImportMC = True
                                                End If

                                                Dim rsIMportUC = If(canImportUC, 1, 0)
                                                Dim rsIMportMC = If(canImportMC, 1, 0)
                                                Dim sumImports As Integer = rsIMportMC + rsIMportUC
                                                If sumImports = 2 Then
                                                    dsResult.Tables(0).ImportRow(item)
                                                    i += 1
                                                Else
                                                    dsError.Tables(0).ImportRow(item)
                                                    errorMessagee = message5
                                                    dsError.Tables(0).Rows(j).Item("ErrorDesc") = errorMessagee
                                                    j += 1
                                                End If

                                            End If
                                        End If
                                    Else
                                        dsError.Tables(0).ImportRow(item)
                                        errorMessagee = message4
                                        dsError.Tables(0).Rows(j).Item("ErrorDesc") = errorMessagee
                                        j += 1
                                    End If
                                End If
                            End If

                        Next

                        LikeSession.dsErrorSession = dsError
                        LikeSession.dsResultsSession = dsResult

                        If dsError.Tables(0).Rows.Count > 0 Then
                            MessageBox.Show("Some project references has errors. You can check them by clicking in the Check Errors button.", "CTP System", MessageBoxButtons.OK)
                        End If

                        If dsResult.Tables(0).Rows.Count = 0 And dsError.Tables(0).Rows.Count = 0 Then
                            MessageBox.Show("There is not data to load. Please check the excel file that you uploaded.", "CTP System", MessageBoxButtons.OK)
                        Else
                            If dsResult.Tables(0).Rows.Count > 0 Then
                                fillcell1(dsResult.Tables(0), 0, dsResult.Namespace)
                            End If

                            If dsError.Tables(0).Rows.Count > 0 Then
                                fillcell1(dsError.Tables(0), 1, dsError.Namespace)
                            End If
                        End If

                        btnSuccess.Enabled = If(dsResult.Tables(0).Rows.Count > 0, True, False)
                        btnCheck.Enabled = If(dsError.Tables(0).Rows.Count > 0, True, False)

                        If dsResult.Tables(0).Rows.Count > 0 Then
                            setSplitContainerVisualization(1, False)
                        Else
                            setSplitContainerVisualization(2, False)
                        End If

                    End If
                Else
                    MessageBox.Show("Error reading excel data.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("Error reading excel data.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Sub setSplitContainerVisualization(index As Integer, value As Boolean)
        Dim exMessage As String = ""
        Try
            '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX={2}'"
            'conStr = String.Format(Excel03ConString, filePath, "YES", 1)
            SplitContainer1.Visible = Not value
            SplitContainer1.Enabled = Not value
            Dim buildedName = "Panel" & index & "Collapsed"
            Dim buildNameReverse As String = Nothing
            Dim pi As PropertyInfo = SplitContainer1.GetType().GetProperty(buildedName)
            pi.SetValue(SplitContainer1, Convert.ChangeType(value, pi.PropertyType), Nothing)
            If index.Equals(1) Then
                btnCheck.Enabled = Not value
                btnSuccess.Enabled = value
                DataGridView1.Visible = Not value
                DataGridView1.Enabled = Not value
                cmdExcel.Visible = value
                lblExcel.Visible = value
                buildNameReverse = "Panel" & index + 1 & "Collapsed"
                Dim pi2 As PropertyInfo = SplitContainer1.GetType().GetProperty(buildNameReverse)
                pi2.SetValue(SplitContainer1, Convert.ChangeType(Not value, pi2.PropertyType), Nothing)
            Else
                btnSuccess.Enabled = Not value
                btnCheck.Enabled = value
                cmdExcel.Visible = Not value
                lblExcel.Visible = Not value
                'cmdExcel.Enabled = Not value
                DataGridView2.Visible = Not value
                DataGridView2.Enabled = Not value
                buildNameReverse = "Panel" & index - 1 & "Collapsed"
                Dim pi1 As PropertyInfo = SplitContainer1.GetType().GetProperty(buildNameReverse)
                pi1.SetValue(SplitContainer1, Convert.ChangeType(Not value, pi1.PropertyType), Nothing)
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub fillcell1(dt As DataTable, flag As Integer, dsName As String, Optional ByVal stopPag As Boolean = False, Optional ByVal NoInsert As Boolean = False)
        Dim exMessage As String = " "
        Try
            If NoInsert = False Then
                If (dsName.Equals("dsResult") Or dsName.Equals("dsGrig1")) Then
                    DataGridView1.Columns.Clear()
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 7

                    'Add Columns
                    DataGridView1.Columns(0).Name = "clPRHCOD"
                    DataGridView1.Columns(0).HeaderText = "Project No."
                    DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

                    DataGridView1.Columns(1).Name = "clPRDPTN"
                    DataGridView1.Columns(1).HeaderText = "Part No."
                    DataGridView1.Columns(1).DataPropertyName = "PartNo"

                    DataGridView1.Columns(2).Name = "clVMVNUM"
                    DataGridView1.Columns(2).HeaderText = "Vendor No."
                    DataGridView1.Columns(2).DataPropertyName = "VMVNUM"

                    DataGridView1.Columns(3).Name = "clPRDMFR"
                    DataGridView1.Columns(3).HeaderText = "Manufacturer No."
                    DataGridView1.Columns(3).DataPropertyName = "MFRNo"

                    DataGridView1.Columns(4).Name = "clPQPRC"
                    DataGridView1.Columns(4).HeaderText = "Unit Cost"
                    DataGridView1.Columns(4).DataPropertyName = "UnitCost"

                    DataGridView1.Columns(5).Name = "clPQMIN"
                    DataGridView1.Columns(5).HeaderText = "Min Qty"
                    DataGridView1.Columns(5).DataPropertyName = "MOQ"

                    DataGridView1.Columns(6).Name = "clPRDSTS"
                    DataGridView1.Columns(6).HeaderText = "Status"
                    DataGridView1.Columns(6).DataPropertyName = "PRDSTS"

                    'FILL GRID

                    DataGridView1.DataSource = dt
                    DataGridView1.Refresh()

                    'If String.IsNullOrEmpty(txtProjectNo.Text) Then
                    If flag.Equals(0) Then
                        btnInsert_Click(Nothing, Nothing)
                    End If
                    'btnInsert_Click(Nothing, Nothing)
                    'End If
                    'DataGridView1.Refresh()

#Region "Checkbox Column"

                    'Dim headerCellLocation As Point = Me.DataGridView1.GetCellDisplayRectangle(0, -1, True).Location

                    ''Place the Header CheckBox in the Location of the Header Cell.
                    'Dim headerCheckBox As New CheckBox
                    'headerCheckBox.Location = New Point(headerCellLocation.X + 8, headerCellLocation.Y + 2)
                    'headerCheckBox.BackColor = Color.White
                    'headerCheckBox.Size = New Size(18, 18)

                    ''Assign Click event to the Header CheckBox.
                    'AddHandler headerCheckBox.Click, AddressOf HeaderCheckBox_Clicked
                    'DataGridView1.Controls.Add(headerCheckBox)

                    ''Add a CheckBox Column to the DataGridView at the first position.
                    'Dim checkBoxColumn As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                    'checkBoxColumn.HeaderText = "All"
                    'checkBoxColumn.Width = 50
                    'checkBoxColumn.Name = "checkBoxColumn"
                    'DataGridView1.Columns.Insert(0, checkBoxColumn)


                    'If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                    '    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                    'End If

#End Region

                    If DataGridView1.Rows.Count > 0 And Not stopPag Then
                        toPaginateDs(DataGridView1, dt)
                    End If
                Else
                    Dim dsError = LikeSession.dsErrorSession
                    DataGridView2.DataSource = Nothing
                    DataGridView2.Refresh()
                    DataGridView2.AutoGenerateColumns = False
                    DataGridView2.ColumnCount = 8

                    'Add Columns
                    DataGridView2.Columns(0).Name = "EditReference"
                    DataGridView2.Columns(0).HeaderText = "Edit"
                    DataGridView2.Columns(0).DataPropertyName = ""

                    DataGridView2.Columns(1).Name = "AddReference"
                    DataGridView2.Columns(1).HeaderText = "Add"
                    DataGridView2.Columns(1).DataPropertyName = ""

                    DataGridView2.Columns(2).Name = "clPRDPTN2"
                    DataGridView2.Columns(2).HeaderText = "Part Number"
                    DataGridView2.Columns(2).DataPropertyName = "PartNo"

                    DataGridView2.Columns(3).Name = "clVMVNUM2"
                    DataGridView2.Columns(3).HeaderText = "Vendor Number"
                    DataGridView2.Columns(3).DataPropertyName = "VMVNUM"

                    DataGridView2.Columns(4).Name = "clPRDMFR2"
                    DataGridView2.Columns(4).HeaderText = "Manufacturer No."
                    DataGridView2.Columns(4).DataPropertyName = "MFRNo"

                    DataGridView2.Columns(5).Name = "clPQPRC2"
                    DataGridView2.Columns(5).HeaderText = "Unit Cost"
                    DataGridView2.Columns(5).DataPropertyName = "UnitCost"

                    DataGridView2.Columns(6).Name = "clPQMIN2"
                    DataGridView2.Columns(6).HeaderText = "Min Qty"
                    DataGridView2.Columns(6).DataPropertyName = "MOQ"

                    DataGridView2.Columns(7).Name = "clError"
                    DataGridView2.Columns(7).HeaderText = "Error Description"
                    DataGridView2.Columns(7).DataPropertyName = "ErrorDesc"

                    If Not dt.Columns.Contains("VMVNUM") Then
                        'Add vendor column
                        Dim dtError = dt.Copy()
                        dtError.Columns.Add("VMVNUM", GetType(Integer)).SetOrdinal(1)

                        For Each dw1 As DataRow In dtError.Rows
                            dw1.Item("VMVNUM") = Trim(txtVendorNo.Text)
                        Next
                        dtError.AcceptChanges()

                        dsError.Tables.RemoveAt(0)
                        dsError.Tables.Add(dtError)
                        dsError.AcceptChanges()

                        LikeSession.dsErrorSession = dsError

                        'FILL GRID
                        DataGridView2.DataSource = dsError.Tables(0)
                    Else
                        'FILL GRID
                        DataGridView2.DataSource = dt
                    End If

                    If DataGridView2.Rows.Count > 0 Then
                        Dim cellAmount = DataGridView2.Rows(0).Cells.Count - 1
                        Dim numbers(cellAmount) As Integer
                        Dim lstVal = New List(Of Integer)()

                        For value As Integer = 0 To cellAmount
                            lstVal.Add(value)
                        Next

                        For Each item As DataGridViewRow In DataGridView2.Rows
                            For Each val As Integer In lstVal
                                If Not (val.Equals(0) Or val.Equals(1)) Then

                                    If item.Cells(val).Value IsNot Nothing And Not IsDBNull(item.Cells(val).Value) Then
                                        If Not String.IsNullOrEmpty(item.Cells(val).Value) Then
                                            item.Cells(val).ReadOnly = True
                                        End If
                                    End If
                                End If
                            Next
                        Next

                        DataGridView2.Columns(cellAmount).ReadOnly = True
                        DataGridView2.Refresh()

                        'btnCheck_Click(Nothing, Nothing)

                        If DataGridView2.Rows.Count > 0 And Not stopPag Then
                            toPaginateDs(DataGridView2, dt)
                        End If
                    End If
                End If
            End If


        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            DataGridView2.DataSource = Nothing
            DataGridView2.Refresh()
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

#Region "Not in use gridview events"

    'Private Sub DataGridView2_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    '    If e.ColumnIndex = 0 Then
    '        DataGridView2.Rows(e.RowIndex).Cells(2).ReadOnly = False
    '        DataGridView2.Rows(e.RowIndex).Cells(3).ReadOnly = True
    '        Dim value = DataGridView2.Rows(e.RowIndex).Cells(0).FormattedValue
    '        If value.Equals("Edit") Then
    '            If DataGridView2.Rows(e.RowIndex).Cells(4).Value <> "The Price and Minimun QT must have a number." Then
    '                DataGridView2.BeginEdit(True)
    '                LikeSession.acceptChanges = True
    '                DataGridView2.RefreshEdit()
    '            Else
    '                DataGridView2.BeginEdit(True)
    '                LikeSession.acceptChanges = False
    '                DataGridView2.RefreshEdit()
    '            End If
    '        Else
    '            DataGridView2.BeginEdit(True)
    '            LikeSession.acceptChanges = False
    '            DataGridView2.RefreshEdit()
    '        End If
    '    ElseIf e.ColumnIndex = 1 Then
    '        Dim partValue = DataGridView2.Rows(e.RowIndex).Cells(2).Value.ToString()
    '        Dim vendorValue = DataGridView2.Rows(e.RowIndex).Cells(3).Value.ToString()
    '        If Not String.IsNullOrEmpty(partValue) Then
    '            'And Not String.IsNullOrEmpty(vendorValue) Then
    '            'Dim vendorOk = gnr.isVendorAccepted(vendorValue)
    '            Dim partOk = gnr.isPartInExistence(partValue)
    '            'If (vendorOk) Then
    '            If partOk Then
    '                Dim myProjectNo = If(String.IsNullOrEmpty(txtProjectNo.Text), "", txtProjectNo.Text)
    '                If String.IsNullOrEmpty(myProjectNo) Then
    '                    'InsertOnDemand(partValue, vendorValue, e.RowIndex)
    '                    InsertOnDemand(partValue, txtVendorNo.Text, e.RowIndex)
    '                Else
    '                    'InsertOnDemand(partValue, vendorValue, e.RowIndex, myProjectNo)
    '                    InsertOnDemand(partValue, txtVendorNo.Text, e.RowIndex, txtProjectNo.Text)
    '                End If
    '            Else
    '                DataGridView2.Rows(e.RowIndex).Cells(4).Value = "The Part Number is not available at this moment."
    '                MessageBox.Show("The Part Number is not available at this moment.", "CTP System", MessageBoxButtons.OK)
    '            End If
    '            'Else
    '            '    DataGridView2.Rows(e.RowIndex).Cells(4).Value = "The Vendor Number is not accepted as a valid vendor."
    '            '    MessageBox.Show("The Vendor Number is not accepted as a valid vendor.", "CTP System", MessageBoxButtons.OK)
    '            'End If
    '        Else
    '            DataGridView2.Rows(e.RowIndex).Cells(4).Value = "There is an error in the input values that prevent the insert process."
    '            MessageBox.Show("You must fill the value for the part for this reference.", "CTP System", MessageBoxButtons.OK)
    '        End If
    '    Else
    '        If LikeSession.acceptChanges = True Then
    '            DataGridView2.Rows(e.RowIndex).Cells(2).ReadOnly = False
    '        Else
    '            DataGridView2.Rows(e.RowIndex).Cells(2).ReadOnly = True
    '            DataGridView2.Rows(e.RowIndex).Cells(3).ReadOnly = True
    '            DataGridView2.Rows(e.RowIndex).Cells(4).ReadOnly = True
    '        End If

    '        'DataGridView1_DoubleClick(sender, e)
    '    End If
    'End Sub

    'Private Sub DataGridView2_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged

    '    Dim exMessage As String = " "
    '    Try
    '        If e.RowIndex >= 0 Then
    '            If e.ColumnIndex = 2 Then
    '                Dim inputText = If(DataGridView2.EditingControl IsNot Nothing, DataGridView2.EditingControl.Text, DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
    '                'Dim inputText = DataGridView2.EditingControl.Text
    '                If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) And gnr.isPartInExistence(inputText) Then
    '                    'DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value = Nothing
    '                    DataGridView2.EndEdit()
    '                    LikeSession.acceptChanges = True
    '                Else
    '                    'DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
    '                    DataGridView2.CancelEdit()
    '                    'DataGridView2.RefreshEdit()
    '                    If (Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)) Then
    '                        DataGridView2.Rows(e.RowIndex).Cells(4).Value = "The Part Number must have existences in stock."
    '                        MessageBox.Show("The Part Number must have existences in stock..", "CTP System", MessageBoxButtons.OK)
    '                    End If
    '                    LikeSession.acceptChanges = True
    '                End If
    '            Else
    '                If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
    '                    DataGridView2.EndEdit()
    '                    LikeSession.acceptChanges = True
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        exMessage = ex.Message + ". " + ex.ToString
    '    End Try

    'End Sub

    'Private Sub DataGridView2_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit

    '    Dim exMessage As String = " "
    '    Try
    '        If e.RowIndex >= 0 Then
    '            If e.ColumnIndex = 2 Then
    '                'Dim inputText = DataGridView2.EditingControl.Text
    '                If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) And gnr.isPartInExistence(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) Then
    '                    DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).Value = Nothing 'clear error description
    '                    DataGridView2.EndEdit()
    '                    LikeSession.acceptChanges = True
    '                ElseIf Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) And Not gnr.isPartInExistence(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) Then
    '                    DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
    '                    DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).Value = "The Part Number must have existences in stock."
    '                    DataGridView2.EndEdit()
    '                    LikeSession.acceptChanges = True
    '                End If
    '            Else
    '                'check for part validation
    '            End If
    '        End If
    '    Catch ex As Exception
    '        exMessage = ex.Message + ". " + ex.ToString
    '    End Try
    'End Sub

    'Private Sub dataGridView2_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) Handles DataGridView2.CellBeginEdit


    '    Dim exMessage As String = " "
    '    Try
    '        If Not LikeSession.acceptChanges Then
    '            If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) And Not e.ColumnIndex.Equals(0) And Not e.ColumnIndex.Equals(1) Then
    '                e.Cancel = True
    '                LikeSession.acceptChanges = False
    '            End If
    '        Else
    '            e.Cancel = False
    '        End If

    '    Catch ex As Exception
    '        exMessage = ex.Message + ". " + ex.ToString
    '    End Try

    'End Sub

    'Private Sub DataGridView2_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

    '    Dim exMessage As String = " "
    '    Try
    '        If e.ColumnIndex = 2 Then
    '            Dim value = DataGridView2(e.ColumnIndex, e.RowIndex).Value.ToString()
    '            Dim inputText = DataGridView2.EditingControl.Text
    '            If Not Regex.IsMatch(inputText, "^[a-zA-Z0-9]{6,19}$") Then
    '                DataGridView2.CancelEdit()
    '                DataGridView2.RefreshEdit()
    '                MessageBox.Show("The Part Number must be setted for a numeric value!", "CTP System", MessageBoxButtons.OK)
    '            End If
    '            'ElseIf e.ColumnIndex = 3 Then
    '            '    DataGridView2.CancelEdit()
    '            '    DataGridView2.RefreshEdit()
    '            '    Dim inputText = If(DataGridView2.EditingControl IsNot Nothing, DataGridView2.EditingControl.Text, DataGridView2(e.ColumnIndex, e.RowIndex).Value.ToString())
    '            '    If Not String.IsNullOrEmpty(inputText) Then
    '            '        If Not Regex.IsMatch(inputText, "^[0-9]{1,6}$") Then
    '            '            MessageBox.Show("The Vendor Number must match with an accepted vendor!", "CTP System", MessageBoxButtons.OK)
    '            '        End If
    '            '    End If
    '        End If
    '    Catch ex As Exception
    '        exMessage = ex.Message + ". " + ex.ToString
    '    End Try
    'End Sub

#End Region

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged
        Dim exMessage As String = " "
        Try
            If e.RowIndex >= 0 Then
                Dim pepe = "pepe"

            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Dim exMessage As String = " "
        Dim CurrentState As String = ""
        Dim NewState As String = ""
        Dim dsResult = LikeSession.dsResultsSession
        Try
            If e.ColumnIndex = 0 Then
                'Dim partNo = DataGridView1.Rows(e.RowIndex).Cells("clPRDPTN").Value
                'Dim vendorNo = DataGridView1.Rows(e.RowIndex).Cells("clVMVNUM").Value
                'Dim result = checkIfPartAndVdrExist(partNo, vendorNo)
                'If result Then
                '    DataGridView1.Rows(e.RowIndex).Cells(0).ReadOnly = False
                'Else
                '    Dim cell As DataGridViewCheckBoxCell = DataGridView1.Rows(e.RowIndex).Cells(0)
                '    cell.Value = True
                '    DataGridView1.Rows(e.RowIndex).Cells(0).ReadOnly = True
                'End If
            ElseIf e.ColumnIndex = 3 Then
                Dim status = If(cmbStatusMore.SelectedIndex = 0, "E", cmbStatusMore.SelectedValue.ToString())
                Dim valueField = If(e.Value IsNot Nothing, e.Value.ToString(), Nothing)
                CurrentState = If((Not String.IsNullOrEmpty(valueField)), e.Value.ToString, status)
                'CurrentState = If((e.Value IsNot Nothing), e.Value.ToString, "E")
                NewState = buildStatusString(CurrentState)
                If Not String.IsNullOrEmpty(NewState) Then
                    DataGridView1.Rows(e.RowIndex).Cells("clPRDSTS").Value = NewState
                Else
                    Exit Sub
                End If
                'ElseIf e.ColumnIndex = 6 Then
                '    Dim statusRow = DataGridView1.Rows(e.RowIndex).Cells("clPRDSTS").Value
                '    Dim peep = "a"
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub DataGridView2_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting

        Dim exMessage As String = " "
        Try
            If e.ColumnIndex = 0 Then
                'If LikeSession.isPageLoad Then
                '    e.Value = "Edit"
                '    e.FormattingApplied = True
                '    If String.IsNullOrEmpty(DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(0).FormattedValue) Then
                '        LikeSession.isPageLoad = False
                '    End If
                'End If
                'If Not String.IsNullOrEmpty(e.Value) Then
                'If LikeSession.acceptChanges Then
                '    e.Value = "Back"
                '    e.FormattingApplied = True
                'Else
                e.Value = "Edit"
                e.FormattingApplied = True
                '    End If
                'End If
            ElseIf e.ColumnIndex = 1 Then
                e.Value = "Add"
                e.FormattingApplied = True
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Protected Sub toPaginateDs(dgv As DataGridView, dtOk As DataTable)
        Dim exMessage As String = " "
        Try

            Dim dtGrid As New DataTable
            dtGrid = dtOk

            Dim counter As Integer = 0
            Dim dt As DataTable = Nothing

            If dgv.Name.ToString().Equals("DataGridView1") Then

                DataGridView1.Visible = True
                If Tables.Count > 0 Then
                    Tables = New BindingList(Of DataTable)()
                    bs.MoveFirst()
                End If

                For Each item As DataRow In dtGrid.Rows
                    If counter = 0 Then
                        dt = dtGrid.Clone()
                        Tables.Add(dt)
                    End If

                    dt.Rows.Add(item.ItemArray)
                    counter += 1

                    If counter > 9 Then
                        counter = 0
                    End If
                Next

                BindingNavigator1.BindingSource = bs
                bs.DataSource = Tables
                AddHandler bs.PositionChanged, AddressOf bs_PositionChanged
                bs_PositionChanged(bs, Nothing)
            Else
                DataGridView2.Visible = True
                If Tables1.Count > 0 Then
                    Tables1 = New BindingList(Of DataTable)()
                    bs1.MoveFirst()
                End If

                For Each item As DataRow In dtGrid.Rows
                    If counter = 0 Then
                        dt = dtGrid.Clone()
                        Tables1.Add(dt)
                    End If

                    dt.Rows.Add(item.ItemArray)
                    counter += 1

                    If counter > 9 Then
                        counter = 0
                    End If
                Next

                BindingNavigator2.BindingSource = bs1
                bs1.DataSource = Tables1
                AddHandler bs1.PositionChanged, AddressOf bs1_PositionChanged
                bs1_PositionChanged(bs1, Nothing)

            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Protected Sub toPaginate(dgv As DataGridView)
        Dim exMessage As String = " "
        Try
            'dim tables as BindingList<DataTable>  = new BindingList<DataTable>()
            Dim dtGrid As New DataTable
            dtGrid = (DirectCast(dgv.DataSource, DataTable))

            Dim counter As Integer = 0
            Dim dt As DataTable = Nothing

            For Each item As DataRow In dtGrid.Rows
                If counter = 0 Then
                    dt = dtGrid.Clone()
                    Tables.Add(dt)
                End If

                dt.Rows.Add(item.ItemArray)
                counter += 1

                If counter > 9 Then
                    counter = 0
                End If
            Next

            BindingNavigator1.BindingSource = bs
            bs.DataSource = Tables
            AddHandler bs.PositionChanged, AddressOf bs_PositionChanged
            'AddHandler bs.PositionChanged, AddressOf bs_PositionChanged1

            bs_PositionChanged(bs, Nothing)
            'bs_PositionChanged1(bs, Nothing)

        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Protected Sub toPaginate1(dgv As DataGridView)
        Dim exMessage As String = " "
        Try
            'dim tables as BindingList<DataTable>  = new BindingList<DataTable>()
            Dim dtGrid As New DataTable
            dtGrid = (DirectCast(dgv.DataSource, DataTable))

            Dim counter As Integer = 0
            Dim dt As DataTable = Nothing

            For Each item As DataRow In dtGrid.Rows
                If counter = 0 Then
                    dt = dtGrid.Clone()
                    Tables1.Add(dt)
                End If

                dt.Rows.Add(item.ItemArray)
                counter += 1

                If counter > 9 Then
                    counter = 0
                End If
            Next

            BindingNavigator2.BindingSource = bs1
            bs1.DataSource = Tables1
            'AddHandler bs.PositionChanged, AddressOf bs_PositionChanged
            AddHandler bs1.PositionChanged, AddressOf bs1_PositionChanged

            'bs_PositionChanged(bs, Nothing)
            bs1_PositionChanged(bs1, Nothing)

        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Public Sub refreshPagination(partNo As String)
        Dim exMessage As String = Nothing
        Try
            Dim myTables = Tables1
            Dim iterator As Integer = 0
            Dim changeDone As Boolean = False
            For Each dtInnerTable As DataTable In myTables
                For Each item As DataRow In dtInnerTable.Rows
                    Dim lookupValue = item("PRDPTN").ToString()
                    If lookupValue.Equals(partNo) Then
                        Dim rowToDelete = dtInnerTable.Rows(iterator)
                        rowToDelete.Delete()
                        dtInnerTable.AcceptChanges()
                        changeDone = True
                        Exit For
                    End If
                    iterator += 1
                Next
                If changeDone Then
                    Exit For
                End If
            Next

            Tables1 = myTables
            Dim epep = Nothing
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub bs_PositionChanged(ByVal sender As Object, ByVal e As EventArgs)
        DataGridView1.DataSource = Tables(bs.Position)
    End Sub

    Private Sub bs1_PositionChanged(ByVal sender As Object, ByVal e As EventArgs)
        DataGridView2.DataSource = Tables1(bs1.Position)
    End Sub

    Public Sub handleDataGridColumnsOnDemand(dgvHandle As DataGridView, listToChange As List(Of Integer), index As Integer, flag As Boolean)
        Dim exMessage As String = " "
        Try
            For Each item As Integer In listToChange
                dgvHandle.Rows(index).Cells(item).ReadOnly = flag
            Next
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Public Sub handleDataGridColumns(handleDataRow As DataGridViewRow, listToChange As List(Of Integer), flag As Boolean)
        Dim exMessage As String = " "
        Try
            For Each item As Integer In listToChange
                handleDataRow.Cells(item).ReadOnly = flag
            Next
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
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
                    conStr = String.Format(Excel03ConString, filePath, "YES", 1)
                    Exit Select

                Case ".xlsx"
                    'Excel 07
                    conStr = String.Format(Excel07ConString, filePath, "YES", 1)
                    Exit Select
            End Select

            If String.IsNullOrEmpty(conStr) Then
                MessageBox.Show("File not valid. You must upload only excel files.", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If

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
                        dt.Columns.Add("PartNo", GetType(String))
                        dt.Columns.Add("UnitCost", GetType(String))
                        dt.Columns.Add("MOQ", GetType(String))
                        dt.Columns.Add("MFRNo", GetType(String))
                        dt.AcceptChanges()
                        cmd.CommandText = (Convert.ToString("SELECT * From [") & sheetName) + "]"
                        cmd.Connection = con
                        con.Open()
                        oda.SelectCommand = cmd
                        'oda.TableMappings.Add("Table", "Net-informations.com")
                        oda.Fill(dt)

                        Dim result = xlsDataSchemaValidation(dt)
                        If String.IsNullOrEmpty(result) Then
                            LikeSession.dsData = dt
                            fillData(dt)
                        Else
                            Dim message = If(result.Equals("No XML Data."), "Error in the xml document structure.", result)
                            MessageBox.Show(message, "CTP System", MessageBoxButtons.OK)
                        End If

                        'LoadThread()
                        'ExecuteFillData(dt)
                        con.Close()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Threads"

    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) _
        Handles BackgroundWorker2.RunWorkerCompleted
        LoadingExcel.Close()
    End Sub

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) _
        Handles BackgroundWorker2.DoWork
        ExecuteCombos(sender, e)
        'ExecuteFillData()
    End Sub

    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) _
        Handles BackgroundWorker2.ProgressChanged
        'txtMfrNoSearch.Text = e.ProgressPercentage.ToString()
    End Sub

    'Private Sub ExecuteFillData(Optional dt As DataTable = Nothing)
    '    fillData(dt)
    'End Sub

#End Region

#Region "button methods"


    'Private Sub txtVendorNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) _
    '    Handles txtVendorNo.KeyPress
    '    If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
    '        btnValidVendor_Click(sender, Nothing)
    '    End If
    'End Sub

    Private Sub txtVendorNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVendorNo.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            btnValidVendor_Click(sender, Nothing)
        End If
    End Sub

    Private Sub btnValidVendor_Click(sender As Object, e As EventArgs) Handles btnValidVendor.Click
        Dim exMessage As String = " "
        Try
            Dim vendorNoValue = Trim(txtVendorNo.Text)

            If Regex.IsMatch(vendorNoValue, "^[0-9]{1,6}$") Then
                Dim validVendor = gnr.isVendorAccepted(vendorNoValue)
                If Not validVendor Then
                    lblVendorDesc.Text = txtVendorNo.Text & ": It is not a valid vendor number."
                    txtVendorNo.Text = Nothing
                    ComboBox1.SelectedIndex = -1
                    ac1.Text = Nothing
                    ComboBox2.DataSource = Nothing
                Else
                    txtVendorNo_TextChanged_1(Nothing, Nothing)

                    Dim dtHandle = DirectCast(ComboBox1.DataSource, DataTable)

                    If dtHandle IsNot Nothing Then
                        ComboBox2.DataSource = dtHandle
                        ComboBox2.DisplayMember = "VMNAME"
                        ComboBox2.ValueMember = "VMVNUM"

                        Dim selIndex = ComboBox2.FindStringExact(lblVendorDesc.Text)
                        ComboBox2.SelectedIndex = If(selIndex <> -1, selIndex, 0)
                    End If

                    'Dim setCombo = If(ComboBox2.DataSource IsNot Nothing, True, False)
                    'If setCombo Then
                    '    ComboBox2.SelectedIndex = 0
                    'End If

                    'If ComboBox1.SelectedIndex > 0 Then
                    '    ac1.Text = lblVendorDesc.Text
                    'End If

                End If
            Else
                txtVendorNo.Text = Nothing
                lblVendorDesc.Text = Nothing
                ComboBox1.SelectedIndex = -1
                MessageBox.Show("The vendor number must have only numeric values and less than 6 characters.", "CTP System", MessageBoxButtons.OK)
            End If
            btnValidVendor.Enabled = False
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click

        'If Not String.IsNullOrEmpty(txtProjectName.Text) Then

        'End If
        'cleanFormValues()
        OpenFileDialog1.ShowDialog()
        'OpenFileDialog1.Dispose()
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Dim exMessage As String = " "
        Dim countErrors As Integer = 0
        'Dim Qry As New DataTable
        Dim iterator As Integer = 0
        Dim arraySuccess As New List(Of Integer)
        Dim arrayError As New List(Of Integer)
        Dim vendorNo = Trim(txtVendorNo.Text)
        Dim projectNo As Integer = 0
        Dim acumulativeFailure As Integer = 0

        Dim objData = New ProductClass()
        Dim objHeader = New ProductHeader()
        objData.Header = objHeader
        Dim objDetails = New Details()
        objData.Header.Detail = objDetails
        Dim objProdDetail = New ProductDetails()
        objData.Header.Detail.Details = objProdDetail
        Dim lstPDet = New List(Of ProductDetails)()
        lstPDet.Add(objProdDetail)
        objData.Header.Detail.LstProdDetails = lstPDet

        Try
            If String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtProjectNo.Text) Then
                MessageBox.Show("The Project Name is a required field.", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If

            'Dim dt As New DataTable
            'dt = (DirectCast(DataGridView1.DataSource, DataTable))

            Dim dsResult = LikeSession.dsResultsSession
            If dsResult IsNot Nothing Then
                If dsResult.Tables(0).Rows.Count <= 0 Then
                    SplitContainer1.Panel2Collapsed = False
                    SplitContainer1.Panel1Collapsed = True
                    MessageBox.Show("There is an error in the data.", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                Else
                    If Not dsResult.Tables(0).Columns.Contains("PRDSTS") Then
                        dsResult.Tables(0).Columns.Add("PRDSTS", GetType(String))
                    End If

                    If Not dsResult.Tables(0).Columns.Contains("VMVNUM") Then
                        dsResult.Tables(0).Columns.Add("VMVNUM", GetType(String))
                    End If

                    dsResult.Tables(0).Columns(0).DataType = GetType(String)
                    dsResult.AcceptChanges()

                    For Each dw As DataRow In dsResult.Tables(0).Rows
                        'Dim pepe = dw.Item("VmVNUM").ToString()
                        dw.Item("VmVNUM") = If(String.IsNullOrEmpty(dw.Item("VmVNUM").ToString()), vendorNo, dw.Item("VmVNUM"))
                        dw.Item("PRDSTS") = If(String.IsNullOrEmpty(cmbStatusMore.SelectedValue), "Entered", Trim(cmbStatusMore.GetItemText(cmbStatusMore.SelectedItem).Split("--")(2)))
                    Next
                End If
            Else
                SplitContainer1.Panel2Collapsed = False
                SplitContainer1.Panel1Collapsed = True
                MessageBox.Show("There is an error in the data.", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If

            Dim queryResult As Integer = 0
            Dim ProjectNoCurrent
            Dim projectPerCharge As String = Nothing
            Dim existProject As Boolean
            If String.IsNullOrEmpty(txtProjectNo.Text) Then
                Dim maxProjectNo = gnr.getmax("PRDVLH", "PRHCOD")
                ProjectNoCurrent = CInt(maxProjectNo) + 1
                existProject = False
            Else
                ProjectNoCurrent = CInt(txtProjectNo.Text)
                existProject = True
            End If

            'validation for create a project or retrieve project data from database
            objData.Header.Detail.Details.VendorNumber = txtVendorNo.Text
            If Not existProject Then

                projectPerCharge = If(cmbPerCharge.SelectedIndex = 0, userid, cmbPerCharge.SelectedValue)

                Dim dsExistsProject = gnr.GetExistByPRNAME(txtProjectName.Text)
                If dsExistsProject IsNot Nothing Then
                    Dim msgResult As DialogResult =
                        MessageBox.Show("The name " & txtProjectName.Text & " is in use in project number: " & dsExistsProject.Tables(0).Rows(0).ItemArray(0).ToString() & ". Please change the project name entered.", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                Else
                    objData.Header.creationUser = userid
                    objData.Header.modificationUser = userid
                    objData.Header.creationDate = Today().ToShortDateString()
                    objData.Header.modificationDate = Today().ToShortDateString()
                    objData.Header.projectDate = Today().ToShortDateString()
                    objData.Header.personInCharge = userid
                    objData.Header.projectInfo = txtDesc.Text
                    objData.Header.projectName = txtProjectName.Text

                    cmbStatus.SelectedIndex = If(cmbStatus.SelectedIndex = 0 Or cmbStatus.SelectedIndex = -1, 1, cmbStatus.SelectedIndex)
                    objData.Header.projectStat = Trim(cmbStatus.GetItemText(cmbStatus.SelectedItem).Split("-")(0))

                    queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, txtDesc.Text, txtProjectName.Text, cmbStatus, projectPerCharge)
                    'queryResult = 0

                    '---------------- End Of Project Header Insertion if new reference ---------------------------------------

                End If
            Else
                Dim ds = gnr.GetDataByPRHCOD(ProjectNoCurrent)
                For Each item As DataRow In ds.Tables(0).Rows
                    txtProjectName.Text = Trim(item.ItemArray(ds.Tables(0).Columns("PRNAME").Ordinal).ToString())
                    cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(item.ItemArray(ds.Tables(0).Columns("PRPECH").Ordinal).ToString()))
                    cmbStatus.SelectedIndex = cmbStatus.FindString(Trim(item.ItemArray(ds.Tables(0).Columns("PRSTAT").Ordinal).ToString()))
                    txtDesc.Text = Trim(item.ItemArray(ds.Tables(0).Columns("PRINFO").Ordinal).ToString())
                    dtProjectDate.Value = CDate(item.ItemArray(ds.Tables(0).Columns("PRDATE").Ordinal)).ToShortDateString()

                    objData.Header.creationUser = Trim(item.ItemArray(ds.Tables(0).Columns("CRUSER").Ordinal).ToString())
                    objData.Header.modificationUser = userid
                    objData.Header.creationDate = Trim(item.ItemArray(ds.Tables(0).Columns("CRDATE").Ordinal).ToString())
                    objData.Header.modificationDate = Today().ToShortDateString()
                    objData.Header.projectDate = Today().ToShortDateString()
                    objData.Header.personInCharge = Trim(item.ItemArray(ds.Tables(0).Columns("PRPECH").Ordinal).ToString())
                    objData.Header.projectInfo = txtDesc.Text
                    objData.Header.projectName = txtProjectName.Text
                    objData.Header.projectStat = cmbStatus.SelectedText

                    objData.Header.Detail.Details.VendorNumber = txtVendorNo.Text
                    'cmbStatusMore.SelectedIndex = 1
                    objData.Header.Detail.Details.Status = cmbStatusMore.SelectedValue

                    '-------------------------- Prepare the data for the update is existed project  ------------------------
                Next
            End If

            If queryResult < 0 Then
                'error message insertion
            Else
                txtProjectNo.Text = ProjectNoCurrent
                objData.Header.projectNo = ProjectNoCurrent

                DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

                For Each row As DataGridViewRow In DataGridView1.Rows
                    'save
                    Dim partNo = row.Cells("clPRDPTN").Value
                    Dim manufNo = row.Cells("clPRDMFR").Value
                    Dim price = row.Cells("clPQPRC").Value
                    Dim minimun = row.Cells("clPQMIN").Value

                    vendorNo = objData.Header.Detail.Details.VendorNumber
                    If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                        dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                    End If

                    dsResult.Tables(0).Rows(iterator).Item("PRHCOD") = ProjectNoCurrent
                    'dsResult.Tables(0).Rows(iterator).Item("VMVNUM") = txtVendorNo.Text

                    'preguntar si parte ya existe en el proyecto
                    Dim dsExist = gnr.GetDataByCodeAndPartNoProdDesc1(ProjectNoCurrent, partNo)
                    If dsExist Is Nothing Then

                        'Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
                        '                 .Where(Function(x) Trim(UCase(x.Field(Of String)("VMVNUM")).ToString()) = Trim(UCase(vendorNo)) And
                        '                 Trim(UCase(x.Field(Of String)("PartNo")).ToString()) = Trim(UCase(partNo)))
                        'Dim Qry1 As DataRowCollection

                        Dim goAhead As Boolean = False
                        Dim selectedRow As DataRow = Nothing
                        For Each item As DataRow In dsResult.Tables(0).Rows
                            If Trim(UCase(item("VMVNUM").ToString())) = Trim(UCase(vendorNo)) And Trim(UCase(item("PartNo").ToString())) = Trim(UCase(partNo)) Then
                                goAhead = True
                                selectedRow = item
                                'solo accede si tiene el mismo vendor asignado
                                Exit For
                            End If
                        Next

                        If goAhead Then
                            'Dim Qry = Qry1.CopyToDataTable
                            Dim personInChargeValue = userid

                            Dim rsInsert = InsertProductDetails(partNo, ProjectNoCurrent, personInChargeValue)
                            'Dim rsInsert = 0

                            '------------------ Insertion in product details first data --------------------------------

                            'add to error dataset if insertion fails
                            If rsInsert < 0 Then
                                Dim dsError = LikeSession.dsErrorSession

                                Dim dtError = dsError.Tables(0).Copy()
                                dtError.Columns.Add("VMVNUM", GetType(String))

                                For Each dw1 As DataRow In dtError.Rows
                                    dw1.Item("VMVNUM") = vendorNo
                                Next

                                Dim row1 As DataRow = dtError.NewRow()
                                row1(0) = partNo
                                row1(2) = txtVendorNo.Text
                                row1(1) = "Error inserting the project reference."

                                dtError.Rows.Add(row1)
                                dtError.AcceptChanges()

                                dsError.Tables.RemoveAt(0)
                                dsError.Tables.Add(dtError)
                                dsError.AcceptChanges()
                                LikeSession.dsErrorSession = dsError
                            Else
                                'right insertion
                                '----- Get data forthe part from dvinva or imnsta, insert or update into poqota ----------------------

                                objData.Header.Detail.Details.ProjectNo = objData.Header.projectNo

                                Dim zeroValue = 0
                                objData.Header.Detail.Details.UnitCostNew = If(price Is Nothing Or IsDBNull(price), 0, price.ToString())
                                objData.Header.Detail.Details.ManufactNo = If(manufNo Is Nothing Or IsDBNull(manufNo), zeroValue.ToString(), manufNo.ToString())
                                objData.Header.Detail.Details.MinQty = If(minimun Is Nothing Or IsDBNull(minimun), 0, minimun.ToString)
                                objData.Header.Detail.Details.PartNo = partNo

                                objData.Header.Detail.Details.NewOrSupplier = If(itemCategory(partNo, txtVendorNo.Text) = 2, 1, 0)

                                Dim qotaObj = GetDataByVendorAndPartNo(txtVendorNo.Text, partNo, True, objData)

                                If qotaObj IsNot Nothing And objData.Header.Detail.Details.PoqotaValidation = 0 Then
                                    'update product development detail
                                    Dim rsUpdProdDet = gnr.UpdateProductDetail1("", qotaObj.Header.Detail.Details.MinorCode, 0, Today(), "", qotaObj.Header.Detail.Details.VendorNumber,
                                                                                qotaObj.Header.Detail.Details.NewOrSupplier, Today(), 0, 0,
                                                                                qotaObj.Header.personInCharge, Today(), userid, qotaObj.Header.Detail.Details.CTPNo,
                                                                                 0, qotaObj.Header.Detail.Details.Qty, "", qotaObj.Header.Detail.Details.ManufactNo, qotaObj.Header.Detail.Details.UnitCost,
                                                                                 qotaObj.Header.Detail.Details.UnitCostNew, "", Today(), qotaObj.Header.Detail.Details.Status,
                                                                                 "", "", qotaObj.Header.Detail.Details.ProjectNo, qotaObj.Header.Detail.Details.PartNo)

                                    If rsUpdProdDet < 0 Then
                                        'error message
                                    End If

                                Else
                                    'error message
                                End If
                                'If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                                '    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                                'End If

                                'If Not dsResult.Tables(0).Columns.Contains("PRDSTS") Then
                                '    dsResult.Tables(0).Columns.Add("PRDSTS", GetType(String))
                                'End If

                                'dsResult.Tables(0).Rows(iterator).Item("PRDSTS") = ProjectNoCurrent


                                'dsResult.Tables(0).Rows(iterator).Item("PRHCOD") = ProjectNoCurrent
                                'txtProjectNo.Text = ProjectNoCurrent
                                If cmbPerCharge.FindStringExact(Trim(projectPerCharge)) Then
                                    cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(projectPerCharge))
                                End If
                                'arraySuccess.Add(ProjectNoCurrent)
                            End If
                            'countErrors += InsertProductDetails(Qry)
                        Else
                            btnSuccess.Enabled = False
                            acumulativeFailure += 1
                            DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)

                            'DataGridView1.Rows.Remove(row)
                            'MessageBox.Show("The vendor selected must be the vendor configured in the project. The right vendor is " & vendorNo, "CTP System", MessageBoxButtons.OK)
                        End If

                    End If
                    iterator += 1
                Next
                dsResult.AcceptChanges()

                Dim rsReferences = gnr.GetReferencesInProject(ProjectNoCurrent)
                If rsReferences = 0 Then
                    Dim rsDeletion = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
                    If rsDeletion < 0 Then
                        'error deleting go to dsError
                    End If
                End If

            End If

            If countErrors > 0 Then
                MessageBox.Show("The insertion process finished with some fails inserting data.", "CTP System", MessageBoxButtons.OK)
            ElseIf acumulativeFailure > 0 Then
                MessageBox.Show("The vendor selected must be the vendor configured in the project. The right vendor is " & vendorNo, "CTP System", MessageBoxButtons.OK)
            Else
                MessageBox.Show("The insertion process finished successfully.", "CTP System", MessageBoxButtons.OK)
                disableAfterInsert(False)
                LikeSession.gridEnable = True
                DataGridView2.Enabled = LikeSession.gridEnable
                DataGridView2.Refresh()

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

#Region "not use"

            '                    For Each tt As DataRow In dsResult.Tables(0).Rows
            '#Region "not in use validate"

            '                        'If dsExistsProject.Tables(0).Rows.Count > 0 Then
            '                        '    'update

            '                        'Else
            '                        '    'insert
            '                        '    Dim maxProjectNo = gnr.getmax("PRDVLH", "PRHCOD")
            '                        '    Dim ProjectNoCurrent = CInt(maxProjectNo) + 1



            '                        '    Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
            '                        '                 .Where(Function(x) Trim(UCase(x.Field(Of String)("PRNAME")).ToString()) = Trim(UCase(txtProjectName.Text)) And
            '                        '                 Trim(UCase(x.Field(Of Double)("PRDPTN"))) = Trim(UCase(partNo)))

            '                        '    If Qry1.Count > 0 Then
            '                        '        Qry = Qry1.CopyToDataTable

            '                        '        Dim projectNameValue = txtProjectName.Text
            '                        '        Dim personInChargeValue = Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString()
            '                        '        Dim detailsValue = txtDesc.Text

            '                        '        Dim queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, detailsValue, projectNameValue, cmbStatus, personInChargeValue)
            '                        '        If queryResult < 0 Then
            '                        '            'error message insertion
            '                        '        Else
            '                        '            Dim rsInsert = InsertProductDetails(Qry, ProjectNoCurrent)
            '                        '            If rsInsert > 0 Then
            '                        '                'delete project no
            '                        '                Dim rsDelete = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
            '                        '                If rsDelete < 0 Then
            '                        '                    'error
            '                        '                End If
            '                        '                countErrors += rsInsert
            '                        '                arrayError.Add(ProjectNoCurrent)
            '                        '            Else
            '                        '                If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
            '                        '                    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
            '                        '                End If

            '                        '                tt("PRHCOD") = ProjectNoCurrent
            '                        '                dsResult.AcceptChanges()
            '                        '                arraySuccess.Add(ProjectNoCurrent)
            '                        '            End If
            '                        '            'countErrors += InsertProductDetails(Qry)
            '                        '        End If
            '                        '    Else
            '                        '        MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
            '                        '    End If


            '                        '    'If Qry IsNot Nothing Then
            '                        '    '    If Qry.Rows.Count > 0 Then

            '                        '    '    Else
            '                        '    '        MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
            '                        '    '    End If
            '                        '    'Else
            '                        '    '    MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
            '                        '    'End If
            '                        'End If

            '#End Region
            '                        'insert
            '                        Dim partNo = tt.Item(dsResult.Tables(0).Columns("PRDPTN").Ordinal).ToString()
            '                        Dim vendorNo = tt.Item(dsResult.Tables(0).Columns("VMVNUM").Ordinal).ToString()

            '                        Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
            '                                             .Where(Function(x) Trim(UCase(x.Field(Of Double)("VMVNUM")).ToString()) = Trim(UCase(vendorNo)) And
            '                                             Trim(UCase(x.Field(Of Double)("PRDPTN"))) = Trim(UCase(partNo)))

            '                        If Qry1.Count > 0 Then
            '                            Qry = Qry1.CopyToDataTable
            '                            Dim personInChargeValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString()), userid, Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString())

            '                            Dim rsInsert = InsertProductDetails(Qry, ProjectNoCurrent)
            '                            If rsInsert > 0 Then
            '                                'delete project no
            '                                Dim rsDelete = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
            '                                If rsDelete < 0 Then
            '                                    'error borrando
            '                                End If
            '                                countErrors += rsInsert
            '                                arrayError.Add(ProjectNoCurrent)
            '                            Else
            '                                'right insertion
            '                                If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
            '                                    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
            '                                End If

            '                                tt("PRHCOD") = ProjectNoCurrent
            '                                dsResult.AcceptChanges()

            '                                txtProjectNo.Text = ProjectNoCurrent
            '                                If cmbPerCharge.FindStringExact(Trim(projectPerCharge)) Then
            '                                    cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(projectPerCharge))
            '                                End If


            '                                arraySuccess.Add(ProjectNoCurrent)
            '                            End If
            '                            'countErrors += InsertProductDetails(Qry)

            '                        Else
            '                            MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
            '                        End If
            '                    Next

#End Region
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub InsertOnDemand(partNo As String, vendorNo As String, position As Integer, Optional ByVal projectNo As String = Nothing)
        Dim exMessage As String = " "
        Dim countErrors As Integer = 0
        'Dim Qry As New DataTable
        Dim arraySuccess As New List(Of Integer)
        Dim arrayError As New List(Of Integer)
        Try
            'test grid
            Dim dtest1 = (DirectCast(DataGridView1.DataSource, DataTable))
            Dim dtest2 = (DirectCast(DataGridView2.DataSource, DataTable))

            If String.IsNullOrEmpty(txtProjectName.Text) Then
                MessageBox.Show("The Project Name is a required field.", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If

            Dim queryResult As Integer = 0
            Dim ProjectNoCurrent As Integer = 0

            If String.IsNullOrEmpty(projectNo) Then

                Dim maxProjectNo = gnr.getmax("PRDVLH", "PRHCOD")
                ProjectNoCurrent = CInt(maxProjectNo) + 1
                Dim projectPerCharge = If(cmbPerCharge.SelectedIndex = 0, userid, cmbPerCharge.SelectedValue)

                Dim dsExistsProject = gnr.GetExistByPRNAME(txtProjectName.Text)
                If dsExistsProject IsNot Nothing Then
                    'decirlo y preguntar que hacer, puede actualizar o puede dejarlo
                    Dim msgResult As DialogResult =
                        MessageBox.Show("The name " & txtProjectName.Text & " is in use in project number: " & dsExistsProject.Tables(0).Rows(0).ItemArray(0).ToString() & ". Please change the project name entered.", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                    'Dim projectNo1 = dsExistsProject.Tables(0).Rows(0).ItemArray(0).ToString()
                Else
                    queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, txtDesc.Text, txtProjectName.Text, cmbStatus, projectPerCharge)
                End If
            Else
                ProjectNoCurrent = CInt(projectNo)
            End If

            If queryResult < 0 Then
                'error message insertion
            Else
                Dim dsResult As DataSet = New DataSet()
                Dim dt As New DataTable
                Dim dsInsert As New DataSet
                Dim dtInsert As New DataTable

                dt = (DirectCast(DataGridView2.DataSource, DataTable))
                dtInsert = dt.Clone()
                Dim dtUse = dt.Copy()
                dsResult.Tables.Add(dtUse)

                Dim sourceRow = dsResult.Tables(0).Rows(position)
                dsInsert.Tables.Add(dtInsert)
                dsInsert.Tables(0).ImportRow(sourceRow)

                Dim strCompare = "This project reference for this part number and vendor already exist."
                Dim strDetail = dsInsert.Tables(0).Rows(0).Item("ErrorDesc").ToString()
                If strDetail.Equals(strCompare) Then
                    Dim msgProceed As DialogResult = MessageBox.Show("This part number and vendor number are present in project number: " & LikeSession.referencedExistence & ". Do you want to create a new project with that reference?", "CTP System", MessageBoxButtons.YesNo)
                    If msgProceed = DialogResult.No Then
                        Exit Sub
                    End If
                End If

                'save
                'Dim partNo = row.Cells("clPRDPTN2").Value
                'Dim vendorNo = row.Cells("clVMVNUM2").Value 

                Dim personInChargeValue = userid
                'Dim personInChargeValue = If(String.IsNullOrEmpty(dsInsert.Tables(0).Rows(0).ItemArray(dsInsert.Tables(0).Columns("PRPECH").Ordinal).ToString()), userid, dsInsert.Tables(0).Rows(0).ItemArray(dsInsert.Tables(0).Columns("PRPECH").Ordinal).ToString())

                Dim rsInsert = InsertProductDetails(partNo, ProjectNoCurrent, personInChargeValue)
                If rsInsert > 0 Then
                    'delete project no
                    'Dim rsDelete = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
                    'If rsDelete < 0 Then
                    '    'error borrando
                    'End If
                    countErrors += rsInsert
                    arrayError.Add(projectNo)
                Else
                    'right insertion
                    Dim dtGrig1 As New DataTable
                    Dim dtGrig2 As New DataTable
                    Dim dtGrig1Ok As New DataTable
                    Dim dtGrig2Ok As New DataTable
                    Dim dsGrig1 As New DataSet
                    Dim dsGrig2 As New DataSet

                    If DataGridView1.DataSource Is Nothing Then
                        dtGrig1 = (DirectCast(LikeSession.dsResultsSession.Tables(0), DataTable))
                        dtGrig1Ok = dtGrig1.Clone()
                    Else
                        dtGrig1 = (DirectCast(DataGridView1.DataSource, DataTable))
                        dtGrig1Ok = dtGrig1.Copy()
                    End If

                    If DataGridView2.DataSource Is Nothing Then
                        dtGrig2 = (DirectCast(LikeSession.dsErrorSession.Tables(0), DataTable))
                        dtGrig2Ok = dtGrig2.Clone()
                    Else
                        dtGrig2 = (DirectCast(DataGridView2.DataSource, DataTable))
                        dtGrig2Ok = dtGrig2.Copy()
                    End If

                    dsGrig2.Tables.Add(dtGrig2Ok)
                    dsGrig1.Namespace = "dsGrig1"
                    dsGrig2.Namespace = "dsGrig2"

                    If Not dtGrig1Ok.Columns.Contains("VMVNUM") Then
                        dtGrig1Ok.Columns.Add("VMVNUM", GetType(Integer))
                    End If

                    If Not dtGrig1Ok.Columns.Contains("PRDSTS") Then
                        dtGrig1Ok.Columns.Add("PRDSTS", GetType(String))
                    End If

                    Dim newRow As DataRow = dtGrig1Ok.NewRow
                    newRow("PRDPTN") = dsGrig2.Tables(0).Rows(position).Item("PRDPTN").ToString()
                    newRow("VMVNUM") = dsGrig2.Tables(0).Rows(position).Item("VMVNUM").ToString()
                    Dim status = If(cmbStatusMore.SelectedIndex = 0, "E", cmbStatusMore.SelectedValue.ToString())
                    newRow("PRDSTS") = status
                    dtGrig1Ok.Rows.Add(newRow)
                    dsGrig1.Tables.Add(dtGrig1Ok)
                    'dsGrig1.AcceptChanges()
                    'dsGrig1.Tables(0).ImportRow(dsGrig2.Tables(0).Rows(position))

                    dsGrig2.Tables(0).Rows.Remove(dsGrig2.Tables(0).Rows(position))
                    dsGrig2.AcceptChanges()

                    'DataGridView2.DataSource = dsGrig2
                    'DataGridView2.Refresh()
                    LikeSession.dsErrorSession = dsGrig2

                    If Not (dsGrig1.Tables(0).Columns.Contains("PRHCOD")) Then
                        dsGrig1.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                    End If

                    dsGrig1.Tables(0).Rows(dsGrig1.Tables(0).Rows.Count - 1).Item("PRHCOD") = ProjectNoCurrent
                    dsGrig1.AcceptChanges()

                    'DataGridView1.DataSource = dsGrig1
                    'DataGridView1.Refresh()
                    LikeSession.dsResultsSession = dsGrig1

                    fillcell1(dsGrig1.Tables(0), 1, dsGrig1.Namespace, True)
                    fillcell1(dsGrig2.Tables(0), 1, dsGrig2.Namespace, True)

                    refreshPagination(newRow("PRDPTN").ToString())

                    bs.ResetBindings(False)
                    bs1.ResetBindings(False)

                    setSplitContainerVisualization(1, False)

                    'txtProjectNo.Text = projectNo
                    'If cmbPerCharge.FindStringExact(Trim(projectPerCharge)) Then
                    '    cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(projectPerCharge))
                    'End If

                    arraySuccess.Add(projectNo)

                    If String.IsNullOrEmpty(txtProjectNo.Text) Then
                        Dim rsReferences = gnr.GetReferencesInProject(ProjectNoCurrent)
                        txtProjectNo.Text = If(rsReferences > 0, ProjectNoCurrent, Nothing)
                    End If
                End If
#Region "not use"

                '                    For Each tt As DataRow In dsResult.Tables(0).Rows
                '#Region "not in use validate"

                '                        'If dsExistsProject.Tables(0).Rows.Count > 0 Then
                '                        '    'update

                '                        'Else
                '                        '    'insert
                '                        '    Dim maxProjectNo = gnr.getmax("PRDVLH", "PRHCOD")
                '                        '    Dim ProjectNoCurrent = CInt(maxProjectNo) + 1



                '                        '    Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
                '                        '                 .Where(Function(x) Trim(UCase(x.Field(Of String)("PRNAME")).ToString()) = Trim(UCase(txtProjectName.Text)) And
                '                        '                 Trim(UCase(x.Field(Of Double)("PRDPTN"))) = Trim(UCase(partNo)))

                '                        '    If Qry1.Count > 0 Then
                '                        '        Qry = Qry1.CopyToDataTable

                '                        '        Dim projectNameValue = txtProjectName.Text
                '                        '        Dim personInChargeValue = Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString()
                '                        '        Dim detailsValue = txtDesc.Text

                '                        '        Dim queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, detailsValue, projectNameValue, cmbStatus, personInChargeValue)
                '                        '        If queryResult < 0 Then
                '                        '            'error message insertion
                '                        '        Else
                '                        '            Dim rsInsert = InsertProductDetails(Qry, ProjectNoCurrent)
                '                        '            If rsInsert > 0 Then
                '                        '                'delete project no
                '                        '                Dim rsDelete = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
                '                        '                If rsDelete < 0 Then
                '                        '                    'error
                '                        '                End If
                '                        '                countErrors += rsInsert
                '                        '                arrayError.Add(ProjectNoCurrent)
                '                        '            Else
                '                        '                If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                '                        '                    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                '                        '                End If

                '                        '                tt("PRHCOD") = ProjectNoCurrent
                '                        '                dsResult.AcceptChanges()
                '                        '                arraySuccess.Add(ProjectNoCurrent)
                '                        '            End If
                '                        '            'countErrors += InsertProductDetails(Qry)
                '                        '        End If
                '                        '    Else
                '                        '        MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                '                        '    End If


                '                        '    'If Qry IsNot Nothing Then
                '                        '    '    If Qry.Rows.Count > 0 Then

                '                        '    '    Else
                '                        '    '        MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                '                        '    '    End If
                '                        '    'Else
                '                        '    '    MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                '                        '    'End If
                '                        'End If

                '#End Region
                '                        'insert
                '                        Dim partNo = tt.Item(dsResult.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                '                        Dim vendorNo = tt.Item(dsResult.Tables(0).Columns("VMVNUM").Ordinal).ToString()

                '                        Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
                '                                             .Where(Function(x) Trim(UCase(x.Field(Of Double)("VMVNUM")).ToString()) = Trim(UCase(vendorNo)) And
                '                                             Trim(UCase(x.Field(Of Double)("PRDPTN"))) = Trim(UCase(partNo)))

                '                        If Qry1.Count > 0 Then
                '                            Qry = Qry1.CopyToDataTable
                '                            Dim personInChargeValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString()), userid, Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString())

                '                            Dim rsInsert = InsertProductDetails(Qry, ProjectNoCurrent)
                '                            If rsInsert > 0 Then
                '                                'delete project no
                '                                Dim rsDelete = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
                '                                If rsDelete < 0 Then
                '                                    'error borrando
                '                                End If
                '                                countErrors += rsInsert
                '                                arrayError.Add(ProjectNoCurrent)
                '                            Else
                '                                'right insertion
                '                                If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                '                                    dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                '                                End If

                '                                tt("PRHCOD") = ProjectNoCurrent
                '                                dsResult.AcceptChanges()

                '                                txtProjectNo.Text = ProjectNoCurrent
                '                                If cmbPerCharge.FindStringExact(Trim(projectPerCharge)) Then
                '                                    cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(projectPerCharge))
                '                                End If


                '                                arraySuccess.Add(ProjectNoCurrent)
                '                            End If
                '                            'countErrors += InsertProductDetails(Qry)

                '                        Else
                '                            MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                '                        End If
                '                    Next

#End Region
            End If

            If countErrors > 0 Then
                MessageBox.Show("The insertion process fail.", "CTP System", MessageBoxButtons.OK)
            Else
                MessageBox.Show("The insertion process finished successfully.", "CTP System", MessageBoxButtons.OK)
                disableAfterInsert(False)
            End If

        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub btnCheck_Click(sender As Object, e As EventArgs) Handles btnCheck.Click
        Dim exMessage As String = " "
        Try
            Dim dsValue = LikeSession.dsErrorSession
            fillcell1(dsValue.Tables(0), 1, dsValue.Namespace, True)
            setSplitContainerVisualization(2, False)
            'btnSuccess.Enabled = True
            'btnCheck.Enabled = False
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub btnSuccess_Click(sender As Object, e As EventArgs) Handles btnSuccess.Click
        Dim exMessage As String = " "
        Try
            Dim dsValue = LikeSession.dsResultsSession
            fillcell1(dsValue.Tables(0), 0, dsValue.Namespace, True, True)
            setSplitContainerVisualization(1, False)
            'btnSuccess.Enabled = False
            'btnCheck.Enabled = True
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdExcel_Click_1(sender As Object, e As EventArgs) Handles cmdExcel.Click
        Dim exMessage As String = " "
        Try
            Dim userPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim folderPath As String = userPath & "\PD-Bulk-Errors\"
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If

            Dim dt As New DataTable
            dt = (DirectCast(DataGridView2.DataSource, DataTable))
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim fileExtension As String = Determine_OfficeVersion()
                    If String.IsNullOrEmpty(fileExtension) Then
                        Exit Sub
                    End If

                    Dim fileName As String
                    If Not String.IsNullOrEmpty(txtProjectNo.Text) Then
                        fileName = "Project number " & txtProjectNo.Text & " - " & DateTime.Now.ToString("d") & " - Errors." & fileExtension
                    Else
                        fileName = "Project Name " & txtProjectName.Text & " - Errors. The project does not have a number yet." & fileExtension
                    End If

                    Dim fullPath = folderPath & Convert.ToString(fileName)
                    Using wb As New XLWorkbook()
                        wb.Worksheets.Add(dt, "Project")
                        wb.SaveAs(fullPath)
                    End Using

                    If File.Exists(fullPath) Then
                        MessageBox.Show("The file was created successfully in this path " & folderPath, "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        disableAfterInsert(True)
    End Sub

    Private Sub cmbStatusMore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStatusMore.SelectedIndexChanged

    End Sub

#End Region

#Region "Get All Data by Vendor and Part No"

    Private Function GetDataByVendorAndPartNo(vendor As String, partNo As String, flag As Boolean, objData As ProductClass) As ProductClass
        Dim exMessage As String = " "
        Dim spacepoqota1 = "                               DEV"
        Dim statusquote As String = Nothing

        Try
            If flag Then 'no existe referencia para la combinacion
                Dim validation As Integer = 0

                objData.Header.Detail.Details.PartNo = partNo
                objData.Header.Detail.Details.VendorNumber = vendor

                'busco en dvinva
                Dim dsGetDataFromDualInv = gnr.GetDataFromDualInventory(partNo)
                If Not dsGetDataFromDualInv Is Nothing Then
                    If dsGetDataFromDualInv.Tables(0).Rows.Count > 0 Then
                        objData.Header.Detail.Details.MinorCode = Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC2").Ordinal).ToString())

                        'If Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString()) <> "" Then
                        '    Dim dsGetVendorQuey = gnr.GetVendorQuey(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString())
                        '    If Not dsGetVendorQuey Is Nothing Then
                        '        If dsGetVendorQuey.Tables(0).Rows.Count > 0 Then
                        '            'objData.Header.Detail.Details.VendorNumber = dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString()
                        '            objData.Header.Detail.Details.VendorNumber = vendor
                        '            'prdDetData.vendor = Trim(dsGetVendorQuey.Tables(0).Rows(0).ItemArray(dsGetVendorQuey.Tables(0).Columns("VMNAME").Ordinal).ToString())
                        '        Else
                        '            objData.Header.Detail.Details.VendorNumber = ""
                        '            'txtvendornamea.Text = ""
                        '        End If
                        '    End If
                        'Else
                        '    objData.Header.Detail.Details.VendorNumber = ""
                        '    'txtvendornamea.Text = ""
                        'End If

                        Dim dsGetCTPPartRef = gnr.GetCTPPartRef(partNo)
                        If Not String.IsNullOrEmpty(dsGetCTPPartRef) Then
                            objData.Header.Detail.Details.CTPNo = dsGetCTPPartRef
                            objData.Header.Detail.Details.ManufactNo = dsGetCTPPartRef
                        Else
                            objData.Header.Detail.Details.CTPNo = ""
                            objData.Header.Detail.Details.ManufactNo = ""
                        End If

                        '----- Quizas no poner aqui ------

                        'Dim dsGetAssignedVendor = gnr.GetAssignedVendor(vendor, partNo)
                        'If dsGetAssignedVendor IsNot Nothing Then
                        '    If dsGetAssignedVendor.Tables(0).Rows.Count > 0 Then
                        '        objData.Header.Detail.Details.UnitCost = dsGetAssignedVendor.Tables(0).Rows(0).ItemArray(dsGetAssignedVendor.Tables(0).Columns("PQPRC").Ordinal).ToString()
                        '        '?
                        '        objData.Header.Detail.Details.MinQty = dsGetAssignedVendor.Tables(0).Rows(0).ItemArray(dsGetAssignedVendor.Tables(0).Columns("PQMIN").Ordinal).ToString()
                        '        'txtminqty.Text = dsGetAssignedVendor.Tables(0).Rows(0).ItemArray(dsGetAssignedVendor.Tables(0).Columns("PQMIN").Ordinal).ToString()
                        '    Else
                        '        objData.Header.Detail.Details.UnitCost = 0
                        '        objData.Header.Detail.Details.MinQty = 0
                        '    End If
                        '    objData.Header.Detail.Details.VendorNumber = vendor
                        'End If

                        '----- Quizas no poner aqui ------

                    End If
                Else
                    'sino esta en dvinva busco en imnsta
                    Dim dsGetDataFromDualInventory1 = gnr.GetDataByPartNoVendor(partNo)
                    If Not dsGetDataFromDualInventory1 Is Nothing Then
                        If dsGetDataFromDualInventory1.Tables(0).Rows.Count > 0 Then
                            objData.Header.Detail.Details.MinorCode = Trim(dsGetDataFromDualInventory1.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInventory1.Tables(0).Columns("IMPC2").Ordinal).ToString())

                            Dim dsGetCTPPartRef = gnr.GetCTPPartRef(partNo)
                            If Not String.IsNullOrEmpty(dsGetCTPPartRef) Then
                                objData.Header.Detail.Details.CTPNo = dsGetCTPPartRef
                                objData.Header.Detail.Details.ManufactNo = dsGetCTPPartRef
                            Else
                                objData.Header.Detail.Details.CTPNo = ""
                                objData.Header.Detail.Details.ManufactNo = ""
                            End If
                        Else

                        End If
                    Else

                    End If
                End If

                'test purpose
                'Dim testPartNo = "5257106"
                'Dim dsGetDataFromProdHeaderAndDetail = gnr.GetDataFromProdHeaderAndDetail(partNo)
                Dim dtpDate = New DateTimePicker()
                Dim dtpDate1 = New DateTimePicker()
                Dim dt = DateTime.Now

                Dim iDate As String = "1900-01-01"
                Dim oDate As DateTime = DateTime.Parse(iDate)
                dtpDate.Value = dt
                dtpDate1.Value = oDate
                Dim code As String
                Dim name As String
                Dim ResultQuery As Integer

                'If Not dsGetDataFromProdHeaderAndDetail Is Nothing Then
                'If dsGetDataFromProdHeaderAndDetail.Tables(0).Rows.Count > 0 Then

                Dim Qry As New DataTable
                Dim strQueryAdd1 As String = "WHERE PQVND = " & Trim(vendor) & " AND PQPTN = '" & Trim(UCase(partNo)) & "'"

                'busco en poqota si hay referencia para la parte y el vendor
                Dim dsPoQota = gnr.GetPOQotaData(vendor, partNo)
                If dsPoQota IsNot Nothing Then

                    'ya existe en poqota esta parte con este vendor. Permite actualizar los valores?

                    Dim rowMax = dsPoQota.Tables(0).AsEnumerable().Where(Function(row) row.ItemArray(1).ToString() = partNo And row.ItemArray(2).ToString() = vendor).Max(Function(row) row.ItemArray(3))
                    Dim rowOk1 = dsPoQota.Tables(0).AsEnumerable().Where(Function(row) row.ItemArray(1).ToString() = partNo And row.ItemArray(2).ToString() = vendor And row.ItemArray(2).ToString() = rowMax)
                    Dim dtRow As New DataTable

                    If rowOk1.Count = 1 Then
                        dtRow = rowOk1.CopyToDataTable()
                        objData.Header.Detail.Details.Qty = dtRow.Rows(0).ItemArray(11).ToString()
                        'objData.Header.Detail.Details.UnitCost = dtRow.Rows(0).ItemArray(3).ToString()
                        'objData.Header.Detail.Details.UnitCostNew = 0
                        objData.Header.Detail.Details.Status = cmbStatusMore.SelectedValue
                        objData.Header.Detail.Details.MinQty = 0

                        statusquote = "D-" & cmbStatusMore.SelectedText

                        'objData.Header(0).Detail.Add(prdDetData)

                        'recuperar datos de poqota y actualizar
                        'ResultQuery = gnr.InsertNewPOQota1(prdDetData.PartNo, prdDetData.VendorNumber, 1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), prdDetData.ManufactNo,
                        'DateTime.Now.Day.ToString(), statusquote, spacepoqota1)

                    Else
                        objData.Header.Detail.Details.Qty = 0
                        objData.Header.Detail.Details.UnitCost = 0
                        objData.Header.Detail.Details.UnitCostNew = 0
                        objData.Header.Detail.Details.Status = cmbStatusMore.SelectedValue
                        objData.Header.Detail.Details.MinQty = 0

                        statusquote = "D-" & cmbStatusMore.SelectedText

                        'objData.Header(0).Detail.Add(prdDetData)

                        'insertar en poqota valores iniciales en cero
                        ResultQuery = gnr.InsertNewPOQota1(objData.Header.Detail.Details.PartNo, objData.Header.Detail.Details.VendorNumber, 1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                                                                   objData.Header.Detail.Details.ManufactNo, DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                    End If

                Else
                    objData.Header.Detail.Details.Qty = 0
                    objData.Header.Detail.Details.UnitCost = 0
                    If cmbStatusMore.SelectedIndex = 0 Or cmbStatusMore.SelectedIndex = -1 Then
                        cmbStatusMore.SelectedIndex = 1
                        objData.Header.Detail.Details.Status = cmbStatusMore.SelectedValue
                    Else
                        objData.Header.Detail.Details.Status = cmbStatusMore.SelectedValue
                    End If
                    'objData.Header.Detail.Details.Status = If(cmbStatusMore.SelectedIndex = 0 Or cmbStatusMore.SelectedIndex = -1, "E", cmbStatusMore.SelectedValue.ToString())

                    statusquote = "D-" & Trim(cmbStatusMore.GetItemText(cmbStatusMore.SelectedItem).Split("--")(2))

                    'insertar en poqota valores iniciales en cero
                    ResultQuery = gnr.InsertNewPOQota(objData.Header.Detail.Details.PartNo, objData.Header.Detail.Details.VendorNumber, 1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                                                               objData.Header.Detail.Details.ManufactNo, DateTime.Now.Day.ToString(), statusquote, spacepoqota1,
                                                                objData.Header.Detail.Details.UnitCostNew, objData.Header.Detail.Details.MinQty)

                    objData.Header.Detail.Details.PoqotaValidation = ResultQuery.ToString()

                    If ResultQuery < 0 Then
                        'error message
                    End If
                End If
                'End If
                'End If
            Else
                Dim result1 As DialogResult = MessageBox.Show("Part No. cannot be changed when is already created.", "CTP System", MessageBoxButtons.OK)
            End If

            Return objData

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
            Return Nothing
        End Try
    End Function

#End Region

#Region "Utils"

    'Public Static Object GetCellValueFromColumnHeader(this DataGridViewCellCollection CellCollection, String HeaderText)
    '{
    '    Return CellCollection.Cast < DataGridViewCell > ().First(c >= c.OwningColumn.HeaderText == HeaderText).Value;            
    '}

    'Public Shared Function GetCellValueFromColumnHeader(CellCollection As DataGridViewCellCollection, HeaderText As String) As Object
    '    Dim exMessage As String = ""
    '    Try

    '        CellCollection.Cast(Of DataGridViewCell)().First(Function(c) c.o)
    '    Catch ex As Exception
    '        exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
    '    End Try

    'End Function


    Private Function itemCategory(partNo As String, vendorNo As String) As Integer
        Dim exMessage As String = " "
        Dim result As Integer = -1
        Try
            If String.IsNullOrEmpty(partNo) Then
                Return 2
            Else
                Dim listItemCat = gnr.VendorWhiteFlagMethod.Split(",")

                Dim dsResult1 = gnr.getItemCategoryByVendorAndPart(vendorNo, partNo)
                If dsResult1 IsNot Nothing Then
                    If dsResult1.Tables(0).Rows.Count > 0 Then
                        For Each item As String In listItemCat
                            If Trim(item).Equals(Trim(vendorNo)) Then
                                Return 2
                            End If
                        Next
                        Return -1
                    Else
                        Return 2
                    End If
                Else
                    Return 2
                End If

                Return result
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return result
        End Try

    End Function

    Private Function FillDDlMinorCode() As Dictionary(Of String, String)
        Dim exMessage As String = " "
        Dim dictionary As New Dictionary(Of String, String)
        Try
            Dim dsMinCodes = gnr.FillDDlMinorCode()

            dsMinCodes.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsMinCodes.Tables(0).Rows.Count - 1
                If dsMinCodes.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsMinCodes.Tables(0).Rows(i).Item(2).ToString() + " -- " + dsMinCodes.Tables(0).Rows(i).Item(3).ToString()
                    'dsMinCodes = Trim(dsMinCodes.Tables(0).Rows(i).Item(0).ToString())
                    dsMinCodes.Tables(0).Rows(i).Item(5) = fllValueName
                    'dsMinCodes.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                    dictionary.Add(dsMinCodes.Tables(0).Rows(i).Item(2).ToString(), dsMinCodes.Tables(0).Rows(i).Item(5).ToString())
                End If
            Next

            Return dictionary
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Function

    Private Function FillDDlMajorCode() As Dictionary(Of String, String)
        Dim exMessage As String = " "
        Dim dictionary As New Dictionary(Of String, String)
        Try
            Dim dsMMajCodes = gnr.FillDDlMajorCode()

            dsMMajCodes.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsMMajCodes.Tables(0).Rows.Count - 1
                If dsMMajCodes.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsMMajCodes.Tables(0).Rows(i).Item(2).ToString() + " -- " + dsMMajCodes.Tables(0).Rows(i).Item(3).ToString()
                    'dsMinCodes = Trim(dsMinCodes.Tables(0).Rows(i).Item(0).ToString())
                    dsMMajCodes.Tables(0).Rows(i).Item(5) = fllValueName
                    'dsMinCodes.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                    dictionary.Add(dsMMajCodes.Tables(0).Rows(i).Item(2).ToString(), dsMMajCodes.Tables(0).Rows(i).Item(5).ToString())
                End If
            Next

            Return dictionary
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Function


    'Public Sub InsertProductDetails(projectNo As String, partstoshow As String, partNo As String)
    '    Dim dtTime As DateTimePicker = New DateTimePicker()
    '    Dim dtTime1 As DateTimePicker = New DateTimePicker()
    '    Dim dtTime2 As DateTimePicker = New DateTimePicker()
    '    Dim dtTime3 As DateTimePicker = New DateTimePicker()
    '    Dim dtTime4 As DateTimePicker = New DateTimePicker()
    '    Dim dtTime5 As DateTimePicker = New DateTimePicker()
    '    Dim QueryDetailResult As Integer = -1
    '    Dim exMessage As String = " "
    '    Try
    '        dtTime5.Value = New DateTime(1900, 1, 1)
    '        dtTime5.CustomFormat = "yyyy/MM/dd/"

    '        Dim strCheck = gnr.checkFields(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, txtqty.Text,
    '                                                            "", txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbStatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
    '                                                            cmbuser.SelectedValue, chknew, dtTime3, txtsample.Text, txttcost.Text, txtVendorNo.Text, partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, dtTime4,
    '                                                            dtTime5.Value.ToShortDateString(), txtsampleqty.Text)
    '        If String.IsNullOrEmpty(strCheck) Then
    '            QueryDetailResult = gnr.InsertProductDetail(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, txtqty.Text,
    '                                "", txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbStatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
    '                                cmbuser.SelectedValue, chknew, dtTime3, txtsample.Text, txttcost.Text, txtVendorNo.Text, partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, dtTime4,
    '                                dtTime5, CInt(txtsampleqty.Text))
    '            If QueryDetailResult <> 0 Then
    '                'show message error
    '            End If
    '        Else
    '            Dim arrayCheck As New List(Of String)
    '            arrayCheck = strCheck.Split(",").ToList()
    '            For Each item As String In arrayCheck
    '                If item = "Project Number" Then
    '                    'show error message must have data
    '                    Exit For
    '                ElseIf item = "Quantity" Then
    '                    txtqty.Text = "0"
    '                ElseIf item = "Unit Cost" Then
    '                    txtunitcost.Text = "0"
    '                ElseIf item = "Unit Cost New" Then
    '                    txtunitcostnew.Text = "0"
    '                ElseIf item = "Sample Cost" Then
    '                    txtsample.Text = "0"
    '                ElseIf item = "Misc. Cost" Then
    '                    txttcost.Text = "0"
    '                ElseIf item = "Vendor Number" Then
    '                    Exit For
    '                    'txtvendorno.Text = "0"  must have data
    '                ElseIf item = "Tooling Cost" Then
    '                    txttoocost.Text = "0"
    '                ElseIf item = "Sample Quantity" Then
    '                    txtsampleqty.Text = "0"
    '                End If
    '            Next

    '            If txtVendorNo.Text <> "" And projectNo <> 0 Then
    '                QueryDetailResult = gnr.InsertProductDetail(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, CInt(txtqty.Text),
    '                                "", txtmfrno.Text, CInt(txtunitcost.Text), CInt(txtunitcostnew.Text), txtpo.Text, dtTime2, cmbStatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
    '                                cmbuser.SelectedValue, chknew, dtTime3, CInt(txtsample.Text), CInt(txttcost.Text), CInt(txtVendorNo.Text), partstoshow, cmbminorcode.SelectedValue, CInt(txttoocost.Text), dtTime4,
    '                                dtTime5, CInt(txtsampleqty.Text))
    '            Else
    '                QueryDetailResult = -1
    '                MessageBox.Show("The project number an d vendor number must have value.", "CTP System", MessageBoxButtons.OK)
    '            End If

    '            If QueryDetailResult < 0 Then
    '                MessageBox.Show("Ann error ocurred inserting data in database.", "CTP System", MessageBoxButtons.OK)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
    '        MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
    '    End Try
    'End Sub

    Private Sub PoQotaFunction(Status2 As String, partNo As String)
        Dim exMessage As String = " "
        Dim statusquote As String
        'Dim Status2 As String = ""

        Try
            statusquote = "D-" & Status2
            Dim mpnopo As String = String.Empty
            Dim spacepoqota As String = String.Empty
            Dim unitCostNew As String = String.Empty
            Dim minQty As String = String.Empty
            Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtVendorNo.Text) & " AND PQPTN = '" & Trim(UCase(partNo)) & "'"
            Dim dsPoQota = gnr.GetPOQotaData(txtVendorNo.Text, partNo)

            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    'mpnopo = Trim(UCase(txtmfrno.Text))
                    Dim maxValue = 0
                    Dim dsUpdatedData As Integer

                    Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(partNo, txtVendorNo.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                        DateTime.Now.Day.ToString(), statusquote, spacepoqota, unitCostNew, minQty)
                    If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                        dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, minQty, unitCostNew, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                            txtVendorNo.Text, partNo)
                        If dsUpdatedData <> 0 Then
                            MessageBox.Show("An error ocurred updating fields.", "CTP System", MessageBoxButtons.OK)
                        End If
                    Else
                        Dim arrayCheck As New List(Of String)
                        arrayCheck = strCheckPoQoteIns.Split(",").ToList()
                        For Each item As String In arrayCheck
                            If item = "Sequencial" Then
                                'show error message
                                Exit For
                            ElseIf item = "Vendor Number" Then
                                txtVendorNo.Text = "0" 'ask for vendor??
                            ElseIf item = "Unit Cost New" Then
                                unitCostNew = "0"
                            ElseIf item = "Min Quantity" Then
                                minQty = "0"
                            End If
                        Next
                        dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, minQty, unitCostNew, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                            txtVendorNo.Text, partNo)

                        If dsUpdatedData <> 0 Then
                            'show message error
                        End If
                    End If
                Else
                    'warning message
                End If
            Else
                Dim maxValue = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)
                If Not String.IsNullOrEmpty(maxValue) Then
                    maxValue += 1
                Else
                    maxValue = 1
                End If
                spacepoqota = "                               DEV"
                'mpnopo = Trim(UCase(txtmfrno.Text))
                Dim ResultQuery As String = String.Empty

                Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(partNo, txtVendorNo.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                        DateTime.Now.Day.ToString(), statusquote, spacepoqota, unitCostNew, minQty)
                If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                    ResultQuery = gnr.InsertNewPOQota(partNo, txtVendorNo.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                       DateTime.Now.Day.ToString(), statusquote, spacepoqota, unitCostNew, minQty)
                    If ResultQuery <> 0 Then
                        'show message error
                    End If
                Else
                    Dim arrayCheck As New List(Of String)
                    arrayCheck = strCheckPoQoteIns.Split(",").ToList()
                    For Each item As String In arrayCheck
                        If item = "Sequencial" Then
                            'show error message
                            Exit For
                        ElseIf item = "Vendor Number" Then
                            txtVendorNo.Text = "0"
                        ElseIf item = "Unit Cost New" Then
                            unitCostNew = "0"
                        ElseIf item = "Min Qty" Then
                            minQty = "0"
                        End If
                    Next

                    ResultQuery = gnr.InsertNewPOQota(partNo, txtVendorNo.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                       DateTime.Now.Day.ToString(), statusquote, spacepoqota, unitCostNew, minQty)
                    If ResultQuery <> 0 Then
                        'show message error
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub LoadCombos(Optional ByVal sender As Object = Nothing, Optional ByVal e As EventArgs = Nothing)

        BackgroundWorker2.RunWorkerAsync()
        Loading.ShowDialog()
        Loading.BringToFront()

    End Sub

    Private Sub ExecuteCombos(Optional ByVal sender As Object = Nothing, Optional ByVal e As EventArgs = Nothing)
        Dim exMessage As String = " "

        Try
            cmbStatus.Items.Add("-- Select Status --")
            cmbStatus.Items.Add("I - In Process")
            cmbStatus.Items.Add("F - Finished")
            cmbStatus.SelectedIndex = 1

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub

    Private Sub setValues()

        Application.CurrentCulture = New CultureInfo("EN-US")

        cmdExcel.BackgroundImageLayout = ImageLayout.Stretch

        btnSuccess.Enabled = False
        btnInsert.Enabled = False
        btnCheck.Enabled = False
        btnSelect.Enabled = False
        'dtProjectDate.Value = Now
        DataGridView1.ReadOnly = True
        cmdExcel.Visible = False
        SplitContainer1.Visible = False

        txtProjectNo.SetWatermark("Project Number")
        txtProjectName.SetWatermark("Project Name")
        txtVendorNo.SetWatermark("Vendor Number")
        txtDesc.SetWatermark("Description")

        cmbStatus.SetWatermark("Project Status")
        cmbPerCharge.SetWatermark("Person In Charge")
        cmbStatusMore.SetWatermark("Project Status")

        ac1.SetWatermark("Vendor Name")

        txtVendorNo.Text = ""

        lblUsrLog.Text += userid

        DataGridView2.Enabled = LikeSession.gridEnable

        'cmbStatus.Items.Add("-- Select Status --")
        'cmbStatus.Items.Add("I - In Process")
        'cmbStatus.Items.Add("F - Finished")
        'cmbStatus.SelectedIndex = 1

        FillDDLStatus1()
        FillDDlUser1()

    End Sub

    Private Function utilDT(value As String, dtvalues As DataTable) As String
        Dim exMessage As String = Nothing
        Dim code As String = Nothing
        Try
            If value IsNot Nothing And dtvalues IsNot Nothing Then
                For Each item As DataRow In dtvalues.Rows
                    If item.ItemArray(0).ToString().Equals(value) Then
                        code = item.ItemArray(1).ToString()
                        Exit For
                    End If
                Next
            End If
            Return code
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function xlsDataSchemaValidation(dt As DataTable) As String
        Dim exMessage As String = " "
        'Dim blResult As Boolean = False
        Dim strResult As String = Nothing
        Try
            Dim userPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim rsPath As String = userPath & "\Excel_validation\"
            If Not Directory.Exists(rsPath) Then
                Directory.CreateDirectory(rsPath)
                'copiar archivo xsd del server
            End If

            deleteFilesInPath(rsPath)
            'If Not flagDelete Then
            Dim result = xmlConvertClass.CreateXltoXML(dt, rsPath, "MainNode", "reference")
            If result Then
                'blResult = If(String.IsNullOrEmpty(validationSchema(rsPath)), True, False)
                'Return blResult
                strResult = validationSchema(LikeSession.fullFilePath)
            Else
                strResult = "No XML Data."
            End If
            'Else
            '    MessageBox.Show("Please close the file previously created to process a new one.", "CTP System", MessageBoxButtons.OK)
            'End If

            xmlConvertClass.Dispose()
            'Dim rsPath = New Uri(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)).LocalPath
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            strResult = exMessage
            'Return blResult
        End Try
        Return strResult
    End Function

    Public Function validationSchema(rsPath As String) As String
        Dim exMessage As String = " "
        Dim blResult As Boolean = False
        Try
            Dim schema As XmlSchemaSet = New XmlSchemaSet()
            schema.Add("", gnr.UrlPathXsdFileMethod)
            Dim rd As XmlReader = XmlReader.Create(rsPath)
            Dim doc As XDocument = XDocument.Load(rd)
            doc.Validate(schema, AddressOf XSDErrors)
            Dim outMessage As String = Nothing
            outMessage = If(errors, "Not Validated. " & schemaErrorDesc, "")

            'blResult = If(outMessage.Equals("Validated"), True, False)
            Return outMessage
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return "Not Validated. " & ex.Message
        End Try
    End Function

    Private Sub XSDErrors(ByVal o As Object, ByVal e As ValidationEventArgs)
        Dim exMessage As String = " "
        Try
            Dim Type As XmlSeverityType = XmlSeverityType.Warning
            If [Enum].TryParse(Of XmlSeverityType)("Error", Type) Then
                If (Type = XmlSeverityType.Error) Then
                    errors = True
                    schemaErrorDesc = e.Message
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    'Private Sub cmdClearFilters_Click(sender As Object, e As EventArgs) Handles cmdClearFilters.Click

    'End Sub

    Protected Function IsFileLocked(file As FileInfo) As Boolean
        Dim exMessage As String = Nothing
        Dim stream As FileStream = Nothing
        Try
            stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None)
        Catch ex As IOException
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return True
        Finally
            If stream IsNot Nothing Then
                stream.Close()
            End If
        End Try
        Return False
    End Function

    Private Function areFilesInPath(strpath As String) As Boolean
        Dim exMessage As String = Nothing
        Try
            Dim myDir As DirectoryInfo = New DirectoryInfo(strpath)
            If myDir.EnumerateFiles().Any() Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return False
        End Try
    End Function

    Private Sub deleteFilesInPath(strpath As String)
        Dim exMessage As String = Nothing
        Dim deletedFiles As Boolean = False
        Try
            Dim directoryName As String = strpath
            For Each deleteFile In Directory.GetFiles(directoryName, "*.*", SearchOption.TopDirectoryOnly)
                Dim fi2 = New FileInfo(deleteFile)
                If Not IsFileLocked(fi2) Then
                    File.Delete(deleteFile)
                End If
            Next
            'deletedFiles = If(areFilesInPath(strpath) = False, True, False)
            'Return deletedFiles
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Return deletedFiles
        End Try
    End Sub

    Private Function checkIfPartAndVdrExist(partNo As String, vendorNo As String) As Boolean
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Dim rsReturn As Boolean = False
        Try
            ds = gnr.GetDataByVendorAndPartNoProdDesc(partNo, vendorNo)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    LikeSession.referencedExistence = ds.Tables(0).Rows(0).ItemArray(0).ToString()
                    rsReturn = True
                    Return rsReturn
                End If
                Return False
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            Return rsReturn
        End Try
    End Function

    Private Function InsertProductDetails(partNo As String, code As String, personInCharge As String, Optional ByVal vendorNoo As String = Nothing) As Integer
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
            'Dim strCheck = Nothing
            'If String.IsNullOrEmpty(strCheck) Then

#Region "Variable assign"

            Dim projectNoValue = code
            Dim PartNoValue = partNo
            Dim chkControl = New CheckBox()

#Region "Old Data"

            'Dim vendorNoValue = If(Qry.Columns.Contains("VMVNUM"), Qry.Rows(0).ItemArray(Qry.Columns("VMVNUM").Ordinal).ToString(), Nothing)
            'Dim CTPNoValue = If(Qry.Columns.Contains("PRDCTP"), Qry.Rows(0).ItemArray(Qry.Columns("PRDCTP").Ordinal).ToString(), Nothing)
            'Dim qtyValue = If(Qry.Columns.Contains("PRDQTY"), Qry.Rows(0).ItemArray(Qry.Columns("PRDQTY").Ordinal).ToString(), 0)
            'Dim MFRValue = ""
            'Dim MFRNoValue = If(Qry.Columns.Contains("PRDMFR#"), Qry.Rows(0).ItemArray(Qry.Columns("PRDMFR#").Ordinal).ToString(), Nothing)
            'Dim unitcostValue = If(Qry.Columns.Contains("PRDCOS"), Qry.Rows(0).ItemArray(Qry.Columns("PRDCOS").Ordinal).ToString(), 0)
            'Dim unitcostVValue = If(Qry.Columns.Contains("PRDCON"), Qry.Rows(0).ItemArray(Qry.Columns("PRDCON").Ordinal).ToString(), 0)
            'Dim chkControl = New CheckBox()
            'Dim cnew = If(Qry.Columns.Contains("PRDNEW"), Qry.Rows(0).ItemArray(Qry.Columns("PRDNEW").Ordinal).ToString(), Nothing)
            'Dim chkValue = If(String.IsNullOrEmpty(cnew),
            '    "0", Qry.Rows(0).ItemArray(Qry.Columns("PRDNEW").Ordinal).ToString())
            'Dim chkSelection = If(chkValue = "0", Not chkControl.Checked, chkControl.Checked)
            'chkControl.Checked = chkSelection

            ''Dim statusValue2 = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRDSTS").Ordinal).ToString()),
            ''    "E", Qry.Rows(0).ItemArray(Qry.Columns("PRDSTS").Ordinal).ToString())
            'Dim sts = If(Qry.Columns.Contains("PRDSTS"), Qry.Rows(0).ItemArray(Qry.Columns("PRDSTS").Ordinal).ToString(), Nothing)
            'Dim tempStatus = gnr.GetAllStatusesReturn(sts, "cntde1")
            'Dim statusValue = If(String.IsNullOrEmpty(sts),
            '    "E", If(String.IsNullOrEmpty(tempStatus), "E", tempStatus))

#End Region

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

#Region "Guidance"

            'PRHCOD,PRDPTN,PRDDAT,CRUSER,CRDATE,MOUSER,MODATE,PRDCTP,PRDQTY,PRDMFR,PRDMFR#,PRDCOS,PRDCON,PRDPO#,PODATE,PRDSTS,PRDBEN,PRDINF,PRDUSR,PRDNEW,
            'PRDEDD, PRDSCO, PRDTTC, VMVNUM, PRDPTS, PRDMPC, PRDTCO, PRDERD, PRDPDA, PRDSQTY

            'QueryDetailResult = gnr.InsertProductDetail(projectNoValue, PartNoValue, dtTime,
            '                    userid, dtTime1, userid, dtTime2, CTPNoValue, qtyValue,
            '                    MFRValue, MFRNoValue, unitcostValue, unitcostVValue,
            '                    poNoValue, dtTime3, statusValue, benefitsValue,
            '                    DetailsValue, personChValue, chkControl, dtTime4, samplecostValue,
            '                    misccostValue, vendorNoValue, partstoshow, minorcodeValue, toolingcostValue, dtTime5,
            '                    dtTime6, If(Not String.IsNullOrEmpty(sampleQtyValue), CInt(sampleQtyValue), 0))

#End Region

            Dim generalStatus = If(cmbStatusMore.SelectedIndex = 0 Or cmbStatusMore.SelectedIndex = -1, "E", cmbStatusMore.SelectedValue.ToString())
            QueryDetailResult = gnr.InsertProductDetail(projectNoValue, PartNoValue, dtTime,
                                    userid, dtTime1, userid, dtTime2, "", 0,
                                    "", "", 0, 0,
                                    "", dtTime3, generalStatus, "",
                                    "", personInCharge, chkControl, dtTime4, "0",
                                    "0", Trim(txtVendorNo.Text), "", "", "0", dtTime5,
                                    dtTime6, If(Not String.IsNullOrEmpty(""), CInt(""), "0"))

            If QueryDetailResult < 0 Then
                'MessageBox.Show("An error ocurred in the process.", "CTP System", MessageBoxButtons.OK)
                Return 1
            Else
                Return 0
            End If
            'End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
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

            DataGridView2.DataSource = Nothing
            DataGridView2.Refresh()

        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Function buildStatusString(status As String) As String
        Dim exMessage As String = ""
        Dim newValue As String = ""
        Try
            Dim dsStatuses = gnr.GetAllStatuses()

            'dsStatuses.Tables(0).Columns.Add("FullValue", GetType(String))

            'For i As Integer = 0 To dsStatuses.Tables(0).Rows.Count - 1
            '    If dsStatuses.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
            '        Dim fllValueName = dsStatuses.Tables(0).Rows(i).Item(2).ToString() + " -- " + dsStatuses.Tables(0).Rows(i).Item(3).ToString()
            '        dsStatuses.Tables(0).Rows(i).Item(5) = fllValueName
            '    End If
            'Next

            Dim dwResult = dsStatuses.Tables(0).AsEnumerable() _
                          .Where(Function(x) Trim(UCase(x.Field(Of String)("CNT03"))) = Trim(UCase(status)))
            Dim rowLenght = dwResult.LongCount
            If rowLenght > 0 Then
                newValue = Trim(dwResult(0).ItemArray(3).ToString())
                Return newValue
            Else
                Exit Function
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Sub disableAfterInsert(flag As Boolean)
        Dim exMessage As String = " "
        Dim myTableLayout As TableLayoutPanel
        Dim myTableLayout4 As TableLayoutPanel
        Try
            If flag Then
                myTableLayout = Me.TableLayoutPanel2
                For Each tt In myTableLayout.Controls
                    If TypeOf tt Is Windows.Forms.TextBox Then
                        tt.Enabled = flag
                        If flag Then
                            tt.Text = Nothing
                        End If
                    ElseIf TypeOf tt Is Autocomplete_Textbox Then
                        tt.Enabled = flag
                        If flag Then
                            tt.Text = Nothing
                        End If
                    ElseIf TypeOf tt Is Windows.Forms.ComboBox Then
                        tt.Enabled = flag
                        If flag Then
                            tt.selectedIndex = 0
                        End If
                    ElseIf TypeOf tt Is Windows.Forms.DateTimePicker Then
                        tt.Enabled = flag
                    ElseIf TypeOf tt Is Windows.Forms.Button Then
                        If tt.Name = "btnSuccess" Or tt.Name = "btnInsert" Then
                            tt.Enabled = flag
                        Else
                            tt.Enabled = Not flag
                        End If
                    ElseIf TypeOf tt Is Windows.Forms.SplitContainer Then
                        If tt.Name = "SplitContainer1" Then
                            If Not flag Then
                                Dim tlp As TableLayoutPanel = tt.Panel1.Controls("TableLayoutPanel6")
                                For Each ttt In tlp.Controls
                                    If TypeOf ttt Is Windows.Forms.DataGridView Then
                                        Dim dgv As DataGridView = ttt
                                        'dgv.ReadOnly = True
                                        For Each t4 As DataGridViewRow In dgv.Rows
                                            If t4.Cells("clPRHCOD").ToString() IsNot Nothing Then
                                                Dim index = t4.Index
                                                dgv.Rows(index).ReadOnly = Not flag
                                                'ttt.ReadOnly = False
                                            End If
                                        Next
                                    End If
                                Next
                            Else

                                tt.Visible = Not flag
                                Dim tlp1 As TableLayoutPanel = tt.Panel1.Controls("TableLayoutPanel6")
                                Dim tlp2 As TableLayoutPanel = tt.Panel1.Controls("TableLayoutPanel6")

                                For Each tttt In tlp1.Controls
                                    If TypeOf tttt Is Windows.Forms.DataGridView Then
                                        tttt.Datasource = Nothing
                                        tttt.Visible = Not flag
                                    End If
                                Next

                                For Each tttt In tlp2.Controls
                                    If TypeOf tttt Is Windows.Forms.DataGridView Then
                                        tttt.Datasource = Nothing
                                        tttt.Visible = Not flag
                                    End If
                                Next

                            End If
                        End If
                    End If

                    myTableLayout4 = Me.TableLayoutPanel4
                    For Each tt4 In myTableLayout4.Controls
                        If TypeOf tt4 Is Windows.Forms.TextBox Then
                            tt4.Enabled = flag
                            If flag Then
                                tt4.Text = Nothing
                            End If
                        ElseIf TypeOf tt4 Is Windows.Forms.Button Then
                            tt4.Enabled = flag
                            If flag Then
                                tt4.Text = Nothing
                            End If
                        End If
                    Next

                Next
            Else
                'Me.hide()
                'Me.ShowDialog()
            End If

        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub copyAlltoClipboard()
        DataGridView1.SelectAll()
        Dim dataObj As DataObject = DataGridView1.GetClipboardContent()
        If (dataObj IsNot Nothing) Then
            Clipboard.SetDataObject(dataObj)
        End If
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
            MessageBox.Show("Exception Occured while releasing object " + ex.ToString())
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Function Determine_OfficeVersion() As String
        Dim exMessage As String = " "
        Dim strExt As String = Nothing
        Try
            Dim strEVersionSubKey As String = "\Excel.Application\CurVer" '/HKEY_CLASSES_ROOT/Excel.Application/Curver

            Dim strValue As String 'Value Present In Above Key
            Dim strVersion As String 'Determines Excel Version
            Dim strExtension() As String = {"xls", "xlsx"}

            Dim rkVersion As RegistryKey = Nothing 'Registry Key To Determine Excel Version
            rkVersion = Registry.ClassesRoot.OpenSubKey(name:=strEVersionSubKey, writable:=False) 'Open Registry Key

            If Not rkVersion Is Nothing Then 'If Key Exists
                strValue = rkVersion.GetValue(String.Empty) 'get Value
                strValue = strValue.Substring(strValue.LastIndexOf(".") + 1) 'Store Value

                Select Case strValue 'Determine Version
                    Case "7"
                        strVersion = "95"
                        strExt = strExtension(0)
                    Case "8"
                        strVersion = "97"
                        strExt = strExtension(0)
                    Case "9"
                        strVersion = "2000"
                        strExt = strExtension(0)
                    Case "10"
                        strVersion = "2002"
                        strExt = strExtension(0)
                    Case "11"
                        strVersion = "2003"
                        strExt = strExtension(0)
                    Case "12"
                        strVersion = "2007"
                        strExt = strExtension(1)
                    Case "14"
                        strVersion = "2010"
                        strExt = strExtension(1)
                    Case "15"
                        strVersion = "2013"
                        strExt = strExtension(1)
                    Case "16"
                        strVersion = "2016"
                        strExt = strExtension(1)
                    Case Else
                        strExt = strExtension(1)
                End Select

                Return strExt
            Else
                MessageBox.Show("Microsoft Excel is not installed or corrupt in this computer.", "CTP System", MessageBoxButtons.OK)
                Return strExt
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            Return strExt
        End Try
    End Function

    Private Sub ac1_TextChanged(sender As Object, e As EventArgs)

    End Sub


#End Region

#Region "Not Used Now"

    'Private Sub HeaderCheckBox_Clicked(ByVal sender As Object, ByVal e As EventArgs)
    '    'Necessary to end the edit mode of the Cell.
    '    DataGridView1.EndEdit()

    '    'Loop and check and uncheck all row CheckBoxes based on Header Cell CheckBox.
    '    For Each row As DataGridViewRow In DataGridView1.Rows
    '        Dim checkBox As DataGridViewCheckBoxCell = (TryCast(row.Cells(0), DataGridViewCheckBoxCell))

    '        Dim myItem As CheckBox = CType(sender, CheckBox)
    '        'If myItem.ena Then

    '        'End If
    '        If Not checkBox.ReadOnly Then
    '            checkBox.Value = myItem.Checked
    '            'DataGridView1.CurrentCell = Nothing
    '        End If
    '    Next
    'End Sub

    'Private Sub Datagridview1_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) _
    '    Handles DataGridView1.CellBeginEdit
    '    Try
    '        'Dim value = DataGridView1(e.ColumnIndex, e.RowIndex).Value.ToString()

    '    Catch ex As Exception

    '    End Try

    'End Sub

    'Private Sub Datagridview1_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
    '    Handles DataGridView1.CellContentClick
    '    Try
    '        If e.ColumnIndex = 0 Then
    '            'Dim value = DataGridView1(e.ColumnIndex, e.RowIndex).Value.ToString()
    '            'Dim inputText = DataGridView1.EditingControl.Text

    '            'DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
    '            'If CBool(DataGridView1.CurrentCell.Value) = True Then
    '            '    Dim ppe = ""
    '            '    Dim calros = "1"

    '            '    Dim ok = ppe & " - " & calros
    '            'Else
    '            '    Dim ppe = ""
    '            '    Dim calros = "1"

    '            '    Dim ok = ppe & " - " & calros
    '            'End If
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub DataGridView1_CellMouseUp(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) _
    '    Handles DataGridView1.CellMouseUp
    '    Dim exMessage As String = " "
    '    Try
    '        If e.ColumnIndex = 0 Then
    '            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
    '            row.Cells(0).Value = Convert.ToBoolean(row.Cells(0).EditedFormattedValue)
    '            If Convert.ToBoolean(row.Cells(0).Value) Then

    '                DataGridView1(0, e.RowIndex).ReadOnly = True
    '            Else
    '                DataGridView1(0, e.RowIndex).ReadOnly = False
    '            End If
    '            'DataGridView1.CurrentCell = Nothing
    '        End If
    '    Catch ex As Exception
    '        exMessage = ex.Message + ". " + ex.ToString
    '    End Try

    'End Sub

    'Private Sub Datagridview1_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
    '    Handles DataGridView1.CellContentClick
    '    Try
    '        If e.ColumnIndex = 0 Then
    '            Dim value = DataGridView1(e.ColumnIndex, e.RowIndex).Value.ToString()
    '            'Dim inputText = DataGridView1.EditingControl.Text

    '            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
    '            If CBool(DataGridView1.CurrentCell.Value) = True Then
    '                Dim ppe = ""
    '                Dim calros = "1"

    '                Dim ok = ppe & " - " & calros
    '            Else
    '                Dim ppe = ""
    '                Dim calros = "1"

    '                Dim ok = ppe & " - " & calros
    '            End If
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub DataGridView1_CellMouseUp(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) _
    '    Handles DataGridView1.CellMouseUp
    '    Dim exMessage As String = " "
    '    Try
    '        If e.ColumnIndex = 0 Then
    '            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
    '            row.Cells("checkBoxColumn").Value = Convert.ToBoolean(row.Cells("checkBoxColumn").EditedFormattedValue)
    '            If Convert.ToBoolean(row.Cells("checkBoxColumn").Value) Then
    '                Dim value = DataGridView1(3, e.RowIndex).Value.ToString()
    '                LikeSession.flyingValue = value
    '                DataGridView1(3, e.RowIndex).ReadOnly = False
    '            Else
    '                DataGridView1(3, e.RowIndex).ReadOnly = True
    '            End If
    '        End If
    '    Catch ex As Exception
    '        exMessage = ex.Message + ". " + ex.ToString
    '    End Try

    'End Sub

    'Private Sub trimmMethod(dt As DataTable)
    '    Try
    '        For Each itemR As DataRow In dt.Rows
    '            For Each itemC As DataColumn In dt.Columns
    '                If TypeOf itemC.DataType Is String Then

    '                End If
    '            Next
    '        Next
    '    Catch ex As Exception

    '    End Try
    'End Sub

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
    '        exMessage = ex.Message + ". " + ex.ToString
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