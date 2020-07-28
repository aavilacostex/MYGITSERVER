Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports ExcelTools = Microsoft.Office

Imports Excel = Microsoft.Office.Interop.Excel

Imports ClosedXML.Excel
Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.Reflection
Imports System.Xml.Schema
Imports System.Xml
'Dim ac As New Autocomplete__module()

Public Class frmLoadExcel

    Private Excel03ConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX={2}'"
    Private Excel07ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX={2}'"
    Dim gnr As Gn1 = New Gn1()
    Dim xmlConvertClass As ConvertXml = New ConvertXml()
    Public userid As String
    Public flagallow As Integer

    Private Const totalRecords As Integer = 43
    Private Const pageSize As Integer = 10

    Dim bs As BindingSource = New BindingSource()
    Dim bs1 As BindingSource = New BindingSource()
    Dim Tables = New BindingList(Of DataTable)()
    Dim Tables1 = New BindingList(Of DataTable)()
    Dim errors As Boolean = False
    Dim schemaErrorDesc As String = Nothing

#Region "Page Load"

    Private Sub frmLoadExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim exMessage As String = " "
        Try
            userid = Trim(UCase(frmLogin.txtUserName.Text))
            If gnr.getFlagAllow(userid) = 1 Then
                flagallow = 1
            End If

            cmdExcel.BackgroundImageLayout = ImageLayout.Stretch

            btnSuccess.Enabled = False
            btnInsert.Enabled = False
            btnCheck.Enabled = False
            btnSelect.Enabled = False
            'dtProjectDate.Value = Now
            DataGridView1.ReadOnly = True
            cmdExcel.Visible = False
            SplitContainer1.Visible = False

            DataGridView2.Enabled = LikeSession.gridEnable

            cmbStatus.Items.Add("-- Select Status --")
            cmbStatus.Items.Add("I - In Process")
            cmbStatus.Items.Add("F - Finished")
            cmbStatus.SelectedIndex = 1
            FillDDlUser1()

            txtProjectNo.SetWatermark("Project Number")
            txtProjectName.SetWatermark("Project Name")
            txtVendorNo.SetWatermark("Vendor Number")
            txtDesc.SetWatermark("Description")

            cmbStatus.SetWatermark("Project Status")
            cmbPerCharge.SetWatermark("Person In Charge")

            'Autocomplete__module.create_textAutocomplete(txtVendorName)
            'Autocomplete__module.create_ddlAutocomplete(ComboBox1)

            ComboBox1.AutoCompleteMode = AutoCompleteMode.Append
            ComboBox1.DropDownStyle = ComboBoxStyle.DropDown
            ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems

            'Then Set ComboBox AutoComplete properties
            Dim ds = gnr.getVendorNoAndNameByNameDS()
            'Dim ds1 = gnr.getVendorsAccepted(ds)
            Dim bs = New BindingSource()
            bs.DataSource = ds.Tables(0)
            Dim dataview = New DataView(ds.Tables(0))
            Dim myTable As DataTable = dataview.ToTable(False, "VMNAME", "VMVNUM")


            Dim newRow As DataRow = myTable.NewRow
            newRow("VMNAME") = ""
            newRow("VMVNUM") = -1
            'dsUser.Tables(0).Rows.Add(newRow)
            myTable.Rows.InsertAt(newRow, 0)

            With ComboBox1
                .DisplayMember = "VMNAME"
                .ValueMember = "VMVNUM"
                .DataSource = myTable
                .DropDownStyle = ComboBoxStyle.DropDown
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With




        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Gridview,  dropdowns and textboxes methods"

    Public Function xlsDataSchemaValidation(dt As DataTable) As Boolean
        Dim exMessage As String = " "
        Dim blResult As Boolean = False
        Try
            Dim userPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim rsPath As String = userPath & "\Excel_validation\"
            If Not Directory.Exists(rsPath) Then
                Directory.CreateDirectory(rsPath)
                'copiar archivo xsd del server
            End If

            Dim result = xmlConvertClass.CreateXltoXML(dt, rsPath & "Input.xml", "MainNode")
            If result Then
                blResult = validationSchema(rsPath)
                Return blResult
            End If
            'Dim rsPath = New Uri(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)).LocalPath
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return blResult
        End Try
    End Function

    Public Function validationSchema(rsPath As String) As Boolean
        Dim exMessage As String = " "
        Dim blResult As Boolean = False
        Try
            Dim schema As XmlSchemaSet = New XmlSchemaSet()
            schema.Add("", rsPath + "xsdSchema.xsd")
            Dim rd As XmlReader = XmlReader.Create(rsPath + "Input.xml")
            Dim doc As XDocument = XDocument.Load(rd)
            doc.Validate(schema, AddressOf XSDErrors)
            Dim outMessage As String = Nothing
            outMessage = If(errors, "Not Validated. " & schemaErrorDesc, "Validated")

            blResult = If(outMessage.Equals("Validated"), True, False)
            Return blResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
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
        Dim aditionMessage As String = ""
        Try
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then

                    Dim dictionary As New Dictionary(Of String, String)
                    'dictionary.Add("PRNAME", "Project Name")
                    dictionary.Add("PRDPTN", "Part Number")
                    'dictionary.Add("VMVNUM", "Vendor Number")
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
                        dtError.Columns.Add("ErrorDesc", GetType(String))
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
                            If checkIfPartAndVdrExist(item.ItemArray(dt.Columns("PRDPTN").Ordinal).ToString(), txtVendorNo.Text) Then
                                dsError.Tables(0).ImportRow(item)
                                errorMessagee = message3
                                dsError.Tables(0).Rows(j).Item("ErrorDesc") = errorMessagee
                                j += 1
                            Else
                                If gnr.isPartInExistence(item.ItemArray(dt.Columns("PRDPTN").Ordinal).ToString()) Then
                                    Dim checkDuplicates = From data In dsResult.Tables(0).AsEnumerable()
                                                          Where Trim(UCase(data.Item("PRDPTN").ToString())) = Trim(UCase(item.ItemArray(dt.Columns("PRDPTN").Ordinal).ToString()))

                                    If checkDuplicates IsNot Nothing Then
                                        If Not checkDuplicates.Any() Then
                                            dsResult.Tables(0).ImportRow(item)
                                            i += 1
                                        End If
                                    End If
                                Else
                                    dsError.Tables(0).ImportRow(item)
                                    errorMessagee = message4
                                    dsError.Tables(0).Rows(j).Item("ErrorDesc") = errorMessagee
                                    j += 1
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            pi.SetValue(SplitContainer1, Convert.ChangeType(value, pi.PropertyType))
            If index.Equals(1) Then
                btnCheck.Enabled = Not value
                btnSuccess.Enabled = value
                DataGridView1.Visible = Not value
                DataGridView1.Enabled = Not value
                cmdExcel.Visible = value
                lblExcel.Visible = value
                buildNameReverse = "Panel" & index + 1 & "Collapsed"
                Dim pi2 As PropertyInfo = SplitContainer1.GetType().GetProperty(buildNameReverse)
                pi2.SetValue(SplitContainer1, Convert.ChangeType(Not value, pi2.PropertyType))
            Else
                btnCheck.Enabled = value
                btnSuccess.Enabled = Not value
                cmdExcel.Visible = Not value
                lblExcel.Visible = Not value
                'cmdExcel.Enabled = Not value
                DataGridView2.Visible = Not value
                DataGridView2.Enabled = Not value
                buildNameReverse = "Panel" & index - 1 & "Collapsed"
                Dim pi1 As PropertyInfo = SplitContainer1.GetType().GetProperty(buildNameReverse)
                pi1.SetValue(SplitContainer1, Convert.ChangeType(Not value, pi1.PropertyType))
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub txtProjectNo_TextChanged(sender As Object, e As EventArgs) Handles txtProjectNo.TextChanged
        If Not String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtProjectNo.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        ElseIf Not String.IsNullOrEmpty(txtProjectNo.Text) And String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        Else
            btnSelect.Enabled = False
        End If
    End Sub

    Private Sub txtProjectName_TextChanged(sender As Object, e As EventArgs) Handles txtProjectName.TextChanged
        If Not String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtProjectNo.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        ElseIf Not String.IsNullOrEmpty(txtProjectNo.Text) And String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtVendorNo.Text) Then
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox1.TextChanged
        'Dim result = gnr.getVendorNoAndNameByNameLike(txtVendorName.Text)
        'ComboBox1.DataSource = result
        'ComboBox1.Refresh()
        If ComboBox1.SelectedValue IsNot Nothing Then
            txtVendorNo.Text = ComboBox1.SelectedValue.ToString()
        End If
    End Sub

    Private Sub txtVendorNo_TextChanged_1(sender As Object, e As EventArgs) Handles txtVendorNo.TextChanged
        If Not String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtProjectNo.Text) And Not String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        ElseIf Not String.IsNullOrEmpty(txtProjectNo.Text) And String.IsNullOrEmpty(txtProjectName.Text) And String.IsNullOrEmpty(txtVendorNo.Text) Then
            btnSelect.Enabled = True
        Else
            btnSelect.Enabled = False
        End If
        btnValidVendor.Enabled = True
        txtVendorNo.Text = txtVendorNo.Text.Replace(Environment.NewLine, "")

        If txtVendorNo.Text = "-1" Then
            txtVendorNo.Text = ""
        End If
    End Sub

    Private Sub fillcell1(dt As DataTable, flag As Integer, dsName As String, Optional ByVal stopPag As Boolean = False)
        Dim exMessage As String = " "
        Try
            If (dsName.Equals("dsResult") Or dsName.Equals("dsGrig1")) Then
                DataGridView1.Columns.Clear()
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()
                DataGridView1.AutoGenerateColumns = False
                DataGridView1.ColumnCount = 4

                'Add Columns
                DataGridView1.Columns(0).Name = "clPRHCOD"
                DataGridView1.Columns(0).HeaderText = "Project No."
                DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

                DataGridView1.Columns(1).Name = "clPRDPTN"
                DataGridView1.Columns(1).HeaderText = "Part No."
                DataGridView1.Columns(1).DataPropertyName = "PRDPTN"

                DataGridView1.Columns(2).Name = "clVMVNUM"
                DataGridView1.Columns(2).HeaderText = "Vendor No."
                DataGridView1.Columns(2).DataPropertyName = "VMVNUM"

                DataGridView1.Columns(3).Name = "clPRDSTS"
                DataGridView1.Columns(3).HeaderText = "Status"
                DataGridView1.Columns(3).DataPropertyName = "PRDSTS"

                'FILL GRID
                DataGridView1.DataSource = dt

                'If String.IsNullOrEmpty(txtProjectNo.Text) Then
                If flag.Equals(0) Then
                    btnInsert_Click(Nothing, Nothing)
                End If
                'btnInsert_Click(Nothing, Nothing)
                'End If
                DataGridView1.Refresh()

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
                    toPaginate(DataGridView1)
                End If
            Else
                Dim dsError = LikeSession.dsErrorSession
                DataGridView2.DataSource = Nothing
                DataGridView2.Refresh()
                DataGridView2.AutoGenerateColumns = False
                DataGridView2.ColumnCount = 5

                'Add Columns
                DataGridView2.Columns(0).Name = "EditReference"
                DataGridView2.Columns(0).HeaderText = "Edit"
                DataGridView2.Columns(0).DataPropertyName = ""

                DataGridView2.Columns(1).Name = "AddReference"
                DataGridView2.Columns(1).HeaderText = "Add"
                DataGridView2.Columns(1).DataPropertyName = ""

                DataGridView2.Columns(2).Name = "clPRDPTN2"
                DataGridView2.Columns(2).HeaderText = "Part Number"
                DataGridView2.Columns(2).DataPropertyName = "PRDPTN"

                DataGridView2.Columns(3).Name = "clVMVNUM2"
                DataGridView2.Columns(3).HeaderText = "Vendor Number"
                DataGridView2.Columns(3).DataPropertyName = "VMVNUM"

                DataGridView2.Columns(4).Name = "clError"
                DataGridView2.Columns(4).HeaderText = "Error Description"
                DataGridView2.Columns(4).DataPropertyName = "ErrorDesc"

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
                                If Not String.IsNullOrEmpty(item.Cells(val).Value.ToString()) Then
                                    item.Cells(val).ReadOnly = True
                                End If
                            End If
                        Next
                    Next

                    DataGridView2.Columns(cellAmount).ReadOnly = True
                    DataGridView2.Refresh()

                    'btnCheck_Click(Nothing, Nothing)

                    If DataGridView2.Rows.Count > 0 And Not stopPag Then
                        toPaginate1(DataGridView2)
                    End If
                End If
            End If

        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            DataGridView2.DataSource = Nothing
            DataGridView2.Refresh()
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
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
                'Dim valueField = e.Value.ToString()
                CurrentState = If((e.Value IsNot Nothing), e.Value.ToString, "E")
                NewState = buildStatusString(CurrentState)
                If Not String.IsNullOrEmpty(NewState) Then
                    DataGridView1.Rows(e.RowIndex).Cells("clPRDSTS").Value = NewState
                Else
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

        If e.ColumnIndex = 0 Then
            DataGridView2.Rows(e.RowIndex).Cells(2).ReadOnly = False
            'DataGridView2.Rows(e.RowIndex).Cells(3).ReadOnly = False
            Dim value = DataGridView2.Rows(e.RowIndex).Cells(0).FormattedValue
            If value.Equals("Edit") Then
                DataGridView2.BeginEdit(True)
                LikeSession.acceptChanges = True
                DataGridView2.RefreshEdit()
            Else
                DataGridView2.BeginEdit(True)
                LikeSession.acceptChanges = False
                DataGridView2.RefreshEdit()
            End If
        ElseIf e.ColumnIndex = 1 Then
            Dim partValue = DataGridView2.Rows(e.RowIndex).Cells(2).Value.ToString()
            Dim vendorValue = DataGridView2.Rows(e.RowIndex).Cells(3).Value.ToString()
            If Not String.IsNullOrEmpty(partValue) Then
                'And Not String.IsNullOrEmpty(vendorValue) Then
                'Dim vendorOk = gnr.isVendorAccepted(vendorValue)
                Dim partOk = gnr.isPartInExistence(partValue)
                'If (vendorOk) Then
                If partOk Then
                    Dim myProjectNo = If(String.IsNullOrEmpty(txtProjectNo.Text), "", txtProjectNo.Text)
                    If String.IsNullOrEmpty(myProjectNo) Then
                        'InsertOnDemand(partValue, vendorValue, e.RowIndex)
                        InsertOnDemand(partValue, txtVendorNo.Text, e.RowIndex)
                    Else
                        'InsertOnDemand(partValue, vendorValue, e.RowIndex, myProjectNo)
                        InsertOnDemand(partValue, txtVendorNo.Text, e.RowIndex, txtProjectNo.Text)
                    End If
                Else
                    DataGridView2.Rows(e.RowIndex).Cells(4).Value = "The Part Number is not available at this moment."
                    MessageBox.Show("The Part Number is not available at this moment.", "CTP System", MessageBoxButtons.OK)
                End If
                'Else
                '    DataGridView2.Rows(e.RowIndex).Cells(4).Value = "The Vendor Number is not accepted as a valid vendor."
                '    MessageBox.Show("The Vendor Number is not accepted as a valid vendor.", "CTP System", MessageBoxButtons.OK)
                'End If
            Else
                DataGridView2.Rows(e.RowIndex).Cells(4).Value = "There is an error in the input values that prevent the insert process."
                MessageBox.Show("You must fill the value for the part for this reference.", "CTP System", MessageBoxButtons.OK)
            End If
        Else
            'DataGridView1_DoubleClick(sender, e)
        End If
    End Sub

    Private Sub DataGridView2_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged

        Dim exMessage As String = " "
        Try
            If e.RowIndex >= 0 Then
                If e.ColumnIndex = 2 Then
                    Dim inputText = If(DataGridView2.EditingControl IsNot Nothing, DataGridView2.EditingControl.Text, DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString())
                    'Dim inputText = DataGridView2.EditingControl.Text
                    If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) And gnr.isPartInExistence(inputText) Then
                        'DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value = Nothing
                        DataGridView2.EndEdit()
                        LikeSession.acceptChanges = True
                    Else
                        'DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
                        DataGridView2.CancelEdit()
                        'DataGridView2.RefreshEdit()
                        If (Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString())) Then
                            DataGridView2.Rows(e.RowIndex).Cells(4).Value = "The Part Number must have existences in stock."
                            MessageBox.Show("The Part Number must have existences in stock..", "CTP System", MessageBoxButtons.OK)
                        End If
                        LikeSession.acceptChanges = True
                    End If
                Else
                    If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) Then
                        DataGridView2.EndEdit()
                        LikeSession.acceptChanges = True
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub

    Private Sub DataGridView2_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit

        Dim exMessage As String = " "
        Try
            If e.RowIndex >= 0 Then
                If e.ColumnIndex = 2 Then
                    'Dim inputText = DataGridView2.EditingControl.Text
                    If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) And gnr.isPartInExistence(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) Then
                        DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).Value = Nothing 'clear error description
                        DataGridView2.EndEdit()
                        LikeSession.acceptChanges = True
                    ElseIf Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) And Not gnr.isPartInExistence(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) Then
                        DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
                        DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).Value = "The Part Number must have existences in stock."
                        DataGridView2.EndEdit()
                        LikeSession.acceptChanges = True
                    End If
                Else
                    'check for part validation
                End If
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub dataGridView2_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) Handles DataGridView2.CellBeginEdit


        Dim exMessage As String = " "
        Try
            If Not LikeSession.acceptChanges Then
                If Not String.IsNullOrEmpty(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()) And Not e.ColumnIndex.Equals(0) And Not e.ColumnIndex.Equals(1) Then
                    e.Cancel = True
                    LikeSession.acceptChanges = False
                End If
            Else
                e.Cancel = False
            End If

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub

    Private Sub DataGridView2_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

        Dim exMessage As String = " "
        Try
            If e.ColumnIndex = 2 Then
                Dim value = DataGridView2(e.ColumnIndex, e.RowIndex).Value.ToString()
                Dim inputText = DataGridView2.EditingControl.Text
                If Not Regex.IsMatch(inputText, "^[a-zA-Z0-9]{6,19}$") Then
                    DataGridView2.CancelEdit()
                    DataGridView2.RefreshEdit()
                    MessageBox.Show("The Part Number must be setted for a numeric value!", "CTP System", MessageBoxButtons.OK)
                End If
                'ElseIf e.ColumnIndex = 3 Then
                '    DataGridView2.CancelEdit()
                '    DataGridView2.RefreshEdit()
                '    Dim inputText = If(DataGridView2.EditingControl IsNot Nothing, DataGridView2.EditingControl.Text, DataGridView2(e.ColumnIndex, e.RowIndex).Value.ToString())
                '    If Not String.IsNullOrEmpty(inputText) Then
                '        If Not Regex.IsMatch(inputText, "^[0-9]{1,6}$") Then
                '            MessageBox.Show("The Vendor Number must match with an accepted vendor!", "CTP System", MessageBoxButtons.OK)
                '        End If
                '    End If
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Excel process"

    'Private Function GetTableDataXl(sender As Object, e As System.ComponentModel.CancelEventArgs) As DataTable
    '    Dim exMessage As String = " "
    '    Dim dt = New DataTable()
    '    Try
    '        Dim filePath As String = OpenFileDialog1.FileName
    '        Dim extension As String = Path.GetExtension(filePath)
    '        'Dim header As String = If(rbHeaderYes.Checked, "YES", "NO")
    '        Dim conStr As String, sheetName As String
    '        conStr = String.Empty
    '        Select Case extension

    '            Case ".xls"
    '                'Excel 97-03
    '                conStr = String.Format(Excel03ConString, filePath, "YES", 1)
    '                Exit Select

    '            Case ".xlsx"
    '                'Excel 07
    '                conStr = String.Format(Excel07ConString, filePath, "YES", 1)
    '                Exit Select
    '        End Select

    '        If String.IsNullOrEmpty(conStr) Then
    '            MessageBox.Show("File not valid. You must upload only excel files.", "CTP System", MessageBoxButtons.OK)
    '            Exit Sub
    '        End If

    '        'Get the name of the First Sheet.
    '        Using con As New OleDbConnection(conStr)
    '            Using cmd As New OleDbCommand()
    '                cmd.Connection = con
    '                con.Open()
    '                Dim dtExcelSchema As DataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
    '                sheetName = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
    '                con.Close()
    '            End Using
    '        End Using

    '        'Read Data from the First Sheet.
    '        Using con As New OleDbConnection(conStr)
    '            Using cmd As New OleDbCommand()
    '                Using oda As New OleDbDataAdapter()
    '                    Dim dt As New DataTable()
    '                    dt.Columns.Add("PRDPTN", GetType(String))
    '                    dt.AcceptChanges()
    '                    cmd.CommandText = (Convert.ToString("SELECT * From [") & sheetName) + "]"
    '                    cmd.Connection = con
    '                    con.Open()
    '                    oda.SelectCommand = cmd
    '                    'oda.TableMappings.Add("Table", "Net-informations.com")
    '                    oda.Fill(dt)
    '                    LikeSession.dsData = dt
    '                    fillData(dt)
    '                    'LoadThread()
    '                    'ExecuteFillData(dt)
    '                    con.Close()
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
    '    End Try
    'End Function

    Private Sub openFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim exMessage As String = " "
        Dim dt As New DataTable()
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

                    Dim dtSheetname As DataTable = New DataTable()

                    Dim cmd1 As OleDbCommand = Nothing
                    dtSheetname = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                    Using oda As New OleDbDataAdapter()
                        For Each dw As DataRow In dtSheetname.Rows

                            Dim query1 As String = "SELECT count(*) FROM [" + dw("TABLE_NAME").ToString() + "]"
                            cmd.CommandText = query1
                            cmd1 = New OleDbCommand(cmd.CommandText, con)
                            If (CInt(cmd1.ExecuteScalar()) > 0) Then
                                Dim query As String = "SELECT * FROM [" + dw("TABLE_NAME").ToString() + "]"
                                Dim data As OleDbDataAdapter = New OleDbDataAdapter(query, con)
                                data.Fill(dt)
                            End If
                        Next
                        con.Close()
                    End Using
                End Using
            End Using
            'Dim dtExcelSchema As DataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            'sheetName = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()



            'Read Data from the First Sheet.


            'Dim dt As New DataTable()
            'dt.Columns.Add("PRDPTN", GetType(String))
            'dt.AcceptChanges()
            'cmd.CommandText = (Convert.ToString("SELECT * From [") & sheetName) + "]"
            'cmd.Connection = con
            'con.Open()
            'oda.SelectCommand = cmd
            'oda.TableMappings.Add("Table", "Net-informations.com")
            'oda.Fill(dt)
            LikeSession.dsData = dt

            Dim validate = xlsDataSchemaValidation(dt)
            If validate Then
                fillData(dt)
            Else
                MessageBox.Show("File not valid.", "CTP System", MessageBoxButtons.OK)
            End If
            'LoadThread()
            'ExecuteFillData(dt)
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Threads"

    'Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) _
    '    Handles BackgroundWorker1.RunWorkerCompleted
    '    LoadingExcel.Close()
    'End Sub

    'Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) _
    '    Handles BackgroundWorker1.DoWork
    '    ExecuteFillData()
    'End Sub

    'Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) _
    '    Handles BackgroundWorker1.ProgressChanged
    '    'txtMfrNoSearch.Text = e.ProgressPercentage.ToString()
    'End Sub

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
                Else
                    txtVendorNo_TextChanged_1(Nothing, Nothing)
                    ComboBox1.SelectedIndex = ComboBox1.FindString(Trim(lblVendorDesc.Text))
                End If
            Else
                txtVendorNo.Text = Nothing
                lblVendorDesc.Text = Nothing
                ComboBox1.SelectedIndex = -1
                MessageBox.Show("The vendor number must have only numeric values and less than 6 characters.", "CTP System", MessageBoxButtons.OK)
            End If
            btnValidVendor.Enabled = False
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click

        'If Not String.IsNullOrEmpty(txtProjectName.Text) Then

        'End If
        'cleanFormValues()
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Dim exMessage As String = " "
        Dim countErrors As Integer = 0
        Dim Qry As New DataTable
        Dim iterator As Integer = 0
        Dim arraySuccess As New List(Of Integer)
        Dim arrayError As New List(Of Integer)
        Dim vendorNo = Trim(txtVendorNo.Text)
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
                    dsResult.Tables(0).Columns.Add("VMVNUM", GetType(String))

                    dsResult.Tables(0).Columns(0).DataType = GetType(String)

                    dsResult.AcceptChanges()

                    For Each dw As DataRow In dsResult.Tables(0).Rows
                        dw.Item("VmVNUM") = vendorNo
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
            If Not existProject Then
                projectPerCharge = If(cmbPerCharge.SelectedIndex = 0, userid, cmbPerCharge.SelectedValue)

                Dim dsExistsProject = gnr.GetExistByPRNAME(txtProjectName.Text)
                If dsExistsProject IsNot Nothing Then
                    Dim msgResult As DialogResult =
                        MessageBox.Show("The name " & txtProjectName.Text & " is in use in project number: " & dsExistsProject.Tables(0).Rows(0).ItemArray(0).ToString() & ". Do you want to create a new project with that name?", "CTP System", MessageBoxButtons.YesNo)
                    If msgResult = DialogResult.No Then
                        Exit Sub
                    End If
                End If
                queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, txtDesc.Text, txtProjectName.Text, cmbStatus, projectPerCharge)
            Else
                Dim ds = gnr.GetDataByPRHCOD(ProjectNoCurrent)
                For Each item As DataRow In ds.Tables(0).Rows
                    txtProjectName.Text = Trim(item.ItemArray(ds.Tables(0).Columns("PRNAME").Ordinal).ToString())
                    cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(item.ItemArray(ds.Tables(0).Columns("PRPECH").Ordinal).ToString()))
                    cmbStatus.SelectedIndex = cmbStatus.FindString(Trim(item.ItemArray(ds.Tables(0).Columns("PRSTAT").Ordinal).ToString()))
                    txtDesc.Text = Trim(item.ItemArray(ds.Tables(0).Columns("PRINFO").Ordinal).ToString())
                    dtProjectDate.Value = CDate(item.ItemArray(ds.Tables(0).Columns("PRDATE").Ordinal)).ToShortDateString()
                Next
                '?
            End If

            If queryResult < 0 Then
                'error message insertion
            Else
                For Each row As DataGridViewRow In DataGridView1.Rows
                    'save
                    Dim partNo = row.Cells("clPRDPTN").Value
                    'Dim vendorNo = row.Cells("clVMVNUM").Value

                    Dim Qry1 = dsResult.Tables(0).AsEnumerable() _
                                         .Where(Function(x) Trim(UCase(x.Field(Of String)("VMVNUM")).ToString()) = Trim(UCase(vendorNo)) And
                                         Trim(UCase(x.Field(Of String)("PRDPTN"))).ToString() = Trim(UCase(partNo)))

                    If Qry1.Count > 0 Then
                        Qry = Qry1.CopyToDataTable
                        'Dim personInChargeValue = If(String.IsNullOrEmpty(Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString()), userid, Qry.Rows(0).ItemArray(Qry.Columns("PRPECH").Ordinal).ToString())
                        Dim personInChargeValue = userid

                        Dim rsInsert = InsertProductDetails(Qry, ProjectNoCurrent, personInChargeValue)

                        'add to error dataset if insertion fails
                        If rsInsert < 0 Then
                            Dim dsError = LikeSession.dsErrorSession

                            Dim dtError = dsError.Tables(0).Copy()
                            dtError.Columns.Add("VMVNUM", GetType(String))

                            For Each dw1 As DataRow In dtError.Rows
                                dw1.Item("VMVNUM") = vendorNo
                            Next

                            Dim row1 As DataRow = dtError.NewRow()
                            row1(0) = Qry.Rows(0).ItemArray(Qry.Columns("PRDPTN").Ordinal).ToString()
                            row1(2) = Qry.Rows(0).ItemArray(Qry.Columns("VMVNUM").Ordinal).ToString()
                            row1(1) = "Error inserting the project reference."

                            dtError.Rows.Add(row1)
                            dtError.AcceptChanges()

                            dsError.Tables.RemoveAt(0)
                            dsError.Tables.Add(dtError)
                            dsError.AcceptChanges()
                            LikeSession.dsErrorSession = dsError
                        Else
                            'right insertion
                            If Not (dsResult.Tables(0).Columns.Contains("PRHCOD")) Then
                                dsResult.Tables(0).Columns.Add("PRHCOD", GetType(Integer))
                            End If

                            dsResult.Tables(0).Rows(iterator).Item("PRHCOD") = ProjectNoCurrent
                            dsResult.AcceptChanges()
                            iterator += 1

                            txtProjectNo.Text = ProjectNoCurrent
                            If cmbPerCharge.FindStringExact(Trim(projectPerCharge)) Then
                                cmbPerCharge.SelectedIndex = cmbPerCharge.FindString(Trim(projectPerCharge))
                            End If
                            'arraySuccess.Add(ProjectNoCurrent)
                        End If
                        'countErrors += InsertProductDetails(Qry)
                    Else
                        btnSuccess.Enabled = False
                        MessageBox.Show("The data has errors.", "CTP System", MessageBoxButtons.OK)
                    End If
                Next

                Dim rsReferences = gnr.GetReferencesInProject(ProjectNoCurrent)
                If rsReferences = 0 Then
                    Dim rsDeletion = gnr.DeleteDataFromProdHead(ProjectNoCurrent)
                    If rsDeletion < 0 Then
                        'error deleting go to dsError
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
                MessageBox.Show("The insertion process finished with some fails inserting data.", "CTP System", MessageBoxButtons.OK)
            Else
                MessageBox.Show("The insertion process finished successfully.", "CTP System", MessageBoxButtons.OK)
                disableAfterInsert()
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
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            Dim ProjectNoCurrent As Integer

            If String.IsNullOrEmpty(projectNo) Then

                Dim maxProjectNo = gnr.getmax("PRDVLH", "PRHCOD")
                ProjectNoCurrent = CInt(maxProjectNo) + 1
                Dim projectPerCharge = If(cmbPerCharge.SelectedIndex = 0, userid, cmbPerCharge.SelectedValue)

                Dim dsExistsProject = gnr.GetExistByPRNAME(txtProjectName.Text)
                If dsExistsProject IsNot Nothing Then
                    'decirlo y preguntar que hacer, puede actualizar o puede dejarlo
                    Dim msgResult As DialogResult =
                    MessageBox.Show("The name " & txtProjectName.Text & " is in use in project number: " & dsExistsProject.Tables(0).Rows(0).ItemArray(0).ToString() & ". Do you want to create a new project with that name?", "CTP System", MessageBoxButtons.YesNo)
                    If msgResult = DialogResult.Yes Then
                        queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, dtProjectDate, txtDesc.Text, txtProjectName.Text, cmbStatus, projectPerCharge)
                    Else
                        Exit Sub
                    End If
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

                Dim rsInsert = InsertProductDetails(dsInsert.Tables(0), ProjectNoCurrent, personInChargeValue)
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
                    newRow("PRDSTS") = "E"
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
                disableAfterInsert()
            End If

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub btnSuccess_Click(sender As Object, e As EventArgs) Handles btnSuccess.Click
        Dim exMessage As String = " "
        Try
            Dim dsValue = LikeSession.dsResultsSession
            fillcell1(dsValue.Tables(0), 0, dsValue.Namespace, True)
            setSplitContainerVisualization(1, False)
            'btnSuccess.Enabled = False
            'btnCheck.Enabled = True
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
                        MessageBox.Show("The file was created successfully.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Utils"

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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            AddHandler bs1.PositionChanged, AddressOf bs_PositionChanged1

            'bs_PositionChanged(bs, Nothing)
            bs_PositionChanged1(bs1, Nothing)

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub bs_PositionChanged(ByVal sender As Object, ByVal e As EventArgs)
        DataGridView1.DataSource = Tables(bs.Position)
    End Sub

    Private Sub bs_PositionChanged1(ByVal sender As Object, ByVal e As EventArgs)
        DataGridView2.DataSource = Tables1(bs1.Position)
    End Sub

    Public Sub handleDataGridColumnsOnDemand(dgvHandle As DataGridView, listToChange As List(Of Integer), index As Integer, flag As Boolean)
        Dim exMessage As String = " "
        Try
            For Each item As Integer In listToChange
                dgvHandle.Rows(index).Cells(item).ReadOnly = flag
            Next
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Public Sub handleDataGridColumns(handleDataRow As DataGridViewRow, listToChange As List(Of Integer), flag As Boolean)
        Dim exMessage As String = " "
        Try
            For Each item As Integer In listToChange
                handleDataRow.Cells(item).ReadOnly = flag
            Next
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return rsReturn
        End Try
    End Function

    Private Function InsertProductDetails(Qry As DataTable, code As String, personInCharge As String) As Integer
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
            Dim PartNoValue = Qry.Rows(0).ItemArray(Qry.Columns("PRDPTN").Ordinal).ToString()
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

            QueryDetailResult = gnr.InsertProductDetail(projectNoValue, PartNoValue, dtTime,
                                    userid, dtTime1, userid, dtTime2, "", 0,
                                    "", "", 0, 0,
                                    "", dtTime3, "E", "",
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

            DataGridView2.DataSource = Nothing
            DataGridView2.Refresh()

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Sub disableAfterInsert()
        Dim exMessage As String = " "
        Dim myTableLayout As TableLayoutPanel
        Dim myTableLayout4 As TableLayoutPanel
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
                ElseIf TypeOf tt Is Windows.Forms.SplitContainer Then
                    If tt.Name = "SplitContainer1" Then
                        Dim tlp As TableLayoutPanel = tt.Panel1.Controls("TableLayoutPanel6")
                        For Each ttt In tlp.Controls
                            If TypeOf ttt Is Windows.Forms.DataGridView Then
                                Dim dgv As DataGridView = ttt
                                'dgv.ReadOnly = True
                                For Each t4 As DataGridViewRow In dgv.Rows
                                    If t4.Cells("clPRHCOD").ToString() IsNot Nothing Then
                                        Dim index = t4.Index
                                        dgv.Rows(index).ReadOnly = True
                                        'ttt.ReadOnly = False
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If

                myTableLayout4 = Me.TableLayoutPanel4
                For Each tt4 In myTableLayout4.Controls
                    If TypeOf tt4 Is Windows.Forms.TextBox Then
                        tt4.Enabled = False
                    ElseIf TypeOf tt4 Is Windows.Forms.Button Then
                        tt4.Enabled = False
                    End If
                Next

            Next
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return strExt
        End Try
    End Function

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
    '        exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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
    '        exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
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