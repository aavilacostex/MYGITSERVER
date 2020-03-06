Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls

Public Class frmProductsDevelopment

    Dim gnr As Gn1 = New Gn1()

    Public flagdeve As Long '1 is new
    'Public filepicture As New clsReadWrite
    Public strwhere As String
    Public userid As String
    Public flagnewpart As Integer
    Public flagallow As Integer
    Public puragent As Integer
    Dim sql As String

    'the userid is burned. Need to fix!!!!!!!!!Importatnt!!!!!

    Private Sub frmProductsDevelopment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SSTab1.ItemSize = (New Size(SSTab1.Width / SSTab1.TabCount, 0))
        SSTab1.Padding = New System.Drawing.Point(300, 10)
        SSTab1.Appearance = TabAppearance.FlatButtons
        'TabControl1.ItemSize = New Size(0, 1)
        SSTab1.SizeMode = TabSizeMode.Fixed

        cmdsearch.FlatStyle = FlatStyle.Flat
        cmdsearchcode.FlatStyle = FlatStyle.Flat
        cmdsearch1.FlatStyle = FlatStyle.Flat
        cmdsearchpart.FlatStyle = FlatStyle.Flat
        cmdsearchctp.FlatStyle = FlatStyle.Flat
        cmdsearchstatus.FlatStyle = FlatStyle.Flat
        cmdall.FlatStyle = FlatStyle.Flat

        DataGridView1.RowHeadersVisible = False
        dgvProjectDetails.RowHeadersVisible = False

        'Button12.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\doc.PNG")
        cmdnew1.ImageAlign = ContentAlignment.MiddleRight
        cmdnew1.TextAlign = ContentAlignment.MiddleLeft

        ' Button13.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\save.PNG")
        cmdSave1.ImageAlign = ContentAlignment.MiddleRight
        cmdSave1.TextAlign = ContentAlignment.MiddleLeft

        'Button14.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\exit.PNG")
        cmdexit1.ImageAlign = ContentAlignment.MiddleRight
        cmdexit1.TextAlign = ContentAlignment.MiddleLeft

        'AddHandler DataGridView1.SelectionChanged, AddressOf dataGridView1_SelectionChanged
        'AddHandler dgvProjectDetails.DataBindingComplete, AddressOf dgvProjectDetails_DataBindingComplete

        'Datepickers customization

        DTPicker1.Format = DateTimePickerFormat.Custom
        DTPicker1.CustomFormat = "MM/dd/yyyy"

        DTPicker2.Format = DateTimePickerFormat.Custom
        DTPicker2.CustomFormat = "MM/dd/yyyy"

        DTPicker3.Format = DateTimePickerFormat.Custom
        DTPicker3.CustomFormat = "MM/dd/yyyy"

        DTPicker4.Format = DateTimePickerFormat.Custom
        DTPicker4.CustomFormat = "MM/dd/yyyy"

        'test purpose
        'Dim dss = gnr.GetTestData()

        'dropdownlist default fill section

        FillDDlUser() 'Fill user cmb
        FillDDlUser1()
        FillDDLStatus()
        FillDDlMinorCode()

        cmbprstatus.Items.Add("-- Select Status --")
        cmbprstatus.Items.Add("I - In Process")
        cmbprstatus.Items.Add("F - Finished")
        cmbprstatus.SelectedIndex = 1

        Dim posValue As Integer = 0
        For Each obj As DataRowView In cmbstatus.Items
            Dim VarQuery = "E"
            Dim VarCombo = Trim(obj.Item(2).ToString())
            If VarQuery = VarCombo Then
                cmbstatus.SelectedIndex = posValue
                Exit For
            Else
                posValue += 1
            End If
        Next

        Panel4.Enabled = False
        Panel1.Enabled = False

        optCTP.Checked = True
        optVENDOR.Checked = False
        optboth.Checked = False

        flagdeve = 1
        flagnewpart = 1


    End Sub

#Region "Combobox load Region"

    Private Sub FillDDlUser()
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

            cmbuser1.DataSource = dsUser.Tables(0)
            cmbuser1.DisplayMember = "FullValue"
            cmbuser1.ValueMember = "USUSER"


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

            cmbuser.DataSource = dsUser.Tables(0)
            cmbuser.DisplayMember = "FullValue"
            cmbuser.ValueMember = "USUSER"


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

            'Dim newRow As DataRow = dsStatuses.Tables(0).NewRow
            'newRow("CNT01") = "N/A"
            'newRow("CNT02") = "N/A"
            'newRow("CNT03") = "N/A"
            'newRow("CNTDE1") = "N/A -- NO NAME"
            'newRow("CNTDE2") = "NO STATUS"
            'newRow("FullValue") = "N/A"
            'dsStatuses.Tables(0).Rows.Add(newRow)
            'dsStatuses.Tables(0).Rows.InsertAt(newRow, 0)

            cmbstatus.DataSource = dsStatuses.Tables(0)
            cmbstatus.DisplayMember = "FullValue"
            cmbstatus.ValueMember = "CNT03"

            cmbstatus1.DataSource = dsStatuses.Tables(0)
            cmbstatus1.DisplayMember = "FullValue"
            cmbstatus1.ValueMember = "CNT03"

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub FillDDlMinorCode()
        Dim exMessage As String = " "
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
                End If
            Next

            cmbminorcode.DataSource = dsMinCodes.Tables(0)
            cmbminorcode.DisplayMember = "FullValue"
            cmbminorcode.ValueMember = "CNT03"

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Grid Events"

    Private Sub fillcell1(strwhere)
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT * FROM PRDVLH " & strwhere & " ORDER BY PRDATE DESC"   'DELETE BURNED REFERENCE
            'get the query results
            ds = gnr.FillGrid(sql)

            DataGridView1.AutoGenerateColumns = False
            DataGridView1.ColumnCount = 5

            'Add Columns
            DataGridView1.Columns(0).Name = "ProjectNo"
            DataGridView1.Columns(0).HeaderText = "Project No."
            DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

            DataGridView1.Columns(1).Name = "ProjectName"
            DataGridView1.Columns(1).HeaderText = "Project Name"
            DataGridView1.Columns(1).DataPropertyName = "PRNAME"

            DataGridView1.Columns(2).Name = "DateEnt"
            DataGridView1.Columns(2).HeaderText = "Date Entered"
            DataGridView1.Columns(2).DataPropertyName = "PRDATE"

            DataGridView1.Columns(3).Name = "PersonInCharge"
            DataGridView1.Columns(3).HeaderText = "Person In Charge"
            DataGridView1.Columns(3).DataPropertyName = "PRPECH"

            DataGridView1.Columns(4).Name = "Status"
            DataGridView1.Columns(4).HeaderText = "Status"
            DataGridView1.Columns(4).DataPropertyName = "PRSTAT"

            'FILL GRID
            DataGridView1.DataSource = ds.Tables(0)

            'changeControlAccess(True)
            'Exit Sub
        Catch ex As Exception
            Dim example As String = ex.Message
            Call gnr.gotoerror("frmproductsdevelopment", "fillcell1", Err.Number, Err.Description, Err.Source)
        End Try
    End Sub

    Private Sub fillcell1LastOne(strwhere)
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT * FROM PRDVLH " & strwhere & " ORDER BY PRHCOD DESC FETCH FIRST 1 ROW ONLY"   'DELETE BURNED REFERENCE
            'get the query results
            ds = gnr.FillGrid(sql)

            DataGridView1.AutoGenerateColumns = False
            DataGridView1.ColumnCount = 5

            'Add Columns
            DataGridView1.Columns(0).Name = "ProjectNo"
            DataGridView1.Columns(0).HeaderText = "Project No."
            DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

            DataGridView1.Columns(1).Name = "ProjectName"
            DataGridView1.Columns(1).HeaderText = "Project Name"
            DataGridView1.Columns(1).DataPropertyName = "PRNAME"

            DataGridView1.Columns(2).Name = "DateEnt"
            DataGridView1.Columns(2).HeaderText = "Date Entered"
            DataGridView1.Columns(2).DataPropertyName = "PRDATE"

            DataGridView1.Columns(3).Name = "PersonInCharge"
            DataGridView1.Columns(3).HeaderText = "Person In Charge"
            DataGridView1.Columns(3).DataPropertyName = "PRPECH"

            DataGridView1.Columns(4).Name = "Status"
            DataGridView1.Columns(4).HeaderText = "Status"
            DataGridView1.Columns(4).DataPropertyName = "PRSTAT"

            'FILL GRID
            DataGridView1.DataSource = ds.Tables(0)

            'changeControlAccess(True)
            'Exit Sub
        Catch ex As Exception
            Dim example As String = ex.Message
            Call gnr.gotoerror("frmproductsdevelopment", "fillcell1", Err.Number, Err.Description, Err.Source)
        End Try
    End Sub

    Private Sub fillcell2(code As String)
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT PRDDAT,PRDPTN,PRDCTP,PRDMFR#,PRDVLD.VMVNUM,VMNAME,PRDSTS FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & code & " "  'DELETE BURNED REFERENCE
            'get the query results
            ds = gnr.FillGrid(sql)

            dgvProjectDetails.AutoGenerateColumns = False
            'dgvProjectDetails.ColumnCount = 8

            'Add Columns
            dgvProjectDetails.Columns(0).Name = "Date"
            dgvProjectDetails.Columns(0).HeaderText = "Date"
            dgvProjectDetails.Columns(0).DataPropertyName = "PRDDAT"

            dgvProjectDetails.Columns(1).Name = "PartNo"
            dgvProjectDetails.Columns(1).HeaderText = "Part#"
            dgvProjectDetails.Columns(1).DataPropertyName = "PRDPTN"

            dgvProjectDetails.Columns(2).Name = "CTPNo"
            dgvProjectDetails.Columns(2).HeaderText = "CTP#"
            dgvProjectDetails.Columns(2).DataPropertyName = "PRDCTP"

            dgvProjectDetails.Columns(3).Name = "MFRNo"
            dgvProjectDetails.Columns(3).HeaderText = "MFR#"
            dgvProjectDetails.Columns(3).DataPropertyName = "PRDMFR#"

            dgvProjectDetails.Columns(4).Name = "Vendor"
            dgvProjectDetails.Columns(4).HeaderText = "Vendor"
            dgvProjectDetails.Columns(4).DataPropertyName = "VMVNUM"

            dgvProjectDetails.Columns(5).Name = "VendorName"
            dgvProjectDetails.Columns(5).HeaderText = "Vendor Name"
            dgvProjectDetails.Columns(5).DataPropertyName = "VMNAME"

            dgvProjectDetails.Columns(6).Name = "Status"
            dgvProjectDetails.Columns(6).HeaderText = "Status"
            dgvProjectDetails.Columns(6).DataPropertyName = "PRDSTS"



            'FILL GRID
            dgvProjectDetails.DataSource = ds.Tables(0)
            'dgvProjectDetails_DataBindingComplete(Nothing, Nothing)

            Exit Sub
        Catch ex As Exception
            Dim example As String = ex.Message
            Call gnr.gotoerror("frmproductsdevelopment", "fillcell2", Err.Number, Err.Description, Err.Source)
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) _
    Handles DataGridView1.CellFormatting
        Dim CurrentState As String = ""
        If e.ColumnIndex = 4 Then
            If e.Value IsNot Nothing Then
                CurrentState = e.Value.ToString
                If CurrentState = "I" Then
                    DataGridView1.Rows(e.RowIndex).Cells("Status").Value = "In Process"
                ElseIf CurrentState = "F" Then
                    e.CellStyle.ForeColor = Color.Red
                    e.Value = "Finished"
                    'DataGridView1.Rows(e.RowIndex).Cells("Status").Value = "Finished"
                End If
            End If
        End If
    End Sub

    Private Sub dgvProjectDetails_DataBindingComplete(ByVal sender As Object, ByVal e As DataGridViewBindingCompleteEventArgs) _
       Handles dgvProjectDetails.DataBindingComplete

        Dim CurrentState As String = " "
        Dim NewState As String = " "

        For Each row As DataGridViewRow In dgvProjectDetails.Rows
            CurrentState = row.Cells(6).Value.ToString()
            If CurrentState.Length <= 4 Then
                NewState = gnr.GetProjectStatusDescription(CurrentState)
                row.Cells(6).Value = NewState
            Else
                Exit For
            End If
        Next

        dgvProjectDetails.AutoResizeColumns()

    End Sub

    Private Sub dgvProjectDetails_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles dgvProjectDetails.DoubleClick
        Dim Index As Integer
        Dim ds As New DataSet()
        Dim ds1 As New DataSet()
        Dim RowDs As DataRow
        ds.Locale = CultureInfo.InvariantCulture
        ds1.Locale = CultureInfo.InvariantCulture
        Dim exMessage As String = " "
        Dim code As String = txtCode.Text
        Dim partDescription As String
        Dim dtSecondTb As DataTable = New DataTable()
        Dim columnToChange As String = "DVPRMG"
        Dim columnToChange1 As String = "VMNAME"

        Try
            For Each row As DataGridViewRow In dgvProjectDetails.SelectedRows
                Index = dgvProjectDetails.CurrentCell.RowIndex
                If dgvProjectDetails.Rows(Index).Selected = True Then
                    Dim part As String = row.Cells(1).Value.ToString()
                    ds = gnr.GetDataByCodeAndPartNo(code, part)
                    partDescription = gnr.GetDataByPartNo(part)
                    ds1 = gnr.GetDataByPartNo2(part)
                    If ds.Tables(0).Rows.Count > 0 Then
                        SSTab1.SelectedTab = TabPage3
                        For Each RowDs In ds.Tables(0).Rows

                            Dim CleanDateString As String = Regex.Replace(RowDs.Item(2).ToString(), "/[^0-9a-zA-Z:]/g", "")
                            Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
                            DTPicker2.Value = dtChange.ToShortDateString()

                            txtvendorno.Text = RowDs.Item(ds.Tables(0).Columns("VMVNUM").Ordinal).ToString()
                            txtvendorname.Text = RowDs.Item(ds.Tables(0).Columns("VMNAME").Ordinal).ToString()
                            txtpartno.Text = RowDs.Item(ds.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                            txtctpno.Text = RowDs.Item(ds.Tables(0).Columns("PRDCTP").Ordinal).ToString()
                            txtqty.Text = RowDs.Item(ds.Tables(0).Columns("PRDQTY").Ordinal).ToString()
                            txtmfr.Text = RowDs.Item(ds.Tables(0).Columns("PRDMFR").Ordinal).ToString()
                            txtmfrno.Text = RowDs.Item(ds.Tables(0).Columns("PRDMFR#").Ordinal).ToString()
                            txtsampleqty.Text = RowDs.Item(ds.Tables(0).Columns("PRDSQTY").Ordinal).ToString()
                            'txtminqty.Text = RowDs.Item(ds.Tables(0).Columns("PQMIN").Ordinal).ToString()
                            txtunitcostnew.Text = RowDs.Item(ds.Tables(0).Columns("PRDCON").Ordinal).ToString()
                            txtunitcost.Text = RowDs.Item(ds.Tables(0).Columns("PRDCOS").Ordinal).ToString()
                            txtsample.Text = RowDs.Item(ds.Tables(0).Columns("PRDSCO").Ordinal).ToString()
                            txttcost.Text = RowDs.Item(ds.Tables(0).Columns("PRDTTC").Ordinal).ToString()
                            txttoocost.Text = RowDs.Item(ds.Tables(0).Columns("PRDTCO").Ordinal).ToString()
                            txtpo.Text = RowDs.Item(ds.Tables(0).Columns("PRDPO#").Ordinal).ToString()
                            txtBenefits.Text = RowDs.Item(ds.Tables(0).Columns("PRDBEN").Ordinal).ToString()

                            txtminqty.Text = gnr.GetDataByVendorAndPartNo(txtvendorno.Text, txtpartno.Text)
                            flagdeve = 0
                            flagnewpart = 0

                            If cmbuser.FindStringExact(Trim(RowDs.Item(18).ToString())) Then
                                cmbuser.SelectedIndex = cmbuser.FindString(Trim(RowDs.Item(18).ToString()))
                            End If

                            Dim posValue As Integer = 0
                            For Each obj As DataRowView In cmbstatus.Items
                                Dim VarQuery = Trim(RowDs.Item(ds.Tables(0).Columns("PRDSTS").Ordinal).ToString())
                                Dim VarCombo = Trim(obj.Item(2).ToString())
                                If VarQuery = VarCombo Then
                                    cmbstatus.SelectedIndex = posValue
                                    Exit For
                                Else
                                    posValue += 1
                                End If
                            Next

                            Dim posValueMin As Integer = 0
                            For Each obj As DataRowView In cmbminorcode.Items
                                Dim VarQuery = Trim(RowDs.Item(ds.Tables(0).Columns("PRDMPC").Ordinal).ToString())
                                Dim VarCombo = Trim(obj.Item(2).ToString())
                                If VarQuery = VarCombo Then
                                    cmbminorcode.SelectedIndex = posValueMin
                                    Exit For
                                Else
                                    posValueMin += 1
                                End If
                            Next

                            txtpartdescription.Text = partDescription

                            Dim rdValue = RowDs.Item(ds.Tables(0).Columns("PRDPTS").Ordinal).ToString()
                            If rdValue = "1" Then
                                optCTP.Checked = True
                                optVENDOR.Checked = False
                                optboth.Checked = False
                            ElseIf rdValue = "2" Then
                                optCTP.Checked = False
                                optVENDOR.Checked = True
                                optboth.Checked = False
                            ElseIf rdValue = "" Then
                                optCTP.Checked = False
                                optVENDOR.Checked = False
                                optboth.Checked = True
                            End If

                        Next

                        If cmbuser.SelectedIndex = -1 Then
                            cmbuser.SelectedIndex = cmbuser1.Items.Count - 1
                        End If

                        Dim ctIndex = ds1.Tables(0).Columns(columnToChange).Ordinal
                        Dim ctIndex1 = ds1.Tables(0).Columns(columnToChange1).Ordinal
                        txtvendornoa.Text = ds1.Tables(0).Rows(0).ItemArray(ctIndex).ToString()
                        txtvendornamea.Text = ds1.Tables(0).Rows(0).ItemArray(ctIndex1).ToString()
                    End If


                End If
            Next

            changeControlAccess(True)

            cmbminorcode.Enabled = False
            txtminor.Enabled = False
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView1.DoubleClick
        Dim Index As Integer
        Dim ds As New DataSet()
        Dim RowDs As DataRow
        ds.Locale = CultureInfo.InvariantCulture
        Dim exMessage As String = " "

        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                Index = DataGridView1.CurrentCell.RowIndex
                If DataGridView1.Rows(Index).Selected = True Then
                    Dim code As String = row.Cells(0).Value.ToString()

                    ds = gnr.GetDataByPRHCOD(code)
                    If ds.Tables(0).Rows.Count = 1 Then

                        SSTab1.SelectedTab = TabPage2
                        For Each RowDs In ds.Tables(0).Rows
                            txtCode.Text = Trim(RowDs.Item(0).ToString())
                            txtname.Text = Trim(RowDs.Item(3).ToString()) ' format date
                            TabPage2.Text = "Project: " + txtname.Text

                            Dim CleanDateString As String = Regex.Replace(RowDs.Item(1).ToString(), "/[^0-9a-zA-Z:]/g", "")
                            'Dim dtChange As DateTime = DateTime.ParseExact(CleanDateString, "MM/dd/yyyy HH:mm:ss tt", CultureInfo.InvariantCulture)
                            Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
                            DTPicker1.Value = dtChange.ToShortDateString()

                            If cmbuser1.FindStringExact(Trim(RowDs.Item(9).ToString())) Then
                                cmbuser1.SelectedIndex = cmbuser1.FindString(Trim(RowDs.Item(9).ToString()))
                            End If

                            If cmbuser1.SelectedIndex = -1 Then
                                cmbuser1.SelectedIndex = cmbuser1.Items.Count - 1
                            End If


                            If Trim(RowDs.Item(4).ToString()) = "I" Then
                                cmbprstatus.SelectedIndex = 1
                            ElseIf Trim(RowDs.Item(4).ToString()) = "F" Then
                                cmbprstatus.SelectedIndex = 2
                            Else
                                cmbprstatus.SelectedIndex = 2
                            End If
                            'Dim Test1 = RowDs.Item(1).ToString() get the value begans with 0 pos
                            'Dim test2 = ds.Tables(0).Columns.Item(1).ColumnName  get the grid header
                        Next
                    Else
                        'message box warning
                    End If

                    'fill second grid process
                    fillcell2(code)
                Else
                    'is is not selected
                End If
            Next
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try


    End Sub

#End Region

#Region "Textbox events"

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles txtsearch.GotFocus
        txtsearch.SelectionStart = 0
        txtsearch.SelectionLength = Len(Trim(txtsearch.Text))
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles txtsearch1.GotFocus
        txtsearch1.SelectionStart = 0
        txtsearch1.SelectionLength = Len(Trim(txtsearch.Text))
    End Sub

#End Region

#Region "Button Events"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdsearch.Click
        Try
            If Trim(txtsearch.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(txtsearch.Text)), "'", "") & "%'"
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "')) AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(txtsearch.Text)), "'", "") & "%'"
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdall_Click()
        Try
            strwhere = CustomStrWhereResult()
            fillcell1(strwhere)
            Exit Sub
        Catch ex As Exception
            'Call gnr.gotoerror("frmproductsdevelopment", "cmdall_click", Err.Number, Err.Description, Err.Source)
            gnr.gotoerror("frmproductsdevelopment", "cmdall_click", ex.HResult, ex.Message + ". " + ex.ToString, ex.Source)
        End Try
    End Sub

    Private Sub cmdall_Click(sender As Object, e As EventArgs) Handles cmdall.Click
        cmdall_Click()
    End Sub

    Private Sub cmdnew3_Click(sender As Object, e As EventArgs) Handles cmdnew3.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("No pressed")
        ElseIf result = DialogResult.Yes Then
            'MessageBox.Show("Yes pressed")
            gotonew()
        End If
    End Sub

    Private Sub cmdnew2_Click(sender As Object, e As EventArgs) Handles cmdnew3.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("No pressed")
        ElseIf result = DialogResult.Yes Then
            'MessageBox.Show("Yes pressed")
            gotonew()
        End If
    End Sub

    Private Sub cmdnew1_Click(sender As Object, e As EventArgs) Handles cmdnew3.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("No pressed")
        ElseIf result = DialogResult.Yes Then
            'MessageBox.Show("Yes pressed")
            gotonew()
        End If
    End Sub

    Private Sub cmdexit1_Click(sender As Object, e As EventArgs) Handles cmdexit1.Click
        Me.Close()
    End Sub

    Private Sub cmdexit2_Click(sender As Object, e As EventArgs) Handles cmdexit2.Click
        Me.Close()
    End Sub

    Private Sub cmdexit3_Click(sender As Object, e As EventArgs) Handles cmdexit3.Click
        Me.Close()
    End Sub

    Private Sub gotonew()
        SSTab1.SelectedTab = TabPage2
        cleanValues()
    End Sub

    Private Sub save()
        Dim exMessage As String = " "
        Try

            Dim insertYear As String = (Year(Now()) - 2000)
            'Dim test As String
            'insertYear = insertYear.Substring(1, 2)
            insertYear = CInt(insertYear)
            Dim insertMonth = Date.Today.Month
            Dim insertDay = Date.Today.Day
            Dim partstoshow As String = displayPart()
            Dim dtTime As DateTimePicker = New DateTimePicker()
            Dim dtTime1 As DateTimePicker = New DateTimePicker()
            Dim dtTime2 As DateTimePicker = New DateTimePicker()
            Dim dtTime3 As DateTimePicker = New DateTimePicker()
            Dim dtTime4 As DateTimePicker = New DateTimePicker()
            Dim dtTime5 As DateTimePicker = New DateTimePicker()
            dtTime5.Value = New DateTime(1900, 1, 1)
            dtTime5.CustomFormat = "yyyy/MM/dd/"

            If flagdeve = 1 Then
                Dim ProjectNo = gnr.getmax("PRDVLH", "PRHCOD") + 1
                Dim queryResult = gnr.InsertNewProject(ProjectNo, "LREDONDO", DTPicker1, txtainfo.Text, txtname.Text, cmbprstatus, cmbuser1)
                If queryResult < 0 Then
                    'error message
                Else
                    txtCode.Text = ProjectNo
                    flagdeve = 0

                    strwhere = CustomStrWhereResult()
                    fillcell1LastOne(strwhere)

                    If flagnewpart = 1 Then
                        If Trim(txtpartno.Text) <> " " Then '?????
                            Dim Status2 As String = ""
                            'Status2 = If(gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()) <> "", gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()), "")
                            Status2 = gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString())
                            Dim dsProjectNoResult As DataSet = gnr.GetCodeAndNameByPartNo("1527554")
                            Dim strProjectNo = If(String.IsNullOrEmpty(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()), 0, CInt(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()))
                            Dim strProjectName = Trim(dsProjectNoResult.Tables(0).Rows(0).ItemArray(1).ToString())

                            'test purpose
                            strProjectNo = ProjectNo
                            'test purpose

                            If (ProjectNo = strProjectNo) Then
                                Dim result As DialogResult = MessageBox.Show("This part no. already exists in this project. :" & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                If result = DialogResult.No Then
                                    MessageBox.Show("No pressed")
                                ElseIf result = DialogResult.Yes Then
                                    Dim QueryDetailResult = gnr.InsertProductDetail(ProjectNo, txtpartno.Text, DTPicker2, "LREDONDO", dtTime, "LREDONDO", dtTime1, txtctpno.Text, txtqty.Text,
                                                            txtmfr.Text, txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbstatus, txtBenefits.Text, txtcomm.Text,
                                                            cmbuser, chknew, dtTime3, txtsample.Text, txttcost.Text, txtvendorno.Text, partstoshow, cmbminorcode, txttoocost.Text, dtTime4,
                                                            dtTime5.Value.ToShortDateString(), CInt(txtsampleqty.Text))

                                    If QueryDetailResult < 0 Then
                                        'error message
                                    Else

                                    End If

                                End If
                            End If
                        End If
                        End If
                    End If
                End If


        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try



    End Sub

    Private Sub cmdSave2_Click(sender As Object, e As EventArgs) Handles cmdSave2.Click
        save()
    End Sub

    Private Sub cmdSave3_Click(sender As Object, e As EventArgs) Handles cmdSave3.Click
        save()
    End Sub

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click
        save()
    End Sub

#End Region

#Region "Utils"

    Private Sub changeControlAccess(value As Boolean)
        txtvendorno.ReadOnly = value
        txtvendorname.ReadOnly = value
        txtpartno.ReadOnly = value
        txtpartdescription.ReadOnly = value
        txtvendornoa.ReadOnly = value
        txtvendornamea.ReadOnly = value
        txtminor.ReadOnly = value
        txtCode.ReadOnly = value
    End Sub

    Private Sub cleanValues()

        txtCode.Text = ""
        txtname.Text = ""
        txtainfo.Text = ""
        txtpartno.Text = ""
        txtvendornoa.Text = ""
        txtvendornamea.Text = ""
        txtpo.Text = ""
        txtcomm.Text = ""
        txtBenefits.Text = ""
        txttoocost.Text = 0
        txtpartdescription.Text = ""
        txtvendorno.Text = ""
        txtvendorname.Text = ""
        txtctpno.Text = ""
        txtqty.Text = 0
        txtsampleqty.Text = 0
        txtmfr.Text = ""
        txtmfrno.Text = ""
        txtunitcost.Text = 0
        txtminqty.Text = 0
        txtsample.Text = 0
        txttcost.Text = 0
        txtunitcostnew.Text = 0

        dgvProjectDetails.DataSource = Nothing
        DataGridView1.DataSource = Nothing

        optCTP.Checked = True
        optVENDOR.Checked = False
        optboth.Checked = False
        chknew.Checked = False

        DTPicker1.Value = Format(Now, "MM/dd/yyyy")
        'DTPicker5.Value = "01/01/1900"
        DTPicker2.Value = Format(Now, "MM/dd/yyyy")
        'DTPicker3.Value = "01/01/1900"
        'DTPicker4.Value = "01/01/1900"

        FillDDlUser()
        cmbuser1.SelectedIndex = 0

        FillDDlUser1()
        cmbuser.SelectedIndex = 0

        FillDDLStatus()
        cmbstatus.SelectedIndex = 0

        cmbminorcode.Items.Clear()

        'cmbminorcode.Clear
        'cmbprstatus.ListIndex = 0
        'cmbstatus.ListIndex = 0

        TabPage2.Text = ""

        flagdeve = 1
        flagnewpart = 1

    End Sub

    Private Function displayPart() As String
        Dim result As String = "-1"
        If optCTP.Checked = True Then
            result = "1"
        ElseIf optVENDOR.Checked = True Then
            result = "2"
        ElseIf optboth.Checked = True Then
            result = ""
        End If
        Return result
    End Function

    Private Function CustomStrWhereResult() As String
        'If flagallow = 1 Then
        strwhere = ""
        'Else
        'TEST QUERY
        'strwhere = "WHERE PRPECH = 'LREDONDO' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = 'LREDONDO') "
        'strwhere = "WHERE PRPECH = '" & userid & "' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "') "
        'strwhere = "WHERE PRPECH = '" & UserID & "'
        'End If
        Return strwhere
    End Function

#End Region

    'Protected Sub OnRowCommand(ByVal sender As Object, ByVal e As GridViewCommandEventArgs)
    'Dim index As Integer = Convert.ToInt32(e.CommandArgument)
    'Dim gvRow As DataGridViewRow = DataGridView1.Rows(index)
    'End Sub'

    'Private Sub dataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
    'For Each row As DataGridViewRow In DataGridView1.SelectedRows
    'Dim value11 As String = row.Cells(0).Value.ToString()
    'Next
    'End Sub

End Class