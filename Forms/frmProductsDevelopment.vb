﻿Imports System.Globalization
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

        cmdSave1.Enabled = False

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
        'Dim dss = gnr.GetTestData("1527554")
        'Dim dss = gnr.GetPOQotaData()
        'dropdownlist default fill section
        'Dim varvar = 1439
        'Dim dstest = gnr.DeleteDataByProjectNo(varvar)


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

        'extra method
        Panel4.Enabled = False
        Panel1.Enabled = False
        txtCode.Enabled = False
        txtvendorno.ReadOnly = True
        txtvendorname.ReadOnly = True
        txtvendornamea.ReadOnly = True
        txtvendornoa.ReadOnly = True
        txtminor.ReadOnly = True
        txtpartno.ReadOnly = True
        txtpartdescription.ReadOnly = True
        cmbminorcode.Enabled = False

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

    Private Sub SSTab1_Selected(ByVal sender As Object, ByVal e As TabControlEventArgs) _
    Handles SSTab1.Selected

        If SSTab1.SelectedTab.Name = "TabPage1" Then
            cmdSave1.Enabled = False
        ElseIf SSTab1.SelectedTab.Name = "TabPage3" Then
            Dim rsValue As Integer = -1
            rsValue = mandatoryFields("new", "TabPage2")
            If rsValue = 0 Then
                flagdeve = 0
                flagnewpart = 1
            Else
                Dim rsMessage As DialogResult = MessageBox.Show("All the fields in the Project Tab must be filled before add parts!", "CTP System", MessageBoxButtons.OK)
                If rsMessage = DialogResult.OK Then
                    SSTab1.SelectedIndex = 1
                End If
            End If
        End If
    End Sub


#End Region

#Region "Grid Events"

    Private Sub fillcell1(strwhere)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT * FROM PRDVLH " & strwhere & " ORDER BY PRDATE DESC"   'DELETE BURNED REFERENCE
            'get the query results

            ds = gnr.FillGrid(sql)
            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then

                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
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
                Else
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub fillcell1LastOne(strwhere)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT * FROM PRDVLH " & strwhere & " ORDER BY PRHCOD DESC FETCH FIRST 1 ROW ONLY"   'DELETE BURNED REFERENCE
            'get the query results

            ds = gnr.FillGrid(sql)
            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
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
                Else
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub fillcell2(code As String)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT PRDDAT,PRDPTN,PRDCTP,PRDMFR#,PRDVLD.VMVNUM,VMNAME,PRDSTS FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & code & " "  'DELETE BURNED REFERENCE
            'get the query results
            ds = gnr.FillGrid(sql)

            If Not ds Is Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
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
                Else
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub fillcelldetail(strwhere)
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture

        Try
            sql = "SELECT distinct(prdvlh.prhcod),prname,prdate,prpech,prstat FROM PRDVLH INNER JOIN PRDVLD ON PRDVLH.PRHCOD = PRDVLD.PRHCOD " & strwhere & " ORDER BY PRDATE DESC"

            ds = gnr.FillGrid(sql)

            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then

                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
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
                Else
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
        Exit Sub
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
                    'clean all other fields
                    flagdeve = 0
                    flagnewpart = 1
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

    'Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles txtsearch.GotFocus
    '    txtsearch.SelectionStart = 0
    '    txtsearch.SelectionLength = Len(Trim(txtsearch.Text))
    'End Sub

    'Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles txtsearch1.GotFocus
    '    txtsearch1.SelectionStart = 0
    '    txtsearch1.SelectionLength = Len(Trim(txtsearch.Text))
    'End Sub

#End Region

#Region "Button Events"

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

        Dim validationResult = mandatoryFields("new", SSTab1.SelectedTab.Name)
        If validationResult.Equals(0) Then
            Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                'MessageBox.Show("No pressed")
            ElseIf result = DialogResult.Yes Then
                'MessageBox.Show("Yes pressed")
                cleanFormValues(SSTab1.SelectedTab.Name)
                gotonew()
            End If
        Else
            Dim resultNew As DialogResult = MessageBox.Show("You have data in the form. You could missing if continue. Do you want to proceed?", "CTP System", MessageBoxButtons.YesNo)
            If resultNew = DialogResult.Yes Then
                cleanFormValues(SSTab1.SelectedTab.Name)
                gotonew()
            End If
        End If

    End Sub

    Private Sub cmdnew2_Click(sender As Object, e As EventArgs) Handles cmdnew2.Click

        Dim validationResult = mandatoryFields("new", SSTab1.SelectedTab.Name)
        If validationResult.Equals(0) Then
            Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                'MessageBox.Show("No pressed")
            ElseIf result = DialogResult.Yes Then
                'MessageBox.Show("Yes pressed")
                cleanFormValues(SSTab1.SelectedTab.Name)
                gotonew()
            End If
        Else
            Dim resultNew As DialogResult = MessageBox.Show("You have data in the form. You could missing if continue. Do you want to proceed?", "CTP System", MessageBoxButtons.YesNo)
            If resultNew = DialogResult.Yes Then
                cleanFormValues(SSTab1.SelectedTab.Name)
                gotonew()
            End If
        End If

    End Sub

    Private Sub cmdnew1_Click(sender As Object, e As EventArgs) Handles cmdnew1.Click

        Dim validationResult = mandatoryFields("new", SSTab1.SelectedTab.Name)
        If validationResult.Equals(0) Then
            Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                'MessageBox.Show("No pressed")
            ElseIf result = DialogResult.Yes Then
                'MessageBox.Show("Yes pressed")
                cleanFormValues(SSTab1.SelectedTab.Name)
                gotonew()
            End If
        Else
            Dim resultNew As DialogResult = MessageBox.Show("You have data in the form. You could missing if continue. Do you want to proceed?", "CTP System", MessageBoxButtons.YesNo)
            If resultNew = DialogResult.Yes Then
                cleanFormValues(SSTab1.SelectedTab.Name)
                gotonew()
            End If
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
        'cleanValues()
    End Sub

    Private Sub PoQotaFunction()
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = ""
        Try
            statusquote = "D-" & Status2
            Dim mpnopo As String = String.Empty
            Dim spacepoqota As String = String.Empty
            Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
            Dim dsPoQota = gnr.GetPOQotaData(txtvendorno.Text, txtpartno.Text)

            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    mpnopo = Trim(UCase(txtmfrno.Text))
                    Dim maxValue = 0
                    Dim dsUpdatedData As Integer

                    Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                        DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                    If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                        dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, txtunitcost.Text, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                            txtvendorno.Text, txtpartno.Text)
                        If dsUpdatedData <> 0 Then
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
                                txtvendorno.Text = "0" 'ask for vendor??
                            ElseIf item = "Unit Cost New" Then
                                txtunitcostnew.Text = "0"
                            ElseIf item = "Min Quantity" Then
                                txtminqty.Text = "0"
                            End If
                        Next
                        dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, txtunitcost.Text, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                            txtvendorno.Text, txtpartno.Text)

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
                    maxValue = 1 'preguntar duda
                End If
                spacepoqota = "                               DEV"
                mpnopo = Trim(UCase(txtmfrno.Text))
                Dim ResultQuery As String = String.Empty

                Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                        DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                    ResultQuery = gnr.InsertNewPOQota(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                       DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
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
                            txtvendorno.Text = "0"
                        ElseIf item = "Unit Cost New" Then
                            txtunitcostnew.Text = "0"
                        ElseIf item = "Min Quantity" Then
                            txtminqty.Text = "0"
                        End If
                    Next

                    ResultQuery = gnr.InsertNewPOQota(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                       DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                    If ResultQuery <> 0 Then
                        'show message error
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub InsertProductDetails(projectNo As String, partstoshow As String)
        Dim dtTime As DateTimePicker = New DateTimePicker()
        Dim dtTime1 As DateTimePicker = New DateTimePicker()
        Dim dtTime2 As DateTimePicker = New DateTimePicker()
        Dim dtTime3 As DateTimePicker = New DateTimePicker()
        Dim dtTime4 As DateTimePicker = New DateTimePicker()
        Dim dtTime5 As DateTimePicker = New DateTimePicker()
        Dim QueryDetailResult As Integer = -1
        Dim exMessage As String = " "
        Try
            dtTime5.Value = New DateTime(1900, 1, 1)
            dtTime5.CustomFormat = "yyyy/MM/dd/"

            Dim strCheck = gnr.checkFields(projectNo, txtpartno.Text, DTPicker2, "LREDONDO", dtTime, "LREDONDO", dtTime1, txtctpno.Text, txtqty.Text,
                                                                txtmfr.Text, txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                                                cmbuser.SelectedValue, chknew, dtTime3, txtsample.Text, txttcost.Text, txtvendorno.Text, partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, dtTime4,
                                                                dtTime5.Value.ToShortDateString(), txtsampleqty.Text)
            If String.IsNullOrEmpty(strCheck) Then
                QueryDetailResult = gnr.InsertProductDetail(projectNo, txtpartno.Text, DTPicker2, "LREDONDO", dtTime, "LREDONDO", dtTime1, txtctpno.Text, txtqty.Text,
                                    txtmfr.Text, txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                    cmbuser.SelectedValue, chknew, dtTime3, txtsample.Text, txttcost.Text, txtvendorno.Text, partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, dtTime4,
                                    dtTime5, CInt(txtsampleqty.Text))
                If QueryDetailResult <> 0 Then
                    'show message error
                End If
            Else
                Dim arrayCheck As New List(Of String)
                arrayCheck = strCheck.Split(",").ToList()
                For Each item As String In arrayCheck
                    If item = "Project Number" Then
                        'show error message must have data
                        Exit For
                    ElseIf item = "Quantity" Then
                        txtqty.Text = "0"
                    ElseIf item = "Unit Cost" Then
                        txtunitcost.Text = "0"
                    ElseIf item = "Unit Cost New" Then
                        txtunitcostnew.Text = "0"
                    ElseIf item = "Sample Cost" Then
                        txtsample.Text = "0"
                    ElseIf item = "Misc. Cost" Then
                        txttcost.Text = "0"
                    ElseIf item = "Vendor Number" Then
                        Exit For
                        'txtvendorno.Text = "0"  must have data
                    ElseIf item = "Tooling Cost" Then
                        txttoocost.Text = "0"
                    ElseIf item = "Sample Quantity" Then
                        txtsampleqty.Text = "0"
                    End If
                Next

                If txtvendorno.Text <> "" And projectNo <> 0 Then
                    QueryDetailResult = gnr.InsertProductDetail(projectNo, txtpartno.Text, DTPicker2, "LREDONDO", dtTime, "LREDONDO", dtTime1, txtctpno.Text, CInt(txtqty.Text),
                                    txtmfr.Text, txtmfrno.Text, CInt(txtunitcost.Text), CInt(txtunitcostnew.Text), txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                    cmbuser.SelectedValue, chknew, dtTime3, CInt(txtsample.Text), CInt(txttcost.Text), CInt(txtvendorno.Text), partstoshow, cmbminorcode.SelectedValue, CInt(txttoocost.Text), dtTime4,
                                    dtTime5, CInt(txtsampleqty.Text))
                Else
                    'message answering for a vendor
                    QueryDetailResult = -1
                End If

                If QueryDetailResult < 0 Then
                    'error message
                End If
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub ProdDetailAndAllCommentHelper(strUser As String, flag As Integer)
        Dim exMessage As String = " "
        Try
            If flag = 0 Then
                Dim queryProdDetail = gnr.UpdateProductDetail(txtCode.Text, txtpartno.Text)
                If queryProdDetail <> 0 Then
                    'error message
                End If
            End If

            Dim codComment = gnr.getmax("PRDCMH", "PRDCCO")
            Dim queryProdComments = gnr.InsertProductComment(txtCode.Text, txtpartno.Text, codComment, userid)
            If queryProdComments <> 0 Then
                'ERROR MESSAGE  
            End If
            Dim codDetComment = 1
            'Dim messcomm = "Person in charge changed assigned " & Trim(cmbuser.SelectedValue)
            Dim messcomm = strUser
            Dim queryProdCommentsDet = gnr.InsertProductCommentDetail(txtCode.Text, txtpartno.Text, codComment, codDetComment, messcomm)
            If queryProdCommentsDet <> 0 Then
                'error message
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
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
            Dim flagustatus As Integer
            Dim partstoshow As String = displayPart()
            Dim QueryDetailResult As Integer = -1
            Dim statusquote As String

            If flagdeve = 1 Then 'new
                Dim ProjectNo = gnr.getmax("PRDVLH", "PRHCOD") + 1
                Dim queryResult = gnr.InsertNewProject(ProjectNo, "LREDONDO", DTPicker1, txtainfo.Text, txtname.Text, cmbprstatus, cmbuser1)
                If queryResult < 0 Then
                    'error message
                Else
                    txtCode.Text = ProjectNo
                    flagdeve = 0
                    'strwhere = CustomStrWhereResult()
                    'fillcell1LastOne(strwhere)

                    If flagnewpart = 1 Then
                        If Trim(txtpartno.Text) <> "" Then '?????
                            Dim Status2 As String = ""
                            'Status2 = If(gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()) <> "", gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()), "")
                            Status2 = gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString())
                            Dim dsProjectNoResult As DataSet = gnr.GetCodeAndNameByPartNo("1527554") 'burned
                            Dim strProjectNo = If(String.IsNullOrEmpty(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()), 0, CInt(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()))
                            Dim strProjectName = Trim(dsProjectNoResult.Tables(0).Rows(0).ItemArray(1).ToString())

                            'test purpose
                            'strProjectNo = ProjectNo
                            'test purpose

                            If dsProjectNoResult.Tables(0).Rows.Count > 0 Then
                                If (ProjectNo = strProjectNo) Then
                                    Dim resultAlert As DialogResult = MessageBox.Show("This part no. already exists in this project. :" & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.OK)
                                Else
                                    Dim result As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & ProjectNo & " - " & strProjectName & ". Do you want to create it?.", "CTP System", MessageBoxButtons.YesNo)
                                    If result = DialogResult.No Then
                                        MessageBox.Show("No pressed")
                                    ElseIf result = DialogResult.Yes Then
                                        InsertProductDetails(ProjectNo, partstoshow)

                                        If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                            'send email
                                        End If

                                        'burned data test
                                        'txtvendorno.Text = "261747" 'has results
                                        'txtpartno.Text = "CABLE14B"

                                        'txtvendorno.Text = "261138"
                                        'txtpartno.Text = "99983"

                                        'end burned data test

                                        PoQotaFunction()

                                        If cmbuser.SelectedValue <> "N/A " Then
                                            ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                                        End If

                                        Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                        If resultMsgUser = DialogResult.Yes Then
                                            'save files
                                        End If
                                    End If
                                End If
                            Else
                                InsertProductDetails(ProjectNo, partstoshow)
                                If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                    'send email
                                End If

                                'burned data test
                                'txtvendorno.Text = "261747" 'has results
                                'txtpartno.Text = "CABLE14B"

                                'txtvendorno.Text = "261138"
                                'txtpartno.Text = "99983"

                                'end burned data test

                                PoQotaFunction()

                                If cmbuser.SelectedValue <> "N/A" Then
                                    ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                                End If

                                Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                If resultMsgUser = DialogResult.Yes Then
                                    'save files
                                End If
                            End If
                        End If
                    End If
                    SSTab1.TabPages(1).Text = "Project No." & Trim(txtCode.Text)
                    'SSTab1.tex = "Project No." & Trim(txtCode.Text)
                    txtsearchcode.Text = Trim(txtCode.Text)
                    cmdsearchcode_Click()
                    Dim resultDone As DialogResult = MessageBox.Show("project created successfully", "CTP System", MessageBoxButtons.OK)
                    flagdeve = 0
                    flagnewpart = 0
                End If
            Else 'update
                Dim Status2 As String = ""
                If Not (String.IsNullOrEmpty(gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()))) Then
                    Status2 = gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString())
                End If

                If cmbprstatus.FindString("F") Then
                    Dim dsProdDet = gnr.GetProdDetByCodeAndExc(txtCode.Text)
                    If Not dsProdDet Is Nothing Then
                        If dsProdDet.Tables(0).Rows.Count <= 0 Then
                            Dim rsProdClosedParts = gnr.UpdateProdClosedParts(userid, DTPicker1.Value.ToString(), Trim(cmbuser1.SelectedValue.ToString()), Trim(txtainfo.Text), Trim(txtname.Text),
                                                                              Trim(cmbprstatus.SelectedValue.ToString()), Trim(txtCode.Text))
                            If Not rsProdClosedParts.Equals(0) Then
                                'error message
                            End If
                        Else

                            Dim resultOpenParts As DialogResult = MessageBox.Show("All Items must be closed if you want to finish the project.", "CTP System", MessageBoxButtons.OK)
                            Dim rsProdOpenParts = gnr.UpdateProdOpenParts(userid, DTPicker1.Value.ToString(), Trim(cmbuser1.SelectedValue.ToString()), Trim(txtainfo.Text), Trim(txtname.Text),
                                                                          Trim(txtCode.Text))
                            If Not rsProdOpenParts.Equals(0) Then
                                'error message
                            End If

                            'Dim resultError As DialogResult = MessageBox.Show("An error ocurred. Call to an administrator.", "CTP System", MessageBoxButtons.OK)

                        End If
                    End If
                Else
                    Dim rsProdClosedParts = gnr.UpdateProdClosedParts(userid, DTPicker1.Value.ToString(), Trim(cmbuser1.SelectedValue.ToString()), Trim(txtainfo.Text), Trim(txtname.Text),
                                                                          Trim(cmbprstatus.SelectedValue.ToString()), Trim(txtCode.Text))
                    If Not rsProdClosedParts.Equals(0) Then
                        'error message
                    End If
                End If
                flagdeve = 0
                If flagnewpart = 1 Then
                    If Trim(txtpartno.Text) <> "" And Trim(txtvendorno.Text) <> "" Then
                        Dim dsProjectNoResult As DataSet = gnr.GetCodeAndNameByPartNo("1527554") 'burned
                        Dim strProjectNo = If(String.IsNullOrEmpty(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()), 0, CInt(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()))
                        Dim strProjectName = Trim(dsProjectNoResult.Tables(0).Rows(0).ItemArray(1).ToString())

                        Dim ProjectNo = txtCode.Text

                        If dsProjectNoResult.Tables(0).Rows.Count > 0 Then
                            If (ProjectNo = strProjectNo) Then
                                Dim resultAlert As DialogResult = MessageBox.Show("This part no. already exists in this project. :" & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.OK)
                            Else
                                Dim result As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & ProjectNo & " - " & strProjectName & ". Do you want to create it?.", "CTP System", MessageBoxButtons.YesNo)
                                If result = DialogResult.No Then
                                    Exit Sub
                                ElseIf result = DialogResult.Yes Then

                                    InsertProductDetails(ProjectNo, partstoshow)
                                    If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                        'send email
                                    End If

                                    'burned data test
                                    'txtvendorno.Text = "261747" 'has results
                                    'txtpartno.Text = "CABLE14B"

                                    'txtvendorno.Text = "261138"
                                    'txtpartno.Text = "99983"

                                    'end burned data test

                                    PoQotaFunction()


                                    If cmbuser.SelectedValue <> "N/A " Then
                                        ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                                    End If

                                    Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                    If resultMsgUser = DialogResult.Yes Then
                                        'save files
                                    End If
                                End If
                            End If
                        Else

                            InsertProductDetails(ProjectNo, partstoshow)

                            If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                'send email
                            End If

                            'burned data test
                            'txtvendorno.Text = "261747" 'has results
                            'txtpartno.Text = "CABLE14B"

                            'txtvendorno.Text = "261138"
                            'txtpartno.Text = "99983"

                            'end burned data test

                            PoQotaFunction()

                            If cmbuser.SelectedValue <> "N/A" Then
                                ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                            End If

                            Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                            If resultMsgUser = DialogResult.Yes Then
                                'save files
                            End If
                        End If
                    End If
                Else
                    If Trim(txtpartno.Text) <> "" And Trim(txtvendorno.Text) <> "" Then
                        Dim dsGetProdDesc = gnr.GetDataByCodeAndPartNoProdDesc(txtCode.Text, txtpartno.Text)
                        If dsGetProdDesc.Tables(0).Rows.Count > 0 Then
                            If Trim(cmbuser.SelectedValue) <> Trim(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDUSR").Ordinal)) Then
                                Dim messcomm = "Person in charge changed from " & Trim(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDUSR").Ordinal)) & " to " & Trim(cmbuser.SelectedValue)
                                ProdDetailAndAllCommentHelper(messcomm, 0)
                            End If
                            If cmbstatus.SelectedValue = "CS" Or cmbstatus.SelectedValue = "CN" Then
                                Dim rsAddPart As DialogResult = MessageBox.Show("Do you want to add this part to the Wish List?", "CTP System", MessageBoxButtons.YesNo)
                                If rsAddPart = DialogResult.Yes Then
                                    Dim dsGetWLByPartNo = gnr.GetWLDataByPartNo(txtpartno.Text)
                                    If dsGetWLByPartNo.Tables(0).Rows.Count > 0 Then
                                        Dim rsPartExist As DialogResult = MessageBox.Show("This part # is already included in the wish list.", "CTP System", MessageBoxButtons.OK)
                                    Else
                                        Dim maxItem = gnr.getmax("PRDWL", "PRWCOD")
                                        Dim rsInsWishListPart = gnr.InsertWishListProduct(maxItem, userid, txtpartno.Text)
                                        If rsInsWishListPart < 0 Then
                                            'error message
                                        End If
                                    End If
                                End If
                            End If
                            Dim status1 = ""
                            status1 = gnr.GetProjectStatusDescription(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDUSR").Ordinal))
                            Status2 = gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString())

                            If Trim(cmbstatus.SelectedValue) <> dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDSTS").Ordinal) Then
                                If (Trim(Status2) = "Closed without Negotiation") Or (Trim(Status2) = "Closed (Demand/cost/material)") Then
                                    Dim rsEnterComm As DialogResult = MessageBox.Show("Enter Comment.", "CTP System", MessageBoxButtons.OK)
                                    Dim seeaddprocomments = 5
                                    'frmproductsdevelopmentcomments.Show 1
                                End If
                                If (Trim(Status2) = "Approved") Or (Trim(Status2) = "Approved with advice") Then
                                    Dim rsAssignVendor As DialogResult = MessageBox.Show("Do you want to change the assigned vendor?", "CTP System", MessageBoxButtons.YesNo)
                                    If rsAssignVendor = DialogResult.Yes Then
                                        Dim dsGetPartVendor = gnr.GetDataByPartVendor(txtpartno.Text)
                                        If dsGetPartVendor.Tables(0).Rows.Count > 0 Then
                                            'change vendor method
                                        Else
                                            Dim dsGetPartNoVendor = gnr.GetDataByPartNoVendor(txtpartno.Text)
                                            If dsGetPartNoVendor.Tables(0).Rows.Count > 0 Then
                                                Dim rsInsertNewInv = gnr.InsertNewInv("", txtpartno.Text, dsGetPartNoVendor.Tables(0).Rows(0).ItemArray(dsGetPartNoVendor.Tables(0).Columns("impc1").Ordinal),
                                                                                      dsGetPartNoVendor.Tables(0).Rows(0).ItemArray(dsGetPartNoVendor.Tables(0).Columns("impc2").Ordinal), "", txtunitcostnew.Text, "", "", txtvendorno.Text)
                                                If rsInsertNewInv <> 0 Then
                                                    'error message
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                If (Trim(cmbstatus.SelectedValue) = "R") And (dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDSTS").Ordinal) = "RP") Then
                                    Dim flagchangevendor = 1
                                    'frmchangevendor.Show 1  check in vb net
                                End If
                                If Trim(Status2) = "Closed Successfully" Then
                                    'send email
                                End If
                                If (Trim(Status2) = "Technical Documentation") Or (Trim(Status2) = "Analysis of Samples") Or (Trim(Status2) = "Pending from Supplier") Then
                                    'send email
                                End If
                                If (Trim(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDSTS").Ordinal) = "AS") And (Trim(cmbstatus.SelectedValue) <> "AS")) Then
                                    If (Trim(cmbstatus.SelectedValue) = "R") Or Trim(cmbstatus.SelectedValue) = "A" Or Trim(cmbstatus.SelectedValue) = "AA" Then
                                        flagustatus = 1
                                    Else
                                        flagustatus = 0
                                        Dim rsStatusGet As DialogResult = MessageBox.Show("Status can not be changed. You must change this item to Approved with advice, Approved or Rejected.", "CTP System", MessageBoxButtons.OK)
                                    End If
                                Else
                                    flagustatus = 1
                                End If
                                If flagustatus = 1 Then
                                    Dim messcomm = "Status changed from " & status1 & " to " & Status2
                                    ProdDetailAndAllCommentHelper(messcomm, flagustatus)

                                    'burned data test
                                    'txtvendorno.Text = "261747" 'has results
                                    'txtpartno.Text = "CABLE14B"

                                    'txtvendorno.Text = "261138"
                                    'txtpartno.Text = "99983"

                                    'end burned data test

                                    PoQotaFunction()

                                End If
                            End If
                        End If

                        If flagustatus = 1 Then
                            Dim rsUpdProdDet = gnr.UpdateProductDetail1(partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, DateTime.Now.ToString(), "", txtvendorno.Text, chknew.Checked,
                                                                        DTPicker4.Value.ToString(), txtsample.Text, txttcost.Text, cmbuser.SelectedValue, DTPicker2.Value.ToString(), userid,
                                                                        txtctpno.Text, txtsampleqty.Text, txtqty.Text, txtmfr.Text, txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text,
                                                                        DTPicker3.Value.ToString(), cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text, txtCode.Text, txtpartno.Text)
                            If rsUpdProdDet <> 0 Then
                                'error message
                            End If
                        Else
                            Dim rsUpdProdDet = gnr.UpdateProductDetail2(partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, DateTime.Now.ToString(), txtvendorno.Text, chknew.Checked,
                                                                       DTPicker4.Value.ToString(), txtsample.Text, txttcost.Text, cmbuser.SelectedValue, DTPicker2.Value.ToString(), userid,
                                                                       txtctpno.Text, txtsampleqty.Text, txtqty.Text, txtmfr.Text, txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text,
                                                                       DTPicker3.Value.ToString(), cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text, txtpartno.Text)
                            If rsUpdProdDet <> 0 Then
                                'error message
                            End If
                        End If

                        Dim mpnopo As String = String.Empty
                        Dim spacepoqota As String = String.Empty
                        Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                        mpnopo = Trim(UCase(txtmfrno.Text))
                        Dim maxValue = 0
                        Dim dsUpdatedData As Integer

                        Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                            DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                        If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                            dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, txtunitcost.Text, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                txtvendorno.Text, txtpartno.Text)
                            If dsUpdatedData <> 0 Then
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
                                    txtvendorno.Text = "0" 'ask for vendor??
                                ElseIf item = "Unit Cost New" Then
                                    txtunitcostnew.Text = "0"
                                ElseIf item = "Min Quantity" Then
                                    txtminqty.Text = "0"
                                End If
                            Next
                            dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, txtunitcost.Text, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                txtvendorno.Text, txtpartno.Text)

                            If dsUpdatedData <> 0 Then
                                'show message error
                            End If
                        End If
                    End If
                End If
                txtsearchcode.Text = Trim(txtCode.Text)
                'Call cmdsearchcode_Click

                Dim dsGetProdDetByCodeAndExc = gnr.GetProdDetByCodeAndExc(txtCode.Text)
                If Not dsGetProdDetByCodeAndExc Is Nothing Then
                    If dsGetProdDetByCodeAndExc.Tables(0).Rows.Count = 0 Then
                        Dim dspMsg As DialogResult = MessageBox.Show("All parts for this project are closed. Do you want to finish the project?", "CTP System", MessageBoxButtons.YesNo)
                        If dspMsg = DialogResult.Yes Then
                            Dim rsUpdProdDevHeader = gnr.UpdateProductDevHeader(txtCode.Text)
                            If rsUpdProdDevHeader <> 0 Then
                                'error message
                            End If
                        End If
                    End If
                End If

                Dim dspUpdMess As DialogResult = MessageBox.Show("Record updated", "CTP System", MessageBoxButtons.OK)
            End If

            If SSTab1.SelectedIndex = 2 Then
                If Trim(txtpartno.Text) <> "" Then

                End If
            End If
            fillcell1LastOne("")
            fillcell2(txtCode.Text)

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdSave2_Click(sender As Object, e As EventArgs) Handles cmdSave2.Click

        Dim validationResult = mandatoryFields("save", SSTab1.SelectedTab.Name)
        If validationResult.Equals(0) Then
            'Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
            'If result = DialogResult.No Then
            'MessageBox.Show("No pressed")
            'ElseIf result = DialogResult.Yes Then
            'MessageBox.Show("Yes pressed")
            save()
            If SSTab1.SelectedIndex = 1 Then
                Dim result1 As DialogResult = MessageBox.Show("Please proceed to the project tab to add parts?", "CTP System", MessageBoxButtons.OK)
                If result1 = DialogResult.OK Then
                    SSTab1.SelectedTab = TabPage3
                End If
            End If
            'End If
        Else
            Dim resultSave As DialogResult = MessageBox.Show("Error in Data Validation. Mandatory fields must be filled!!", "CTP System", MessageBoxButtons.OK)
        End If

    End Sub

    Private Sub cmdSave3_Click(sender As Object, e As EventArgs) Handles cmdSave3.Click

        Dim DtUseTime = New DateTimePicker()
        DtUseTime.Value = DateTime.Now
        Dim rsValidation = gnr.checkFields(txtCode.Text, txtpartno.Text, DTPicker2, "LREDONDO", DtUseTime, "LREDONDO", DtUseTime, txtctpno.Text, txtqty.Text,
                                                                txtmfr.Text, txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, DtUseTime, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                                                cmbuser.SelectedValue, chknew, DtUseTime, txtsample.Text, txttcost.Text, txtvendorno.Text, 0, cmbminorcode.SelectedValue, txttoocost.Text, DtUseTime,
                                                                DateTime.Now.ToShortDateString(), txtsampleqty.Text)

        Dim validationResult = mandatoryFields("save", SSTab1.SelectedTab.Name)
        If validationResult.Equals(0) Then
            Dim result As DialogResult = MessageBox.Show("Ff click yes the part will be added to the project. Do you want to proceed?", "CTP System", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                'MessageBox.Show("No pressed")
            ElseIf result = DialogResult.Yes Then
                'MessageBox.Show("Yes pressed")
                save()
            End If
        Else
            Dim resultSave As DialogResult = MessageBox.Show("Error in Data Validation. Mandatory fields must be filled!!", "CTP System", MessageBoxButtons.OK)
        End If

    End Sub

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click

        Dim resultNew As DialogResult = MessageBox.Show("Operation not allowed from this tab screen!!! ", "CTP System", MessageBoxButtons.OK)

    End Sub

    Private Sub cmdpartno_Click(sender As Object, e As EventArgs) Handles cmdpartno.Click
        Dim exMessage As String = " "
        Try
            If (flagdeve = 0 And flagnewpart = 1) Or (flagdeve) = 1 Then
                If Trim(txtvendorno.Text) <> "" Then
                    Dim partno = InputBox("Enter Part No. :", "Select Part No.")
                    If Trim(partno) <> "" Then
                        Dim dsGetDataFromProdHeadAndDet = gnr.GetDataFromProdHeaderAndDetail(partno)
                        Dim codeTemp As String
                        Dim nameTemp As String
                        If Not dsGetDataFromProdHeadAndDet Is Nothing Then
                            If dsGetDataFromProdHeadAndDet.Tables(0).Rows.Count <> 0 Then
                                codeTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                nameTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                If txtCode.Text = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal).ToString() Then
                                    Dim result1 As DialogResult = MessageBox.Show("This part no. already exists in this project. : " & codeTemp & " - " & Trim(nameTemp), "CTP System", MessageBoxButtons.OK)
                                Else
                                    codeTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                    nameTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                    Dim result2 As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & codeTemp & " - " & Trim(nameTemp), "CTP System", MessageBoxButtons.OK)
                                End If
                            End If
                        End If

                        Dim dsGetDataFromDualInv = gnr.GetDataFromDualInventory(partno)
                        If Not dsGetDataFromDualInv Is Nothing Then
                            txtpartno.Text = partno
                            txtpartdescription.Text = Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMDSC").Ordinal).ToString())

                            If cmbminorcode.FindStringExact(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC2").Ordinal).ToString())) Then
                                cmbminorcode.SelectedIndex = cmbminorcode.FindString(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC2").Ordinal).ToString()))
                            End If

                            If Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString()) <> "" Then
                                Dim dsGetVendorQuey = gnr.GetVendorQuey(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString())
                                If Not dsGetVendorQuey Is Nothing Then
                                    If dsGetVendorQuey.Tables(0).Rows.Count > 0 Then
                                        txtvendornoa.Text = dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString()
                                        txtvendornamea.Text = Trim(dsGetVendorQuey.Tables(0).Rows(0).ItemArray(dsGetVendorQuey.Tables(0).Columns("VMNAME").Ordinal).ToString())
                                    Else
                                        txtvendornoa.Text = ""
                                        txtvendornamea.Text = ""
                                    End If
                                End If
                            Else
                                txtvendornoa.Text = ""
                                txtvendornamea.Text = ""
                            End If

                            Dim dsGetCTPPartRef = gnr.GetCTPPartRef(partno)
                            If Not String.IsNullOrEmpty(dsGetCTPPartRef) Then
                                txtctpno.Text = dsGetCTPPartRef
                                txtmfrno.Text = dsGetCTPPartRef
                            Else
                                txtctpno.Text = ""
                                txtmfrno.Text = ""
                            End If

                            If Trim(txtvendornoa.Text) <> "" Then
                                Dim dsGetAssignedVendor = gnr.GetAssignedVendor(txtvendornoa.Text, partno)
                                If Not dsGetAssignedVendor Is Nothing Then
                                    If dsGetAssignedVendor.Tables(0).Rows.Count > 0 Then
                                        txtunitcost.Text = dsGetAssignedVendor.Tables(0).Rows(0).ItemArray(dsGetAssignedVendor.Tables(0).Columns("PQPRC").Ordinal).ToString()
                                        txtminqty.Text = dsGetAssignedVendor.Tables(0).Rows(0).ItemArray(dsGetAssignedVendor.Tables(0).Columns("PQMIN").Ordinal).ToString()
                                    Else
                                        txtunitcost.Text = 0
                                        txtminqty.Text = 0
                                    End If
                                End If
                            Else
                                txtunitcost.Text = 0
                                txtminqty.Text = 0
                            End If

                            'Call searchpart
                            'txtctpno.SetFocus
                        Else
                            Dim dsGetDataFromDualInventory1 = gnr.GetDataByPartNoVendor(partno)
                            If Not dsGetDataFromDualInventory1 Is Nothing Then
                                If dsGetDataFromDualInventory1.Tables(0).Rows.Count > 0 Then
                                    txtpartno.Text = partno
                                    txtpartdescription.Text = Trim(dsGetDataFromDualInventory1.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInventory1.Tables(0).Columns("IMDSC").Ordinal).ToString())

                                    If cmbminorcode.FindStringExact(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC2").Ordinal).ToString())) Then
                                        cmbminorcode.SelectedIndex = cmbminorcode.FindString(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC2").Ordinal).ToString()))
                                    End If

                                    txtvendornoa.Text = ""
                                    txtvendornamea.Text = ""

                                    Dim dsGetCTPPartRef = gnr.GetCTPPartRef(partno)
                                    If Not String.IsNullOrEmpty(dsGetCTPPartRef) Then
                                        txtctpno.Text = dsGetCTPPartRef
                                        txtmfrno.Text = dsGetCTPPartRef
                                    Else
                                        txtctpno.Text = ""
                                        txtmfrno.Text = ""
                                    End If
                                    'Call searchpart
                                    'txtctpno.SetFocus
                                Else
                                    Dim result3 As DialogResult = MessageBox.Show("Part No. not found.", "CTP System", MessageBoxButtons.OK)
                                End If
                            End If
                        End If

                        'test purpose
                        Dim testPartNo = "5257106"
                        'Dim dsGetPartInWishList = gnr.GetPartInWishList(partno)
                        Dim dsGetPartInWishList = gnr.GetPartInWishList(testPartNo)
                        If Not dsGetPartInWishList Is Nothing Then
                            If dsGetPartInWishList.Tables(0).Rows.Count > 0 Then
                                Dim wlcode = dsGetPartInWishList.Tables(0).Rows(0).ItemArray(dsGetPartInWishList.Tables(0).Columns("WHLCODE").Ordinal).ToString()

                                'test purpose
                                Dim tetsVendorNo = "120138"
                                'Dim dsGetDataByVendorAndPartNoProdDesc = gnr.GetDataByVendorAndPartNoProdDesc(txtvendorno.Text, partno)
                                Dim dsGetDataByVendorAndPartNoProdDesc = gnr.GetDataByVendorAndPartNoProdDesc(tetsVendorNo, testPartNo)
                                If Not dsGetDataByVendorAndPartNoProdDesc Is Nothing Then
                                    If dsGetDataByVendorAndPartNoProdDesc.Tables(0).Rows.Count > 0 Then
                                        'Dim dsGetDataByCodAndPartProdAndComm =
                                        'gnr.GetDataByCodAndPartProdAndComm(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Rows(0).ItemArray(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Columns("PRHCOD").Ordinal).ToString(), partno)
                                        'test purposes
                                        Dim dsGetDataByCodAndPartProdAndComm =
                                            gnr.GetDataByCodAndPartProdAndComm(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Rows(0).ItemArray(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Columns("PRHCOD").Ordinal).ToString(), testPartNo)
                                        If Not dsGetDataByCodAndPartProdAndComm Is Nothing Then
                                            If dsGetDataByCodAndPartProdAndComm.Tables(0).Rows.Count > 0 Then
                                                Dim result4 As DialogResult = MessageBox.Show("This part# : " & Trim(UCase(partno)) & " has been quoted with this vendor# : " & Trim(txtvendorno.Text) & " before. Do you want to continue?", "CTP System", MessageBoxButtons.YesNo)
                                                If result4 = DialogResult.No Then
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                                Dim dsGetDataFromProdHeaderAndDetail = gnr.GetDataFromProdHeaderAndDetail(partno)
                                Dim dtpDate = New DateTimePicker()
                                Dim dtpDate1 = New DateTimePicker()
                                Dim dt = DateTime.Now

                                Dim iDate As String = "1900-01-01"
                                Dim oDate As DateTime = DateTime.Parse(iDate)
                                dtpDate.Value = dt
                                dtpDate1.Value = oDate

                                If Not dsGetDataFromProdHeaderAndDetail Is Nothing Then
                                    If dsGetDataFromProdHeaderAndDetail.Tables(0).Rows.Count > 0 Then
                                        If Trim(txtCode.Text) = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRHCOD").Ordinal).ToString() Then
                                            Dim code = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                            Dim name = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                            Dim result5 As DialogResult = MessageBox.Show("This part no. already exists in this project. : " & code & "-" & name & " ", "CTP System", MessageBoxButtons.OK)
                                        Else
                                            Dim code = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                            Dim name = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                            Dim result6 As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & code & "-" & name & ". Do you want to create it?.", "CTP System", MessageBoxButtons.YesNo)
                                            If result6 = DialogResult.Yes Then
                                                Dim rsInsertProductDetailv2 = gnr.InsertProductDetailv2(txtCode.Text, txtpartno.Text, dtpDate, userid, dtpDate, userid, dtpDate, txtctpno.Text,
                                                                                                        0, "", "", txtunitcost.Text, 0, "", dtpDate1, "E", "", "", userid, chknew, dtpDate1, 0, 0, txtvendorno.Text,
                                                                                                        "", cmbminorcode.SelectedValue, 0, dtpDate1, dtpDate1, DTPicker2, 1)
                                                If rsInsertProductDetailv2 <> 0 Then
                                                    'error message
                                                End If

                                                Dim statusquote = "D-Entered"
                                                Dim mpnopo1 As String
                                                Dim spacepoqota1 As String = String.Empty
                                                Dim strQueryAdd1 As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                                                Dim dsPoQota = gnr.GetPOQotaData(txtvendorno.Text, txtpartno.Text) 'aqui llegue full

                                                'separate here in other methods--------------------------------

                                                'burned data test
                                                'txtvendorno.Text = "261747" 'has results
                                                'txtpartno.Text = "CABLE14B"

                                                'txtvendorno.Text = "261138"
                                                'txtpartno.Text = "99983"

                                                'end burned data test

                                                If dsPoQota IsNot Nothing Then
                                                    If dsPoQota.Tables(0).Rows.Count > 0 Then
                                                        mpnopo1 = Trim(UCase(txtmfrno.Text))
                                                        Dim maxValue1 = 0
                                                        Dim dsUpdatedData1 As Integer

                                                        Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                            DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                                        If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                            dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                                txtvendorno.Text, txtpartno.Text)
                                                            If dsUpdatedData1 <> 0 Then
                                                                'show message error
                                                            End If
                                                        Else
                                                            Dim arrayCheck As New List(Of String)
                                                            arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                            For Each item As String In arrayCheck
                                                                If item = "Sequencial" Then
                                                                    'show error message
                                                                    Exit For
                                                                ElseIf item = "Vendor Number" Then
                                                                    txtvendorno.Text = "0" 'ask for vendor??
                                                                ElseIf item = "Unit Cost New" Then
                                                                    txtunitcostnew.Text = "0"
                                                                End If
                                                            Next
                                                            dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                                txtvendorno.Text, txtpartno.Text)

                                                            If dsUpdatedData1 <> 0 Then
                                                                'show message error
                                                            End If
                                                        End If
                                                    Else
                                                        'warning message
                                                    End If
                                                Else
                                                    Dim maxValue1 = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd1)
                                                    If Not String.IsNullOrEmpty(maxValue1) Then
                                                        maxValue1 += 1
                                                    Else
                                                        maxValue1 = 1 'preguntar duda
                                                    End If
                                                    spacepoqota1 = "                               DEV"
                                                    mpnopo1 = Trim(UCase(txtmfrno.Text))
                                                    Dim ResultQuery As String = String.Empty

                                                    Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                            DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                                    If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                        ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                           DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                        If ResultQuery <> 0 Then
                                                            'show message error
                                                        End If
                                                    Else
                                                        Dim arrayCheck As New List(Of String)
                                                        arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                        For Each item As String In arrayCheck
                                                            If item = "Sequencial" Then
                                                                'show error message
                                                                Exit For
                                                            ElseIf item = "Vendor Number" Then
                                                                txtvendorno.Text = "0"
                                                            End If
                                                        Next

                                                        ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                           DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                        If ResultQuery <> 0 Then
                                                            'show message error
                                                        End If
                                                    End If

                                                    Dim rsDeletion = gnr.DeleteDataByWSCod(txtCode.Text)
                                                    If rsDeletion = 1 Then
                                                        'deletion ok
                                                    End If

                                                    'Call gotonew()
                                                    'Call cmdall_Click()
                                                    Dim result7 As DialogResult = MessageBox.Show("Part # added to Project : " & Trim(txtCode.Text), "CTP System", MessageBoxButtons.OK)

                                                End If

                                                'aqui continuo
                                            End If
                                        End If
                                    Else
                                        Dim rsInsertProductDetailv2 = gnr.InsertProductDetailv2(txtCode.Text, txtpartno.Text, dtpDate, userid, dtpDate, userid, dtpDate, txtctpno.Text,
                                                                                                        0, "", "", txtunitcost.Text, 0, "", dtpDate1, "E", "", "", userid, chknew, dtpDate1, 0, 0, txtvendorno.Text,
                                                                                                        "", cmbminorcode.SelectedValue, 0, dtpDate1, dtpDate1, DTPicker2, 1)
                                        If rsInsertProductDetailv2 <> 0 Then
                                            'error message
                                        End If

                                        Dim statusquote = "D-Entered"
                                        Dim mpnopo1 As String
                                        Dim spacepoqota1 As String = String.Empty
                                        Dim strQueryAdd1 As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                                        Dim dsPoQota = gnr.GetPOQotaData(txtvendorno.Text, txtpartno.Text)

                                        'separate here in other methods--------------------------------

                                        'burned data test
                                        'txtvendorno.Text = "261747" 'has results
                                        'txtpartno.Text = "CABLE14B"

                                        'txtvendorno.Text = "261138"
                                        'txtpartno.Text = "99983"

                                        'end burned data test

                                        If dsPoQota IsNot Nothing Then
                                            If dsPoQota.Tables(0).Rows.Count > 0 Then
                                                mpnopo1 = Trim(UCase(txtmfrno.Text))
                                                Dim maxValue1 = 0
                                                Dim dsUpdatedData1 As Integer

                                                Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                    DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                                If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                    dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                        txtvendorno.Text, txtpartno.Text)
                                                    If dsUpdatedData1 <> 0 Then
                                                        'show message error
                                                    End If
                                                Else
                                                    Dim arrayCheck As New List(Of String)
                                                    arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                    For Each item As String In arrayCheck
                                                        If item = "Sequencial" Then
                                                            'show error message
                                                            Exit For
                                                        ElseIf item = "Vendor Number" Then
                                                            txtvendorno.Text = "0" 'ask for vendor??
                                                        End If
                                                    Next
                                                    dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                        txtvendorno.Text, txtpartno.Text)

                                                    If dsUpdatedData1 <> 0 Then
                                                        'show message error
                                                    End If
                                                End If
                                            Else
                                                'warning message
                                            End If
                                        Else
                                            Dim maxValue1 = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd1)
                                            If Not String.IsNullOrEmpty(maxValue1) Then
                                                maxValue1 += 1
                                            Else
                                                maxValue1 = 1 'preguntar duda
                                            End If
                                            spacepoqota1 = "                               DEV"
                                            mpnopo1 = Trim(UCase(txtmfrno.Text))
                                            Dim ResultQuery As String = String.Empty

                                            Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                    DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                            If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                   DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                If ResultQuery <> 0 Then
                                                    'show message error
                                                End If
                                            Else
                                                Dim arrayCheck As New List(Of String)
                                                arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                For Each item As String In arrayCheck
                                                    If item = "Sequencial" Then
                                                        'show error message
                                                        Exit For
                                                    ElseIf item = "Vendor Number" Then
                                                        txtvendorno.Text = "0"
                                                    End If
                                                Next

                                                ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                   DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                If ResultQuery <> 0 Then
                                                    'show message error
                                                End If
                                            End If
                                            Dim rsDeletion = gnr.DeleteDataByWSCod(txtCode.Text)
                                            If rsDeletion = 1 Then
                                                'deletion ok
                                            End If
                                            'Call gotonew()
                                            'Call cmdall_Click()
                                            Dim result7 As DialogResult = MessageBox.Show("Part # added to Project : " & Trim(txtCode.Text), "CTP System", MessageBoxButtons.OK)
                                        End If

                                    End If
                                End If
                            End If
                        End If
                        txtsample.Text = "0"
                        txttcost.Text = "0"
                        txttoocost.Text = "0"
                        txtmfr.Text = " "
                        txtsampleqty.Text = "0"
                        txtBenefits.Text = " "
                        txtainfo.Text = " "
                        txtqty.Text = "0"

                    End If
                Else
                    Dim result As DialogResult = MessageBox.Show("Enter Vendor.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                Dim result1 As DialogResult = MessageBox.Show("Part No. cannot be changed when is already created.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdvendor_Click(sender As Object, e As EventArgs) Handles cmdvendor.Click
        Dim exMessage As String = " "
        Dim partstoshow As String = displayPart()
        Try
            Dim oldvendorno = Trim(txtvendorno.Text)
            Dim vendorno = InputBox("Enter Vendor No. :", "Change Vendor")
            '889912

            If Not IsNumeric(vendorno) Then
                Dim result As DialogResult = MessageBox.Show("Enter just numbers.", "CTP System", MessageBoxButtons.OK)
            Else
                Dim dsGetVendorByVendorNo = gnr.GetVendorByVendorNo(vendorno)
                If Not dsGetVendorByVendorNo Is Nothing Then
                    If (dsGetVendorByVendorNo.Tables(0).Rows.Count > 0) Then
                        txtvendorno.Text = vendorno
                        txtvendorname.Text = dsGetVendorByVendorNo.Tables(0).Rows(0).ItemArray(dsGetVendorByVendorNo.Tables(0).Columns("VMNAME").Ordinal).ToString()
                        partstoshow = ""

                        optCTP.Checked = True
                        optVENDOR.Checked = False
                        optboth.Checked = False
                        partstoshow = "1"
                        Dim strQueryAdd As String = "WHERE PQVND = " & Trim(vendorno) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                        If flagnewpart = 0 And Trim(txtpartno.Text) <> "" Then
                            Dim dsGetDataByVendorAndPartNo = gnr.GetDataByVendorAndPartNoDst(oldvendorno, txtpartno.Text)
                            If Not dsGetDataByVendorAndPartNo Is Nothing Then
                                If dsGetDataByVendorAndPartNo.Tables(0).Rows.Count > 0 Then
                                    Dim rsUpdatePoQotaByVendorAndPart = gnr.UpdatePoQotaByVendorAndPart(vendorno, oldvendorno, txtpartno.Text,
                                                                        dsGetDataByVendorAndPartNo.Tables(0).Rows(0).ItemArray(dsGetDataByVendorAndPartNo.Tables(0).Columns("PQSEQ").Ordinal).ToString())
                                    If rsUpdatePoQotaByVendorAndPart <> 0 Then
                                        'error message
                                    End If
                                Else
                                    Dim maxValue = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)
                                    If Not String.IsNullOrEmpty(maxValue) Then
                                        maxValue += 1
                                    Else
                                        Dim spacepoqota = "                               DEV"
                                        Dim rsInsertNewPOQota = gnr.InsertNewPOQotaLess(txtpartno.Text, vendorno, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), "", DateTime.Now.Day.ToString(), "", spacepoqota, 0)
                                        If rsInsertNewPOQota <> 0 Then
                                            'error message
                                        End If
                                        maxValue = 1 'preguntar duda
                                    End If

                                End If
                                Dim rsUpdProdDetVend = gnr.UpdateProdDetailVendor(partstoshow, vendorno, txtCode.Text, txtpartno.Text)
                                If rsUpdProdDetVend <> 0 Then
                                    'erro message
                                End If
                                fillcell2(txtCode.Text)
                            End If
                            Dim result2 As DialogResult = MessageBox.Show("Vendor Changed.", "CTP System", MessageBoxButtons.OK)
                        End If

                    Else
                        Dim result3 As DialogResult = MessageBox.Show("Vendor not found.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    Dim result4 As DialogResult = MessageBox.Show("Vendor not found.", "CTP System", MessageBoxButtons.OK)
                End If
            End If

        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdgenerate_Click(sender As Object, e As EventArgs) Handles cmdgenerate.Click
        Dim exMessage As String = " "
        Try
            If Trim(txtpartno.Text) <> "" Then
                If Trim(txtctpno.Text) Then
                    Dim result As DialogResult = MessageBox.Show("CTP # has been already generated.", "CTP System", MessageBoxButtons.OK)
                Else
                    Dim PartNo = Trim(UCase(txtpartno.Text)).Substring(0, 19) & "                   "
                    'PartNo = Left(PartNo, 19)
                    Dim ctppartno = "                   "
                    Dim flagctp = "9"
                    Dim dsctpValue = gnr.CallForCtpNumber(PartNo, ctppartno, flagctp)
                    If Not dsctpValue Is Nothing Then
                        If dsctpValue.Tables(0).Rows.Count > 0 Then
                            txtctpno.Text = Trim(UCase(dsctpValue.Tables(0).Rows(0).ItemArray(0).ToString()))
                            txtmfrno.Text = Trim(UCase(dsctpValue.Tables(0).Rows(0).ItemArray(0).ToString()))
                        Else
                            txtctpno.Text = ""
                            txtmfrno.Text = ""
                        End If
                    End If
                End If
            Else
                Dim result1 As DialogResult = MessageBox.Show("Select Part No.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdsearch.Click
        cmdSearch_Click()
    End Sub

    Private Sub cmdSearch_Click()
        Dim exMessage As String = " "
        userid = "LREDONDO"
        Try
            If Trim(txtsearch.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(txtsearch.Text)), "'", "") & "%'"
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "')) AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(txtsearch.Text)), "'", "") & "%'"
                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(txtsearch.Text)), "'", "") & "%'"
                End If
                fillcell1(strwhere)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
        'Call gotoerror("frmproductsdevelopment", "cmdsearch_click", Err.Number, Err.Description, Err.Source)
    End Sub

    Private Sub cmdsearch1_Click(sender As Object, e As EventArgs) Handles cmdsearch1.Click
        cmdsearch1_Click()
    End Sub

    Private Sub cmdsearch1_Click()
        Dim exMessage As String = " "
        userid = "LREDONDO"
        Try
            If Trim(txtsearch1.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE PRDVLD.VMVNUM = " & Trim(txtsearch1.Text)
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND PRDVLD.VMVNUM = " & Trim(txtsearch1.Text)
                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND PRDVLD.VMVNUM = " & Trim(txtsearch1.Text)
                End If
                fillcelldetail(strwhere)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdsearchpart_Click(sender As Object, e As EventArgs) Handles cmdsearchpart.Click
        cmdSearchPart_Click()
    End Sub

    Private Sub cmdSearchPart_Click()
        Dim exMessage As String = " "
        userid = "LREDONDO"
        Try
            If Trim(txtsearchpart.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(txtsearchpart.Text)) & "' "
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(txtsearchpart.Text)) & "' "
                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(txtsearchpart.Text)) & "' "
                End If
                fillcelldetail(strwhere)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdsearchctp_Click(sender As Object, e As EventArgs) Handles cmdsearchctp.Click
        cmdsearchctp_Click()
    End Sub

    Private Sub cmdsearchctp_Click()
        Dim exMessage As String = " "
        userid = "LREDONDO"
        Try
            If Trim(txtsearchctp.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(txtsearchctp.Text)) & "' "
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(txtsearchctp.Text)) & "' "
                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(txtsearchctp.Text)) & "' "
                End If
                fillcelldetail(strwhere)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdsearchcode_Click(sender As Object, e As EventArgs) Handles cmdsearchcode.Click
        cmdsearchcode_Click()
    End Sub

    Private Sub cmdsearchcode_Click()
        Dim exMessage As String = " "
        userid = "LREDONDO"
        Try
            If Trim(txtsearchcode.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE PRHCOD = " & Trim(txtsearchcode.Text)
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "')) AND PRHCOD = " & Trim(txtsearchcode.Text)
                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND PRHCOD = " & Trim(txtsearchcode.Text)
                End If
                fillcell1(strwhere)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdsearchstatus_Click(sender As Object, e As EventArgs) Handles cmdsearchstatus.Click
        cmdsearchstatus_Click()
    End Sub

    Private Sub cmdsearchstatus_Click()
        Dim exMessage As String = " "
        userid = "LREDONDO"
        Try
            If Trim(cmbstatus1.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE TRIM(UCASE(PRDSTS)) = '" & Trim(cmbstatus1.SelectedValue) & "' "
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDSTS)) = '" & Trim(cmbstatus1.SelectedValue) & "' "
                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRDSTS)) = '" & Trim(Left(cmbstatus1.Text, 2)) & "' "
                End If
                fillcelldetail(strwhere)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try


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

    Private Function mandatoryFields(flag As String, tab As String) As Integer

        Dim methodResult As Integer = 0
        Dim myTableLayout As TableLayoutPanel

        If tab = "TabPage1" Then
            myTableLayout = Me.TableLayoutPanel1
        ElseIf tab = "TabPage2" Then
            myTableLayout = Me.TableLayoutPanel3
        Else
            myTableLayout = Me.TableLayoutPanel4
        End If

        If flag = "new" Then
            Dim TextboxQty As Integer
            Dim TextboxQtyEmpty As Integer
            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Then
                    TextboxQty += 1
                    If tt.Text = "" Then
                        TextboxQtyEmpty += 1
                        'MsgBox("Complete Entry!")
                        'Exit Sub
                        'Exit For
                    End If
                End If
            Next

            If TextboxQtyEmpty <> 0 Then
                If TextboxQty > TextboxQtyEmpty Then
                    methodResult = 1
                End If
            Else
                methodResult = 0
            End If

        Else

            If tab = "TabPage2" Then
                txtCode.Text = " "

                Dim empty = myTableLayout.Controls.OfType(Of Windows.Forms.TextBox)().Where(Function(txt) txt.Text.Length = 0)
                If empty.Any Then
                    methodResult = 1
                    'MessageBox.Show(String.Format("Please fill following textboxes: {0}", String.Join(",", empty.Select(Function(txt) txt.Name))))
                End If
            End If

            'Dim empties As Integer
            ''let optional empty values
            'For Each Val As Windows.Forms.TextBox In myTableLayout.Controls.OfType(Of Windows.Forms.TextBox)
            '    If String.IsNullOrEmpty(Val.Text) Then
            '        empties += 1
            '    End If
            'Next
        End If

        Return methodResult

    End Function

    Private Sub cleanFormValues(tab As String)
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel

            If tab = "TabPage1" Then
                myTableLayout = Me.TableLayoutPanel1
            ElseIf tab = "TabPage2" Then
                myTableLayout = Me.TableLayoutPanel3
            Else
                myTableLayout = Me.TableLayoutPanel4
            End If

            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Then
                    tt.Text = ""
                ElseIf TypeOf tt Is Windows.Forms.ComboBox Then
                    tt.selectedIndex = 0
                ElseIf TypeOf tt Is Windows.Forms.DateTimePicker Then
                    tt.Value = DateTime.Now
                End If
            Next
            'myTableLayout.Controls.OfType(Of Windows.Forms.TextBox)().Select(Function(ctx) ctx.Text = "")
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

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

    Private Sub txtpartno_TextChanged(sender As Object, e As EventArgs) Handles txtpartno.TextChanged
        If Not String.IsNullOrEmpty(txtpartno.Text) Then
            TabPage3.Name = "Part No. " & txtpartno.Text
        End If
    End Sub

    Private Sub txtsearchcode_TextChanged(sender As Object, e As EventArgs) Handles txtsearchcode.TextChanged

    End Sub











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