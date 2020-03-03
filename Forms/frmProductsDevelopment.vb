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


    Private Sub frmProductsDevelopment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SSTab1.ItemSize = (New Size(SSTab1.Width / SSTab1.TabCount, 0))
        SSTab1.Padding = New System.Drawing.Point(300, 10)
        SSTab1.Appearance = TabAppearance.FlatButtons
        'TabControl1.ItemSize = New Size(0, 1)
        SSTab1.SizeMode = TabSizeMode.Fixed

        Button1.FlatStyle = FlatStyle.Flat
        Button2.FlatStyle = FlatStyle.Flat
        Button3.FlatStyle = FlatStyle.Flat
        Button4.FlatStyle = FlatStyle.Flat
        Button5.FlatStyle = FlatStyle.Flat
        Button6.FlatStyle = FlatStyle.Flat
        cmdall.FlatStyle = FlatStyle.Flat

        DataGridView1.RowHeadersVisible = False
        dgvProjectDetails.RowHeadersVisible = False

        'Button12.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\doc.PNG")
        Button12.ImageAlign = ContentAlignment.MiddleRight
        Button12.TextAlign = ContentAlignment.MiddleLeft

        ' Button13.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\save.PNG")
        Button13.ImageAlign = ContentAlignment.MiddleRight
        Button13.TextAlign = ContentAlignment.MiddleLeft

        'Button14.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\exit.PNG")
        Button14.ImageAlign = ContentAlignment.MiddleRight
        Button14.TextAlign = ContentAlignment.MiddleLeft

        AddHandler DataGridView1.SelectionChanged, AddressOf dataGridView1_SelectionChanged
        'AddHandler dgvProjectDetails.DataBindingComplete, AddressOf dgvProjectDetails_DataBindingComplete

        'Datepickers customization

        DTPicker1.Format = DateTimePickerFormat.Custom
        DTPicker1.CustomFormat = "MM/dd/yyyy"


        'dropdownlist default fill section

        FillDDlUser() 'Fill user cmb
        FillDDlUser1()

        cmbprstatus.Items.Add("-- Select Status --")
        cmbprstatus.Items.Add("I - In Process")
        cmbprstatus.Items.Add("F - Finished")
        cmbprstatus.SelectedItem = "-- Select Status --"


    End Sub

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
            newRow("USUSER") = "NO NAME"
            newRow("USUSER") = "N/A -- NO NAME"
            dsUser.Tables(0).Rows.Add(newRow)

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
            newRow("USUSER") = "NO NAME"
            newRow("USUSER") = "N/A -- NO NAME"
            dsUser.Tables(0).Rows.Add(newRow)

            cmbuser.DataSource = dsUser.Tables(0)
            cmbuser.DisplayMember = "FullValue"
            cmbuser.ValueMember = "USUSER"


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


            'Dim newRow As DataRow = dsMinCodes.Tables(0).NewRow
            'newRow("USUSER") = "N/A"
            'newRow("USUSER") = "NO NAME"
            'newRow("USUSER") = "N/A -- NO NAME"
            'dsMinCodes.Tables(0).Rows.Add(newRow)

            cmbuser1.DataSource = dsMinCodes.Tables(0)
            cmbuser1.DisplayMember = "FullValue"
            cmbuser1.ValueMember = "CNTDE1"


        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = Len(Trim(TextBox1.Text))
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        TextBox3.SelectionStart = 0
        TextBox3.SelectionLength = Len(Trim(TextBox1.Text))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If Trim(TextBox1.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = "WHERE TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(TextBox1.Text)), "'", "") & "%'"
                Else
                    strwhere = "WHERE (PRPECH = '" & userid & "' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "')) AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(TextBox1.Text)), "'", "") & "%'"
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProjectDetails.CellContentClick

    End Sub

    Private Sub TableLayoutPanel4_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel4.Paint
        'TableLayoutPanel4.CellBorderStyle = TableLayoutPanelCellBorderStyle.Outset
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

    End Sub

    'here class traduction begans---------------------------------------------------------

    'the userid is burned. Need to fix!!!!!!!!!Importatnt!!!!!

    Private Sub cmdall_Click()
        Try
            If flagallow = 1 Then
                strwhere = ""
            Else
                'TEST QUERY
                strwhere = "WHERE PRPECH = 'LREDONDO' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = 'LREDONDO') "
                'strwhere = "WHERE PRPECH = '" & userid & "' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "') "
                'strwhere = "WHERE PRPECH = '" & UserID & "'
            End If
            fillcell1(strwhere)
            Exit Sub
        Catch ex As Exception
            'Call gnr.gotoerror("frmproductsdevelopment", "cmdall_click", Err.Number, Err.Description, Err.Source)
            gnr.gotoerror("frmproductsdevelopment", "cmdall_click", ex.HResult, ex.Message + ". " + ex.ToString, ex.Source)
        End Try
    End Sub

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
            Exit Sub
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


    'Private Sub dgvProjectDetails_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) _
    'Handles dgvProjectDetails.CellFormatting
    'Dim CurrentState As String = " "
    'Dim NewState As String = " "

    'For Each row As DataGridViewRow In dgvProjectDetails.Rows
    '       CurrentState = row.Cells(6).Value.ToString()
    'If CurrentState.Length <= 4 Then
    '           NewState = gnr.GetProjectStatusDescription(CurrentState)
    '          dgvProjectDetails.Rows(e.RowIndex).Cells("Status").Value = NewState
    'Else
    'Exit For
    'End If
    'Next
    'If e.ColumnIndex = 6 Then
    'If e.Value IsNot Nothing Then
    'CurrentState = e.Value.ToString
    'NewState = gnr.GetProjectStatusDescription(CurrentState)
    'dgvProjectDetails.Rows(e.RowIndex).Cells("Status").Value = NewState

    'End If
    'End If
    'End Sub

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


    Private Sub cmdall_Click(sender As Object, e As EventArgs) Handles cmdall.Click
        cmdall_Click()
    End Sub

    Private Sub dataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            Dim value11 As String = row.Cells(0).Value.ToString()
        Next
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

                            txtvendorno.Text = RowDs.Item(23).ToString()
                            txtvendorname.Text = RowDs.Item(36).ToString()
                            txtpartno.Text = RowDs.Item(1).ToString()
                            txtctpno.Text = RowDs.Item(7).ToString()
                            txtqty.Text = RowDs.Item(8).ToString()
                            txtmfr.Text = RowDs.Item(9).ToString()
                            txtmfrno.Text = RowDs.Item(10).ToString()

                            If cmbuser.FindStringExact(Trim(RowDs.Item(18).ToString())) Then
                                cmbuser.SelectedIndex = cmbuser.FindString(Trim(RowDs.Item(18).ToString()))
                            End If

                            txtpartdescription.Text = partDescription


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

    Protected Sub OnRowCommand(ByVal sender As Object, ByVal e As GridViewCommandEventArgs)
        Dim index As Integer = Convert.ToInt32(e.CommandArgument)
        Dim gvRow As DataGridViewRow = DataGridView1.Rows(index)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
End Class