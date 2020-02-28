Imports System.Globalization
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
        TabControl1.ItemSize = (New Size(TabControl1.Width / TabControl1.TabCount, 0))
        TabControl1.Padding = New System.Drawing.Point(300, 10)
        TabControl1.Appearance = TabAppearance.FlatButtons
        'TabControl1.ItemSize = New Size(0, 1)
        TabControl1.SizeMode = TabSizeMode.Fixed

        Button1.FlatStyle = FlatStyle.Flat
        Button2.FlatStyle = FlatStyle.Flat
        Button3.FlatStyle = FlatStyle.Flat
        Button4.FlatStyle = FlatStyle.Flat
        Button5.FlatStyle = FlatStyle.Flat
        Button6.FlatStyle = FlatStyle.Flat
        cmdall.FlatStyle = FlatStyle.Flat

        DataGridView1.RowHeadersVisible = False
        DataGridView2.RowHeadersVisible = False

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
        'DataGridView1. SelectionChanged += New EventHandler(dataGridView1_SelectionChanged)


    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

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

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub TableLayoutPanel4_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel4.Paint
        'TableLayoutPanel4.CellBorderStyle = TableLayoutPanelCellBorderStyle.Outset
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

    End Sub

    'here class traduction begans---------------------------------------------------------

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

    Private Sub cmdall_Click(sender As Object, e As EventArgs) Handles cmdall.Click
        cmdall_Click()
    End Sub

    Private Sub dataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            Dim value11 As String = row.Cells(0).Value.ToString()
        Next
    End Sub

End Class