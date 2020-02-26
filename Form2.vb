Public Class Form2

    Public flagdeve As Long '1 is new
    'Public filepicture As New clsReadWrite
    Public strwhere As String
    Public userid As String
    Public flagnewpart As Integer
    Public flagallow As Integer
    Public puragent As Integer

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Button7.FlatStyle = FlatStyle.Flat

        DataGridView1.RowHeadersVisible = False

        Button12.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\doc.PNG")
        Button12.ImageAlign = ContentAlignment.MiddleRight
        Button12.TextAlign = ContentAlignment.MiddleLeft

        Button13.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\save.PNG")
        Button13.ImageAlign = ContentAlignment.MiddleRight
        Button13.TextAlign = ContentAlignment.MiddleLeft

        Button14.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\exit.PNG")
        Button14.ImageAlign = ContentAlignment.MiddleRight
        Button14.TextAlign = ContentAlignment.MiddleLeft


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


End Class