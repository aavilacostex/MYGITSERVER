Public Class Form3
    Public intrespond As Long
    Public LoginSucceeded As Boolean
    Public s As String

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub txtUserName_GotFocus(sender As Object, e As EventArgs) Handles txtUserName.GotFocus
        txtUserName.SelectionStart = 0
        txtUserName.SelectionLength = Len(txtUserName.Text)
    End Sub

    Private Sub txtPassword_TextChanged(sender As Object, e As EventArgs) Handles txtPassword.TextChanged
        txtPassword.SelectionStart = 0
        txtPassword.SelectionLength = Len(txtPassword.Text)
    End Sub

    Private Sub txtpassword_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            'Call cmdok_Click
        End If
    End Sub

    Private Sub txtusername_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys.Send("{tab}")
            'SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
        End If
    End Sub

    Private Sub cmdok_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys.Send("{tab}")
            'SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
        End If
    End Sub

    Private Sub atoolbar(dkey)
        On Error GoTo errhandler
        'j = 1
        'For j = 1 To MDIMain.toolbar1.Buttons.Count
        '    If MDIMain.toolbar1.Buttons.Item(j).Key = dkey Then
        '        MDIMain.toolbar1.Buttons.Item(j).Enabled = False
        '        j = MDIMain.toolbar1.Buttons.Count + 1
        '    End If
        'Next j
        Exit Sub
errhandler:
        'Call gotoerror("frmlogin", "atoolbar", Err.Number, Err.Description, Err.Source)
    End Sub


End Class