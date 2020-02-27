Imports System.Windows
Imports System.Windows.Forms.DataFormats

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ProductsDevelopmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ProductsDevelopmentToolStripMenuItem1.Click

        Form2.Show()



    End Sub

    Private Sub LoginToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoginToolStripMenuItem.Click
        Form3.Show()
    End Sub
End Class
