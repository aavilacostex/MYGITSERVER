Imports System.Windows
Imports System.Windows.Forms.DataFormats
Imports System.Threading
Imports System.ComponentModel
Imports System.IO

Public Class MDIMain

    Dim gnr As Gn1 = New Gn1()
    Dim pathpictureparts As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'BackgroundWorker1.WorkerReportsProgress = True
        loadImage()

    End Sub

    Private Sub ProductsDevelopmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ProductsDevelopmentToolStripMenuItem1.Click

        frmProductsDevelopment.Show()
        'Loading.Show()
        'Loading.BringToFront()
    End Sub

    Private Sub LoginToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frmLogin.Show()
    End Sub

    Private Sub SupplierClaimsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SupplierClaimsToolStripMenuItem.Click
        frmclaimsvendor.Show()
    End Sub

    Private Sub Test1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Test1ToolStripMenuItem.Click
        frmLoadExcel.Show()
    End Sub

    Private Sub Test2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Test2ToolStripMenuItem.Click
        test.Show()
    End Sub

    Private Sub loadImage()
        Dim exMessage As String = " "
        Try
            Dim pathPictures = gnr.UrlPathImgNewMethod
            If Not Directory.Exists(pathPictures) Then
                'looking into embeded resorces
                Dim resource = GetType(MDIMain).Assembly.GetManifestResourceNames()

                If GetType(MDIMain).Assembly.GetManifestResourceStream(resource(17)) IsNot Nothing Then
                    PictureBox1.Image = New System.Drawing.Bitmap(GetType(MDIMain).Assembly.GetManifestResourceStream(resource(17)))
                Else
                    PictureBox1.Image = New System.Drawing.Bitmap(GetType(MDIMain).Assembly.GetManifestResourceStream(resource(27)))
                End If
            Else
                pathpictureparts = gnr.PathStartImageMethod
                pathpictureparts = If(File.Exists(pathpictureparts), pathpictureparts, pathPictures & "img_default_logo.jpg")
                If pathpictureparts IsNot Nothing Then
                    PictureBox1.Load(pathpictureparts)
                End If
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
        Exit Sub
    End Sub

End Class
