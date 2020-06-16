Imports System.Windows
Imports System.Windows.Forms.DataFormats
Imports System.Threading
Imports System.ComponentModel

Public Class MDIMain
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        BackgroundWorker1.WorkerReportsProgress = True

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

#Region "Threads"

    'Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) _
    '    Handles BackgroundWorker1.RunWorkerCompleted
    '    Loading.Close()
    'End Sub

    'Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) _
    '    Handles BackgroundWorker1.DoWork
    '    Dim strMethodName = LikeSession.currentAction
    '    'If strMethodName.Equals("cmdsearchcode_Click") Then
    '    launchPDev(sender, e)
    '    'End If
    '    'LoadCombos(sender, e)
    'End Sub

    'Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) _
    '    Handles BackgroundWorker1.ProgressChanged
    '    frmProductsDevelopment.txtMfrNoSearch.Text = e.ProgressPercentage.ToString()
    'End Sub

    'Private Sub launchPDev(ByVal sender As Object, ByVal e As DoWorkEventArgs)

    '    Dim bgWorker = CType(sender, BackgroundWorker)
    '    For index = 0 To 10
    '        bgWorker.ReportProgress(index)
    '        Thread.Sleep(1000)
    '    Next
    'End Sub

#End Region
End Class
