Public Class frmChangeVendor
    Dim gnr As Gn1 = New Gn1()
    Public userid As String
    Public flagchangevendor As Integer

    Private Sub frmChangeVendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmChangeVendor_Load()
    End Sub

    Private Sub frmChangeVendor_Load()
        Dim exMessage As String = " "
        Try
            txtCode.Text = ""
            txtsearch1.Text = ""
            flagchangevendor = 1
            cmbvendor.Items.Clear()
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub cmdSearch_Click()
        Dim strfield As String
        Dim dfield
        Dim exMessage As String = " "
        Try
            txtsearch1.Text = ""
            cmbvendor.DataSource = Nothing
            dfield = Trim(txtCode.Text)
            If Len(Trim(txtCode.Text)) > 0 Then
                cmbvendor.Items.Clear()
                For i = 1 To Len(Trim(dfield))
                    strfield = Mid(dfield, i, 1)
                    If strfield Like "[!0-9]" Then
                        MsgBox("Just numbers", vbOKOnly + vbInformation, "CTP System")
                        i = Len(Trim(dfield)) + 1
                        Exit Sub
                    End If
                Next i
                If Trim(txtCode.Text) > 0 Then
                    Dim dsResult = gnr.GetVendorByVendorNo(txtCode.Text)

                    If dsResult IsNot Nothing Then
                        If dsResult.Tables(0).Rows.Count > 0 Then
                            cmbvendor.DataSource = dsResult.Tables(0)
                            cmbvendor.DisplayMember = "VMNAME"
                            cmbvendor.ValueMember = "VMVNUM"
                            If cmbvendor.Items.Count > 0 Then
                                cmbvendor.SelectedIndex = 0
                            End If
                        Else
                            cmbvendor.Items.Clear()
                            MsgBox("Vendor(s) not found.", vbOKOnly + vbInformation, "CTP System")
                            txtCode.Text = 0
                        End If
                    End If
                Else
                    If Trim(txtCode.Text) = 0 Then
                        cmbvendor.Items.Clear()
                    End If
                End If
            Else
                If Trim(txtCode.Text) = "" Then
                    txtCode.Text = 0
                    Exit Sub
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdsearch1_Click()
        Dim exMessage As String = " "
        Try
            cmbvendor.Items.Clear()
            txtCode.Text = ""
            If Trim(txtsearch1.Text) <> "" Then

                'Dim Sql = "SELECT * FROM VNMAS WHERE TRIM(UCASE(VMNAME)) LIKE '" & Trim(UCase(txtsearch1.Text)) & "%'"
                Dim dsResult = gnr.GetVendorByName(txtsearch1.Text)

                If dsResult IsNot Nothing Then
                    If dsResult.Tables(0).Rows.Count > 0 Then
                        cmbvendor.DataSource = dsResult.Tables(0)
                        cmbvendor.DisplayMember = "VMNAME"
                        cmbvendor.ValueMember = "VMVNUM"
                        If cmbvendor.Items.Count > 0 Then
                            cmbvendor.SelectedIndex = 0
                        End If
                    Else
                        cmbvendor.Items.Clear()
                        MsgBox("Vendor(s) not found.", vbOKOnly + vbInformation, "CTP System")
                        txtCode.Text = 0
                    End If
                End If
            Else
                If Trim(txtCode.Text) = 0 Then
                    cmbvendor.Items.Clear()
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub cmdchange_Click()

        If Trim(txtCode.Text) <> "" Then
            If flagchangevendor = 1 Then
                Call gnr.changeVendor(frmProductsDevelopment.txtpartno.Text, txtCode.Text, userid)
            End If
            If flagchangevendor = 2 Then
                'Call gnr.changeVendor(frmproductsdevelopmentTS.txtpartno.Text, txtCode.Text, userid)
            End If
            If flagchangevendor = 3 Then
                ' Call gnr.changeVendor(frmproductsdevelopmentpur.txtpartno.Text, txtCode.Text, userid)
            End If
            MsgBox("Vendor Changed.", vbOKOnly + vbInformation, "CTP System")
        End If
        Exit Sub
    End Sub

    Private Sub cmdexit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdchange_Click(sender As Object, e As EventArgs) Handles cmdchange.Click
        cmdchange_Click()
    End Sub

    Private Sub cmdsearch_Click(sender As Object, e As EventArgs) Handles cmdsearch.Click
        cmdSearch_Click()
    End Sub

    Private Sub cmdsearch1_Click(sender As Object, e As EventArgs) Handles cmdsearch1.Click
        cmdsearch1_Click()
    End Sub
End Class