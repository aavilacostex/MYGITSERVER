Public Class frmLogin

#Region "Variables"

    Public intrespond As Long
    Public LoginSucceeded As Boolean
    Public s As String
    Public sql As String

    Dim frm As frmLogin
    Dim gnr As Gn1 = New Gn1()

    Public Conn As New ADODB.Connection
    Public ConnSql As New ADODB.Connection
    Public codloginctp As Long
    Public Versionctp As String '= gnr.Versionctp
    Public rs As ADODB.Recordset '= gnr.rs
    Dim CurrentCTPVersion As Version = My.Application.Info.Version
    Dim userid As String '= gnr.userid
    Dim passcomm As String '= gnr.passcomm

#End Region

#Region "Key Events"

    Private Sub txtUserName_GotFocus(sender As Object, e As EventArgs) Handles txtUserName.GotFocus
        txtUserName.SelectionStart = 0
        txtUserName.SelectionLength = Len(txtUserName.Text)
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

#End Region

#Region "Control Events"

    Private Sub txtPassword_TextChanged(sender As Object, e As EventArgs) Handles txtPassword.TextChanged
        'txtPassword.SelectionStart = 0
        'txtPassword.SelectionLength = Len(txtPassword.Text)
    End Sub

    Private Sub cmdcancel_Click()
        Try
            'sql = "CALL CTPINV.RECLAIM"
            'Conn.Execute (sql)
            sql = "delete from loginctp where codlogin = " & codloginctp
            Conn.Execute(sql)
            If Conn.State = 1 Then
                Conn.Close()
            End If
            If ConnSql.State = 1 Then
                ConnSql.Close()
            End If
            LoginSucceeded = False
            Me.Close()
        Catch ex As Exception
            LoginSucceeded = False
            Me.Close()
        End Try

    End Sub

    Private Sub cmdok_Click()
        'Dim WinWnd As Long
        Dim check As String
        On Error GoTo errhandler
        Dim totaldays As Integer

        'WinWnd = FindWindow(vbNullString, "CTPSystem " & Version)
        'If WinWnd <> 0 Then
        '    End
        '    Exit Sub
        'End If

        Dim colorbackcolor As Integer = 0
        Dim initialwindow As String = "Main Menu"

        'servername = Conn.DefaultDatabase
        Dim servername As String = ""
        check = gnr.checkusr(Trim(UCase(txtUserName.Text)), Trim(UCase(txtPassword.Text)))
        If check = "U" Then
            MsgBox("Username not valid.", vbInformation + vbOKOnly, "CTP System")
            'txtUserName.SetFocus
        Else
            If check = "N" Then
                MsgBox("User not authorized.", vbOKOnly + vbInformation, "CTP System")
                Exit Sub
            End If
            If check = "E" Then
                userid = Trim(UCase(txtUserName.Text))
                MsgBox("Your password has expired; please change it.", vbOKOnly + vbInformation, "CTP System")
                'frmpasschange.Show 1
                If passcomm = "" Then
                    MsgBox("Password expired.", vbOKOnly + vbInformation, "CTP System")
                    Exit Sub
                Else
                    check = "0"
                End If
            End If
            If check = "0" And Len(Trim(UCase(txtPassword.Text))) = 0 Then
                userid = Trim(UCase(txtUserName.Text))
                MsgBox("You need to set a new password; please change it.", vbOKOnly + vbInformation, "CTP System")
                'frmpasschange.Show 1
                If passcomm = "" Then
                    MsgBox("Password not set.", vbOKOnly + vbInformation, "CTP System")
                    Exit Sub
                Else
                    check = "0"
                End If
            End If
            If check = "0" Then
                userid = Trim(UCase(txtUserName.Text))
                'pass = Trim(UCase(txtpassword.Text)))

                Dim dsUsrData = gnr.getUserDataByUsername(userid)
                If Not dsUsrData Is Nothing Then
                    If dsUsrData.Tables(0).Rows.Count > 0 Then
                        initialwindow = "Main Menu"
                        LoginSucceeded = True
                        Me.Hide()
                        MDIMain.Show()
                        'MDIMain.toolbar1.Visible = True
                        If dsUsrData.Tables(0).Rows(0).ItemArray(dsUsrData.Tables(0).Columns("DECODE").Ordinal) = 14 Then 'Or userid = "JDMERCADO" Then
                            Dim dsMarktData = gnr.getMarketingDataByDate()
                            If Not dsMarktData Is Nothing Then
                                If dsMarktData.Tables(0).Rows.Count > 0 Then
                                    Dim dbDate = CDate(dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACABD").Ordinal))
                                    Dim todayDate = CDate(Now.ToShortDateString())
                                    totaldays = (dbDate - todayDate).TotalDays
                                    Dim macana = dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACANA").Ordinal)
                                    Dim macabd = dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACABD").Ordinal)
                                    Dim macaco = dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACACO").Ordinal)

                                    If totaldays <= dsUsrData.Tables(0).Rows(0).ItemArray(dsUsrData.Tables(0).Columns("MACADY").Ordinal) Then
                                        Dim rsMessage As DialogResult = MessageBox.Show(" " & Trim(macana) & " Date : " & Format(macabd, "mm/dd/yyyy"), "CTP System", MessageBoxButtons.OKCancel)
                                        If rsMessage = DialogResult.Cancel Then
                                            Dim qryResult = gnr.UpdateMarktCampaignData(macaco)
                                            If qryResult <> 0 Then
                                                'error actualizacion mensaje
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        sql = "delete from loginctp where codlogin = " & codloginctp
                        Conn.Execute(sql)
                        'Versionctp = CurrentCTPVersion.Build & " - " & Strings.Right(ipaddresslocal, 5)
                        codloginctp = gnr.getmax("loginctp", "codlogin")
                        sql = "INSERT INTO LOGINCTP VALUES(" & codloginctp & ",'" & userid & "','" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','" & Versionctp & "')"
                        Conn.Execute(sql)
                        Call amenu()

                        If userid = "CARLOS" Or userid = "JDMERCADO" Or userid = "MVELEZ" Or userid = "KRODRIGUEZ" Or userid = "JDMIRA" Or userid = "HOLIVEROS" Or userid = "LARIAS" Then
                            ConnSql.ConnectionString = gnr.strconnSQL
                            ConnSql.Open()

                            gnr.ConnSqlNOVA.ConnectionString = gnr.strconnSQLNOVA
                            gnr.ConnSqlNOVA.Open()
                        End If
                    End If
                End If





            Else
                        MsgBox("Invalid Password, try again!", vbOKOnly + vbInformation, "CTP System")
                'txtPassword.SetFocus
                SendKeys.Send("{Home}+{End}")
            End If
        End If

        rs = Nothing

        Exit Sub
errhandler:
        rs = Nothing
        If gnr.ConnSqlNOVA.State = 0 Then
            MsgBox("Connection to NOVATIME failed!", vbOKOnly + vbInformation, "CTP System")
        End If
        Exit Sub
        'Call gotoerror("frmlogin", "cmdok_click", Err.Number, Err.Description, Err.Source)
    End Sub

#End Region

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

    Private Sub Form_Load()
        Dim IpAddrs
        Dim find1 As Integer
        Dim find2 As Integer
        Dim find3 As Long
        Dim find4 As String
        Dim printpath As String
        Dim exMessage As String = " "

        'On Error GoTo errhandler
        Try
            Conn.ConnectionString = Gn1.strconnection
            Conn.Open()

            Dim dsControlData = gnr.GetDataByPartMix()
            If Not dsControlData Is Nothing Then
                If dsControlData.Tables(0).Rows.Count > 0 Then
                    If dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "SVR" Then
                        gnr.primaryservername = Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                            Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                    End If
                    If dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "PIC" Then
                        gnr.pathpicture = Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                            Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                    End If
                    If dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "GEN" Then
                        gnr.pathgeneral = Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                            Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                    End If
                    If dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "EMA" Then
                        gnr.emailspath = Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                            Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                    End If
                    If dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "REP" Then
                        gnr.printpath = Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                            Trim(dsControlData.Tables(0).Rows(0).ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                    End If
                End If
            End If
            'Public Const pathpicture = "\\Dellserver\CTP_System\images\Employee ID Pictures\"
            'Public Const pathgeneral = "\\Dellserver\Inetpub_D\"
            'Public Const emailspath = "\\Dellserver\Inetpub_D\CTP_System\Emails"

            IpAddrs = gnr.GetIpAddrTable
            gnr.ipaddresslocal = IpAddrs(1)
            Versionctp = CurrentCTPVersion.Build & " - " & Strings.Right(IpAddrs(1), 5)
            'Versionctp = Version & " - " & Right(IpAddrs(1), 5)
            find1 = InStr(1, Trim(IpAddrs(1)), ".")
            find2 = InStr(find1 + 1, Trim(IpAddrs(1)), ".")
            find3 = InStr(find2 + 1, Trim(IpAddrs(1)), ".")
            find4 = Mid(Trim(IpAddrs(1)), find2 + 1, find3 - find2 - 1)
            If find4 = 12 Then
                printpath = "\\Dalsvr\CTP_System\Reports"
            End If

            codloginctp = gnr.getmax("loginctp", "codlogin")
            sql = "INSERT INTO LOGINCTP VALUES(" & codloginctp & ",'" & userid & "','" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','" & Versionctp & "')"
            Conn.Execute(sql)
            Exit Sub
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub amenu()

    End Sub

    Private Sub cmdok_Click(sender As Object, e As EventArgs) Handles cmdok.Click
        cmdok_Click()
    End Sub

    Private Sub cmdcancel_Click(sender As Object, e As EventArgs) Handles cmdcancel.Click
        cmdcancel_Click()
    End Sub

End Class