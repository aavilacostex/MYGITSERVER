Imports System.Windows
Imports System.Windows.Forms.DataFormats
Imports System.Threading
Imports System.ComponentModel
Imports System.IO

Public Class MDIMain

    Dim gnr As Gn1 = New Gn1()
    Dim pathpictureparts As String
    Public userid As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        loadImage()
        Dim user As String = Nothing
        Dim optionSelection As String = Nothing
        Dim valid As Boolean = False
        Dim lstNewMenus As String() = gnr.NewUserMenuCodes.Split(",")
        Dim currentCode As String = Nothing

        If CInt(gnr.FlagProductionMethod).Equals(1) Then
            Dim args As String() = Environment.GetCommandLineArgs()
            Dim argumentsJoined = String.Join(".", args)

            Dim arrayArgs As String() = argumentsJoined.Split(".")
            optionSelection = UCase(arrayArgs(3).ToString().Replace(",", ""))
            user = UCase(arrayArgs(2).ToString().Replace(",", ""))
            LikeSession.retrieveUser = user
            'MessageBox.Show(optionSelection & " - " & user, "CTP Sytems", MessageBoxButtons.OK)

            currentCode = If(optionSelection.Equals("OPT1"), lstNewMenus(0), lstNewMenus(1))

            If optionSelection.Equals("OPT1") Then
                If String.IsNullOrEmpty(gnr.AuthorizatedUser) Then
                    Dim dsGetUserMenuByUserId = gnr.GetMenuByUser(user)
                    If dsGetUserMenuByUserId IsNot Nothing Then
                        For Each item As DataRow In dsGetUserMenuByUserId.Tables(0).Rows
                            If item("CODDETMENU") = currentCode Then
                                valid = True
                                frmProductsDevelopment.Show()
                                Exit For
                            End If
                        Next
                    End If
                Else
                    If gnr.AuthorizatedUser.Equals("All") Then
                        'LoadCombos(sender, e)
                        'frmProductsDevelopment_load()
                        'MyBase.Hide()
                        frmProductsDevelopment.Show()
                    Else
                        Dim result = CheckCredentials(user)
                        If Not result Then
                            MessageBox.Show("Operation Error", "CTP System", MessageBoxButtons.OK)
                            Me.Close()
                            Exit Sub
                        Else
                            Dim rightAccess As Boolean = False
                            If valid Then
                                Dim lstUsers As String() = gnr.AuthorizatedUser.Split(",")
                                For Each item As String In lstUsers
                                    If UCase(user).Equals(UCase(item)) Then
                                        rightAccess = True
                                        Exit For
                                    End If
                                Next
                            End If

                            If rightAccess Then
                                'MessageBox.Show("right validation for user: " & user, "CTP System", MessageBoxButtons.OK)
                                'LoadCombos(sender, e)
                                'frmProductsDevelopment_load()
                                'MyBase.Hide()
                                frmProductsDevelopment.Show()
                            Else
                                Me.Close()
                            End If
                        End If
                    End If
                End If
            ElseIf optionSelection.Equals("OPT2") Then
                If String.IsNullOrEmpty(gnr.AuthorizatedUser) Then
                    Dim dsGetUserMenuByUserId = gnr.GetMenuByUser(user)
                    If dsGetUserMenuByUserId IsNot Nothing Then

                        'Dim lstNewMenus As String() = gnr.NewUserMenuCodes.Split(",")

                        For Each item As DataRow In dsGetUserMenuByUserId.Tables(0).Rows
                            If item("CODDETMENU") = currentCode Then
                                valid = True
                                frmLoadExcel.Show()
                                Exit For
                            End If
                        Next
                    End If
                Else
                    If gnr.AuthorizatedUser.Equals("All") Then
                        'LoadCombos(sender, e)
                        'frmProductsDevelopment_load()
                        'MyBase.Hide()
                        frmLoadExcel.Show()
                    Else
                        Dim result = CheckCredentials(user)
                        If Not result Then
                            MessageBox.Show("Operation Error", "CTP System", MessageBoxButtons.OK)
                            Me.Close()
                            Exit Sub
                        Else
                            Dim rightAccess As Boolean = False
                            If valid Then
                                Dim lstUsers As String() = gnr.AuthorizatedUser.Split(",")
                                For Each item As String In lstUsers
                                    If UCase(user).Equals(UCase(item)) Then
                                        rightAccess = True
                                        Exit For
                                    End If
                                Next
                            End If

                            If rightAccess Then
                                'MessageBox.Show("right validation for user: " & user, "CTP System", MessageBoxButtons.OK)
                                'LoadCombos(sender, e)
                                'frmProductsDevelopment_load()
                                'MyBase.Hide()
                                frmLoadExcel.Show()
                            Else
                                Me.Close()
                            End If
                        End If
                    End If
                End If
            Else
                MessageBox.Show("OPT3", "CTP Sytems", MessageBoxButtons.OK)
            End If
        Else

            If String.IsNullOrEmpty(gnr.AuthorizatedUser) Then
                user = gnr.AuthorizatedTestUser.Split(",")(0).ToString()
                currentCode = lstNewMenus(0)
                Dim dsGetUserMenuByUserId = gnr.GetMenuByUser(user)
                If dsGetUserMenuByUserId IsNot Nothing Then
                    For Each item As DataRow In dsGetUserMenuByUserId.Tables(0).Rows
                        If item("CODDETMENU") = currentCode Then
                            valid = True
                            frmProductsDevelopment.Show()
                            Exit For
                        End If
                    Next
                End If
            Else
                Dim rightAccess As Boolean = False
                If valid Then
                    Dim lstUsers As String() = gnr.AuthorizatedUser.Split(",")
                    For Each item As String In lstUsers
                        If UCase(user).Equals(UCase(item)) Then
                            rightAccess = True
                            Exit For
                        End If
                    Next
                End If

                If rightAccess Then
                    'MessageBox.Show("right validation for user: " & user, "CTP System", MessageBoxButtons.OK)
                    'LoadCombos(sender, e)
                    'frmProductsDevelopment_load()
                    'MyBase.Hide()
                    frmProductsDevelopment.Show()
                End If
            End If


            'If Not valid Then
            '    user = If(String.IsNullOrEmpty(gnr.AuthorizatedTestUser), "All", UCase(gnr.AuthorizatedTestUser))
            '    LikeSession.retrieveUser = user


            '    'Dim user1 = "aavila"
            '    'Dim pass1 = "alexei20"

            '    'Dim check = gnr.checkusr(Trim(UCase(user1)), Trim(UCase(pass1)))

            '    Dim rightAccess As Boolean = False
            '    Dim rightAccessTest As Boolean = False
            '    Dim lstUsers As String() = gnr.AuthorizatedUser.Split(",")
            '    For Each item As String In lstUsers
            '        If UCase(user).Equals(UCase(item)) Then
            '            rightAccess = True
            '            Exit For
            '        End If
            '    Next

            '    If rightAccess Then
            '        MessageBox.Show("right validation for user: " & user, "CTP System", MessageBoxButtons.OK)
            '        'LoadCombos(sender, e)
            '        'frmProductsDevelopment_load()
            '        'MyBase.Hide()
            '        frmProductsDevelopment.Show()
            '    Else
            '        Dim lstUsersTests As String() = gnr.AuthorizatedTestUser.Split(",")
            '        For Each item As String In lstUsersTests
            '            If UCase(user).Equals(UCase(item)) Then
            '                rightAccessTest = True
            '                Exit For
            '            End If
            '        Next

            '        If rightAccessTest Then
            '            frmProductsDevelopment.Show()
            '        End If

            '    End If
            'End If

        End If
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

    Private Function CheckCredentials(user As String) As Boolean
        Dim exMessage As String = " "
        Try
            Dim dsCheck = gnr.getUserDataByUsername(Trim(UCase(user)))

            If dsCheck IsNot Nothing Then
                userid = Trim(UCase(user))
                'pass = Trim(UCase(txtpassword.Text)))
                Return True
            Else
                Return False
                'MsgBox("Invalid Password, try again!", vbOKOnly + vbInformation, "CTP System")
                'txtPassword.SetFocus
                'SendKeys.Send("{Home}+{End}")
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return False
        End Try
    End Function

End Class
