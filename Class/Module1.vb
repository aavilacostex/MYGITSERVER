Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Public Module TextBoxExtensions

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function SendMessage(ByVal hWnd As HandleRef,
                                        ByVal msg As UInteger,
                                        ByVal wParam As IntPtr,
                                        ByVal lParam As String) As IntPtr
    End Function

    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub SetWatermark(ByVal ctl As Control, ByVal text As String)
        Const EM_SETCUEBANNER As Int32 = &H1501
        Const CB_SETCUEBANNER As Int32 = &H1703

        Dim retainOnFocus As IntPtr = New IntPtr(1)
        Dim msg As UInteger = EM_SETCUEBANNER

        If TypeOf ctl Is ComboBox Then
            msg = CB_SETCUEBANNER
        End If

        SendMessage(New HandleRef(ctl, ctl.Handle), msg, retainOnFocus, text)
    End Sub

End Module