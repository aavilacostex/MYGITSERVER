﻿Imports System.Net.Mail

Public Class Gn1

    Public Const Version = "V.02/20/20"
    Public Const strCompany = "COSTEX"
    Public Const strdatabase = "dbCTPSystem"
    Public pathpicture As String
    Public pathgeneral As String
    Public Const strconnection = "DSN=COSTEX400;UID=INTRANET;PWD=CTP6100"
    Public Const strcrystalconn = "DSN=COSTEX400;UID=INTRANET;PWD=CTP6100;"
    Public Const strconnSQL = "DSN=CTPSystem;UID=sa;PWD=ctp6100;"
    Public Const strcrystalconnSQL = "DSN=CTPSystem;UID=sa;PWD=ctp6100;"
    Public Const strmailhostctp = "mail.costex.com"
    'Public Const strmailhostctp = "mail.costex.com"
    'Public Const strconnSQL = "Driver={SQL Server};Server=dellserver;Database=dbCTPSystem;Uid=sa;Pwd=ctp6100;"
    'Public Const strcrystalconnSQL = "Driver={SQL Server};Server=dellserver;Database=dbCTPSystem;Uid=sa;Pwd=ctp6100;"
    Public Const strconnSQLNOVA = "DSN=NOVATIME;UID=NTI_CS;PWD=csadmin;"
    Public Const strcrystalconnSQLNOVA = "DSN=NOVATIME;UID=NTI_CS;PWD=csadmin;"
    Public emailspath As String
    Public pictureSizeFlag As Integer
    Public filesQuantity As Integer
    Public filesToWrite As Integer
    Public userDepartment As String
    Public flagSalesman As Integer
    Public primaryservername As String
    Public Conn As New ADODB.Connection
    Public ConnSql As New ADODB.Connection
    Public ConnSqlNOVA As New ADODB.Connection
    Public CMD As New ADODB.Command
    Public rs As New ADODB.Recordset
    Public Rs1 As New ADODB.Recordset
    Public Rs2 As New ADODB.Recordset
    Public Rs3 As New ADODB.Recordset
    Public Rs4 As New ADODB.Recordset
    Public RsGeneral As New ADODB.Recordset
    Public userid As String
    Public flagexit As Integer
    Public flaguserrec As Integer
    Public getclaimflag As Integer
    Public claimsplit As Integer
    Public getclaimnosave As Integer
    Public seeaddcomments As Integer
    Public seeaddprocomments As Integer
    Public printpath As String
    Public flagchangevendor As Integer
    Public getclaimno As Long
    Public getclaim As Long
    Public pass As String
    Public check As String
    Public encrpwd As String
    Public passcomm As String
    Public pototalcost As Double
    Public actupdatepo As Long
    Public prpagrid As Integer
    Public fso As New Scripting.FileSystemObject
    Public IP As String
    Public ipaddresslocal As String
    Public Provider As String
    Public projectnoadd As Long
    Public DataSource As String
    Public user As String
    Public password As String
    Public InitialCatalog As String
    'Public objMail As New MailSender
    Public strHost As String, strPort As String, strfrom As String
    Public strFromName As String, strto As String, strSubject As String
    Public strBody As String, stratt As String, stratt1 As String
    Public DirTrabajo, DirLog As String
    Public LoginSucceeded As Boolean
    Public codloginctp As Long
    Public strWarr As String, strNonw As String, strIntr As String
    Public strOpen As String, strClos As String, strwhere As String
    Public strInts As String, strPcus As String, strPoth As String, strFinl As String
    Public countgridrows As Integer
    'variables para convertir un amount en text - begin
    'Set up two arrays to hold string values we
    'will use to convert numbers to words
    Public BigOnes(9) As String
    Public SmallOnes(19) As String
    'Declare variables
    Public Dollars As String
    Public Cents As String
    Public Words As String
    Public Chunk As String
    Public Digits As Integer
    Public LeftDigit As Integer
    Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As String, pdwSize As Long, ByVal bOrder As Long) As Long
    Public RightDigit As Integer
    Public instanceOfModel_ID As Integer
    Public test As String
    Public Const VBObjectError As Integer = -2147221504
    Public Versionctp As String


    Public Sub Valida_DirLog()
        On Error GoTo Valida_DirLog_Err
        ChDir(DirLog)
        ChDir(DirTrabajo)
        Exit Sub
Valida_DirLog_Err:
        Select Case Err()
            'Case 76
            'MkDir(DirTrabajo + "Log")
        End Select
        Err.Clear()
        ChDir(DirTrabajo)
    End Sub

    'Public Sub Generate_Log(Message As String)
    'On Error GoTo Generate_Log_Err
    'Dim LogFile As String
    '   LogFile = ""
    '  LogFile = DirLog + "ErrLog_" + Trim(Str(DatePart("m", Date))) + ".log"
    ' Open LogFile For Append As #2
    'Write #2, Format(Of Date, "mm/dd/yyyy")() + " " + Trim(Str(Time())) + "|" + Message
    'Close #2
    'Exit Sub
    'Generate_Log_Err:
    '       Close #2
    'End Sub

    ' Returns an array with the local IP addresses (as strings).
    ' Author: Christian d'Heureuse, www.source-code.biz

    Public Sub gotoerror(Forms, Events, errnumber, errdescription, errsource)
        'Dim error As String
        Dim sql As String
        Dim intrespond As Long
        On Error GoTo errhandler

        'Error = Forms + "-" + Events + "-" + Trim(Str(errnumber)) + "-" + errdescription + errsource + "-" + Version
        'sql = "INSERT INTO ERRORCTP VALUES('" & Replace(Left(error, 500), "'", "") & "','" & userid & "','" & Format(Now, "yyyy-mm-dd") & "')"
        Conn.Execute(sql)

        intrespond = MsgBox("Error. See Log", vbInformation + vbOKOnly, "CTP System")

        Exit Sub
errhandler:
        'Error = Forms + "-" + "gotoerror" + "-" + Trim(Str(Err.Number)) + "-" + Err.Description + Err.Source + "-" + "Err on gotoerror" + "-" + Version
        'sql = "INSERT INTO ERRORCTP VALUES('" & Replace(Left(error, 500), "'", "") & "','" & userid & "','" & Format(Now, "yyyy-mm-dd") & "')"
        Conn.Execute(sql)
        intrespond = MsgBox("Error. See Log", vbInformation + vbOKOnly, "CTP System")
    End Sub

    Public Sub gotologuse(Progname, Area, Keydata)
        Dim sql As String
        Dim codloguse As Long
        On Error GoTo errhandler

        'codloguse = getmax("logusectp", "codloguse")
        sql = "INSERT INTO LOGUSECTP VALUES(" & codloguse & ",'" & userid & "','" & ipaddresslocal & "','" & Version & "','" & Progname & "','" & Area & "','" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','" & Keydata & "')"
        Conn.Execute(sql)

        Exit Sub
errhandler:
        Call gotoerror("general", "gotologuse", Err.Number, Err.Description, Err.Source)
    End Sub

    Public Function getmax(table, field)
        'Dim error As String
        'Dim intrespond As Long
        'Dim sentence As Variant

        'Set RsGeneral = Nothing
        'sentence = "Select Max(" & field & ") as max from " & table
        'Set RsGeneral = Conn.Execute(sentence)
        'If IsNull(RsGeneral.Fields("max")) Then
        'getmax = 1
        'Else
        'getmax = RsGeneral.Fields("max") + 1
        'End If

        Exit Function
errhandler:
        Call gotoerror("general", "getmax", Err.Number, Err.Description, Err.Source)
    End Function

    Public Function GetIpAddrTable()
        Dim Buf(0 To 511) As Byte
        Dim BufSize As Long : BufSize = UBound(Buf) + 1
        Dim rc As Long
        Dim ArrayOk As Array

        rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
        'If rc <> 0 Then Err.Raise VBObjectError, , "GetIpAddrTable failed with return value " & rc
        If rc <> 0 Then Err.Raise(VBObjectError, , "GetIpAddrTable failed with return value " & rc)
        Dim NrOfEntries As Integer : NrOfEntries = Buf(1) * 256 + Buf(0)
        If NrOfEntries = 0 Then GetIpAddrTable = ArrayOk : Exit Function
        'ReDim IpAddrs(0 To NrOfEntries - 1) As String
        Dim IpAddrs() As String
        ReDim IpAddrs(0 To NrOfEntries - 1)
        Dim i As Integer
        For i = 0 To NrOfEntries - 1
            Dim j As Integer, s As String : s = ""
            For j = 0 To 3 : s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j) : Next
            IpAddrs(i) = s
        Next
        GetIpAddrTable = IpAddrs
    End Function

    Public Function checkstring(StrInput)

        If InStr(1, Trim(StrInput), "'") Or InStr(1, Trim(StrInput), "|") Or InStr(1, Trim(StrInput), "`") Or InStr(1, Trim(StrInput), "~") Or InStr(1, Trim(StrInput), "!") Or InStr(1, Trim(StrInput), "^") Or InStr(1, Trim(StrInput), "_") Or InStr(1, Trim(StrInput), "=") Or InStr(1, Trim(StrInput), "\") Or InStr(1, Trim(StrInput), "%") Or InStr(1, Trim(StrInput), "+") Or InStr(1, Trim(StrInput), "[") Or InStr(1, Trim(StrInput), "]") Or InStr(1, Trim(StrInput), "?") Or InStr(1, Trim(StrInput), "<") Or InStr(1, Trim(StrInput), ">") Then
            checkstring = False
        Else
            checkstring = True
        End If

        Exit Function
errhandler:
        'Call gotoerror("general", "checkstring", Err.Number, Err.Description, Err.Source)
    End Function

    Function checkusr(userid, pass)
        'call routine for encrypting password
        Dim as400 As New cwbx.AS400System
        Dim prog As New cwbx.Program
        Dim parms As New cwbx.ProgramParameters
        Dim server As New cwbx.SystemNames
        Dim stringCvtr As New cwbx.StringConverter
        Dim wuser, wpass, wswvld
        On Error GoTo errhandler

        'Program Parameters
        wuser = Left((Trim(UCase(userid)) & "          "), 10)
        wpass = Left((Trim(UCase(pass)) & "          "), 10)
        wswvld = "0"

        'AS400 Connection Parameters
        as400.Define(server.DefaultSystem)
        as400.UserID = "INTRANET"
        as400.Password = "CTP6100"
        'as400.PromptMode = cwbcoPromptNever   here
        as400.Signon()

        'Program to call
        prog.system = as400
        prog.LibraryName = "CTPINV"
        prog.ProgramName = "PSWVLDR"

        parms.Clear()

        'Parameters Definition
        'parms.Append("USER", cwbrcInput, 10)    here
        'parms.Append("PASS", cwbrcInput, 10)    here
        parms.Append("SWVLD", cwbx.cwbrcParameterTypeEnum.cwbrcInout, 1)

        stringCvtr.CodePage = 37

        'Assign Values to Parameters
        'parms("USER") = stringCvtr.ToBytes(wuser)   here
        'parms("PASS") = stringCvtr.ToBytes(wpass)   here
        'parms("SWVLD") = stringCvtr.ToBytes(wswvld)  here

        prog.Call(parms)

        checkusr = stringCvtr.FromBytes(parms("SWVLD").Value)

        ' as400.Disconnect(cwbcoServiceAll)  here

        Exit Function
errhandler:
        Call gotoerror("general", "checkusr", Err.Number, Err.Description, Err.Source)
    End Function

    'Public Sub Generate_Log(Message As String)
    'On Error GoTo Generate_Log_Err
    'Dim LogFile As String
    '   LogFile = ""
    '   LogFile = DirLog + "ErrLog_" + Trim(Str(DatePart("m", Date))) + ".log"
    '  Open LogFile For Append As #2
    'Write #2, Format(Of Date, "mm/dd/yyyy")() + " " + Trim(Str(Time())) + "|" + Message
    'Close #2
    'Exit Sub
    'Generate_Log_Err:
    'Close #2
    'End Sub






End Class
