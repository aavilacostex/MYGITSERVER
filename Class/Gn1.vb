Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Net.Mail

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
    Public strDate As String = "1900,01,01"
    Public formats() As String = {"M/d/yyyy h:mm:ss tt", "M/d/yyyy h:mm tt", "MM/dd/yyyy hh:mm:ss", "M/d/yyyy h:mm:ss", "M/d/yyyy hh:mm tt",
        "M/d/yyyy hh tt", "M/d/yyyy h:mm", "M/d/yyyy h:mm", "MM/dd/yyyy hh:mm", "M/dd/yyyy hh:mm", "MM/d/yyyy HH:mm:ss.ffffff"}


#Region "Selects"

    Public Function GetDataByPRHCOD(code As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLH WHERE PRHCOD = " & Trim(code)
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodeAndPartNo(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & Trim(code) & " AND trim(ucase(PRDPTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodeAndPartNoProdDesc(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLD WHERE PRHCOD = " & Trim(code) & " AND trim(ucase(PRDPTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetProdDetByCodeAndExc(code As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from prdvld where prhcod = " & Trim(code) & " and prdsts <> 'CS' and prdsts <> 'CN' and prdsts <> 'CD' and prdsts <> 'CL'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetCodeAndNameByPartNo(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT PRDVLH.PRHCOD,PRDVLH.PRNAME FROM PRDVLH INNER JOIN PRDVLD ON PRDVLH.PRHCOD = PRDVLD.PRHCOD WHERE TRIM(PRDPTN) = '" & Trim(UCase(partNo)) & "' ORDER BY PRDVLD.CRDATE DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartNo2(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM DVINVA INNER JOIN VNMAS ON DVINVA.DVPRMG = digits(VNMAS.VMVNUM) WHERE DVPART = '" & Trim(UCase(partNo)) & "' and dvlocn = '01'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartVendor(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM DVINVA WHERE DVPART = '" & Trim(UCase(partNo)) & "' and dvlocn = '01'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartMix() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM CNTRLL WHERE CNT01 = '120' ORDER BY TRIM(CNTDE1)"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByVendorAndPartNo(vendorNo As String, partNo As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim strDescrption As String
        Dim columnToChange = "PQMIN"
        Try
            Sql = "SELECT * FROM POQOTA WHERE PQVND = " & Trim(vendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%' ORDER BY PQQDTY DESC, PQQDTM DESC, PQQDTD DESC"
            strDescrption = GetSingleDataFromDatabase(Sql, columnToChange)
            Return strDescrption
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartNo(partNo As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim strDescrption As String
        Dim columnToChange = "IMDSC"
        Try
            Sql = "SELECT * FROM INMSTA INNER JOIN DVINVA ON INMSTA.IMPTN = DVINVA.DVPART WHERE UCASE(IMPTN) = '" & Trim(UCase(partNo)) & "'"
            strDescrption = GetSingleDataFromDatabase(Sql, columnToChange)
            Return strDescrption
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartNoVendor(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from inmsta where trim(ucase(imptn)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetJiraPath() As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim JiraPath As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM cntrll where cnt01 = 'JIR' and trim(ucase(cnt03)) = 'PRO'"
            ds = GetDataFromDatabase(Sql)
            If ds.Tables(0).Rows.Count = 1 Then
                JiraPath = ds.Tables(0).Rows(0).Item(3).ToString()
            End If
            Return JiraPath
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetAllStatuses() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM cntrll where cnt01 = 'DSI' order by cnt02"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetPOQotaData(vendorNo As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        'burned data
        'vendorNo = "261747"
        'partNo = "CABLE14B"
        'vendorNo = "261138"
        'partNo = "99983"
        'end burned data

        Try
            Sql = "SELECT * FROM POQOTA WHERE PQVND = " & Trim(vendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%' ORDER BY PQQDTY DESC, PQQDTM DESC, PQQDTD DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetProjectStatusDescription(code As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ProjectDescStatus As String = " "
        Dim columnToChange = "CNTDE1"
        Try
            Dim CodeOk As String = Trim(UCase(code))
            Sql = "SELECT * FROM cntrll where cnt01 = 'DSI' and cnt03 = '" & CodeOk & "'"
            ProjectDescStatus = GetSingleDataFromDatabase(Sql, columnToChange)
            Return Trim(ProjectDescStatus)
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function GetWLDataByPartNo(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDWL WHERE TRIM(UCASE(PRWPTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function FillDDLUser() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT USUSER,USNAME FROM CSUSER WHERE USPTY8 = 'X' AND USPTY9 <> 'R' ORDER BY USNAME "
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function FillDDlMinorCode() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM CNTRLL WHERE CNT01 = '120' ORDER BY TRIM(CNTDE1) "
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function


#End Region

#Region "Inserts"

    Public Function InsertNewProject(projectno As String, userid As String, dtValue As DateTimePicker, strInfo As String, strName As String, ddlStatus As ComboBox, ddlUser As ComboBox) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDVLH(PRHCOD,CRUSER,CRDATE,PRDATE,PRINFO,PRNAME,PRSTAT,MOUSER,MODATE,PRPECH) VALUES 
            (" & projectno & ",'" & userid & "','" & Format(Now, "yyyy-MM-dd") & "','" & Format(dtValue.Value, "yyyy-MM-dd") & "',
            '" & Trim(strInfo) & "', '" & Trim(strName) & "','" & Left(ddlStatus.Text, 1) & "','" & userid & "',
            '" & Format(Now, "yyyy-MM-dd") & "','" & Left(Trim(ddlUser.Text), 10) & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductComment(code As String, partNo As String, comment As String, userId As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDCMH(PRHCOD,PRDPTN,PRDCCO,PRDCDA,PRDCTI,PRDCSU,USUSER) 
                    VALUES(" & Trim(code) & ",'" & Trim(partNo) & "'," & comment & ",'" & Format(DateTime.Now, "yyyy-mm-dd") & "','" & Format(DateTime.Now, "hh:mm:ss") & "',
                            'Person in charge changed','" & userId & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductCommentDetail(code As String, partNo As String, comment As String, cod_detcomment As String, messcomm As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDCMD(PRHCOD,PRDPTN,PRDCCO,PRDCDC,PRDCTX) 
                    VALUES(" & Trim(code) & ",'" & Trim(partNo) & "'," & comment & "," & cod_detcomment & ",'" & messcomm & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function InsertWishListProduct(maxItem As String, userId As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDWL(PRWCOD,CRUSER,CRDATE,PRWPTN,PRWAIN,PRWISS) 
                    VALUES(" & maxItem & ",'" & userId & "','" & Format(Now, "yyyy-mm-dd") & "','" & Trim(UCase(partNo)) & "','','No vendor assigned')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function InsertNewPOQota(partNo As String, vendorNo As String, maxValue As String, strYear As String, strMonth As String, mpnPo As String, strDay As String,
                                    strStsQuote As String, strSpace As String, strUnitCostNew As String, strMinQty As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO POQOTA (PQPTN,PQVND,PQSEQ,PQQDTY,PQQDTM,PQMPTN,PQQDTD,PQCOMM,SPACE,PQPRC,PQMIN) VALUES 
            ('" & Trim(UCase(partNo)) & "'," & Trim(vendorNo) & "," & maxValue & "," & strYear.Substring(strYear.Length - 2) & ",
            " & strMonth & ",'" & mpnPo & "'," & strDay & ",'" & strStsQuote & "','" & strSpace & "'," & strUnitCostNew & "," & strMinQty & ")"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductDetail(projectno As String, partNo As String, dtValue As DateTimePicker, userid As String, dtValue1 As DateTimePicker, userid1 As String, dtValue2 As DateTimePicker, ctpNo As String, qty As String,
                                        mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, dtValue3 As DateTimePicker, ddlStatus As String, benefits As String,
                                        comments As String, ddlUser As String, chkNew As CheckBox, dtValue4 As DateTimePicker, sampleCost As String, miscCost As String, vendorNo As String,
                                        partsToShow As String, ddlMinorCode As String, toolingCost As String, dtValue5 As DateTimePicker, dtValue6 As DateTimePicker, sampleQty As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            'dtValue6.Value = New DateTime(strDate)
            Dim chkSelection As Integer = If(chkNew.Checked = False, 0, 1)

            Sql = "INSERT INTO PRDVLD(PRHCOD,PRDPTN,PRDDAT,CRUSER,CRDATE,MOUSER,MODATE,PRDCTP,PRDQTY,PRDMFR,PRDMFR#,PRDCOS,PRDCON,PRDPO#,PODATE,PRDSTS,PRDBEN,PRDINF,PRDUSR,PRDNEW,PRDEDD,PRDSCO,PRDTTC,VMVNUM,PRDPTS,PRDMPC,PRDTCO,PRDERD,PRDPDA,PRDSQTY) 
                   VALUES (" & projectno & ",'" & Trim(UCase(partNo)) & "','" & Format(dtValue.Value, "yyyy-MM-dd") & "','" & userid & "','" & Format(dtValue1.Value, "yyyy-MM-dd") & "','" & userid & "','" & Format(dtValue2.Value, "yyyy-MM-dd") & "','" & Trim(ctpNo) & "'," & qty & ",
            '" & Trim(mfr) & "','" & Trim(mfrNo) & "'," & (unitCost) & "," & (unitCostNew) & ",'" & Trim(poNo) & "','" & Format(dtValue3.Value, "yyyy-MM-dd") & "',
            '" & Trim(ddlStatus) & "','" & Trim(benefits) & "','" & Trim(comments) & "','" & Trim(ddlUser) & "'," & chkSelection & ",'" & Format(dtValue4.Value, "yyyy-MM-dd") & "'," & sampleCost & "," & miscCost & "," & Trim(vendorNo) & ",
            '" & partsToShow & "','" & (ddlMinorCode) & "'," & toolingCost & ",'" & Format(dtValue5.Value, "yyyy-MM-dd") & "', '" & Format(dtValue6.Value, "yyyy-MM-dd") & "' ," & sampleQty & ")"

            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function InsertNewInv(strdvlocn As String, strdvpart As String, strdvmjpc As String, strdvmnpc As String, strdvindt As String, strdvunt As String, strdvslr As String, strdvohr As String,
                                    dvprmg As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "insert into dvinva(dvlocn,dvpart,dvmjpc,dvmnpc,dvindt,dvunt$,dvslr,dvohr,dvprmg) 
                    values('01','" & Trim(UCase(strdvpart)) & "','" & Trim(UCase(strdvmjpc)) & "'," & Trim(UCase(strdvmnpc)) & "','" & Format(DateTime.Now, "yy-MM-dd") & "',
                            " & strdvunt & ",'99999','99999','" & Trim(dvprmg) & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function



#End Region

#Region "Updates"

    Public Function UpdatePoQoraRow(mpnopo As String, minQty As String, unitCostNew As String, statusquote As String, insertYear As String, insertMonth As String, insertDay As String,
                                    vendorNo As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE POQOTA SET PQMPTN = '" & mpnopo & "',PQMIN  = " & minQty & ",PQPRC  = " & unitCostNew & ",PQCOMM = '" & statusquote & "',
                PQQDTY =  " & insertYear.Substring(insertYear.Length - 2) & " ,PQQDTM = " & insertMonth & " ,PQQDTD = " & insertDay & " 
                WHERE PQVND  = " & Trim(vendorNo) & " AND PQPTN  = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' " &
                " AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDetail(code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET PRDPDA = '" & Format(Now, "yyyy-MM-dd") & "' WHERE PRHCOD = " & Trim(code) & " AND PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDetail1(partstoshow As String, minorCode As String, tooCost As String, strDate1 As String, jiraTask As String, vendorNo As String, strChkSel As String,
                                        strDate2 As String, sampleCost As String, miscCost As String, userSelec As String, strDate3 As String, userid As String, tcpNo As String, sampleQty As String,
                                        qty As String, mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, strDate4 As String, status As String,
                                        benefits As String, comments As String, code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET PRDPTS = '" & partstoshow & "',PRDMPC = '" & minorCode & "',PRDTCO = " & tooCost & ",PRDERD = '" & Format(strDate1, "yyyy-mm-dd") & "', 
                    PRDJIRA = '" & Trim(jiraTask) & "', " & "VMVNUM = " & Trim(vendorNo) & ",PRDNEW = " & strChkSel & ",PRDEDD = '" & Format(strDate2, "yyyy-mm-dd") & "',
                    PRDSCO = " & sampleCost & ",PRDTTC = " & miscCost & ",PRDUSR = '" & Trim(userSelec) & "',PRDDAT = '" & Format(strDate3, "yyyy-mm-dd") & "',MOUSER = '" & userid & "',
                    MODATE = '" & Format(Now, "yyyy-mm-dd") & "',PRDCTP = '" & Trim(tcpNo) & "',PRDSQTY = " & sampleQty & ", PRDQTY = " & qty & ",PRDMFR = '" & Trim(mfr) & "',
                    PRDMFR# = '" & Trim(mfrNo) & "',PRDCOS = " & unitCost & ",PRDCON = " & unitCostNew & ",PRDPO# = '" & Trim(poNo) & "',PODATE = '" & Format(strDate4, "yyyy-mm-dd") & "',
                    PRDSTS = '" & Trim(status) & "',PRDBEN = '" & Trim(benefits) & "',PRDINF = '" & Trim(comments) & "' WHERE PRHCOD = " & Trim(code) & " AND
                    PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDetail2(partstoshow As String, minorCode As String, tooCost As String, strDate1 As String, vendorNo As String, strChkSel As String,
                                        strDate2 As String, sampleCost As String, miscCost As String, userSelec As String, strDate3 As String, userid As String, tcpNo As String, sampleQty As String,
                                        qty As String, mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, strDate4 As String,
                                        benefits As String, comments As String, code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET PRDPTS = '" & partstoshow & "',PRDMPC = '" & minorCode & "',PRDTCO = " & tooCost & ",PRDERD = '" & Format(strDate1, "yyyy-mm-dd") & "', 
                     " & "VMVNUM = " & Trim(vendorNo) & ",PRDNEW = " & strChkSel & ",PRDEDD = '" & Format(strDate2, "yyyy-mm-dd") & "',
                    PRDSCO = " & sampleCost & ",PRDTTC = " & miscCost & ",PRDUSR = '" & Trim(userSelec) & "',PRDDAT = '" & Format(strDate3, "yyyy-mm-dd") & "',MOUSER = '" & userid & "',
                    MODATE = '" & Format(Now, "yyyy-mm-dd") & "',PRDCTP = '" & Trim(tcpNo) & "',PRDSQTY = " & sampleQty & ", PRDQTY = " & qty & ",PRDMFR = '" & Trim(mfr) & "',
                    PRDMFR# = '" & Trim(mfrNo) & "',PRDCOS = " & unitCost & ",PRDCON = " & unitCostNew & ",PRDPO# = '" & Trim(poNo) & "',PODATE = '" & Format(strDate4, "yyyy-mm-dd") & "',
                    PRDBEN = '" & Trim(benefits) & "',PRDINF = '" & Trim(comments) & "' WHERE PRHCOD = " & Trim(code) & " AND
                    PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProdClosedParts(userid As String, dtvalue As Date, strUser As String, strInfo As String, strName As String, strStatus As String, code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLH SET MOUSER = '" & userid & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "',PRDATE = '" & Format(dtvalue, "yyyy-MM-dd") & "',PRPECH = '" & strUser & "',
                    PRINFO = '" & strInfo & "',PRNAME = '" & strName & "',PRSTAT = '" & strStatus & "' WHERE PRHCOD = " & code
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProdOpenParts(userid As String, dtvalue As Date, strUser As String, strInfo As String, strName As String, code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLH SET MOUSER = '" & userid & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "',PRDATE = '" & Format(dtvalue, "yyyy-MM-dd") & "',PRPECH = '" & strUser & "',
                    PRINFO = '" & strInfo & "',PRNAME = '" & strName & "' WHERE PRHCOD = " & code
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDevHeader(code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "update prdvlh set prstat = 'F' where prhcod = " & Trim(code)
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return QueryResult
        End Try
    End Function


#End Region

#Region "Utils"

    Public Function checkfieldsPoQote(partNo As String, vendorNo As String, maxValue As String, strYear As String, strMonth As String, mpnPo As String, strDay As String,
                                    strStsQuote As String, strSpace As String, strUnitCostNew As String, strMinQty As String) As String
        Dim strError As String = String.Empty

#Region "NumericFields"
        If String.IsNullOrEmpty(vendorNo) Then
            strError += "Vendor Number,"
        End If
        If String.IsNullOrEmpty(maxValue) Then
            strError += "Sequencial,"
        End If
        If String.IsNullOrEmpty(strYear) Then
            strError += "Year,"
        End If
        If String.IsNullOrEmpty(strMonth) Then
            strError += "Month,"
        End If
        If String.IsNullOrEmpty(strDay) Then
            strError += "Day,"
        End If
        If String.IsNullOrEmpty(strUnitCostNew) Then
            strError += "Unit Cost New,"
        End If
        If String.IsNullOrEmpty(strMinQty) Then
            strError += "Min Quantity,"
        End If
#End Region

        If String.IsNullOrEmpty(strError) Then
            Return ""
        Else
            Return strError
        End If

    End Function

    Public Function checkFields(projectno As String, partNo As String, dtValue As DateTimePicker, userid As String, dtValue1 As DateTimePicker, userid1 As String, dtValue2 As DateTimePicker, ctpNo As String, qty As String,
                                        mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, dtValue3 As DateTimePicker, ddlStatus As String, benefits As String,
                                        comments As String, ddlUser As String, chkNew As CheckBox, dtValue4 As DateTimePicker, sampleCost As String, miscCost As String, vendorNo As String,
                                        partsToShow As String, ddlMinorCode As String, toolingCost As String, dtValue5 As DateTimePicker, strDate As String, sampleQty As String) As String
        Dim strError As String = String.Empty

#Region "TextBoxes"

#End Region

#Region "NumericFields"

        If String.IsNullOrEmpty(projectno) Then
            strError += "Project Number,"
        End If
        If String.IsNullOrEmpty(qty) Then
            strError += "Quantity,"
        End If
        If String.IsNullOrEmpty(unitCost) Then
            strError += "Unit Cost,"
        End If
        If String.IsNullOrEmpty(unitCostNew) Then
            strError += "Unit Cost New,"
        End If
        If String.IsNullOrEmpty(sampleCost) Then
            strError += "Sample Cost,"
        End If
        If String.IsNullOrEmpty(miscCost) Then
            strError += "Misc. Cost,"
        End If
        If String.IsNullOrEmpty(vendorNo) Then
            strError += "Vendor Number,"
        End If
        If String.IsNullOrEmpty(toolingCost) Then
            strError += "Tooling Cost,"
        End If
        If String.IsNullOrEmpty(sampleQty) Then
            strError += "Sample Quantity,"
        End If


#End Region

#Region "ComboBoxes"

#End Region

#Region "SelectionFields"

#End Region


        If String.IsNullOrEmpty(strError) Then
            Return ""
        Else
            Return strError
        End If

    End Function

    Public Function getmax(table, field)
        Dim exMessage As String = " "
        Dim Sql As String = " "
        Try
            Sql = "SELECT " & field & " FROM " & table & " ORDER BY " & field & " DESC FETCH FIRST 1 ROW ONLY"
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Using ObjCmd As Odbc.OdbcCommand = New Odbc.OdbcCommand(Sql, ObjConn)
                    ObjConn.Open()
                    ObjCmd.CommandType = CommandType.Text
                    Dim QueryResult = ObjCmd.ExecuteScalar()
                    Return QueryResult
                End Using
            End Using
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Function getmaxComplex(table, field, strWhereAdd)
        Dim exMessage As String = " "
        Dim Sql As String = " "
        Try
            Sql = "SELECT " & field & " FROM " & table & " " & strWhereAdd & " ORDER BY " & field & " DESC FETCH FIRST 1 ROW ONLY"
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Using ObjCmd As Odbc.OdbcCommand = New Odbc.OdbcCommand(Sql, ObjConn)
                    ObjConn.Open()
                    ObjCmd.CommandType = CommandType.Text
                    Dim QueryResult = ObjCmd.ExecuteScalar()
                    Return QueryResult
                End Using
            End Using
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
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

    Public Function FillGrid(query As String) As Data.DataSet
        Dim exMessage As String = " "
        Try
            Dim ObjConn As New Odbc.OdbcConnection(strconnection)
            Dim dataAdapter As New Odbc.OdbcDataAdapter()
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            ObjConn.Open()
            'Sql = "SELECT COUNT(*) TFIELDS FROM PRDVLH " & strwhere
            Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
            dataAdapter = New Odbc.OdbcDataAdapter(cmd)
            dataAdapter.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            Else
                'message box warning
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

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

#End Region

#Region "Generic Methods"

    Private Function GetDataFromDatabase(query As String) As Data.DataSet
        Dim exMessage As String = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()
                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                dataAdapter.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    Return ds
                Else
                    'message box warning
                    Return Nothing
                End If
            End Using
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Function GetSingleDataFromDatabase(query As String, columnToChange As String) As String
        Dim exMessage As String = " "
        Dim DescriptionCode As String = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()
                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                dataAdapter.Fill(ds)

                Dim index = ds.Tables(0).Columns(columnToChange).Ordinal
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each RowDs In ds.Tables(0).Rows
                        DescriptionCode = ds.Tables(0).Rows(0).Item(index).ToString()
                        Exit For
                    Next
                    Return DescriptionCode
                Else
                    'message box warning
                    Return Nothing
                End If
            End Using
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Function InsertDataInDatabase(query As String) As Integer
        Dim exMessage As String = " "
        Dim result As Integer = -1
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()
                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                result = dataAdapter.Fill(ds)
                Return result
            End Using
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Function UpdateDataInDatabase(query As String) As String
        Dim exMessage As String = " "
        'Dim result As Integer = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()
                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                dataAdapter.Fill(ds)
            End Using
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

#End Region


    Public Function GetTestData(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from inmsta where trim(ucase(imptn)) = '" & Trim(UCase(partNo)) & "' fetch first 10 row only"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.HResult.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
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






End Class
