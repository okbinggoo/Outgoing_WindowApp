Imports System.IO
Imports System.Net.Mail
Imports System.IO.Compression
Imports System.Threading
Imports ADODB
Imports Oracle.DataAccess.Client


Public Class Form1


    Public objRead As New ReadText_module
    Public objFunc As New GetQueryOG
    Public objWrite As New WriteText_Module
    Public objExcel As New Excel_Module
    Public r As ADODB.Recordset
    Public c As New ADODB.Connection
    Public cm As New ADODB.Command



    Private Sub main(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim TNSPath As String
        Dim filenameTNS As String = Format(Today, "yyyyMMdd")

        ' Dim filenameTNS = oDate.Year & oDate.Month & oDate.Day
        TNSPath = Application.StartupPath
        TNSPath = TNSPath + "\ErrorTNS\" + "ERROR_" & filenameTNS & ".txt"
        Try
            Me.Show()
            ProgressBar1.Visible = True
            Dim percent As Integer
            ' percent = 20
            ProgressBar1.Value = 0
            ' ProgressBar1.CreateGraphics().DrawString(percent.ToString() & "%", New Font("Arial", CSng(8.25), FontStyle.Regular), Brushes.Black, New PointF(ProgressBar1.Width / 2 - 10, ProgressBar1.Height / 2 - 7))



            'Me.Show()

            'READ Line
            Dim SQlArray(10) As String
            Dim Num As Integer
            'Excel
            Dim PathFile_name As String
            Dim PathFile As String
            Dim PathFolder As String
            Dim BKPath As String
            Dim PathExcel As String

            Dim H40Folder As String
            Dim StationFolder As String
            'connection
            Dim objConn As New System.Data.OracleClient.OracleConnection
            Dim objCmd As New System.Data.OracleClient.OracleCommand
            Dim dtAdapter As New System.Data.OracleClient.OracleDataAdapter
            Dim objOracle As New ConnectOracle

            Dim db As New DataTable
            Dim textString As String
            Dim TextName As String

            Dim Flag_error As String
            Dim ErrorExcelPath As String
            Dim kill_bat As String

            Dim PathExcel_manual As String
            Dim Pathfile_manual As String
            Dim Pathfilename_manual As String

            Dim PathFile_auto As String
            Dim PathFilename_auto As String
            Dim PathExcel_auto As String
            Dim ErrorExcelPath_auto As String
            'Dim statuss As String
            Dim ErrorExcelPath_Manual As String

            ErrorExcelPath_Manual = Application.StartupPath

            ErrorExcelPath_auto = Application.StartupPath

            kill_bat = Application.StartupPath
            kill_bat = kill_bat + "\killexcel.bat"


            H40Folder = GlobalVariables.H40Code + "\form\"
            StationFolder = GlobalVariables.StrStationName + "\"
            PathFile_auto = Application.StartupPath
            PathFilename_auto = Application.StartupPath
            PathExcel_manual = Application.StartupPath
            Pathfilename_manual = Application.StartupPath
            Pathfile_manual = Application.StartupPath




            PathExcel_auto = PathFile_auto + "\Form\" + GlobalVariables.PathOfPart
            PathFile_auto = PathFile_auto + "\Form\STA\" + GlobalVariables.FileRecord
            PathFilename_auto = PathFilename_auto + "\Form\STA\" + GlobalVariables.FileRecord_filename
            ErrorExcelPath_auto = ErrorExcelPath_auto + "\Form\STA\"

            PathExcel_manual = PathExcel_manual + "\Form_manual\"
            Pathfile_manual = Pathfile_manual + "\Form_manual\STA\" + GlobalVariables.FileRecord
            Pathfilename_manual = Pathfilename_manual + "\Form_manual\STA\" + GlobalVariables.FileRecord_filename
            ErrorExcelPath_Manual = ErrorExcelPath_Manual + "\Form_manual\STA\"
            Dim SqlCom As String


            Dim j As Integer = 0

Reconnect:
            Try

                'Step : 1 Get data
                ProgressBar1.Value = 20
                GlobalVariables.Conn = objOracle.ConnectOracle(GlobalVariables.IM_Host, GlobalVariables.IM_server, GlobalVariables.IM_user, GlobalVariables.IM_pass, GlobalVariables.Port)

                SqlCom = "  Select IMFR_UD_FAC as fac , IMFR_UD_INSP_NO as insp, IMFR_UD_CUST_CD as CustCD , IMFR_UD_MURATA_TYPE as MurataType ,
                            IMFR_UD_INV_NO as invoice , IMFR_UD_ITEM_NO as item ,
                            IMFR_UD_H40 as PRODUCTCODE_H40, IMFR_UD_FROM as OUTGOINGINSP_FORM  , IMFR_UD_CUST_NAME as cust_name ,  IMFR_UD_FILENAME as file_name , IMFR_UD_NUM_TRANSFER as outgingflag,
                            to_char(IMFR_UD_SHIPPING_DATE,'yyyymmdd') as shipdate  ,IMFR_UD_CUST_PURCH_NO as cust_purch , IMFR_UD_CUST_PART as cust_part , IMFR_UD_CARD_NO as card_no, IMFR_UD_CUST_PART_NAME as cust_part_name,
                            IMFR_UD_MURATA_TYPE_ALLOCATE as murata_allocate , IMFR_UD_SHIP_QTY AS shipQty, IMFR_UD_UPADTE_FLAG as status
                            from imfr_ut_mtloutgoing_g where  IMFR_UD_INV_NO =  '2323006126'  and IMFR_UD_ITEM_NO = '001'  order by IMFR_UD_SHIPPING_DATE,IMFR_UD_INV_NO,IMFR_UD_ITEM_NO  asc "
                '(IMFR_UD_UPADTE_FLAG = 'MANUAL' or  IMFR_UD_UPADTE_FLAG = 'UPDATE') and 

                db = objOracle.SelectData(GlobalVariables.Conn, SqlCom)

            Catch ex As Exception
                j += 1
                If j > 5 Then
                    GoTo Endselect
                End If
                GoTo Reconnect

            End Try


            If db.Rows.Count > 0 Then
                For Each row As DataRow In db.Rows




                    Thread.Sleep(5000)
                    'For cn = 0 To GlobalVariables.RS.Fields.Count
                    GlobalVariables.Fcty_CD = row.Item("fac").ToString
                    GlobalVariables.CustCD = row.Item("CustCD").ToString
                    GlobalVariables.Inspection = row.Item("insp").ToString
                    GlobalVariables.Invoice = row.Item("invoice").ToString
                    GlobalVariables.ItemNo = row.Item("item").ToString
                    GlobalVariables.MurataType = row.Item("MurataType").ToString
                    GlobalVariables.H40Code = row.Item("PRODUCTCODE_H40").ToString
                    GlobalVariables.FormNo = row.Item("OUTGOINGINSP_FORM").ToString
                    GlobalVariables.CustName = row.Item("cust_name").ToString
                    GlobalVariables.fileName = row.Item("file_name").ToString
                    GlobalVariables.outging_flag = row.Item("outgingflag").ToString
                    GlobalVariables.shipdate = row.Item("shipdate").ToString
                    GlobalVariables.cust_purch = row.Item("cust_purch").ToString
                    GlobalVariables.cust_part = row.Item("cust_part").ToString
                    GlobalVariables.card_no = row.Item("card_no").ToString
                    GlobalVariables.cust_part_name = row.Item("cust_part_name").ToString
                    GlobalVariables.murata_allocate = row.Item("murata_allocate").ToString
                    GlobalVariables.ShipQty = row.Item("shipQty").ToString
                    GlobalVariables.statuss = row.Item("status").ToString

                    If GlobalVariables.statuss = "MANUAL" Then
                        PathExcel = PathExcel_manual
                        PathFile = Pathfile_manual
                        PathFile_name = Pathfilename_manual
                        ErrorExcelPath = ErrorExcelPath_Manual
                    Else
                        PathExcel = PathExcel_auto
                        PathFile = PathFile_auto
                        PathFile_name = PathFilename_auto
                        ErrorExcelPath = ErrorExcelPath_auto

                    End If


                    TextName = GlobalVariables.fileName + ";" + GlobalVariables.CustName + ";" + GlobalVariables.outging_flag + ";" +
                        GlobalVariables.shipdate + ";" + GlobalVariables.cust_purch + ";" + GlobalVariables.cust_part + ";" + GlobalVariables.card_no + ";" + GlobalVariables.cust_part_name + ";" + GlobalVariables.murata_allocate
                    'db.MoveNext()
                    '1.1 Find Query for each product
                    'Sql = ""
                    GlobalVariables.SQLStr1 = ""
                    GlobalVariables.SQLStr2 = ""
                    GlobalVariables.SQLStr3 = ""
                    GlobalVariables.SQLStr4 = ""
                    GlobalVariables.SQLStr5 = ""
                    GlobalVariables.SQLStr6 = ""
                    GlobalVariables.SQLStr7 = ""
                    GlobalVariables.SQLStr8 = ""
                    GlobalVariables.SQLStr9 = ""
                    GlobalVariables.SQLStr10 = ""
                    'Step2 : Find Query OG DB
                    ProgressBar1.Value = 40

                    Call objFunc.FindQuery()


                    '1.2 Make Data Sting for Each Record
                    SQlArray(1) = GlobalVariables.SQLStr1
                    SQlArray(2) = GlobalVariables.SQLStr2
                    SQlArray(3) = GlobalVariables.SQLStr3
                    SQlArray(4) = GlobalVariables.SQLStr4
                    SQlArray(5) = GlobalVariables.SQLStr5
                    SQlArray(6) = GlobalVariables.SQLStr6
                    SQlArray(7) = GlobalVariables.SQLStr7
                    SQlArray(8) = GlobalVariables.SQLStr8
                    SQlArray(9) = GlobalVariables.SQLStr9
                    SQlArray(10) = GlobalVariables.SQLStr10
                    GlobalVariables.TextRecord = ";"
                    Num = 1
                    GlobalVariables.InspectionRecords = GlobalVariables.InspectionRecords + 1
                    'Step3 : Prepear Data from Query OG DB for write text file
                    ProgressBar1.Value = 60
                    While SQlArray(Num) <> ""
                        textString = objFunc.Get_BaseData(GlobalVariables.Inspection, SQlArray(Num))
                        If Num = 1 Then
                            If textString = "" Or textString = ";|" Or textString = "|" Then
                                'MsgBox("กรุณาตรวจสอบไม่พบ Inspection :" + GlobalVariables.Inspection)
                                Call sendmail("Error Outgoing data sheet not found OG inspection data!!!!!!!", "Please check data " + GlobalVariables.Fcty_CD + " Inspection :" + GlobalVariables.Inspection + " Invoice: " + GlobalVariables.Invoice + " Item:" + GlobalVariables.ItemNo + " H40 :" + GlobalVariables.H40Code + " Query SQL : " + GlobalVariables.SQLStr1, GlobalVariables.Tomail, GlobalVariables.CCmail)
                                Flag_error = Update_Error("NOT FOUND INSPECTION DATA IN OG DB", GlobalVariables.Invoice, GlobalVariables.ItemNo, GlobalVariables.Inspection, "ERROR")
                                If Flag_error <> "OK" Then
                                    Call objWrite.WriteTextFile(Flag_error, TNSPath)
                                    GoTo Reconnect
                                End If
                                GlobalVariables.InspectionRecords = GlobalVariables.InspectionRecords - 1


                                GoTo ErrLabel
                            End If
                        End If
                        GlobalVariables.TextRecord = GlobalVariables.TextRecord + textString
                        Num = Num + 1
                    End While


                    If GlobalVariables.TextRecord = ";" Then
                        GlobalVariables.TextRecord = ";" + GlobalVariables.Invoice + ";" + GlobalVariables.ItemNo + ";" + GlobalVariables.Inspection + ";" + GlobalVariables.CustCD + ";"
                    End If
                    'Step4 : Write text file
                    ProgressBar1.Value = 80

                    Call objWrite.WriteTextFile(GlobalVariables.TextRecord, PathFile)
                    Call objWrite.WriteTextFile_FileName(TextName, PathFile_name)


                    '***** Run Excel Form each product
                    'Step5 : run excel
                    ProgressBar1.Value = 100

                    If GlobalVariables.FormNo = "" Then
                        Flag_error = Update_Error("NOT FOUND FORM No. PLEASE CHECK REGISTER FORM FOR CUSTOMER ", GlobalVariables.Invoice, GlobalVariables.ItemNo, GlobalVariables.Inspection, "ERROR")
                        If Flag_error <> "OK" Then
                            Call objWrite.ClearTextFile(PathFile)
                            Call objWrite.ClearTextFile(PathFile_name)
                            Call objWrite.WriteTextFile(Flag_error, TNSPath)
                            GoTo Reconnect
                        End If
                        Call sendmail("Error Outgoing data sheet create PDF file !!!!!!!", "Pls check and Register Form " + GlobalVariables.Fcty_CD + " for INVOICE :" + GlobalVariables.Invoice + " ITEM : " + GlobalVariables.ItemNo + " INSPECTION :" + GlobalVariables.Inspection + " H40 :" + GlobalVariables.H40Code + " Query SQL : " + GlobalVariables.SQLStr1, GlobalVariables.Tomail, GlobalVariables.CCmail)

                    Else
                        If GlobalVariables.InspectionRecords > 0 Then
                            Dim runExcel = objExcel.runExcel(PathExcel, GlobalVariables.FormNo)
                            If runExcel <> "OK" Then
                                Flag_error = Update_Error(runExcel, GlobalVariables.Invoice, GlobalVariables.ItemNo, GlobalVariables.Inspection, "ERROR")
                                If Flag_error <> "OK" Then
                                    Call objWrite.ClearTextFile(PathFile)
                                    Call objWrite.ClearTextFile(PathFile_name)
                                    Call objWrite.WriteTextFile(Flag_error, TNSPath)
                                    GoTo Reconnect
                                End If
                                Call sendmail("Error Outgoing data sheet create PDF file !!!!!!!", " Not found Excel form !!! " + GlobalVariables.Fcty_CD + " Invoice : " + GlobalVariables.Invoice + " Item : " + GlobalVariables.ItemNo + " H40 : " + GlobalVariables.H40Code + " Form : " + GlobalVariables.FormNo + " Query SQL : " + GlobalVariables.SQLStr1, GlobalVariables.Tomail, GlobalVariables.CCmail)
                            Else
                                Flag_error = Update_Error("COMPLETE", GlobalVariables.Invoice, GlobalVariables.ItemNo, GlobalVariables.Inspection, "COMPLETE")
                                If Flag_error <> "OK" Then
                                    Call objWrite.ClearTextFile(PathFile)
                                    Call objWrite.ClearTextFile(PathFile_name)
                                    Call objWrite.WriteTextFile(Flag_error, TNSPath)
                                    GoTo Reconnect
                                End If
                            End If

                        End If
                        ''Step6 : clear text
                        ProgressBar1.Value = 120


                        Call objWrite.ClearTextFile(PathFile)
                        Call objWrite.ClearTextFile(PathFile_name)
                        Call read_textFileUpdateErrorexelform(ErrorExcelPath, GlobalVariables.FileTNSerror)
                        GoTo EndLoop
                    End If

ErrLabel:
                    Call objWrite.ClearTextFile(PathFile)
                    Call objWrite.ClearTextFile(PathFile_name)
                    Call read_textFileUpdateErrorexelform(ErrorExcelPath, GlobalVariables.FileTNSerror)
EndLoop:
                    Thread.Sleep(3000)

                    System.Diagnostics.Process.Start(kill_bat)

                Next row

            End If
Endselect:
            Call objWrite.ClearTextFile(PathFile)
            Call objWrite.ClearTextFile(PathFile_name)
            Call read_textFileUpdateErrorexelform(ErrorExcelPath, GlobalVariables.FileTNSerror)
        Catch ex As Exception

            Call objWrite.WriteTextFile(ex.ToString, TNSPath)
        End Try

        Me.Close()


    End Sub



    Function Update_Error(ByVal message As String, ByVal invoice As String, ByVal item As String, ByVal inspection As String, ByVal status As String)
        Dim text_string As String
        Dim db As New DataTable
        Dim sql As String
        Dim objOracle As New ConnectOracle
        Dim i As Integer = 0

reconnect:
        Try
            GlobalVariables.Conn = objOracle.ConnectOracle(GlobalVariables.IM_Host, GlobalVariables.IM_server, GlobalVariables.IM_user, GlobalVariables.IM_pass, GlobalVariables.Port)
            sql = "Update IMFR_UT_MTLOUTGOING_g set IMFR_UD_UPADTE_FLAG = '" & status & "' , IMFR_UD_ERROR = '" & message & "' 
               where  IMFR_UD_INV_NO = '" & invoice & "' and  IMFR_UD_ITEM_NO = '" & item & "' and IMFR_UD_INSP_NO = '" & inspection & "' "

            db = objOracle.SelectData(GlobalVariables.Conn, sql)

        Catch ex As Exception
            i += 1
            If i > 5 Then
                GoTo Endreconnect
            End If
            GoTo reconnect

        End Try
Endreconnect:
        If i > 0 Then
            text_string = "ERROR" & "|" & invoice & "|" & item & "|" & inspection & "|" & "TNS_ERROR_UPDATE_FLAG"
        Else
            text_string = "OK"
        End If
        objOracle.EndConn(GlobalVariables.Conn)
        Return text_string


    End Function

    Function read_textFileUpdateErrorexelform(ByVal filePath As String, ByVal textname As String)
        Dim objOracle As New ConnectOracle
        Dim db As New DataTable
        Dim i As Integer = 0

        Dim CurrentLine As String
        Dim myReader As StreamReader
        Dim ExelPath As String
        Dim strValue() As String

        Dim filename_pdf As String
        Dim sql As String
        Dim message As String

        'ExelPath = Application.StartupPath
        ExelPath = filePath + textname

        Dim di As New IO.DirectoryInfo(filePath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.txt")

        If aryFi.Length > 0 Then
            myReader = New StreamReader(ExelPath)
            While myReader.Peek <> -1
                CurrentLine = myReader.ReadLine
                strValue = CurrentLine.Split(";")

                filename_pdf = Trim(strValue(0))
                message = Trim(strValue(1))
                message = message.Replace("'", " ")


reconnect:
                Try
                    GlobalVariables.Conn = objOracle.ConnectOracle(GlobalVariables.IM_Host, GlobalVariables.IM_server, GlobalVariables.IM_user, GlobalVariables.IM_pass, GlobalVariables.Port)
                    sql = "Update IMFR_UT_MTLOUTGOING_g set IMFR_UD_UPADTE_FLAG = 'ERROR' , IMFR_UD_ERROR = '" & message & "' 
                           where  IMFR_UD_FILENAME = '" & filename_pdf & "'  "
                    db = objOracle.SelectData(GlobalVariables.Conn, sql)



                Catch ex As Exception
                    i += 1
                    If i > 5 Then
                        GoTo Endreconnect
                    End If
                    GoTo reconnect

                End Try
Endreconnect:
            End While
            myReader.Close()
            objOracle.EndConn(GlobalVariables.Conn)
            Call objWrite.ClearTextFile(ExelPath)

        End If


    End Function

    Sub sendmail(ByVal subject As String, ByVal body As String, ByVal mail As String, ByVal cc As String)
        'Dim mailInstance As MailMessage = New MailMessage("connect-admin@murata.com", "xx", subject, body)
        Dim mailInstance As MailMessage = New MailMessage
        mailInstance.From = New MailAddress("no-reply@murata.com", "System create Outgoing Data Sheet PDF file")
        mailInstance.Subject = subject
        mailInstance.IsBodyHtml = True
        'mailInstance.Body = "<h1 style=\text-align:center;\>Test</h1>"
        mailInstance.Body = body
        If mail <> "" Then
            mailInstance.To.Add(mail)
        End If
        If cc <> "" Then
            mailInstance.CC.Add(cc)
        End If
        mailInstance.Priority = MailPriority.High
        mailInstance.IsBodyHtml = True
        'mailInstance.Attachments.Add(New Attachment("filename")) 'Optional
        Dim mailSenderInstance As SmtpClient = New SmtpClient("172.24.128.80", 25) '25 is the port of the SMTP host
        mailSenderInstance.Credentials = New System.Net.NetworkCredential("LoginAccout", "Password")
        mailSenderInstance.Send(mailInstance)
        mailInstance.Dispose() 'Please remember to dispose this object

        'sendmail = True

    End Sub


End Class
