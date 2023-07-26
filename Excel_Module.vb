Imports System.IO

Imports System.Net.Mail
Public Class Excel_Module

    Private Property oExcel As Object
    Private Property ExcelDoc As Object
    Private Property FS As Object


    'function Run Excelfile
    'paremeter : factory code & customer code
    'result : run excelfile follow register in db
    Function runExcel(ByVal PathForm As String, ByVal Style As String)
        Dim pathExcel As String
        Dim flag As String
        oExcel = Nothing
        oExcel = CreateObject("Excel.Application")
        oExcel.Visible = True
        oExcel.AutomationSecurity = 1
        oExcel.WindowState = 1

        '''''''''''''''''''''''''''''''''''''
        ' fix filename standard.xlsm
        '''''''''''''''''''''''''''''''''''''
        '          If pathExcel <> "" Then
        '             pathExcel = "standard.xlsm"
        '          End If
        '''''''''''''''''''''''''''''''''''''
        'pathExcel = PathForm + GlobalVariables.StrStationName + "\" + GlobalVariables.H40Code + "form\" + GlobalVariables.FormNo + ".xlsm"
        pathExcel = PathForm + "STA\" + GlobalVariables.H40Code + "form\" + GlobalVariables.FormNo + ".xlsm"
        Dim fFile As New FileInfo(pathExcel)

        If fFile.Exists Then
            ExcelDoc = oExcel.Workbooks.Open(pathExcel, False, True)
            oExcel = Nothing
            flag = "OK"
        Else
            flag = "Not found Microsoft Excel :Print form " + GlobalVariables.FormNo
            oExcel.Visible = False

            'MsgBox("Not found Microsoft Excel :Print form !!! ")
            ' Call sendmail1("Error Outgoing data sheet create PDF file !!!!!!!", " Not found Excel form !!! Invoice : " + GlobalVariables.Invoice + " Item : " + GlobalVariables.ItemNo + " Folder : " + GlobalVariables.H40Code + " Form : " + GlobalVariables.FormNo, GlobalVariables.Tomail, GlobalVariables.CCmail)

        End If
        Return flag

    End Function

    Sub sendmail1(ByVal subject As String, ByVal body As String, ByVal mail As String, ByVal cc As String)
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
