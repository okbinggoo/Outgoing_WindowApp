Imports System.IO
Public Class ReadText_module
    Public Shared Function readConfigFile() As Hashtable
        'Keep data as Key,Value
        Dim path As String = AppDomain.CurrentDomain.BaseDirectory
        Dim configPath As String : configPath = path + "config\PathServer.txt"
        Dim hashMap As Hashtable = New Hashtable
        Using r As StreamReader = New StreamReader(configPath)
            Dim counter As Integer = 0
            ' Read first line.
            Dim line As String : line = r.ReadLine
            ' Loop over each line in file, While list is Not Nothing.
            Do While (Not line Is Nothing)
                If (line.Contains(";")) Then
                    Dim array(2) As String : array = line.Split(";")
                    hashMap.Add(array(0), array(1))
                End If
                line = r.ReadLine
            Loop
        End Using
        Return hashMap
    End Function
    Sub ReadConfigTxt(ByVal thisPath As String, ByVal StationName As String)
        Dim myReader As StreamReader
        Dim CurrentLine As String
        Dim strValue() As String
        Dim i As Integer

        i = 1
        thisPath = thisPath + "\Config\PathServer.txt"
        myReader = New StreamReader(thisPath, System.Text.UnicodeEncoding.Default)
        While myReader.Peek <> -1
            CurrentLine = myReader.ReadLine
            strValue = CurrentLine.Split(";")
            If i = 1 Then       'Oracle Server
                GlobalVariables.Host = Trim(strValue(0))
            ElseIf i = 2 Then   'Database Name
                GlobalVariables.Database = Trim(strValue(0))
            ElseIf i = 3 Then   'Port
                GlobalVariables.Port = Trim(strValue(0))
            ElseIf i = 4 Then   'UserName
                GlobalVariables.UserName = Trim(strValue(0))
            ElseIf i = 5 Then   'Password
                GlobalVariables.Password = Trim(strValue(0))
            ElseIf i = 6 Then   'FTP Server-> IP Address
                GlobalVariables.FTP_SVR = Trim(strValue(0))

            ElseIf i = 7 Then   'FTP Back Up Text file Path
                'GlobalVariables.FTP_SVR_BK = Trim(Trim(strValue(0)) + StationName + "\")
                GlobalVariables.FTP_SVR_BK = Trim(strValue(0))
            ElseIf i = 8 Then   'SQL Command for QUERY DATA
                GlobalVariables.SqlCMD = Trim(strValue(0))
            ElseIf i = 9 Then   'IM Host
                GlobalVariables.IM_Host = Trim(strValue(0))
            ElseIf i = 10 Then   'IM Server
                GlobalVariables.IM_server = Trim(strValue(0))
            ElseIf i = 11 Then   'AS/400 User
                GlobalVariables.IM_user = Trim(strValue(0))
            ElseIf i = 12 Then   'AS/400 Pass
                GlobalVariables.IM_pass = Trim(strValue(0))

            ElseIf i = 13 Then   'Text File To Excel
                GlobalVariables.FileRecord = Trim(strValue(0))
            ElseIf i = 14 Then   'Text File To Excel
                GlobalVariables.FileRecord_filename = Trim(strValue(0))

            ElseIf i = 15 Then   'USERNAME Local LS
                GlobalVariables.UserNameLogisic = Trim(strValue(0))
            ElseIf i = 16 Then   'PASSWORD Local LS
                GlobalVariables.PasswordLogisic = Trim(strValue(0))
            ElseIf i = 17 Then   'Oracle Server LS
                GlobalVariables.HostLS = Trim(strValue(0))
            ElseIf i = 18 Then   'DATABASE NAME LS
                GlobalVariables.DatabaseLS = Trim(strValue(0))
            ElseIf i = 19 Then   'Port LS
                GlobalVariables.PortLS = Trim(strValue(0))
            ElseIf i = 20 Then   'USERNAME LS
                GlobalVariables.UserNameLS = Trim(strValue(0))
            ElseIf i = 21 Then   'PASSWORD LS
                GlobalVariables.PasswordLS = Trim(strValue(0))
            ElseIf i = 22 Then   'To
                GlobalVariables.Tomail = Trim(strValue(0))
            ElseIf i = 23 Then   'CC
                GlobalVariables.CCmail = Trim(strValue(0))
            ElseIf i = 24 Then   'CC
                GlobalVariables.PathPDF = Trim(strValue(0))
            ElseIf i = 25 Then   'CC
                GlobalVariables.PathZip = Trim(strValue(0))
            ElseIf i = 26 Then   'CC
                GlobalVariables.PathZipBK = Trim(strValue(0))

            End If
            i = i + 1
        End While

        myReader.Close()
        myReader = Nothing
    End Sub
End Class
