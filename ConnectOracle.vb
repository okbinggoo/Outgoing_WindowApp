Imports Oracle.DataAccess.Client
Public Class ConnectOracle
    Public Function ConnectOracle(ByVal Host As String, ByVal Databasez As String, ByVal UserName As String, ByVal Password As String, ByVal Port As String) As OracleConnection
        Dim objConn As New OracleConnection

        '        On Error GoTo L_End
        '        Dim strConOracle, oConOracle, oRsOracle

        '        strConOracle = "Driver={Microsoft ODBC for Oracle}; " &
        '                "CONNECTSTRING=(DESCRIPTION=" &
        '                "(ADDRESS=(PROTOCOL=TCP)" &
        '                "(HOST=" & Host & ")(PORT=" & Port & "))" &
        '                "(CONNECT_DATA=(SERVICE_NAME=" & Databasez & "))); uid=" & UserName & " ;pwd=" & Password & ";"

        '        ConnectOracle = CreateObject("ADODB.Connection")
        '        ConnectOracle.Open(strConOracle)

        '        Exit Function
        'L_End:
        '        MsgBox(CStr(Err.Number) + ":" + Err.Description, vbCritical, "Warning")
        objConn.ConnectionString = "Data Source=" & Host & "/" & Databasez & ";" & "Persist Security Info=True;" & "User ID=" & UserName & ";Password=" & Password
        objConn.Open()
        Return objConn
    End Function

    Public Function SelectData(ByVal conn As OracleConnection, ByVal sqlCmd As String) As DataTable

        Dim dt As New DataTable
        Dim da As OracleDataAdapter


        da = New OracleDataAdapter(sqlCmd, conn)
        da.Fill(dt)


        Return dt
    End Function
    'To end connection
    Public Sub EndConn(ByVal objConn As OracleConnection)
        objConn.Close()
        objConn.Dispose()
    End Sub
End Class
