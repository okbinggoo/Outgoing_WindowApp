Imports ADODB
Imports Oracle.DataAccess.Client
Imports System.Data.OracleClient
Public Class GetQueryOG
    Private Property ExcelDoc As Object
    Dim objOracle As New ConnectOracle
    Sub FindQuery()
        Call objOracle.EndConn(GlobalVariables.Conn)
        GlobalVariables.Conn = objOracle.ConnectOracle(GlobalVariables.Host, GlobalVariables.Database, GlobalVariables.UserName, GlobalVariables.Password, GlobalVariables.Port)
        Dim RS As New DataTable
        Dim i As Integer
        Dim N As Integer
        Dim sqlOutgoing As String
        N = GlobalVariables.RS.Fields.Count
        sqlOutgoing = GlobalVariables.SqlCMD
        GlobalVariables.Seperator = ""
        sqlOutgoing = Replace(sqlOutgoing, "#FCTY#", GlobalVariables.Fcty_CD)
        sqlOutgoing = Replace(sqlOutgoing, "#H40#", GlobalVariables.H40Code)
        RS = objOracle.SelectData(GlobalVariables.Conn, sqlOutgoing)
        'GlobalVariables.RS = GlobalVariables.Conn.Execute(GlobalVariables.SqlCMD)
        GlobalVariables.FlagKeepInv = ""

        If RS.Rows.Count > 0 Then
            For Each row As DataRow In RS.Rows
                GlobalVariables.Database1 = row.Item("DB_Name").ToString
                GlobalVariables.UserName1 = row.Item("User").ToString
                GlobalVariables.Password1 = row.Item("PWD").ToString
                'If IsDBNull(row.Item("SQL_CMD1").ToString) = False Then
                GlobalVariables.SQLStr1 = row.Item("SQL_CMD1").ToString
                GlobalVariables.SQLStr2 = row.Item("SQL_CMD2").ToString
                GlobalVariables.SQLStr3 = row.Item("SQL_CMD3").ToString
                GlobalVariables.SQLStr4 = row.Item("SQL_CMD4").ToString
                GlobalVariables.SQLStr5 = row.Item("SQL_CMD5").ToString
                GlobalVariables.SQLStr6 = row.Item("SQL_CMD6").ToString
                GlobalVariables.SQLStr7 = row.Item("SQL_CMD7").ToString
                GlobalVariables.SQLStr8 = row.Item("SQL_CMD8").ToString
                GlobalVariables.SQLStr9 = row.Item("SQL_CMD9").ToString
                GlobalVariables.SQLStr10 = row.Item("SQL_CMD10").ToString
                GlobalVariables.Seperator = row.Item("Seperate").ToString 'Seperate
                GlobalVariables.FlagKeepInv = row.Item("KeepInv").ToString
                GlobalVariables.FlagMurataType = row.Item("KeepMura").ToString
                'End If
            Next row

            'GlobalVariables.Database1 = GlobalVariables.RS.Fields("DB_Name").Value
            'GlobalVariables.UserName1 = GlobalVariables.RS.Fields("User").Value
            'GlobalVariables.Password1 = GlobalVariables.RS.Fields("PWD").Value

            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD1" + CStr(i + 1)).Value) = False Then
            '    GlobalVariables.SQLStr1 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 1)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 2)).Value) = False Then
            '    GlobalVariables.SQLStr2 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 2)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 3)).Value) = False Then
            '    GlobalVariables.SQLStr3 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 3)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 4)).Value) = False Then
            '    GlobalVariables.SQLStr4 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 4)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 5)).Value) = False Then
            '    GlobalVariables.SQLStr5 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 5)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 6)).Value) = False Then
            '    GlobalVariables.SQLStr6 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 6)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 7)).Value) = False Then
            '    GlobalVariables.SQLStr7 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 7)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 8)).Value) = False Then
            '    GlobalVariables.SQLStr8 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 8)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 9)).Value) = False Then
            '    GlobalVariables.SQLStr9 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 9)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 10)).Value) = False Then
            '    GlobalVariables.SQLStr10 = GlobalVariables.RS.Fields("SQL_CMD" + CStr(i + 10)).Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("Seperate").Value) = False Then
            '    GlobalVariables.Seperator = GlobalVariables.RS.Fields("Seperate").Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("KeepInv").Value) = False Then
            '    GlobalVariables.FlagKeepInv = GlobalVariables.RS.Fields("KeepInv").Value
            'End If
            'If IsDBNull(GlobalVariables.RS.Fields("KeepMura").Value) = False Then
            '    GlobalVariables.FlagMurataType = GlobalVariables.RS.Fields("KeepMura").Value
            'End If

        End If

        'GlobalVariables.RS.Close()
        'GlobalVariables.RS = Nothing
        Call objOracle.EndConn(GlobalVariables.Conn)
    End Sub
    Function Get_BaseData(ByVal Insp As String, ByVal Sql As String) As String
        Dim objConn As New System.Data.OracleClient.OracleConnection
        Dim objCmd As New System.Data.OracleClient.OracleCommand
        Dim dtAdapter As New System.Data.OracleClient.OracleDataAdapter

        Dim ds As New DataSet
        Dim dt As DataTable
        Dim strSQL As String
        Dim teststring As String
        Dim J As Integer = 0
        teststring = ""

reconnect:
        Try
            GlobalVariables.Conn = objOracle.ConnectOracle(GlobalVariables.Host, GlobalVariables.Database1, GlobalVariables.UserName1, GlobalVariables.Password1, GlobalVariables.Port)
            'strConnString = "Data Source=(DESCRIPTION=" _
            '                + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + GlobalVariables.Host + ")(PORT=" + GlobalVariables.Port + ")))" _
            '                + "(CONNECT_DATA=(SEVER=SHARED)(SERVICE_NAME=" + GlobalVariables.Database1 + ")));" _
            '                + "User Id=" + GlobalVariables.UserName1 + ";Password=" + GlobalVariables.Password1 + ";"

            Sql = Replace(Sql, "#INSP#", Insp)
            Sql = Replace(Sql, "#CUSTCD#", GlobalVariables.CustCD)
            Sql = Replace(Sql, "#MURATATYPE#", GlobalVariables.MurataType)
            Sql = Replace(Sql, "#INVOICE#", GlobalVariables.Invoice)
            Sql = Replace(Sql, "#ITEM#", GlobalVariables.ItemNo)
            strSQL = Sql

            dt = objOracle.SelectData(GlobalVariables.Conn, strSQL)
        Catch ex As Exception
            J += 1
            If J > 5 Then
                GoTo Endreconnect
            End If
            GoTo reconnect

        End Try
        'objConn.ConnectionString = strConnString
        'With (objCmd)
        '    .Connection = objConn
        '    .CommandText = strSQL
        '    .CommandType = CommandType.Text
        'End With
        'dtAdapter.SelectCommand = objCmd

        'dtAdapter.Fill(ds)
        'dt = ds.Tables(0)

        'dtAdapter = Nothing
        'objConn.Close()
        'objConn = Nothing
Endreconnect:
        Dim N As Integer
        Dim i As Integer
        Dim CurRow As Integer
        Dim Row As DataRow
        N = dt.Columns.Count

        If dt.Rows.Count > 0 Then
            For CurRow = 0 To dt.Rows.Count - 1
                Row = dt.Rows(CurRow)

                For i = 0 To N - 1
                    If GlobalVariables.FlagMurataType = "Y" Then
                        'Find Murata Type
                        If IsDBNull(Row.Item(i)) = False And CurRow = 0 And i = 2 Then ' Lock Only Murata Type
                            GlobalVariables.MurataType = CStr(Row.Item(i))
                        End If
                    End If

                    If IsDBNull(Row.Item(i)) = False Then
                        teststring = teststring + Trim(CStr(Row.Item(i))) + ";"
                    Else
                        teststring = teststring + ";"
                    End If
                Next i
            Next CurRow
        End If
        objOracle.EndConn(GlobalVariables.Conn)

        If GlobalVariables.Seperator <> "" Then
            teststring = teststring + GlobalVariables.Seperator
        End If

        Return teststring

    End Function
End Class
