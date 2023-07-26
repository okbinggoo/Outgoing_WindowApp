Imports System.IO
Public Class WriteText_Module

    Private Property fs As Object
    Private Property a As Object
    Sub WriteTextFile(ByVal TextRecord As String, ByVal pathfile As String)
        Dim sFound As String
        Dim sContents As String


        sFound = ""
        sContents = GetFileContents(pathfile, sFound)

        TextRecord = Replace(TextRecord, vbCr, "")
        TextRecord = Replace(TextRecord, vbLf, "")

        'A. Keep Inv and item to calcualte in Excel
        If GlobalVariables.FlagKeepInv = "Y" Then
            TextRecord = ";" + GlobalVariables.Invoice + ";" + GlobalVariables.ItemNo + ";" +
                        GlobalVariables.CustCD + ";" + GlobalVariables.cust_part + ";" +
                        CStr(GlobalVariables.ShipQty) + TextRecord
        End If
        Dim sw As StreamWriter
        'B. Write on Textfile
        If sFound <> "" Then 'not found text -> Create New Text File
            My.Computer.FileSystem.WriteAllText(pathfile, TextRecord, True)
            'Carriage Return
            sw = File.AppendText(pathfile)
            sw.WriteLine()
            sw.Flush()
            sw.Close()
        Else ' found text continue write text file
            sw = File.AppendText(pathfile)
            sw.WriteLine(TextRecord)
            'sw.WriteLine()
            sw.Flush()
            sw.Close()

        End If
    End Sub
    Sub WriteTextFile_FileName(ByVal TextRecord As String, ByVal pathfile As String)
        Dim sFound As String
        Dim sContents As String


        sFound = ""
        sContents = GetFileContents(pathfile, sFound)

        TextRecord = Replace(TextRecord, vbCr, "")
        TextRecord = Replace(TextRecord, vbLf, "")

        Dim sw As StreamWriter
        'B. Write on Textfile
        If sFound <> "" Then 'not found text -> Create New Text File
            My.Computer.FileSystem.WriteAllText(pathfile, TextRecord, True)
            'Carriage Return
            sw = File.AppendText(pathfile)
            sw.WriteLine()
            sw.Flush()
            sw.Close()
        Else ' found text continue write text file
            sw = File.AppendText(pathfile)
            sw.WriteLine(TextRecord)
            'sw.WriteLine()
            sw.Flush()
            sw.Close()

        End If
    End Sub
    Public Function GetFileContents(ByVal FullPath As String,
   Optional ByRef ErrInfo As String = "") As String

        Dim strContents As String
        Dim objReader As StreamReader
        Try
            objReader = New StreamReader(FullPath)
            strContents = objReader.ReadToEnd()
            objReader.Close()
            Return strContents
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
    End Function

    Sub ClearTextFile(ByVal PathFile As String)
        Dim sFound As String
        Dim sContents As String

        sFound = ""
        sContents = GetFileContents(PathFile, sFound)

        If sFound = "" Then
            Kill(PathFile)
            'Call CreateAfile(PathFile)
        End If

    End Sub




End Class