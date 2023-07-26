Imports Oracle.DataAccess.Client
Public Class GlobalVariables
    'Global Const Database As String = "MTL1"
    Public Shared Read As New ReadText_module
    Public Shared Conn As New OracleConnection  ' Connection of Oracle
    Public Shared RS As New ADODB.Recordset
    Public Shared RS_IM As New ADODB.Recordset
    'set 10 parameter
    '********************************************************
    Public Shared hashConf As Hashtable = Read.readConfigFile()

    'DB For SQL config

    Public Shared Host As String = hashConf.Item("ORSERVER") '163.50.57.17
    Public Shared Port As String = hashConf.Item("PORT") '1521
    Public Shared FTP_SVR As String = hashConf.Item("FTP_SERVER") '163.50.57.20
    Public Shared Database As String = hashConf.Item("DATABASE_NAME") 'MTL1
    Public Shared UserName As String = hashConf.Item("USERNAME") 'OG DB :INSPMTL
    Public Shared Password As String = hashConf.Item("PASSWORD") 'OG DB :insp1112
    'DB For Outgoing SQL 
    Public Shared Database1 As String
    Public Shared UserName1 As String
    Public Shared Password1 As String

    'IM DB
    Public Shared IM_Host As String = hashConf.Item("IM_HOST")
    Public Shared IM_server As String = hashConf.Item("IM_SERVER")
    Public Shared IM_user As String = hashConf.Item("IM_USER")
    Public Shared IM_pass As String = hashConf.Item("IM_PASS")


    'DB For LS
    Public Shared HostLS As String = hashConf.Item("SPIRIT_HOST")
    Public Shared PortLS As String = hashConf.Item("SPIRIT_PORT")
    Public Shared DatabaseLS As String = hashConf.Item("SPIRIT_SERVER")
    Public Shared UserNameLS As String = hashConf.Item("SPIRIT_USER")
    Public Shared PasswordLS As String = hashConf.Item("SPIRIT_PASS")

    Public Shared UserNameLogisic As String = hashConf.Item("LOGISTIC_USERNAME")
    Public Shared PasswordLogisic As String = hashConf.Item("LOGISTIC_PASSWORD")

    'Email
    Public Shared Tomail As String = hashConf.Item("MAIL_TO")
    Public Shared CCmail As String = hashConf.Item("MAIL_CC")

    'Path PDF
    Public Shared PathPDF As String
    Public Shared PathZip As String
    Public Shared PathZipBK As String
    '*******************************************************
    Public Shared Fcty_CD As String
    Public Shared CustCD As String
    Public Shared CustPart As String
    Public Shared MurataType As String
    Public Shared Invoice As String
    Public Shared ItemNo As String
    Public Shared H40Code As String
    Public Shared Inspection As String
    Public Shared ShipQty As String
    Public Shared statuss As String

    Public Shared PONO As String
    Public Shared FormNo As String
    Public Shared CustName As String
    Public Shared fileName As String
    Public Shared outging_flag As String
    Public Shared shipdate As String
    Public Shared cust_purch As String
    Public Shared cust_part As String
    Public Shared card_no As String
    Public Shared cust_part_name As String
    Public Shared murata_allocate As String

    Public Shared Local1 As String
    Public Shared Local2 As String
    Public Shared Local3 As String
    Public Shared Local4 As String
    Public Shared Local5 As String
    Public Shared Local6 As String
    Public Shared Local7 As String
    Public Shared Local8 As String
    Public Shared Local9 As String
    Public Shared Local10 As String
    '***************************************************
    Public Shared LotNo As String
    Public Shared PathOfPart As String
    Public Shared PD_Spec_No As String
    Public Shared PackingCode As String
    Public Shared FlagCheck As Boolean
    Public Shared PathServer As String
    Public Shared FTP_SVR_Input As String
    Public Shared FTP_SVR_BK As String
    Public Shared FTP_TextFile As String
    Public Shared FileRecord As String = hashConf.Item("TOEXCEL_FILE")
    Public Shared FileRecord_filename As String = hashConf.Item("TOEXCEL_FILENAME")
    Public Shared FileTNSerror As String = hashConf.Item("TNS_FILE")
    Public Shared TextRecord As String
    Public Shared SqlCMD As String = hashConf.Item("SQLCMD_OG")
    Public Shared SQLStr1 As String
    Public Shared SQLStr2 As String
    Public Shared SQLStr3 As String
    Public Shared SQLStr4 As String
    Public Shared SQLStr5 As String
    Public Shared SQLStr6 As String
    Public Shared SQLStr7 As String
    Public Shared SQLStr8 As String
    Public Shared SQLStr9 As String
    Public Shared SQLStr10 As String
    Public Shared Seperator As String
    Public Shared FlagKeepInv As String
    Public Shared FlagMurataType As String
    Public Shared StrStationName As String

    Public Shared InspectionRecords As Integer
    Public Shared ErrorFlag As Boolean
    Public Shared ErrID As Integer
End Class
