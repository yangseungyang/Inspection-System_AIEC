Option Strict On : Option Explicit On

Public Class cls030_LineSerial

    Dim mSerial As String
    Dim mError As String

    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    'Line Serial ===============================
    Public ReadOnly Property LineSerial() As String
        Get
            LineSerial = mSerial
        End Get
    End Property

    'Get Serial Error ==========================
    Public ReadOnly Property ERR_MSG() As String
        Get
            ERR_MSG = mError
        End Get
    End Property

    Public Sub IssueLineSerial()
        '==========================================================================================
        ' F.Name   : IssueLineSerial
        ' Function : Issue LineSerial
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim cSQL As cls010_MDB_Operation

        'Dim dbName As String
        Dim tbName As String
        Dim sSQL As String
        Dim rField As String
        Dim kField As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim dbFactry As String
        Dim dbYear As String
        Dim dbMonth As String
        Dim dbDay As String
        Dim dbSystem As String
        Dim dbNumber As String
        Dim nwFactry As String
        Dim nwYear As String
        Dim nwMonth As String
        Dim nwDay As String
        Dim nwSystem As String
        Dim nwNumber As String

        'Read DB Serial No ===========================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\LineSerial.mdb"
        tbName = "tblNextSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        dbFactry = sData(0) : dbYear = sData(1) : dbMonth = sData(2) : dbDay = sData(3)
        dbSystem = sData(4) : dbNumber = sData(5)


        'Make this time number =======================================
        nwFactry = dbFactry
        nwYear = Right(System.DateTime.Now.ToString("yyyy"), 1)
        'nwYear = Right(Format(System.DateTime.Now.ToString, "yyyy"), 1)
        nwMonth = Mid("ABCDEFGHJKLM", CInt(System.DateTime.Now.Month.ToString), 1)
        'nwMonth = Mid("ABCDEFGHJKLM", CInt(Format(System.DateTime.Now.ToString, "mm")), 1)
        nwDay = System.DateTime.Now.ToString("dd")
        'nwDay = Format(System.DateTime.Now.ToString, "dd")

        'Read SystemNo. --------------------
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        tbName = "tblEquipmentSetting"
        cSQL.TB_NAME = tbName
        rField = "EquipmentNumber" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        nwSystem = sData(0)

        'Check SystemNo. ------------------
        If dbSystem <> nwSystem Then
            mError = "The number of systems is different."
            GoTo ErrorTrap
        End If

        'Check Number ---------------------
        If nwYear <> dbYear Or nwMonth <> dbMonth Or nwDay <> dbDay Then
            nwNumber = "001"
        Else
            nwNumber = dbNumber
        End If

        'This time Serial number ----------
        mSerial = nwFactry & nwYear & nwMonth & nwDay & nwSystem & nwNumber


        'Search past Serial No. =====================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\LineSerial.mdb"
        tbName = "tblPastSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = "PastSerial"

        sSQL = "SELECT " & "*" & " FROM " & tbName & " "
        sSQL = sSQL & "WHERE " & kField & " LIKE " & "'" & mSerial & "' "
        sSQL = sSQL & "ORDER BY " & kField
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.DB_DATA_COUNT <> 0 Then
            mError = mSerial & ":This Serial number has been issued in the past." & vbCr & "発行済番号です。発番できません。"
            GoTo ErrorTrap
        End If

        cSQL = Nothing


        'Addition past Serial No. =====================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\LineSerial.mdb"
        tbName = "tblPastSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = mSerial

        sSQL = "INSERT INTO " & tbName
        sSQL = sSQL & " VALUES ('" & mSerial & "');"
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_EXEC_JET()
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing


        'Issue Next Serial No. ===================================
        nwNumber = Format(CInt(nwNumber) + 1, "000")


        'Rewrite Next Serial No. =================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\LineSerial.mdb"
        tbName = "tblNextSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        sSQL = "UPDATE " & tbName & " SET "
        sSQL = sSQL & fName(0) & " = " & "'" & nwFactry & "' , "
        sSQL = sSQL & fName(1) & " = " & "'" & nwYear & "' , "
        sSQL = sSQL & fName(2) & " = " & "'" & nwMonth & "' , "
        sSQL = sSQL & fName(3) & " = " & "'" & nwDay & "' , "
        sSQL = sSQL & fName(4) & " = " & "'" & nwSystem & "' , "
        sSQL = sSQL & fName(5) & " = " & "'" & nwNumber & "'"
        sSQL = sSQL & " WHERE "
        sSQL = sSQL & fName(0) & " = " & "'" & sData(0) & "'"
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_EXEC_JET()
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        Exit Sub

ErrorTrap:

        If mError = "" Then mError = "Issue Line Serial NG"

    End Sub

    Public Sub IssueLineSerialZ()
        '==========================================================================================
        ' F.Name   : IssueLineSerial
        ' Function : Issue LineSerial
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim cSQL As cls010_MDB_Operation

        'Dim dbName As String
        Dim tbName As String
        Dim sSQL As String
        Dim rField As String
        Dim kField As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim dbFactry As String
        Dim dbYear As String
        Dim dbMonth As String
        Dim dbDay As String
        Dim dbSystem As String
        Dim dbNumber As String
        Dim nwFactry As String
        Dim nwYear As String
        Dim nwMonth As String
        Dim nwDay As String
        Dim nwSystem As String = "Z"
        Dim nwNumber As String

        'Read DB Serial No ===========================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\Z\LineSerial.mdb"
        tbName = "tblNextSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        dbFactry = sData(0) : dbYear = sData(1) : dbMonth = sData(2) : dbDay = sData(3)
        dbSystem = sData(4) : dbNumber = sData(5)


        'Make this time number =======================================
        nwFactry = dbFactry
        nwYear = Right(System.DateTime.Now.ToString("yyyy"), 1)
        'nwYear = Right(Format(System.DateTime.Now.ToString, "yyyy"), 1)
        nwMonth = Mid("ABCDEFGHJKLM", CInt(System.DateTime.Now.Month.ToString), 1)
        'nwMonth = Mid("ABCDEFGHJKLM", CInt(Format(System.DateTime.Now.ToString, "mm")), 1)
        nwDay = System.DateTime.Now.ToString("dd")
        'nwDay = Format(System.DateTime.Now.ToString, "dd")

        ''Read SystemNo. --------------------
        'cSQL = New cls010_MDB_Operation

        'cSQL.DB_PATH = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        'tbName = "tblEquipmentSetting"
        'cSQL.TB_NAME = tbName
        'rField = "EquipmentNumber" : kField = ""

        'sSQL = "SELECT " & rField & " FROM " & tbName
        'cSQL.DB_SQL = sSQL

        'Call cSQL.DB_READ_JET(fName, sData)
        'If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        'cSQL = Nothing

        'nwSystem = sData(0)

        ''Check SystemNo. ------------------
        'If dbSystem <> nwSystem Then
        '    mError = "The number of systems is different."
        '    GoTo ErrorTrap
        'End If

        'Check Number ---------------------
        If nwYear <> dbYear Or nwMonth <> dbMonth Or nwDay <> dbDay Then
            nwNumber = "001"
        Else
            nwNumber = dbNumber
        End If

        'This time Serial number ----------
        mSerial = nwFactry & nwYear & nwMonth & nwDay & nwSystem & nwNumber


        'Search past Serial No. =====================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\Z\LineSerial.mdb"
        tbName = "tblPastSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = "PastSerial"

        sSQL = "SELECT " & "*" & " FROM " & tbName & " "
        sSQL = sSQL & "WHERE " & kField & " LIKE " & "'" & mSerial & "' "
        sSQL = sSQL & "ORDER BY " & kField
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.DB_DATA_COUNT <> 0 Then
            mError = mSerial & ":This Serial number has been issued in the past." & vbCr & "発行済番号です。発番できません。"
            GoTo ErrorTrap
        End If

        cSQL = Nothing


        'Addition past Serial No. =====================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\Z\LineSerial.mdb"
        tbName = "tblPastSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = mSerial

        sSQL = "INSERT INTO " & tbName
        sSQL = sSQL & " VALUES ('" & mSerial & "');"
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_EXEC_JET()
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing


        'Issue Next Serial No. ===================================
        nwNumber = Format(CInt(nwNumber) + 1, "000")


        'Rewrite Next Serial No. =================================
        cSQL = New cls010_MDB_Operation

        cSQL.DB_PATH = mAppPath & "100 System Data\Z\LineSerial.mdb"
        tbName = "tblNextSerial"
        cSQL.TB_NAME = tbName
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        sSQL = "UPDATE " & tbName & " SET "
        sSQL = sSQL & fName(0) & " = " & "'" & nwFactry & "' , "
        sSQL = sSQL & fName(1) & " = " & "'" & nwYear & "' , "
        sSQL = sSQL & fName(2) & " = " & "'" & nwMonth & "' , "
        sSQL = sSQL & fName(3) & " = " & "'" & nwDay & "' , "
        sSQL = sSQL & fName(4) & " = " & "'" & nwSystem & "' , "
        sSQL = sSQL & fName(5) & " = " & "'" & nwNumber & "'"
        sSQL = sSQL & " WHERE "
        sSQL = sSQL & fName(0) & " = " & "'" & sData(0) & "'"
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_EXEC_JET()
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        Exit Sub

ErrorTrap:

        If mError = "" Then mError = "Issue Line Serial NG"

    End Sub

End Class
