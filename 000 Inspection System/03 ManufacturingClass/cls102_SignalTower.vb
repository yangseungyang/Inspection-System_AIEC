Public Class cls102_SignalTower

    Dim cLnk As cls080_F3LinkToLE
    Dim cSQL As cls010_MDB_Operation

    'Dim FAM3Slot As String = "13"
    Dim FAM3Slot As String = "5"

    Dim mSigTower As Boolean
    Dim mComment As String
    Dim mError As String
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Public WriteOnly Property SignalTowerSTATUS() As Boolean
        Set(ByVal Value As Boolean)
            mSigTower = Value
        End Set
    End Property

    Public Sub SignalTower(ByVal color As String, ByVal buzz1 As String, ByVal buzz2 As String, ByVal flash As String)
        '==========================================================================================
        ' F.Name   : SignalTower
        ' Function : Signal tower control
        ' Data     : color,buzz1,buzz2,flash
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim scmd As String = ""
        Dim rCmd As String = ""

        Dim dbName As String
        Dim tbName As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim rField As String
        'Dim kField As String
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        'Signaltower Control Check =====================
        cSQL = New cls010_MDB_Operation

        dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        tbName = "tblEquipmentSetting"
        rField = "SignalTowerControl"
        'kField = "SystemVersion"

        'sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & kField & " in "
        'sSQL = sSQL & "(SELECT " & "Max(" & kField & ") as maxver" & " FROM " & tbName & ")"
        'sSQL = sSQL & " ORDER BY " & kField & " DESC;"

        sSQL = "SELECT " & rField & " FROM " & tbName

        cSQL.TB_NAME = tbName
        cSQL.DB_PATH = dbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        If InStr(sData(0), "N") <> 0 Then Exit Sub

        'Relay OFF =====================================
        cLnk = New cls080_F3LinkToLE

        cLnk.IPAdd = "192.168.1.2"

        If buzz1 = "OFF" And buzz2 = "OFF" And flash = "OFF" Then
            Select Case color
                Case "RED" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,10000000"
                Case "YELLOW" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,01000000"
                Case "GREEN" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00100000"
            End Select
        End If

        If buzz1 = "ON" And buzz2 = "OFF" And flash = "OFF" Then
            Select Case color
                Case "RED" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,10010000"
                Case "YELLOW" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,01010000"
                Case "GREEN" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00110000"
            End Select
        End If

        If buzz1 = "OFF" And buzz2 = "ON" And flash = "OFF" Then
            Select Case color
                Case "RED" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,10001000"
                Case "YELLOW" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,01001000"
                Case "GREEN" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00101000"
            End Select
        End If

        If buzz1 = "OFF" And buzz2 = "OFF" And flash = "ON" Then
            Select Case color
                Case "RED" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00000100"
                Case "YELLOW" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00000010"
                Case "GREEN" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00000001"
            End Select
        End If

        If buzz1 = "ON" And buzz2 = "OFF" And flash = "ON" Then
            Select Case color
                Case "RED" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00010100"
                Case "YELLOW" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00010010"
                Case "GREEN" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00010001"
            End Select
        End If

        If buzz1 = "OFF" And buzz2 = "ON" And flash = "ON" Then
            Select Case color
                Case "RED" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00001100"
                Case "YELLOW" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00001010"
                Case "GREEN" : scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00001001"
            End Select
        End If

        'Sound OFF ---------------------------------
        If color = "" And buzz1 = "OFF" And buzz2 = "OFF" And flash = "" Then
            scmd = "01" & "BWRY" & FAM3Slot & "04" & ",002,00"
        End If

        Call cLnk.F3PCLinkCommunication(scmd, rCmd)
        If cLnk.ERR_COMMENT <> "" Then
            mError = cLnk.ERR_COMMENT
            GoTo ErrorTrap
        End If

        cLnk = Nothing

        Exit Sub

ErrorTrap:

        MsgBox("Signal tower control NG " & mError, MsgBoxStyle.OkOnly)

        cLnk = Nothing
        cSQL = Nothing

    End Sub

    Public Sub SignalTowerReset()
        '==========================================================================================
        ' F.Name   : SignalTowerReset
        ' Function : Signal tower All OFF
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim scmd As String
        Dim rCmd As String = ""

        Dim dbName As String
        Dim tbName As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim rField As String
        'Dim kField As String
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        'Signaltower Control Check =====================
        dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        tbName = "tblEquipmentSetting"
        rField = "SignalTowerControl"
        'kField = "SystemVersion"

        'sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & kField & " in "
        'sSQL = sSQL & "(SELECT " & "Max(" & kField & ") as maxver" & " FROM " & tbName & ")"
        'sSQL = sSQL & " ORDER BY " & kField & " DESC;"

        sSQL = "SELECT " & rField & " FROM " & tbName

        cSQL = New cls010_MDB_Operation
        cSQL.DB_PATH = dbName
        cSQL.TB_NAME = tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        If InStr(sData(0), "N") <> 0 Then Exit Sub

        'Relay OFF =====================================
        cLnk = New cls080_F3LinkToLE

        cLnk.IPAdd = "192.168.1.2"
        
        scmd = "01" & "BWRY" & FAM3Slot & "01" & ",008,00000000"
        Call cLnk.F3PCLinkCommunication(scmd, rCmd)
        If cLnk.ERR_COMMENT <> "" Then
            mError = cLnk.ERR_COMMENT
            GoTo ErrorTrap
        End If

        cLnk = Nothing

        Exit Sub

ErrorTrap:

        MsgBox("Signal tower reset NG " & mError)

        cLnk = Nothing
        cSQL = Nothing

    End Sub


End Class
