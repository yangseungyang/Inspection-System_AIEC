Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls110_ServerConnect
'Function   : Data server connect (for FA-M3) 
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'           :
'Issue      : 2008/12/30  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

Public Class cls110_ServerConnect

    'Class ================================================
    Dim cADO As cls010_MDB_Operation
    Dim cDrv As cls020_Drive_Control

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mErrorComment As String = ""
    Dim mErrorCode As String = "0"

    'Local Property =======================================
    Dim mFormName As System.Windows.Forms.Form
    Dim mServerDrive As String = ""

    'Form name ------------------------
    Public WriteOnly Property FORM_NAME() As System.Windows.Forms.Form
        Set(ByVal Value As System.Windows.Forms.Form)
            mFormName = Value
        End Set
    End Property

    Public ReadOnly Property SERVER_DRIVE() As String
        Get
            SERVER_DRIVE = mServerDrive
        End Get
    End Property

    'Error message --------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public ReadOnly Property ERR_CODE() As String
        Get
            ERR_CODE = mErrorCode
        End Get
    End Property

    Public Sub ServerConnect()
        '==============================================================================================================
        ' Function : Server connect
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        'Read of Network setting ==================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "BackupDriveSign,BackupDrivePath" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Check of setting data ==================================================================== 
        If sData(0) = "" Or sData(1) = "" Then
            mErrorComment = "サーバー情報の設定異常"
            mServerDrive = ""
            Exit Sub
        End If

        'Server connect ===========================================================================
        'Get of drive name --------------------------------------
        Dim DriveInfo() As String = {}

        Try
            cDrv = New cls020_Drive_Control

            Call cDrv.DRIVE_INFPRMATION(DriveInfo)

        Catch ex As Exception
            mErrorComment = cDrv.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cDrv = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Is a set drive used ? ----------------------------------
        For i As Integer = 0 To UBound(DriveInfo)
            If InStr(DriveInfo(i), sData(0) & ":") <> 0 Then mServerDrive = DriveInfo(i)
        Next

        'Connect or Confirm of drive name -----------------------
        If mServerDrive = "" Then
            'Connect ----------------------------------
            Try
                cDrv = New cls020_Drive_Control

                cDrv.FORM_NAME = mFormName
                cDrv.NET_PATH = sData(1)
                cDrv.DRIVE = sData(0)
                cDrv.USER = "aiec"
                cDrv.PASS = "Kme12345"

                Call cDrv.NETWORK_DRIVE_CONNECT()

            Catch ex As Exception
                mErrorComment = cDrv.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()

            Finally
                cDrv = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

            mServerDrive = sData(0) & ":" & sData(1)

        Else
            'Connected drive of name check ------------ 
            If Mid(mServerDrive, InStr(mServerDrive, "\\"), Len(mServerDrive)) = sData(1) Then
                mServerDrive = sData(0) & ":" & sData(1)
            Else
                mErrorComment = "指定ドライブが使用中：" & mServerDrive
                mServerDrive = ""
                Exit Sub
            End If

        End If

    End Sub

    Public Sub ServerDisconnect()
        '==============================================================================================================
        ' Function : Server Disconnect
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        'Read of Network setting ==================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "BackupDriveSign" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Check of drive ===========================================================================
        Dim DriveInfo() As String = {}

        Try
            cDrv = New cls020_Drive_Control

            Call cDrv.DRIVE_INFPRMATION(DriveInfo)

        Catch ex As Exception
            mErrorComment = cDrv.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cDrv = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Is a set drive used ? ----------------------------------
        For i As Integer = 0 To UBound(DriveInfo)
            If InStr(DriveInfo(i), sData(0) & ":") <> 0 Then mServerDrive = DriveInfo(i)
        Next

        'Disconnect ===============================================================================
        If mServerDrive <> "" Then
            Try
                cDrv = New cls020_Drive_Control

                cDrv.DRIVE = sData(0)
                Call cDrv.NETWORK_DRIVE_DISCONNECT()

            Catch ex As Exception
                mErrorComment = cDrv.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()

                Throw New ApplicationException(mErrorComment)

            Finally
                cDrv = Nothing

            End Try
        End If

    End Sub

    Private Sub WaitTimer(ByVal tm As Double)
        '==========================================================================================
        ' Function : Wait timer
        ' Data     : tm : Wait time(Sec)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim sTime As Double
        Dim nTime As Double
        Dim eTime As Double

        sTime = Microsoft.VisualBasic.DateAndTime.Timer()

        Do While tm > eTime
            nTime = Microsoft.VisualBasic.DateAndTime.Timer()
            eTime = nTime - sTime
            If eTime < 0 Then eTime = 24.0# * 3600.0# - sTime + nTime
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

End Class
