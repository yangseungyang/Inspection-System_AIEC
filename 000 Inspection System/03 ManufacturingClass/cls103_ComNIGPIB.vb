Option Strict On : Option Explicit On

'Imports c = CommonCommunicationLibrary
Imports CommonCommunicationLibrary

Public Class cls103_ComNIGPIB
    '----------------------------------------
    '   定数宣言
    '----------------------------------------
    Dim cADO As cls010_MDB_Operation

    Private COMM_OK As Integer = 0
    Private COMM_ERR As Integer = 1

    'Private Const GpAd_FAM3 As String = "30"        'FA-M3
    'Private Const S100 As Integer = 3

    '----------------------------------------
    '   ローカル変数
    '----------------------------------------
    Private mRS As clsCommTmRs
    Private mGPIB As clsCommTmGpib
    Private mADDRESS As String
    Private mPort As String
    Private mSetting As String
    Private mTerminalChar As String
    Private mErrorComment As String

    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Private myErrFlag As Boolean

    Public Property ADDRESS() As String
        Set(ByVal Value As String)
            mADDRESS = Value
        End Set
        Get
            ADDRESS = mADDRESS
        End Get
    End Property

    Public Property PORT() As String
        Set(ByVal Value As String)
            mPort = Value
        End Set
        Get
            PORT = mPort
        End Get
    End Property

    Public Property SETTING() As String
        Set(ByVal Value As String)
            mSetting = Value
        End Set
        Get
            SETTING = mSetting
        End Get
    End Property

    Public Property TERMINAL_CHAR() As String
        Set(ByVal Value As String)
            mTerminalChar = Value
        End Set
        Get
            TERMINAL_CHAR = mTerminalChar
        End Get
    End Property

    'Error ----------------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public Function SettingRead() As Boolean

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        If mADDRESS = "" And mPort = "" Then
            SettingRead = False
            Exit Function
        End If

        cADO = New cls010_MDB_Operation
        dbName = mAppPath & "010 EquipmentData\機器構成一覧.xls"
        If mADDRESS <> "" Then
            tbName = "[GPIB$]"
            sSQL = "SELECT 設定 , デリミタ FROM " & tbName & " WHERE アドレス LIKE '" & mADDRESS & "'"
        ElseIf mPort <> "" Then
            tbName = "[RS232C$]"
            sSQL = "SELECT 設定 , デリミタ FROM " & tbName & " WHERE PORT LIKE '" & mPort & "'"
        End If

        cADO.DB_PATH = dbName
        cADO.TB_NAME = tbName
        cADO.DB_SQL = sSQL

        cADO.DB_READ_EXCEL_JET(fName, sData)

        If cADO.DB_DATA_COUNT = 0 Then
            SettingRead = False
            cADO = Nothing
            mErrorComment = "機器構成一覧に設定が登録されていません。"
            Exit Function
        Else
            mSetting = sData(0)
            mTerminalChar = sData(1)
            SettingRead = True
        End If

        cADO = Nothing

    End Function

    Public Sub Open()

        mErrorComment = ""

        Dim ret As Integer
        Dim SettingData As String = ""
        Dim mDelim As CommonCommunicationLibrary.clsCommBase.COMM_TERM

        Select Case mTerminalChar
            Case "vbcrlf" : mDelim = CommonCommunicationLibrary.clsCommBase.COMM_TERM.COMM_RS_CRLF
            Case "vbcr" : mDelim = CommonCommunicationLibrary.clsCommBase.COMM_TERM.COMM_RS_CR
            Case "vblf" : mDelim = CommonCommunicationLibrary.clsCommBase.COMM_TERM.COMM_RS_LF
            Case "EOI" : mDelim = CommonCommunicationLibrary.clsCommBase.COMM_TERM.COMM_RS_EOI
        End Select

        Select Case mADDRESS
            Case ""                 'RS232C
                Dim SetData() As String = Split(mSetting, "-")
                Select Case mPort
                    Case "COM1" : SettingData = "0,"
                    Case "COM2" : SettingData = "1,"
                    Case "COM3" : SettingData = "2,"
                    Case "COM4" : SettingData = "3,"
                    Case "COM5" : SettingData = "4,"
                    Case "COM6" : SettingData = "5,"
                    Case Else
                        mErrorComment = "Setting Error"
                End Select
                Select Case SetData(0)
                    Case "1200" : SettingData &= "0,"
                    Case "2400" : SettingData &= "1,"
                    Case "4800" : SettingData &= "2,"
                    Case "9600" : SettingData &= "3,"
                    Case "19200" : SettingData &= "4,"
                    Case "38400" : SettingData &= "5,"
                    Case "57600" : SettingData &= "6,"
                    Case Else
                        mErrorComment = "Setting Error"
                End Select
                Select Case SetData(1)
                    Case "N"
                        If SetData(2) = "8" And SetData(3) = "1" Then SettingData &= "0,"
                    Case "O"
                        If SetData(2) = "7" And SetData(3) = "1" Then SettingData &= "1,"
                    Case "E"
                        If SetData(2) = "7" And SetData(3) = "1" Then SettingData &= "2,"
                    Case Else
                        mErrorComment = "Setting Error"
                End Select

                SettingData &= "0"      'ハンドシェーク 無し = 0    XonXoff = 1     CTS-RTS = 2

                If mErrorComment <> "" Then Exit Sub

                mRS = New clsCommTmRs(mDelim)
                ret = mRS.Open(SettingData)

            Case Else               'GPIB NI ONLY

                mGPIB = New clsCommTmGpib(mDelim)
                ret = mGPIB.Open(mADDRESS)

        End Select

        If ret <> COMM_OK Then
            mErrorComment = "OPEN Error!"
        End If

    End Sub

    Public Sub Close()

        mErrorComment = ""

        Dim ret As Integer

        Select Case mADDRESS
            Case ""
                ret = mRS.Close
                mRS = Nothing
            Case Else
                ret = mGPIB.Close
                mGPIB = Nothing
        End Select

        If ret <> COMM_OK Then mErrorComment = "CLOSE Error!"

    End Sub

    Public Sub Send(ByVal cmd As String)

        mErrorComment = ""

        Dim ret As Integer

        Select Case mADDRESS
            Case "" : ret = mRS.Write(cmd)
            Case Else : ret = mGPIB.Write(cmd)
        End Select

        If ret <> COMM_OK Then mErrorComment = "DATA SEND Error!:" & cmd

    End Sub

    Public Sub Receive(ByRef Buff As String)
        Dim len As Integer
        Dim rLen As Integer
        Dim endFlag As Boolean
        Dim ret As Integer

        mErrorComment = ""

        len = 1024      '受信可能文字数
        Buff = Space(1024)
        Select Case mADDRESS
            Case "" : ret = mRS.Read(Buff, len, rLen, endFlag)
            Case Else : ret = mGPIB.Read(Buff, len, rLen, endFlag)
        End Select

        If endFlag = False Or (rLen = len) Or rLen <= 0 Then
            ret = COMM_ERR
        Else
            Buff = Left(Buff, rLen - 1)                                             'Length以降を削除
            If InStr(Buff, vbCr) > 0 Then Buff = Left(Buff, InStr(Buff, vbCr) - 1) 'CR以降を削除
            'If InStr(Buff, vbLf) > 0 Then Buff = Left(Buff, InStr(Buff, vbLf) - 1) 'LF以降を削除
        End If


        If ret <> COMM_OK Then mErrorComment = "DATA RECEIVE Error!"

    End Sub

    Public Sub SendReceive(ByVal cmd As String, ByRef Buff As String)

        Me.Send(cmd)
        Me.Receive(Buff)

    End Sub



    '---------------------------------------------
    '   プライベート関数
    '---------------------------------------------
    Private Sub Wait(ByVal Sec As Integer)
        System.Threading.Thread.Sleep((Sec * 1000))   'Sec秒のWait
    End Sub

    Private Sub Waitms(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)   'mSec秒のWait
    End Sub


End Class
