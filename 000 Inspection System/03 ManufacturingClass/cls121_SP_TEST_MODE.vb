Option Strict On : Option Explicit On



'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm020_InspectionStartEnd

Public Class cls121_SP_TEST_MODE

    'Local ================================================
    Dim mRSComm As System.IO.Ports.SerialPort
    Dim mPort As String
    Dim mSetting As String
    Dim mTerminalChar As String
    Dim mRecvData As String
    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mSendCmd As String
    Dim mTimeOut As Integer
    Dim mSendPARA As String()
    Dim mMODEL As String
    Dim mICMD As Integer
    Dim mIUNIT As Integer
    Dim mISLOT As Integer
    Dim mIPARA As Integer
    Dim mSendText As String
    Dim mLOG As String

    Dim cRS232C As cls101_ComRS232C

    Dim ENQ As String = Chr(&H5)
    Dim EOT As String = Chr(&H4)
    Dim ACK0 As String = Chr(&H1030)
    Dim ACK1 As String = Chr(&H1031)

    'Property =====================================================================================
    'Communication Object -------------   
    Public WriteOnly Property RS_COMM() As System.IO.Ports.SerialPort
        Set(ByVal Value As System.IO.Ports.SerialPort)
            mRSComm = Value
        End Set
    End Property

    'CMD -----------------------------
    Public WriteOnly Property ICMD() As Integer
        Set(ByVal Value As Integer)
            mICMD = Value
        End Set
    End Property

    'UNIT -----------------------------
    Public WriteOnly Property IUNIT() As Integer
        Set(ByVal Value As Integer)
            mIUNIT = Value
        End Set
    End Property

    'SLOT -----------------------------
    Public WriteOnly Property ISLOT() As Integer
        Set(ByVal Value As Integer)
            mISLOT = Value
        End Set
    End Property

    'PARAMETER数 -----------------------------
    Public WriteOnly Property IPARA() As Integer
        Set(ByVal Value As Integer)
            mIPARA = Value
        End Set
    End Property

    'Port -----------------------------
    Public WriteOnly Property PORT() As String
        Set(ByVal Value As String)
            mPort = Value
        End Set
    End Property

    'SendCommand -----------------------------
    Public Property SendCmd() As String
        Set(ByVal Value As String)
            mSendCmd = Value
        End Set
        Get
            SendCmd = mSendCmd
        End Get
    End Property

    'SendParameter -----------------------------
    Public Property SendPARA() As String()
        Set(ByVal Value() As String)
            mSendPARA = Value
        End Set
        Get
            SendPARA = mSendPARA
        End Get
    End Property

    'Receive ---------------------------
    Public ReadOnly Property RECEIVE_DATA() As String
        Get
            RECEIVE_DATA = mRecvData
        End Get
    End Property

    'LOG ---------------------------
    Public ReadOnly Property LOG() As String
        Get
            LOG = mLOG
        End Get
    End Property

    'Error ----------------------------
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

    Public ReadOnly Property SEND_TESTCOMMAND() As String
        Get
            SEND_TESTCOMMAND = mSendText
        End Get
    End Property

    Public Sub OldSpTESTMODEStart()

        mLOG = ""
        'SPをﾃｽﾄﾓｰﾄﾞに移行　SP2/SP3/SP5
        Call WriteLOG("SPをﾃｽﾄﾓｰﾄﾞに移行")

        Dim SendByte(14) As Byte
        Dim i As Integer
        Dim SDISM, IC As Integer

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "9600,E,8,1"  '固定

        '------------------------------------------------------
        'ポートを開く
        Call WriteLOG("RS232Cﾎﾟｰﾄｵｰﾌﾟﾝ")
        Try                             'V3.02
            Call cRS232C.PortSetup()
        Catch ex As Exception
            mErrorComment = "ﾎﾟｰﾄｵｰﾌﾟﾝ失敗"
            Call WriteLOG(mErrorComment)
            Call cRS232C.PortClose()
            cRS232C = Nothing
            Throw New ApplicationException(mErrorComment)
        End Try
        '------------------------------------------------------
        Call WriteLOG("RS232CﾎﾟｰﾄｵｰﾌﾟﾝOK")

        mSecWait(300)

        'ENQ送信
        Call WriteLOG("ENQ（Chr(&H5)）送信")
        SendByte(0) = &H5
        cRS232C.SendCmd = ENQ
        Try
            Call cRS232C.FuncSendCmd()
            'Call cRS232C.FuncRecvCmd()
            Call cRS232C.FuncRecvByteData()
        Catch ex As Exception
            mErrorComment = "RS232C送受信ｴﾗｰ"
            Debug.Print(mErrorComment)
            Call WriteLOG(mErrorComment)
            cRS232C.PortClose()
            cRS232C = Nothing           'V3.02
            Throw New ApplicationException(mErrorComment)
        End Try
        Call WriteLOG("ENQ送信OK")

        'ACK0受信確認
        Call WriteLOG("ACK0（Chr(&H1030)）受信確認")
        If cRS232C.RECEIVE_DATA <> ACK0 Then
            mErrorComment = "ACK0受信ｴﾗｰ"
            Debug.Print(mErrorComment)
            Call WriteLOG(mErrorComment)
            cRS232C.PortClose()
            cRS232C = Nothing           'V3.02
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLOG("ACK0受信確認OK")

        Call mSecWait(300)
        Call cRS232C.RecBufIni()

        '送信データの作成 
        SendByte(0) = &H1
        SendByte(1) = &H4D
        SendByte(2) = &H0
        SendByte(3) = &HA
        SendByte(4) = &H2
        SendByte(5) = &H99
        For i = 6 To 12
            SendByte(i) = &H0
        Next i
        SendByte(13) = &H3

        'チェックサムの計算
        SDISM = 0
        For i = 1 To 13
            IC = SendByte(i) And &HFF
            SDISM = (SDISM + IC) Mod &H100
        Next
        SendByte(14) = CByte(SDISM)

        cRS232C.SendByteData = SendByte

        '送信データ送信
        Call WriteLOG("ﾃﾞｰﾀ送信")
        Try
            Call cRS232C.FuncSendByteData()
            'Call cRS232C.FuncRecvCmd()
            Call cRS232C.FuncRecvByteData()
        Catch ex As Exception
            mErrorComment = "RS232C送受信ｴﾗｰ"
            Debug.Print(mErrorComment)
            Call WriteLOG(mErrorComment)
            cRS232C.PortClose()
            cRS232C = Nothing           'V3.02
            Throw New ApplicationException(mErrorComment)
        End Try

        'ACK1受信確認
        Call WriteLOG("ACK1（Chr(&H1031)）受信確認")
        If cRS232C.RECEIVE_DATA <> ACK1 Then
            mErrorComment = "ACK1受信ｴﾗｰ"
            Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
            Call WriteLOG(mErrorComment)
            cRS232C.PortClose()
            cRS232C = Nothing           'V3.02
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLOG("ACK1受信確認OK")

        'EOT送信
        Call WriteLOG("EOT（Chr(&H4)）送信")
        cRS232C.SendCmd = EOT
        Call cRS232C.FuncSendCmd()

        'ポートクローズ
        Call cRS232C.PortClose()
        cRS232C = Nothing

        Call WriteLOG("ﾃｽﾄﾓｰﾄﾞ移行OK")

    End Sub

    Public Sub TestComSendRecv()

        mLOG = ""

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "9600,N,8,1"  '固定

        Try
            '------------------------------------------------------
            'ポートを開く
            cRS232C.PortSetup()
            '------------------------------------------------------
            mSendText = "TS,"
            mSendText &= Microsoft.VisualBasic.Right(Hex(mICMD), 2) & "," & Microsoft.VisualBasic.Right(Hex(mIUNIT), 2) & ","
            mSendText &= Microsoft.VisualBasic.Right(Hex(mISLOT), 2) & ","

            If mIPARA <> 0 Then
                For i = 1 To mIPARA
                    mSendText &= mSendPARA(i) & ","
                Next i
            End If

            cRS232C.TIMEOUT = 60
            cRS232C.SendCmd = mSendText
            cRS232C.TERMINAL_CHAR = vbCrLf
            Call cRS232C.FuncSendCmd()
            Call cRS232C.FuncRecvCmd()
            mRecvData = cRS232C.RECEIVE_DATA

        Catch ex As Exception
            mErrorComment = "TSｺﾏﾝﾄﾞ送受信ｴﾗｰ"
            Throw New ApplicationException(mErrorComment)
        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Sub

    Public Sub TestComSendRecvTimeout30()

        mLOG = ""

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "9600,N,8,1"  '固定

        Try
            '------------------------------------------------------
            'ポートを開く
            cRS232C.PortSetup()
            '------------------------------------------------------

            mSendText = "TS,"
            mSendText &= Microsoft.VisualBasic.Right(Hex(mICMD), 2) & "," & Microsoft.VisualBasic.Right(Hex(mIUNIT), 2) & ","
            mSendText &= Microsoft.VisualBasic.Right(Hex(mISLOT), 2) & ","

            If mIPARA <> 0 Then
                For i = 1 To mIPARA
                    mSendText &= mSendPARA(i) & ","
                Next i
            End If

            cRS232C.TIMEOUT = 30

            cRS232C.SendCmd = mSendText
            cRS232C.TERMINAL_CHAR = vbCrLf
            Call cRS232C.FuncSendCmd()
            Call cRS232C.FuncRecvCmd()
            mRecvData = cRS232C.RECEIVE_DATA

        Catch ex As Exception
            mErrorComment = "TSｺﾏﾝﾄﾞ送受信ｴﾗｰ"
            Throw New ApplicationException(mErrorComment)

        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Sub

    Public Function ComMode_Change(ByRef GA As String) As Boolean

        Dim sData() As Byte

        ComMode_Change = False
        ReDim sData(28)
        'Header ===============================
        sData(0) = &H1 : sData(1) = &H4D
        'Length of data 4-N(Byte) =============
        sData(2) = &H0 : sData(3) = &H18   'STX$～ETX$のデータ数
        'STX$ =================
        sData(4) = &H2
        'Mコマンド ============
        sData(5) = &HF1
        '通信ID ===============
        sData(6) = &H0 : sData(7) = &H0 : sData(8) = &H0 : sData(9) = &H0 : sData(10) = &H0
        sData(11) = &H0 : sData(12) = &H0 : sData(13) = &H0 : sData(14) = &H0 : sData(15) = &H0
        sData(16) = &H0
        'SlotNo===============
        sData(17) = &H2 'ｽﾛｯﾄは2ｽﾛｯﾄ目
        'UnitNo===============
        sData(18) = &H0
        '要求コマンド
        'Main ===================================
        sData(19) = &H4
        'SUB ====================================
        sData(20) = &H1
        'データサイズ============================
        sData(21) = &H0 : sData(22) = &H4
        '通信モード ================================
        sData(23) = &H0
        If GA = "KP" Then sData(24) = &H9 'KPの時は115200bpsﾊﾟﾘﾃｨなし
        If GA = "PF" Then sData(24) = &H2 'PFの時は19200bps偶数ﾊﾟﾘﾃｨ
        '確定時間 ================================
        sData(25) = &HFF : sData(26) = &HFF
        'ETX$ ===================================
        sData(27) = &H3
        'Sum=====
        Dim intSData As Integer = 0
        For i = 1 To 27
            intSData = intSData + sData(i)
        Next i
        sData(28) = CByte(intSData Mod 256)

        Dim intRetry As Integer = 0

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "9600,E,8,1"  '固定

        Try
            '------------------------------------------------------
            'ポートを開く
            Call WriteLOG("RS232Cﾎﾟｰﾄｵｰﾌﾟﾝ")
            Try                             'V3.02
                Call cRS232C.PortSetup()
            Catch ex As Exception
                mErrorComment = "ﾎﾟｰﾄｵｰﾌﾟﾝ失敗"
                Throw New ApplicationException(mErrorComment)
            End Try
            '------------------------------------------------------
            Call WriteLOG("RS232CﾎﾟｰﾄｵｰﾌﾟﾝOK")

            mSecWait(300)

            'ENQ送信
            Call WriteLOG("ENQ（Chr(&H5)）送信")
            cRS232C.SendCmd = ENQ
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ENQ送信OK")

            'ACK0受信確認
            Call WriteLOG("ACK0（Chr(&H1030)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK0 Then
                mErrorComment = "ACK0受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK0受信確認OK")

            Call mSecWait(300)
            Call cRS232C.RecBufIni()

            '送信ﾃﾞｰﾀ送信
            Call WriteLOG("ﾃﾞｰﾀ送信")
            cRS232C.SendByteData = sData
            Try
                Call cRS232C.FuncSendByteData()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'ACK1受信確認
            Call WriteLOG("ACK1（Chr(&H1031)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK1 Then
                mErrorComment = "ACK1受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK1受信確認OK")

            'EOT送信-----------------------------
            Call WriteLOG("EOT（Chr(&H4)）送信")
            cRS232C.SendCmd = EOT
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("EOT送信OK")

            'ENQ受信確認
            Call WriteLOG("ENQ（Chr(&H5)）受信確認")
            If cRS232C.RECEIVE_DATA <> ENQ Then
                mErrorComment = "ENQ受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ENQ受信確認OK")

            'Call cRS232C.RecBufIni()

            'ACK0送信--------------------------------
            Call WriteLOG("ACK0（Chr(&1030)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H30)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK0送信OK")

            'ﾃﾞｰﾀ受信(KPのとき。PFはindex-1）
            'RData(0):0x0
            'RData(1):0x1
            'RData(2):0x4D
            'RData(3):0x0
            'RData(4):0x16
            'RData(5):0x2----STX
            'RData(6):0xF1---M
            'RData(7):0x0
            'RData(8):0x0
            'RData(9):0x0
            'RData(10):0x0
            'RData(11):0x0
            'RData(12):0x0
            'RData(13):0x0
            'RData(14):0x0
            'RData(15):0x0
            'RData(16):0x0
            'RData(17):0x0
            'RData(18):0x1----Slot no
            'RData(19):0x0----Unit no
            'RData(20):0x4----$0
            'RData(21):0x1----$1
            'RData(22):0x0----$2
            'RData(23):0x2----$3
            'RData(24):0x0----$4:ErrorCode
            'RData(25):0x0----$5:ErrorCode
            'RData(26):0x3----ETX
            'RData(27):0x61---SUM

            Select Case GA
                Case "KP"
                    If cRS232C.RECEIVE_BYTE_DATA(24) <> 0 Or cRS232C.RECEIVE_BYTE_DATA(25) <> 0 Then
                        Debug.Print("エラーコード：" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25)))
                        mErrorComment = "ｴﾗｰｺｰﾄﾞ受信:" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25))
                        Throw New ApplicationException(mErrorComment)
                    Else
                        Call WriteLOG("ｴﾗｰｺｰﾄﾞ0受信OK")
                    End If
                Case "PF"
                    If cRS232C.RECEIVE_BYTE_DATA(23) <> 0 Or cRS232C.RECEIVE_BYTE_DATA(24) <> 0 Then
                        Debug.Print("エラーコード：" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25)))
                        mErrorComment = "ｴﾗｰｺｰﾄﾞ受信:" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25))
                        Throw New ApplicationException(mErrorComment)
                    Else
                        Call WriteLOG("ｴﾗｰｺｰﾄﾞ0受信OK")
                    End If
            End Select

            'ACK1送信--------------------------------
            Call WriteLOG("ACK1（Chr(&1031)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H31)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK1送信OK")

            'EOT受信確認
            Call WriteLOG("EOT（Chr(&H4)）受信確認")
            If cRS232C.RECEIVE_BYTE_DATA(0) <> &H4 Then
                mErrorComment = "EOT受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("EOT受信確認OK")

            Debug.Print("データ受信OK")
            ComMode_Change = True

        Catch
            Call WriteLOG(mErrorComment)
        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Function

    Public Function ComMode_Set(ByRef GA As String) As Boolean

        Dim sData() As Byte

        ComMode_Set = False
        ReDim sData(26)
        'Header ===============================
        sData(0) = &H1 : sData(1) = &H4D
        'Length of data 4-N(Byte) =============
        sData(2) = &H0 : sData(3) = &H16
        'STX$ =================
        sData(4) = &H2
        'Mコマンド ============
        sData(5) = &HF1
        '通信ID ===============
        sData(6) = &H0 : sData(7) = &H0 : sData(8) = &H0 : sData(9) = &H0 : sData(10) = &H0
        sData(11) = &H0 : sData(12) = &H0 : sData(13) = &H0 : sData(14) = &H0 : sData(15) = &H0
        sData(16) = &H0
        'SlotNo===============
        sData(17) = &H2 'ｽﾛｯﾄは2ｽﾛｯﾄ目
        'UnitNo===============
        sData(18) = &H0
        '要求コマンド
        'Main ===================================
        sData(19) = &H4
        'SUB ====================================
        sData(20) = &H2
        'データサイズ============================
        sData(21) = &H0 : sData(22) = &H2
        'リザーブ ================================
        sData(23) = &H0 : sData(24) = &H0
        'ETX$ ===================================
        sData(25) = &H3
        'Sum=====
        Dim intSData As Integer = 0
        For i = 1 To 25
            intSData = intSData + sData(i)
        Next i
        sData(26) = CByte(intSData Mod 256)

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        If GA = "KP" Then cRS232C.SETTING = "115200,N,8,1"
        If GA = "PF" Then cRS232C.SETTING = "19200,E,8,1"

        Try
            '------------------------------------------------------
            'ポートを開く
            Call WriteLOG("RS232Cﾎﾟｰﾄｵｰﾌﾟﾝ")
            Try                             'V3.02
                Call cRS232C.PortSetup()
            Catch ex As Exception
                mErrorComment = "ﾎﾟｰﾄｵｰﾌﾟﾝ失敗"
                Throw New ApplicationException(mErrorComment)
            End Try
            '------------------------------------------------------
            Call WriteLOG("RS232CﾎﾟｰﾄｵｰﾌﾟﾝOK")

            mSecWait(300)

            'ENQ送信
            Call WriteLOG("ENQ（Chr(&H5)）送信")
            cRS232C.SendCmd = ENQ
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ENQ送信OK")

            'ACK0受信確認
            Call WriteLOG("ACK0（Chr(&H1030)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK0 Then
                mErrorComment = "ACK0受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK0受信確認OK")

            Call mSecWait(300)
            Call cRS232C.RecBufIni()

            '送信ﾃﾞｰﾀ送信
            Call WriteLOG("ﾃﾞｰﾀ送信")
            cRS232C.SendByteData = sData
            Try
                Call cRS232C.FuncSendByteData()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'ACK1受信確認
            Call WriteLOG("ACK1（Chr(&H1031)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK1 Then
                mErrorComment = "ACK1受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK1受信確認OK")

            'EOT送信-----------------------------
            Call WriteLOG("EOT（Chr(&H4)）送信")
            cRS232C.SendCmd = EOT
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("EOT送信OK")

            'ENQ受信確認
            Call WriteLOG("ENQ（Chr(&H5)）受信確認")
            If cRS232C.RECEIVE_DATA <> ENQ Then
                mErrorComment = "ENQ受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ENQ受信確認OK")

            'ACK0送信--------------------------------
            Call WriteLOG("ACK0（Chr(&1030)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H30)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK0送信OK")

            'ﾃﾞｰﾀ受信（KPのとき。PFはindex-1）
            'RData(0):0x1
            'RData(1):0x4D
            'RData(2):0x0
            'RData(3):0x16
            'RData(4):0x2----STX
            'RData(5):0xF1---M
            'RData(6):0x0
            'RData(7):0x0
            'RData(8):0x0
            'RData(9):0x0
            'RData(10):0x0
            'RData(11):0x0
            'RData(12):0x0
            'RData(13):0x0
            'RData(14):0x0
            'RData(15):0x0
            'RData(16):0x0
            'RData(17):0x1----Slot no
            'RData(18):0x0----Unit no
            'RData(19):0x4----$0
            'RData(20):0x2----$1
            'RData(21):0x0----$2
            'RData(22):0x2----$3
            'RData(23):0x0----$4:ErrorCode
            'RData(24):0x0----$5:ErrorCode
            'RData(25):0x3----ETX
            'RData(26):0x62---SUM

            Select Case GA
                Case "KP"
                    If cRS232C.RECEIVE_BYTE_DATA(24) <> 0 Or cRS232C.RECEIVE_BYTE_DATA(25) <> 0 Then
                        Debug.Print("エラーコード：" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25)))
                        mErrorComment = "ｴﾗｰｺｰﾄﾞ受信:" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25))
                        Throw New ApplicationException(mErrorComment)
                    Else
                        Call WriteLOG("ｴﾗｰｺｰﾄﾞ0受信OK")
                    End If
                Case "PF"
                    If cRS232C.RECEIVE_BYTE_DATA(23) <> 0 Or cRS232C.RECEIVE_BYTE_DATA(24) <> 0 Then
                        Debug.Print("エラーコード：" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25)))
                        mErrorComment = "ｴﾗｰｺｰﾄﾞ受信:" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25))
                        Throw New ApplicationException(mErrorComment)
                    Else
                        Call WriteLOG("ｴﾗｰｺｰﾄﾞ0受信OK")
                    End If
            End Select
            
            'ACK1送信--------------------------------
            Call WriteLOG("ACK1（Chr(&1031)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H31)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK1送信OK")

            'EOT受信確認
            Call WriteLOG("EOT（Chr(&H4)）受信確認")
            If cRS232C.RECEIVE_BYTE_DATA(0) <> &H4 Then
                mErrorComment = "EOT受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("EOT受信確認OK")

            Debug.Print("データ受信OK")
            ComMode_Set = True

        Catch
            Call WriteLOG(mErrorComment)
        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Function

    Public Function ComMode_Change_forMP30() As Boolean

        Dim sData() As Byte

        ComMode_Change_forMP30 = False
        ReDim sData(28)
        'Header ===============================
        sData(0) = &H1 : sData(1) = &H4D
        'Length of data 4-N(Byte) =============
        sData(2) = &H0 : sData(3) = &H18   'STX$～ETX$のデータ数
        'STX$ =================
        sData(4) = &H2
        'Mコマンド ============
        sData(5) = &HF1
        '通信ID ===============
        sData(6) = &H0 : sData(7) = &H0 : sData(8) = &H0 : sData(9) = &H0 : sData(10) = &H0
        sData(11) = &H0 : sData(12) = &H0 : sData(13) = &H0 : sData(14) = &H0 : sData(15) = &H0
        sData(16) = &H0
        'SlotNo===============
        sData(17) = &H2 'ｽﾛｯﾄは2ｽﾛｯﾄ目
        'UnitNo===============
        sData(18) = &H0
        '要求コマンド
        'Main ===================================
        sData(19) = &H4
        'SUB ====================================
        sData(20) = &H1
        'データサイズ============================
        sData(21) = &H0 : sData(22) = &H4
        '通信モード ================================
        sData(23) = &H0
        sData(24) = &H2 'PFの時は19200bps偶数ﾊﾟﾘﾃｨ
        '確定時間 ================================
        sData(25) = &HFF : sData(26) = &HFF
        'ETX$ ===================================
        sData(27) = &H3
        'Sum=====
        Dim intSData As Integer = 0
        For i = 1 To 27
            intSData = intSData + sData(i)
        Next i
        sData(28) = CByte(intSData Mod 256)

        Dim intRetry As Integer = 0

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "9600,E,8,1"  '固定

        Try
            '------------------------------------------------------
            'ポートを開く
            Call WriteLOG("RS232Cﾎﾟｰﾄｵｰﾌﾟﾝ")
            Try                             'V3.02
                Call cRS232C.PortSetup()
            Catch ex As Exception
                mErrorComment = "ﾎﾟｰﾄｵｰﾌﾟﾝ失敗"
                Throw New ApplicationException(mErrorComment)
            End Try
            '------------------------------------------------------
            Call WriteLOG("RS232CﾎﾟｰﾄｵｰﾌﾟﾝOK")

            mSecWait(300)

            'ENQ送信
            Call WriteLOG("ENQ（Chr(&H5)）送信")
            cRS232C.SendCmd = ENQ
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ENQ送信OK")

            'ACK0受信確認
            Call WriteLOG("ACK0（Chr(&H1030)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK0 Then
                mErrorComment = "ACK0受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK0受信確認OK")

            Call mSecWait(300)
            Call cRS232C.RecBufIni()

            '送信ﾃﾞｰﾀ送信
            Call WriteLOG("ﾃﾞｰﾀ送信")
            cRS232C.SendByteData = sData
            Try
                Call cRS232C.FuncSendByteData()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'ACK1受信確認
            Call WriteLOG("ACK1（Chr(&H1031)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK1 Then
                mErrorComment = "ACK1受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK1受信確認OK")

            'EOT送信-----------------------------
            Call WriteLOG("EOT（Chr(&H4)）送信")
            cRS232C.SendCmd = EOT
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("EOT送信OK")

            'ENQ受信確認
            Call WriteLOG("ENQ（Chr(&H5)）受信確認")
            If cRS232C.RECEIVE_DATA <> ENQ Then
                mErrorComment = "ENQ受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ENQ受信確認OK")

            'Call cRS232C.RecBufIni()

            'ACK0送信--------------------------------
            Call WriteLOG("ACK0（Chr(&1030)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H30)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK0送信OK")

            'ﾃﾞｰﾀ受信(KPのとき。PFはindex-1）
            'RData(0):0x0
            'RData(1):0x1
            'RData(2):0x4D
            'RData(3):0x0
            'RData(4):0x16
            'RData(5):0x2----STX
            'RData(6):0xF1---M
            'RData(7):0x0
            'RData(8):0x0
            'RData(9):0x0
            'RData(10):0x0
            'RData(11):0x0
            'RData(12):0x0
            'RData(13):0x0
            'RData(14):0x0
            'RData(15):0x0
            'RData(16):0x0
            'RData(17):0x0
            'RData(18):0x1----Slot no
            'RData(19):0x0----Unit no
            'RData(20):0x4----$0
            'RData(21):0x1----$1
            'RData(22):0x0----$2
            'RData(23):0x2----$3
            'RData(24):0x0----$4:ErrorCode
            'RData(25):0x0----$5:ErrorCode
            'RData(26):0x3----ETX
            'RData(27):0x61---SUM

          
            If cRS232C.RECEIVE_BYTE_DATA(21) <> 0 Or cRS232C.RECEIVE_BYTE_DATA(22) <> 0 Then
                Debug.Print("エラーコード：" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25)))
                mErrorComment = "ｴﾗｰｺｰﾄﾞ受信:" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25))
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLOG("ｴﾗｰｺｰﾄﾞ0受信OK")
            End If

            'ACK1送信--------------------------------
            Call WriteLOG("ACK1（Chr(&1031)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H31)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK1送信OK")

            'EOT受信確認
            Call WriteLOG("EOT（Chr(&H4)）受信確認")
            If cRS232C.RECEIVE_BYTE_DATA(0) <> &H4 Then
                mErrorComment = "EOT受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("EOT受信確認OK")

            Debug.Print("データ受信OK")
            ComMode_Change_forMP30 = True

        Catch
            Call WriteLOG(mErrorComment)
        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Function

    Public Function ComMode_Set_forMP30() As Boolean

        Dim sData() As Byte

        ComMode_Set_forMP30 = False
        ReDim sData(26)
        'Header ===============================
        sData(0) = &H1 : sData(1) = &H4D
        'Length of data 4-N(Byte) =============
        sData(2) = &H0 : sData(3) = &H16
        'STX$ =================
        sData(4) = &H2
        'Mコマンド ============
        sData(5) = &HF1
        '通信ID ===============
        sData(6) = &H0 : sData(7) = &H0 : sData(8) = &H0 : sData(9) = &H0 : sData(10) = &H0
        sData(11) = &H0 : sData(12) = &H0 : sData(13) = &H0 : sData(14) = &H0 : sData(15) = &H0
        sData(16) = &H0
        'SlotNo===============
        sData(17) = &H1 'ｽﾛｯﾄは1ｽﾛｯﾄ目
        'UnitNo===============
        sData(18) = &H0
        '要求コマンド
        'Main ===================================
        sData(19) = &H4
        'SUB ====================================
        sData(20) = &H2
        'データサイズ============================
        sData(21) = &H0 : sData(22) = &H2
        'リザーブ ================================
        sData(23) = &H0 : sData(24) = &H0
        'ETX$ ===================================
        sData(25) = &H3
        'Sum=====
        Dim intSData As Integer = 0
        For i = 1 To 25
            intSData = intSData + sData(i)
        Next i
        sData(26) = CByte(intSData Mod 256)

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "19200,E,8,1"

        Try
            '------------------------------------------------------
            'ポートを開く
            Call WriteLOG("RS232Cﾎﾟｰﾄｵｰﾌﾟﾝ")
            Try                             'V3.02
                Call cRS232C.PortSetup()
            Catch ex As Exception
                mErrorComment = "ﾎﾟｰﾄｵｰﾌﾟﾝ失敗"
                Throw New ApplicationException(mErrorComment)
            End Try
            '------------------------------------------------------
            Call WriteLOG("RS232CﾎﾟｰﾄｵｰﾌﾟﾝOK")

            mSecWait(300)

            'ENQ送信
            Call WriteLOG("ENQ（Chr(&H5)）送信")
            cRS232C.SendCmd = ENQ
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ENQ送信OK")

            'ACK0受信確認
            Call WriteLOG("ACK0（Chr(&H1030)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK0 Then
                mErrorComment = "ACK0受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK0受信確認OK")

            Call mSecWait(300)
            Call cRS232C.RecBufIni()

            '送信ﾃﾞｰﾀ送信
            Call WriteLOG("ﾃﾞｰﾀ送信")
            cRS232C.SendByteData = sData
            Try
                Call cRS232C.FuncSendByteData()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'ACK1受信確認
            Call WriteLOG("ACK1（Chr(&H1031)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK1 Then
                mErrorComment = "ACK1受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK1受信確認OK")

            'EOT送信-----------------------------
            Call WriteLOG("EOT（Chr(&H4)）送信")
            cRS232C.SendCmd = EOT
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("EOT送信OK")

            'ENQ受信確認
            Call WriteLOG("ENQ（Chr(&H5)）受信確認")
            If cRS232C.RECEIVE_DATA <> ENQ Then
                mErrorComment = "ENQ受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ENQ受信確認OK")

            'ACK0送信--------------------------------
            Call WriteLOG("ACK0（Chr(&1030)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H30)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK0送信OK")

            'ﾃﾞｰﾀ受信（KPのとき。PFはindex-1）
            'RData(0):0x1
            'RData(1):0x4D
            'RData(2):0x0
            'RData(3):0x16
            'RData(4):0x2----STX
            'RData(5):0xF1---M
            'RData(6):0x0
            'RData(7):0x0
            'RData(8):0x0
            'RData(9):0x0
            'RData(10):0x0
            'RData(11):0x0
            'RData(12):0x0
            'RData(13):0x0
            'RData(14):0x0
            'RData(15):0x0
            'RData(16):0x0
            'RData(17):0x1----Slot no
            'RData(18):0x0----Unit no
            'RData(19):0x4----$0
            'RData(20):0x2----$1
            'RData(21):0x0----$2
            'RData(22):0x2----$3
            'RData(23):0x0----$4:ErrorCode
            'RData(24):0x0----$5:ErrorCode
            'RData(25):0x3----ETX
            'RData(26):0x62---SUM

            If cRS232C.RECEIVE_BYTE_DATA(21) <> 0 Or cRS232C.RECEIVE_BYTE_DATA(22) <> 0 Then
                Debug.Print("エラーコード：" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25)))
                mErrorComment = "ｴﾗｰｺｰﾄﾞ受信:" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25))
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLOG("ｴﾗｰｺｰﾄﾞ0受信OK")
            End If

            'ACK1送信--------------------------------
            Call WriteLOG("ACK1（Chr(&1031)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H31)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK1送信OK")

            'EOT受信確認
            Call WriteLOG("EOT（Chr(&H4)）受信確認")
            If cRS232C.RECEIVE_BYTE_DATA(0) <> &H4 Then
                mErrorComment = "EOT受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("EOT受信確認OK")

            Debug.Print("データ受信OK")
            ComMode_Set_forMP30 = True

        Catch
            Call WriteLOG(mErrorComment)
        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Function

    Public Function RTC_SET() As Boolean

        Dim sData() As Byte
        Dim sDate As String
        Dim sTime As String
        Dim sDyy As String
        Dim sDmm As String
        Dim sDdd As String
        Dim sThh As String
        Dim sTmm As String
        Dim sTss As String

        sDate = Now.ToString("yyMMdd") : sTime = Now.ToString("HHmmss")

        sDyy = Mid(sDate, 1, 2) : sDmm = Mid(sDate, 3, 2) : sDdd = Mid(sDate, 5, 2)
        sThh = Mid(sTime, 1, 2) : sTmm = Mid(sTime, 3, 2) : sTss = Mid(sTime, 5, 2)

        RTC_SET = False

        ReDim sData(28)

        'Header ===============================
        sData(0) = &H1 : sData(1) = &H4D
        'Length of data 4-N(Byte) =============
        sData(2) = &H0 : sData(3) = &H18
        'STX$ =================
        sData(4) = &H2
        'Mコマンド ============
        sData(5) = &H78
        '通信ID ===============
        sData(6) = &H0 : sData(7) = &H0 : sData(8) = &H0 : sData(9) = &H0

        sData(10) = &H0 : sData(11) = &H80

        'sData(10) = 0 : sData(11) = CByte(CInt(sDyy))
        sData(12) = CByte(CInt(sDyy)) : sData(13) = 0   '0  'CByte(CInt(sDmm))
        sData(14) = CByte(CInt(sDmm)) : sData(15) = 0   '0  'CByte(CInt(sDdd))
        sData(16) = CByte(CInt(sDdd)) : sData(17) = 0   '0

        sData(18) = 0 : sData(19) = &H80    'CByte(CInt(sThh))

        sData(20) = CByte(CInt(sThh)) : sData(21) = 0   '0       'CByte(CInt(sTmm))
        sData(22) = CByte(CInt(sTmm)) : sData(23) = 0   '0       'CByte(CInt(sTss))
        sData(24) = CByte(CInt(sTss)) : sData(25) = 0   '0

        sData(26) = 0 : sData(27) = &H3

        'Sum=====
        Dim intSData As Integer = 0
        For i = 1 To 27
            intSData = intSData + sData(i)
        Next i
        sData(28) = CByte(intSData Mod 256)


        Dim intRetry As Integer = 0

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "9600,E,8,1"  '固定

        Try
            '------------------------------------------------------
            'ポートを開く
            Call WriteLOG("RS232Cﾎﾟｰﾄｵｰﾌﾟﾝ")
            Try                             'V3.02
                Call cRS232C.PortSetup()
            Catch ex As Exception
                mErrorComment = "ﾎﾟｰﾄｵｰﾌﾟﾝ失敗"
                Throw New ApplicationException(mErrorComment)
            End Try
            '------------------------------------------------------
            Call WriteLOG("RS232CﾎﾟｰﾄｵｰﾌﾟﾝOK")

            mSecWait(300)

            'ENQ送信
            Call WriteLOG("ENQ（Chr(&H5)）送信")
            cRS232C.SendCmd = ENQ
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ENQ送信OK")

            'ACK0受信確認
            Call WriteLOG("ACK0（Chr(&H1030)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK0 Then
                mErrorComment = "ACK0受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK0受信確認OK")

            Call mSecWait(300)
            Call cRS232C.RecBufIni()

            '送信ﾃﾞｰﾀ送信
            Call WriteLOG("ﾃﾞｰﾀ送信")
            cRS232C.SendByteData = sData
            Try
                Call cRS232C.FuncSendByteData()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'ACK1受信確認
            Call WriteLOG("ACK1（Chr(&H1031)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK1 Then
                mErrorComment = "ACK1受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK1受信確認OK")

            'EOT送信-----------------------------
            Call WriteLOG("EOT（Chr(&H4)）送信")
            cRS232C.SendCmd = EOT
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("EOT送信OK")

            'ENQ受信確認
            Call WriteLOG("ENQ（Chr(&H5)）受信確認")
            If cRS232C.RECEIVE_DATA <> ENQ Then
                mErrorComment = "ENQ受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ENQ受信確認OK")

            'Call cRS232C.RecBufIni()

            'ACK0送信--------------------------------
            Call WriteLOG("ACK0（Chr(&1030)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H30)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK0送信OK")

            'ﾃﾞｰﾀ受信(KPのとき。PFはindex-1）
            'RData(0):0x0
            'RData(1):0x1
            'RData(2):0x4D
            'RData(3):0x0
            'RData(4):0x16
            'RData(5):0x2----STX
            'RData(6):0xF1---M
            'RData(7):0x0
            'RData(8):0x0
            'RData(9):0x0
            'RData(10):0x0
            'RData(11):0x0
            'RData(12):0x0
            'RData(13):0x0
            'RData(14):0x0
            'RData(15):0x0
            'RData(16):0x0
            'RData(17):0x0
            'RData(18):0x1----Slot no
            'RData(19):0x0----Unit no
            'RData(20):0x2----$0
            'RData(21):0x12----$1
            'RData(22):0x0----$2
            'RData(23):0x2----$3
            'RData(24):0x0----$4:ErrorCode
            'RData(25):0x0----$5:ErrorCode
            'RData(26):0x3----ETX
            'RData(27):0x61---SUM

            'ACK1送信--------------------------------
            Call WriteLOG("ACK1（Chr(&1031)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H31)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK1送信OK")

            'EOT受信確認
            Call WriteLOG("EOT（Chr(&H4)）受信確認")
            If cRS232C.RECEIVE_BYTE_DATA(0) <> &H4 Then
                mErrorComment = "EOT受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("EOT受信確認OK")

            Debug.Print("データ受信OK")
            RTC_SET = True

        Catch
            Call WriteLOG(mErrorComment)
        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Function

    Public Function RTC_READ() As String

        Dim sData() As Byte

        RTC_READ = ""
        'ReDim sData(24)
        ReDim sData(12)

        'Header ===============================
        sData(0) = &H1 : sData(1) = &H4D
        'Length of data 4-N(Byte) =============
        sData(2) = &H0 : sData(3) = &H8
        'STX$ =================
        sData(4) = &H2
        'Mコマンド ============
        sData(5) = &H79
        '通信ID ===============
        sData(6) = &H0 : sData(7) = &H0 : sData(8) = &H0 : sData(9) = &H0 : sData(10) = &H0
        sData(11) = &H3
        'Sum=====
        Dim intSData As Integer = 0
        For i = 1 To 11
            intSData = intSData + sData(i)
        Next i
        sData(12) = CByte(intSData Mod 256)

        cRS232C = New cls101_ComRS232C
        cRS232C.RS_COMM = mRSComm
        cRS232C.PORT = mPort
        cRS232C.SETTING = "9600,E,8,1"  '固定

        Try
            '------------------------------------------------------
            'ポートを開く
            Call WriteLOG("RS232Cﾎﾟｰﾄｵｰﾌﾟﾝ")
            Try                             'V3.02
                Call cRS232C.PortSetup()
            Catch ex As Exception
                mErrorComment = "ﾎﾟｰﾄｵｰﾌﾟﾝ失敗"
                Throw New ApplicationException(mErrorComment)
            End Try
            '------------------------------------------------------
            Call WriteLOG("RS232CﾎﾟｰﾄｵｰﾌﾟﾝOK")

            mSecWait(300)

            'ENQ送信
            Call WriteLOG("ENQ（Chr(&H5)）送信")
            cRS232C.SendCmd = ENQ
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ENQ送信OK")

            'ACK0受信確認
            Call WriteLOG("ACK0（Chr(&H1030)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK0 Then
                mErrorComment = "ACK0受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK0受信確認OK")

            Call mSecWait(300)
            Call cRS232C.RecBufIni()

            '送信ﾃﾞｰﾀ送信
            Call WriteLOG("ﾃﾞｰﾀ送信")
            cRS232C.SendByteData = sData
            Try
                Call cRS232C.FuncSendByteData()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'ACK1受信確認
            Call WriteLOG("ACK1（Chr(&H1031)）受信確認")
            If cRS232C.RECEIVE_DATA <> ACK1 Then
                mErrorComment = "ACK1受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ACK1受信確認OK")

            'EOT送信-----------------------------
            Call WriteLOG("EOT（Chr(&H4)）送信")
            cRS232C.SendCmd = EOT
            Try
                Call cRS232C.FuncSendCmd()
                Call cRS232C.FuncRecvByteData()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("EOT送信OK")

            'ENQ受信確認
            Call WriteLOG("ENQ（Chr(&H5)）受信確認")
            If cRS232C.RECEIVE_DATA <> ENQ Then
                mErrorComment = "ENQ受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("ENQ受信確認OK")

            'ACK0送信--------------------------------
            Call WriteLOG("ACK0（Chr(&1030)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H30)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK0送信OK")

            'ﾃﾞｰﾀ受信（KPのとき。PFはindex-1）
            'RData(0):0x1
            'RData(1):0x4D
            'RData(2):0x0
            'RData(3):0x18
            'RData(4):0x2----STX
            'RData(5):0x79
            'RData(6):0x0
            'RData(7):0x0
            'RData(8):0x0
            'RData(9):0x0
            'RData(10):0x0
            'RData(11):0x0
            'RData(12):0x*  年
            'RData(13):0x0
            'RData(14):0x*  月
            'RData(15):0x0
            'RData(16):0x*  日
            'RData(17):0x0
            'RData(18):0x0
            'RData(19):0x0
            'RData(20):0x*  HH
            'RData(21):0x0
            'RData(22):0x*  MM
            'RData(23):0x0
            'RData(24):0x*  DD
            'RData(25):0x0
            'RData(26):0x0
            'RData(27):0x3----ETX
            'RData(28):0x*----SUM

            RTC_READ = cRS232C.RECEIVE_BYTE_DATA(12).ToString("D2") & "/" & cRS232C.RECEIVE_BYTE_DATA(14).ToString("D2") & "/" & cRS232C.RECEIVE_BYTE_DATA(16).ToString("D2")
            RTC_READ &= cRS232C.RECEIVE_BYTE_DATA(20).ToString("D2") & ":" & cRS232C.RECEIVE_BYTE_DATA(22).ToString("D2") & ":" & cRS232C.RECEIVE_BYTE_DATA(24).ToString("D2")

            'If cRS232C.RECEIVE_BYTE_DATA(23) <> 0 Or cRS232C.RECEIVE_BYTE_DATA(24) <> 0 Then
            '    Debug.Print("エラーコード：" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25)))
            '    mErrorComment = "ｴﾗｰｺｰﾄﾞ受信:" & Hex(cRS232C.RECEIVE_BYTE_DATA(24)) & Hex(cRS232C.RECEIVE_BYTE_DATA(25))
            '    Throw New ApplicationException(mErrorComment)
            'Else
            '    Call WriteLOG("ｴﾗｰｺｰﾄﾞ0受信OK")
            'End If

            'ACK1送信--------------------------------
            Call WriteLOG("ACK1（Chr(&1031)）送信")
            cRS232C.SendCmd = Chr(&H10) & Chr(&H31)
            Try
                Call cRS232C.FuncSendCmd()
                Call mSecWait(100)
                Call cRS232C.FuncRecvByteData2()
            Catch ex As Exception
                mErrorComment = "RS232C送受信ｴﾗｰ" : Debug.Print(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLOG("ACK1送信OK")

            'EOT受信確認
            Call WriteLOG("EOT（Chr(&H4)）受信確認")
            If cRS232C.RECEIVE_BYTE_DATA(0) <> &H4 Then
                mErrorComment = "EOT受信ｴﾗｰ" : Debug.Print(mErrorComment & cRS232C.RECEIVE_DATA)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLOG("EOT受信確認OK")

            Debug.Print("データ受信OK")

        Catch
            Call WriteLOG(mErrorComment)
        Finally
            cRS232C.PortClose()
            cRS232C = Nothing
        End Try

    End Function

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
            DoEvents()
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

    Private Sub WriteLOG(ByVal mComment As String)
        mLOG &= mComment & vbCrLf
    End Sub

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
    End Sub

End Class
