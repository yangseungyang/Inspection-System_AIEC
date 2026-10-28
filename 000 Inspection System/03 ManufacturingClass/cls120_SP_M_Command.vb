Option Strict On : Option Explicit On


'Doevents
Imports System.Windows.Forms.Application

Public Class cls120_SP_M_Command

    Private objNetStream As System.Net.Sockets.NetworkStream
    Private objTcpClient As System.Net.Sockets.TcpClient

    Dim mErrorComment As String
    Dim mRecvData As String

    'Local -------------------
    Dim mSendByte() As Byte
    Dim mRecvByte() As Byte
    Dim mRecvDataSize As Long
    Dim mRecvErrCode As Long
    Dim bytSendBuffer As Byte()
    Dim bytReceiveBuffer(255) As Byte

    'Property ----------------
    Dim mIPAdd As String
    Dim mRemPort As Integer
    Dim mUnitNo As Integer
    Dim mSlotNo As Integer
    Dim mTimeout As Integer = 0
    Dim mSendText As String
    Dim mRecvText As String
    Dim mError As String

    Dim mModuleType As Integer
    Dim mEndianMode As Integer

    Dim mIO_Write_Byte() As Byte = {&HF1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, &H10, &H8, 0, &HC, 0, 0, 0, 2, 0, 1, 0, 0, 0, 0, 0, 0}
    Dim mIO_Read_Byte() As Byte = {&HF1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, &H10, &H7, 0, &HA, 0, 0, 0, 2, 0, 1, 0, 0, 0, 0}
  
    'IPAddress -----------------------------
    Public WriteOnly Property IPAdd() As String
        Set(ByVal IP As String)
            mIPAdd = IP
        End Set
    End Property

    'Remort Port ====================================
    Public WriteOnly Property RemPort() As Integer
        Set(ByVal rp As Integer)
            mRemPort = rp
        End Set
    End Property

    'UnitNo ============================
    Public WriteOnly Property UnitNo() As Integer
        Set(ByVal un As Integer)
            mUnitNo = un
        End Set
    End Property

    'SlotNo ============================
    Public WriteOnly Property SlotNo() As Integer
        Set(ByVal sl As Integer)
            mSlotNo = sl
        End Set
    End Property

    'Time Out ==============================
    Public WriteOnly Property TimeOut() As Integer
        Set(ByVal NewTimeOut As Integer)
            mTimeout = NewTimeOut
        End Set
    End Property

    'ModuleType ==============================
    Public WriteOnly Property ModuleType() As Integer
        Set(ByVal NewModuleType As Integer)
            mModuleType = NewModuleType
        End Set
    End Property

    'EndianType ==============================
    Public WriteOnly Property EndianMode() As Integer
        Set(ByVal NewEndianMode As Integer)
            mEndianMode = NewEndianMode
        End Set
    End Property


    'Send Text ==============================
    Public ReadOnly Property SendText() As String
        Get
            SendText = mSendText
        End Get
    End Property

    'Recv Text ==============================
    Public ReadOnly Property RecvText() As String
        Get
            RecvText = mRecvText
        End Get
    End Property

    'Error message ===========================
    Public ReadOnly Property ERR_MSG() As String
        Get
            ERR_MSG = mError
        End Get
    End Property

    Public Sub CPU_RESTERT()
        '==========================================================================================
        ' Function : CPU Restart
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(21)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H0
        'SUB ============================================
        mSendByte(15) = &H1
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H4

        'CPU (Myself) ===================================
        mSendByte(18) = &H0
        mSendByte(19) = &H0
        mSendByte(20) = &H0 : mSendByte(21) = &H0

        Call F3MCommandSend()
        If mError <> "" Then Exit Sub

    End Sub

    Public Sub ProductInfoREAD()
        '==========================================================================================
        ' Function : ProductInfoREAD
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(17)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H1
        'SUB ============================================
        mSendByte(15) = &H9
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H0

        Call F3MCommandSendRcv()

    End Sub

    Public Sub RTC_READ()
        '==========================================================================================
        ' Function : RTC Read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(17)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H2
        'SUB ============================================
        mSendByte(15) = &H11
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H0

        '        Dim i As Integer = 0
        '        Try
        'retry:
        '            Call F3MCommandSendRcv()
        '        Catch ex As Exception
        '            If i = 0 Then
        '                Call mSecWait(300) : i += 1 : GoTo retry
        '            End If
        '        End Try

        Call F3MCommandSendRcv()

    End Sub

    Public Sub RTC_SET()
        '==========================================================================================
        ' Function : RTC Set
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  : Present time
        '==========================================================================================

        Dim sDate As String
        Dim sTime As String
        Dim sDyy As String
        Dim sDmm As String
        Dim sDdd As String
        Dim sThh As String
        Dim sTmm As String
        Dim sTss As String
        Dim bDyy(1) As String
        Dim bDmm(1) As String
        Dim bDdd(1) As String
        Dim bThh(1) As String
        Dim bTmm(1) As String
        Dim bTss(1) As String

        sDate = Now.ToString("yyMMdd") : sTime = Now.ToString("HHmmss")

        sDyy = Mid(sDate, 1, 2) : sDmm = Mid(sDate, 3, 2) : sDdd = Mid(sDate, 5, 2)
        sThh = Mid(sTime, 1, 2) : sTmm = Mid(sTime, 3, 2) : sTss = Mid(sTime, 5, 2)

        bDyy(0) = Mid(sDyy, 1, 1) : bDyy(1) = Mid(sDyy, 2, 1)
        bDmm(0) = Mid(sDmm, 1, 1) : bDmm(1) = Mid(sDmm, 2, 1)
        bDdd(0) = Mid(sDdd, 1, 1) : bDdd(1) = Mid(sDdd, 2, 1)

        bThh(0) = Mid(sThh, 1, 1) : bThh(1) = Mid(sThh, 2, 1)
        bTmm(0) = Mid(sTmm, 1, 1) : bTmm(1) = Mid(sTmm, 2, 1)
        bTss(0) = Mid(sTss, 1, 1) : bTss(1) = Mid(sTss, 2, 1)

        ReDim mSendByte(37)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H2
        'SUB ============================================
        mSendByte(15) = &H12
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H14

        'Date Time ======================================
        'yy/
        mSendByte(18) = CByte(Asc(bDyy(0))) : mSendByte(19) = CByte(Asc(bDyy(1))) : mSendByte(20) = &H2F
        'mSendByte(18) = bDyy(0) : mSendByte(19) = bDyy(1) : mSendByte(20) = &H2F
        'mm/
        mSendByte(21) = CByte(Asc(bDmm(0))) : mSendByte(22) = CByte(Asc(bDmm(1))) : mSendByte(23) = &H2F
        'mSendByte(21) = bDmm(0) : mSendByte(22) = bDmm(1) : mSendByte(23) = &H2F
        'dd/
        mSendByte(24) = CByte(Asc(bDdd(0))) : mSendByte(25) = CByte(Asc(bDdd(1)))
        'mSendByte(24) = bDdd(0) : mSendByte(25) = bDdd(1)

        'hh:
        mSendByte(26) = CByte(Asc(bThh(0))) : mSendByte(27) = CByte(Asc(bThh(1))) : mSendByte(28) = &H3A
        'mSendByte(26) = bThh(0) : mSendByte(27) = bThh(1) : mSendByte(28) = &H3A
        'mm:
        mSendByte(29) = CByte(Asc(bTmm(0))) : mSendByte(30) = CByte(Asc(bTmm(1))) : mSendByte(31) = &H3A
        'mSendByte(29) = bTmm(0) : mSendByte(30) = bTmm(1) : mSendByte(31) = &H3A
        'ss
        mSendByte(32) = CByte(Asc(bTss(0))) : mSendByte(33) = CByte(Asc(bTss(1)))
        'mSendByte(32) = bTss(0) : mSendByte(33) = bTss(1)

        '+
        mSendByte(34) = &H2B
        '000
        mSendByte(35) = &H30 : mSendByte(36) = &H30 : mSendByte(37) = &H30

        '        Dim i As Integer = 0
        '        Try
        'retry:
        '            Call F3MCommandSendRcv()
        '        Catch ex As Exception
        '            If i = 0 Then
        '                Call mSecWait(300) : i += 1 : GoTo retry
        '            End If
        '        End Try

        Call F3MCommandSendRcv()

    End Sub

    Public Sub SEQ_RUN()
        '==========================================================================================
        ' Function : SEQ RUN
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(17)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H20
        'SUB ============================================
        mSendByte(15) = &H1
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H0

        'Call F3MCommandSend()
        'If mError <> "" Then Exit Sub

        'Call F3McommRcv()
        'If mError <> "" Then Exit Sub

        Call F3MCommandSendRcv()

    End Sub

    Public Sub SEQ_STOP()
        '==========================================================================================
        ' Function : SEQ STOP
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(17)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H20
        'SUB ============================================
        mSendByte(15) = &H2
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H0

        'Call F3MCommandSend()
        'If mError <> "" Then Exit Sub

        'Call F3McommRcv()
        'If mError <> "" Then Exit Sub

        Call F3MCommandSendRcv()

    End Sub

    Public Sub SEQ_DEBUG()
        '==========================================================================================
        ' Function : SEQ DEBUG Mode
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(17)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H20
        'SUB ============================================
        mSendByte(15) = &H3
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H0

        'Call F3MCommandSend()
        'If mError <> "" Then Exit Sub

        'Call F3McommRcv()
        'If mError <> "" Then Exit Sub

        Call F3MCommandSendRcv()

    End Sub

    Public Sub MEM_CLEAR()
        '==========================================================================================
        ' Function : Memory clear
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(17)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H22
        'SUB ============================================
        mSendByte(15) = &H1
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H0

        'Call F3MCommandSend()
        'If mError <> "" Then Exit Sub

        'Call F3McommRcv()
        'If mError <> "" Then Exit Sub

        Call F3MCommandSendRcv()

    End Sub

    Public Sub DATA_CLEAR()
        '==========================================================================================
        ' Function : Data clear
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(17)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H22
        'SUB ============================================
        mSendByte(15) = &H2
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H0

        'Call F3MCommandSend()
        'If mError <> "" Then Exit Sub

        'Call F3McommRcv()
        'If mError <> "" Then Exit Sub

        Call F3MCommandSendRcv()

    End Sub

    Public Sub IO_READ(ByVal AcsType As Integer, ByVal AcsSize As Integer, ByVal Add As String)
        '==========================================================================================
        ' Function : IO Read
        ' Data     : ｱｸｾｽﾀｲﾌﾟ,ｱｸｾｽｻｲｽﾞ, address
        ' Return   : Nothing
        ' Remarks  : ｱｸｾｽﾀｲﾌﾟ0/ﾒﾓﾘ　1/ﾚｼﾞｽﾀ　:ｱｸｾｽｻｲｽﾞ1/ﾊﾞｲﾄ　2/ﾜｰﾄﾞ　4/ﾛﾝｸﾞﾜｰﾄﾞ
        '==========================================================================================

        ReDim mSendByte(27)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H10
        'SUB ============================================
        mSendByte(15) = &H7
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &HA

        'Access TYPE ====================================
        mSendByte(18) = &H0 : mSendByte(19) = CByte(AcsType)

        'Access SIZE ====================================
        mSendByte(20) = &H0 : mSendByte(21) = CByte(AcsSize)

        'Data Count =====================================
        mSendByte(22) = &H0 : mSendByte(23) = &H1

        'Address ========================================
        mSendByte(24) = CByte("&H" & Mid(Add, 1, 2))
        mSendByte(25) = CByte("&H" & Mid(Add, 3, 2))
        mSendByte(26) = CByte("&H" & Mid(Add, 5, 2))
        mSendByte(27) = CByte("&H" & Mid(Add, 7, 2))

        '        Dim i As Integer = 0
        '        Try
        'retry:
        '            Call F3MCommandSendRcv()
        '        Catch ex As Exception
        '            If i = 0 Then
        '                Call mSecWait(300) : i += 1 : GoTo retry
        '            End If
        '        End Try

        Call F3MCommandSendRcvH()

    End Sub

    Public Sub IO_READ_MEMWORD(ByVal Add As String)
        '==========================================================================================
        ' Function : IO Read
        ' Data     : ｱｸｾｽﾀｲﾌﾟ,ｱｸｾｽｻｲｽﾞ, address
        ' Return   : Nothing
        ' Remarks  : 
        '==========================================================================================
        Dim intDataLength As Integer
        Dim tCount As Double = 0

        mTimeout = 5 'ﾀｲﾑｱｳﾄの設定がなければﾀｲﾑｱｳﾄを5秒とする

        'Address ========================================
        mIO_Read_Byte(24) = CByte("&H" & Mid(Add, 1, 2))
        mIO_Read_Byte(25) = CByte("&H" & Mid(Add, 3, 2))
        mIO_Read_Byte(26) = CByte("&H" & Mid(Add, 5, 2))
        mIO_Read_Byte(27) = CByte("&H" & Mid(Add, 7, 2))

        objNetStream.Write(mIO_Read_Byte, 0, mIO_Read_Byte.Length)
        intDataLength = objNetStream.Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)

        mRecvText = ""
        For i = 12 To intDataLength - 1
            mRecvText &= bytReceiveBuffer(i).ToString("X2")
        Next

    End Sub

    Public Sub IO_WRITE(ByVal AcsType As Integer, ByVal AcsSize As Integer, ByVal Add As String, ByVal DATA As String)
        '==========================================================================================
        ' Function : IO WRITE
        ' Data     : ｱｸｾｽﾀｲﾌﾟ, address
        ' Return   : Nothing
        ' Remarks  : ｱｸｾｽﾀｲﾌﾟ0/ﾒﾓﾘ　1/ﾚｼﾞｽﾀ　　:ｱｸｾｽｻｲｽﾞ1/ﾊﾞｲﾄ　2/ﾜｰﾄﾞ　4/ﾛﾝｸﾞﾜｰﾄﾞ
        '==========================================================================================

        If AcsSize = 1 Then
            ReDim mSendByte(28)
        ElseIf AcsSize = 2 Then
            ReDim mSendByte(29)
        ElseIf AcsSize = 4 Then
            ReDim mSendByte(31)
        End If

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H10
        'SUB ============================================
        mSendByte(15) = &H8
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = CByte(10 + AcsSize)

        'Access TYPE ====================================
        mSendByte(18) = &H0 : mSendByte(19) = CByte(AcsType)

        'Access SIZE=====================================
        mSendByte(20) = &H0 : mSendByte(21) = CByte(AcsSize)

        'DATA COUNT =====================================
        mSendByte(22) = &H0 : mSendByte(23) = &H1

        'Address ========================================
        mSendByte(24) = CByte("&H" & Mid(Add, 1, 2))
        mSendByte(25) = CByte("&H" & Mid(Add, 3, 2))
        mSendByte(26) = CByte("&H" & Mid(Add, 5, 2))
        mSendByte(27) = CByte("&H" & Mid(Add, 7, 2))

        'DATA
        mSendByte(28) = CByte("&H" & Mid(DATA, 1, 2))
        If AcsSize > 1 Then mSendByte(29) = CByte("&H" & Mid(DATA, 3, 2))
        If AcsSize = 4 Then
            mSendByte(30) = CByte("&H" & Mid(DATA, 5, 2))
            mSendByte(31) = CByte("&H" & Mid(DATA, 7, 2))
        End If

        Call F3MCommandSendRcvH()

    End Sub
    Public Sub IO_INTERNEL_WRITE(ByVal AcsType As Integer, ByVal AcsSize As Integer, ByVal Add As String, ByVal DATA As String)
        '==========================================================================================
        ' Function : IO WRITE
        ' Data     : ｱｸｾｽﾀｲﾌﾟ, address
        ' Return   : Nothing
        ' Remarks  : ｱｸｾｽﾀｲﾌﾟ0/ﾒﾓﾘ　1/ﾚｼﾞｽﾀ　　:ｱｸｾｽｻｲｽﾞ1/ﾊﾞｲﾄ　2/ﾜｰﾄﾞ　4/ﾛﾝｸﾞﾜｰﾄﾞ
        '==========================================================================================

        If AcsSize = 1 Then
            ReDim mSendByte(26)
        ElseIf AcsSize = 2 Then
            ReDim mSendByte(27)
        ElseIf AcsSize = 4 Then
            ReDim mSendByte(29)
        End If

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H10
        'SUB ============================================
        mSendByte(15) = &H2
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = CByte(8 + AcsSize)

        'Access TYPE ====================================
        'mSendByte(18) = &H0 : mSendByte(19) = CByte(AcsType)

        'Access SIZE=====================================
        mSendByte(18) = &H0 : mSendByte(19) = CByte(AcsSize)

        'DATA COUNT =====================================
        mSendByte(20) = &H0 : mSendByte(21) = &H1

        ''Address ========================================
        mSendByte(22) = CByte("&H" & Mid(Add, 1, 2))
        mSendByte(23) = CByte("&H" & Mid(Add, 3, 2))
        mSendByte(24) = CByte("&H" & Mid(Add, 5, 2))
        mSendByte(25) = CByte("&H" & Mid(Add, 7, 2))

        'DATA
        mSendByte(26) = CByte("&H" & Mid(DATA, 1, 2))
        If AcsSize > 1 Then mSendByte(27) = CByte("&H" & Mid(DATA, 3, 2))
        If AcsSize = 4 Then
            mSendByte(28) = CByte("&H" & Mid(DATA, 5, 2))
            mSendByte(29) = CByte("&H" & Mid(DATA, 7, 2))
        End If

        Call F3MCommandSendRcv()

    End Sub

    Public Function Connect() As Boolean

        ' ソケット接続
        objTcpClient = New System.Net.Sockets.TcpClient
        objTcpClient.Connect(mIPAdd, mRemPort)
        objNetStream = objTcpClient.GetStream

        Connect = True

    End Function

    Public Function DISConnect() As Boolean

        ' ソケットクローズ
        objNetStream.Close()
        objTcpClient.Close()

        DISConnect = True

    End Function

    Public Sub IO_WRITE_MEMWORD(ByVal Add As String, ByVal DATA As String)
        '==========================================================================================
        ' Function : IO WRITE
        ' Data     : ｱｸｾｽﾀｲﾌﾟ, address
        ' Return   : Nothing
        ' Remarks  : 
        '==========================================================================================

        'Dim tCount As Double = 0

        Try

            With objNetStream

                'Address ========================================
                mIO_Write_Byte(24) = CByte("&H" & Mid(Add, 1, 2))
                mIO_Write_Byte(25) = CByte("&H" & Mid(Add, 3, 2))
                mIO_Write_Byte(26) = CByte("&H" & Mid(Add, 5, 2))
                mIO_Write_Byte(27) = CByte("&H" & Mid(Add, 7, 2))

                'DATA
                mIO_Write_Byte(28) = CByte("&H" & Mid(DATA, 1, 2))
                mIO_Write_Byte(29) = CByte("&H" & Mid(DATA, 3, 2))

                bytSendBuffer = mIO_Write_Byte

                .Write(bytSendBuffer, 0, bytSendBuffer.Length)
                .Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)

            End With

        Catch Ex As Exception
            If mErrorComment = "" Then mErrorComment = "PC Link Communication Error"
            Throw New ApplicationException(mErrorComment)

        Finally

        End Try

    End Sub

    Public Sub MEM_READ(ByVal AcsSize As Integer, ByVal DataCount As Integer, ByVal Add As String)
        '==========================================================================================
        ' Function : Memory Read
        ' Data     : Address
        ' Return   : Nothing
        ' Remarks  : ｱｸｾｽｻｲｽﾞ1/ﾊﾞｲﾄ　2/ﾜｰﾄﾞ　4/ﾛﾝｸﾞﾜｰﾄﾞ
        '==========================================================================================

        ReDim mSendByte(25)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H10
        'SUB ============================================
        mSendByte(15) = &H1
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H8

        'Access SIZE ====================================
        mSendByte(18) = &H0 : mSendByte(19) = CByte(AcsSize)

        'Data Count =====================================
        mSendByte(20) = &H0 : mSendByte(21) = CByte(DataCount)

        'Address ========================================
        mSendByte(22) = CByte("&H" & Mid(Add, 1, 2))
        mSendByte(23) = CByte("&H" & Mid(Add, 3, 2))
        mSendByte(24) = CByte("&H" & Mid(Add, 5, 2))
        mSendByte(25) = CByte("&H" & Mid(Add, 7, 2))

        'Call F3MCommandSend()
        'If mError <> "" Then Exit Sub

        'Call F3McommRcv()
        'If mError <> "" Then Exit Sub

        Call F3MCommandSendRcv()

    End Sub

    Public Sub MEM_WRITE(ByVal AcsSize As Integer, ByVal Add As String, ByVal DATA As String)
        '==========================================================================================
        ' Function : Memory Write
        ' Data     : Address
        ' Return   : Nothing
        ' Remarks  : ｱｸｾｽｻｲｽﾞ1/ﾊﾞｲﾄ　2/ﾜｰﾄﾞ　4/ﾛﾝｸﾞﾜｰﾄﾞ
        '==========================================================================================

        If AcsSize = 1 Then
            ReDim mSendByte(26)
        ElseIf AcsSize = 2 Then
            ReDim mSendByte(27)
        ElseIf AcsSize = 4 Then
            ReDim mSendByte(29)
        End If

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H10
        'SUB ============================================
        mSendByte(15) = &H2
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = CByte(8 + AcsSize)

        'Access SIZE ====================================
        mSendByte(18) = &H0 : mSendByte(19) = CByte(AcsSize)

        'Data Count =====================================
        mSendByte(20) = &H0 : mSendByte(21) = &H1

        'Address ========================================
        mSendByte(22) = CByte("&H" & Mid(Add, 1, 2))
        mSendByte(23) = CByte("&H" & Mid(Add, 3, 2))
        mSendByte(24) = CByte("&H" & Mid(Add, 5, 2))
        mSendByte(25) = CByte("&H" & Mid(Add, 7, 2))

        'DATA
        mSendByte(26) = CByte("&H" & Mid(DATA, 1, 2))
        If AcsSize > 1 Then mSendByte(27) = CByte("&H" & Mid(DATA, 3, 2))
        If AcsSize = 4 Then
            mSendByte(28) = CByte("&H" & Mid(DATA, 5, 2))
            mSendByte(29) = CByte("&H" & Mid(DATA, 7, 2))
        End If

        Call F3MCommandSendRcv()

    End Sub

    Public Sub GATE_ARRAY_READ(ByVal unt As Integer, ByVal slt As Integer, ByVal Add As String)
        '==========================================================================================
        ' Function : Gate array Read
        ' Data     : Address
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(29)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H10
        'SUB ============================================
        mSendByte(15) = &H5
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &HC

        'Access SIZE ====================================
        mSendByte(18) = &H0 : mSendByte(19) = &H1

        'Data Count =====================================
        mSendByte(20) = &H0 : mSendByte(21) = &H1

        'UnitNo(IO Module) ==============================
        mSendByte(22) = &H0 : mSendByte(23) = CByte(unt)

        'SlotNo(IO Module) ==============================
        mSendByte(24) = &H0 : mSendByte(25) = CByte(slt)

        'Address ========================================
        mSendByte(26) = CByte("&H" & Mid(Add, 1, 2))
        mSendByte(27) = CByte("&H" & Mid(Add, 3, 2))
        mSendByte(28) = CByte("&H" & Mid(Add, 5, 2))
        mSendByte(29) = CByte("&H" & Mid(Add, 7, 2))

        'Call F3MCommandSend()
        'If mError <> "" Then Exit Sub

        'Call F3McommRcv()
        'If mError <> "" Then Exit Sub

        Call F3MCommandSendRcv()

    End Sub

    Public Sub GATE_ARRAY_WRITE(ByVal unt As Integer, ByVal slt As Integer, ByVal Add As String, ByVal AcsSize As Integer, ByVal DATA As String)
        '==========================================================================================
        ' Function : Gate array Write
        ' Data     : Address
        ' Return   : Nothing
        ' Remarks  : ｱｸｾｽｻｲｽﾞ1/ﾊﾞｲﾄ　2/ﾜｰﾄﾞ　4/ﾛﾝｸﾞﾜｰﾄﾞ
        '==========================================================================================
        If AcsSize = 1 Then
            ReDim mSendByte(30)
        ElseIf AcsSize = 2 Then
            ReDim mSendByte(31)
        ElseIf AcsSize = 4 Then
            ReDim mSendByte(33)
        End If

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H10
        'SUB ============================================
        mSendByte(15) = &H6
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = CByte(12 + AcsSize)

        'Access SIZE ====================================
        mSendByte(18) = &H0 : mSendByte(19) = CByte(AcsSize)

        'Data Count =====================================
        mSendByte(20) = &H0 : mSendByte(21) = &H1

        'UnitNo(IO Module) ==============================
        mSendByte(22) = &H0 : mSendByte(23) = CByte(unt)

        'SlotNo(IO Module) ==============================
        mSendByte(24) = &H0 : mSendByte(25) = CByte(slt)

        'Address ========================================
        mSendByte(26) = CByte("&H" & Mid(Add, 1, 2))
        mSendByte(27) = CByte("&H" & Mid(Add, 3, 2))
        mSendByte(28) = CByte("&H" & Mid(Add, 5, 2))
        mSendByte(29) = CByte("&H" & Mid(Add, 7, 2))

        'DATA
        mSendByte(30) = CByte("&H" & Mid(DATA, 1, 2))
        If AcsSize > 1 Then mSendByte(31) = CByte("&H" & Mid(DATA, 3, 2))
        If AcsSize = 4 Then
            mSendByte(32) = CByte("&H" & Mid(DATA, 5, 2))
            mSendByte(33) = CByte("&H" & Mid(DATA, 7, 2))
        End If

        Call F3MCommandSendRcv()

    End Sub

    Public Sub WriteRom(ByVal FirmData() As Byte, ByVal WriteAddress As String, ByVal WriteSize As Short)
        '==========================================================================================
        ' F.Name   : WriteRom
        ' Function :
        ' Data     : address,data
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        ReDim mSendByte(29 + WriteSize)

        Dim i As Integer
        'Dim SWriteAddress As SignedLong, UWriteAddress As UnsignedLong
        'Dim SWriteSize As SignedInt, UWriteSize As UnsignedInt
        Dim UWriteSize As UShort

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'ROM Access Write
        mSendByte(12) = &H1F : mSendByte(13) = 2
        'Data Size
        'SWriteSize.n = WriteSize + 14
        'LSet(UWriteSize = SWriteSize)
        'mSendByte(14) = UWriteSize.H : mSendByte(15) = UWriteSize.L
        UWriteSize = CType(WriteSize + 14, UShort)
        mSendByte(14) = CByte(UWriteSize) >> 8 : mSendByte(15) = CByte(UWriteSize And &HFF)
        'Module Type
        mSendByte(16) = 0 : mSendByte(17) = CByte(mModuleType)
        'ROM Size
        mSendByte(18) = 0 : mSendByte(19) = 4
        'Slot No
        mSendByte(20) = 0 : mSendByte(21) = CByte(mSlotNo)
        'Reserve
        mSendByte(22) = 0 : mSendByte(23) = 0
        'Accsess Address
        'SWriteAddress.n = WriteAddress
        'LSet(UWriteAddress = SWriteAddress)
        'mSendByte(24) = UWriteAddress.HH
        'mSendByte(25) = UWriteAddress.HL
        'mSendByte(26) = UWriteAddress.LH
        'mSendByte(27) = UWriteAddress.LL
        mSendByte(24) = CByte("&H" & Mid(WriteAddress, 1, 2))
        mSendByte(25) = CByte("&H" & Mid(WriteAddress, 3, 2))
        mSendByte(26) = CByte("&H" & Mid(WriteAddress, 5, 2))
        mSendByte(27) = CByte("&H" & Mid(WriteAddress, 7, 2))

        'Firm Data size
        'SWriteSize.n = WriteSize
        'LSet(UWriteSize = SWriteSize)
        'mSendByte(28) = UWriteSize.H : mSendByte(29) = UWriteSize.L
        UWriteSize = CType(WriteSize, UShort)
        mSendByte(28) = CByte(UWriteSize) >> 8 : mSendByte(29) = CByte(UWriteSize And &HFF)
        'File Read
        For i = 0 To WriteSize - 1 Step 2
            If mEndianMode = 0 Then
                'BIG
                mSendByte(i + 30) = FirmData(i)
                mSendByte(i + 31) = FirmData(i + 1)
            Else
                'LITTLE
                mSendByte(i + 30) = FirmData(i + 1)
                mSendByte(i + 31) = FirmData(i)
            End If
        Next i

        'Communication
        Call F3MCommandSendRcv()

    End Sub

    Public Sub WriteROMWord(ByVal WriteAddress As String, ByVal WriteData As String)
        '==========================================================================================
        ' F.Name   : WriteROMWord
        ' Function :
        ' Data     : address,data
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================
        ReDim mSendByte(31)
        'Dim res As Integer, i As Integer
        'Dim SWriteAddress As SignedLong, UWriteAddress As UnsignedLong
        'Dim SWriteData As SignedLong, UWriteData As UnsignedLong

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'ROM Access Write
        mSendByte(12) = &H1F : mSendByte(13) = 2
        'Data size
        mSendByte(14) = 0 : mSendByte(15) = 16
        'Module Type
        mSendByte(16) = 0 : mSendByte(17) = CByte(mModuleType)
        'ROM Size
        mSendByte(18) = 0 : mSendByte(19) = 4
        'Slot No
        mSendByte(20) = 0 : mSendByte(21) = CByte(mSlotNo)
        'Reserve
        mSendByte(22) = 0 : mSendByte(23) = 0
        'Access Address
        'SWriteAddress.n = WriteAddress
        'LSet(UWriteAddress = SWriteAddress)
        'mSendByte(24) = UWriteAddress.HH
        'mSendByte(25) = UWriteAddress.HL
        'mSendByte(26) = UWriteAddress.LH
        'mSendByte(27) = UWriteAddress.LL
        mSendByte(24) = CByte("&H" & Mid(WriteAddress, 1, 2))
        mSendByte(25) = CByte("&H" & Mid(WriteAddress, 3, 2))
        mSendByte(26) = CByte("&H" & Mid(WriteAddress, 5, 2))
        mSendByte(27) = CByte("&H" & Mid(WriteAddress, 7, 2))
        'Write Data size
        mSendByte(28) = 0 : mSendByte(29) = 2
        'Data
        'SWriteData.n = WriteData
        'LSet(UWriteData = SWriteData)
        If mEndianMode = 0 Then
            'BIG
            'mSendByte(30) = UWriteData.LH
            'mSendByte(31) = UWriteData.LL
            mSendByte(30) = CByte("&H" & Mid(WriteData, 1, 2))
            mSendByte(31) = CByte("&H" & Mid(WriteData, 3, 2))
        Else
            'LITTLE
            'mSendByte(31) = UWriteData.LH
            'mSendByte(30) = UWriteData.LL
            mSendByte(30) = CByte("&H" & Mid(WriteData, 3, 2))
            mSendByte(31) = CByte("&H" & Mid(WriteData, 1, 2))
        End If

        'Communication
        Call F3MCommandSendRcv()

    End Sub

    Public Sub FALink(ByVal LPslotno As Integer)

        'FAﾘﾝｸﾓｼﾞｭｰﾙ局別設定
        '局番号2（通信相手）にﾘﾝｸﾘﾚｰ、ﾘﾝｸﾚｼﾞｽﾀの割り付け実施し通信状態を確認する。
        '正常なら戻り値"0507000400000000"が返ってくる

        ReDim mSendByte(43)

        'Header =========================================
        mSendByte(0) = &HF1 : mSendByte(1) = &H0 : mSendByte(2) = &H0 : mSendByte(3) = &H0 : mSendByte(4) = &H0
        mSendByte(5) = &H0 : mSendByte(6) = &H0 : mSendByte(7) = &H0 : mSendByte(8) = &H0 : mSendByte(9) = &H0
        mSendByte(10) = &H0 : mSendByte(11) = &H0

        'SlotNo =========================================
        mSendByte(12) = CByte(mSlotNo)
        'UnitNo =========================================
        mSendByte(13) = CByte(mUnitNo)

        'Main ===========================================
        mSendByte(14) = &H5
        'SUB ============================================
        mSendByte(15) = &H7
        'SIZE ===========================================
        mSendByte(16) = &H0 : mSendByte(17) = &H1A

        'ﾕﾆｯﾄNo,ｽﾛｯﾄNo
        mSendByte(18) = &H0 : mSendByte(19) = &H0
        mSendByte(20) = &H0 : mSendByte(21) = CByte(LPslotno)

        mSendByte(22) = &H40 : mSendByte(23) = &H18
        mSendByte(24) = &H0 : mSendByte(25) = &H0
        mSendByte(26) = &H0 : mSendByte(27) = &H2
        mSendByte(28) = &H0 : mSendByte(29) = &HA
        mSendByte(30) = &H0 : mSendByte(31) = &H1
        mSendByte(32) = &H0 : mSendByte(33) = &H95

        '局番号
        mSendByte(34) = &H0 : mSendByte(35) = &H2
        'ﾘﾝｸﾘﾚｰ先頭番号
        mSendByte(36) = &H0 : mSendByte(37) = &H1
        'ﾘﾝｸﾘﾚｰ点数
        mSendByte(38) = &H1 : mSendByte(39) = &H0
        'ﾘﾝｸﾚｼﾞｽﾀ先頭番号
        mSendByte(40) = &H0 : mSendByte(41) = &H1
        'ﾘﾝｸﾚｼﾞｽﾀ点数
        mSendByte(42) = &H1 : mSendByte(43) = &H0

        Call F3MCommandSendRcv()
        If mError <> "" Then Exit Sub

    End Sub

    Private Sub F3MCommandSend()

        'Dim bytReceiveBuffer(255) As Byte
        'Dim bytSendBuffer As Byte()

        Dim tCount As Double = 0

        objTcpClient = New System.Net.Sockets.TcpClient
        With objTcpClient

            Try
                objTcpClient.Connect(mIPAdd, mRemPort)
                objNetStream = .GetStream
                With objNetStream
                    bytSendBuffer = mSendByte
                    .Write(bytSendBuffer, 0, bytSendBuffer.Length)

                    'mSecWait(100)
                    .Close()
                End With

            Catch Ex As Exception
                If mErrorComment = "" Then mErrorComment = "PC Link Communication Error"
                Throw New ApplicationException(mErrorComment)

            Finally
                .Close()

            End Try

        End With

        objTcpClient = Nothing

    End Sub

    Private Sub F3McommRcv()

    End Sub

    Private Sub F3MCommandSendRcv()

        '==========================================================================================
        ' F.Name   : F3MCommandSendRcv
        ' Function : PC Link Command Communication
        ' Data     : Send command
        ' Return   : Recv command
        ' Remarks  :
        '==========================================================================================

        'Dim bytReceiveBuffer(255) As Byte
        'Dim bytSendBuffer As Byte()
        Dim intDataLength As Integer

        Dim rSize As Integer
        Dim tCount As Double = 0
        Dim eCode As Integer
        Dim bText As String
        Dim rDisp As String

        Dim sb As New System.Text.StringBuilder()

        If mTimeout = 0 Then mTimeout = 3 'ﾀｲﾑｱｳﾄの設定がなければﾀｲﾑｱｳﾄを3秒とする

        objTcpClient = New System.Net.Sockets.TcpClient
        With objTcpClient

            Try

                objTcpClient.Connect(mIPAdd, mRemPort)
                objNetStream = .GetStream
                With objNetStream
                    bytSendBuffer = mSendByte
                    .Write(bytSendBuffer, 0, bytSendBuffer.Length)

                    'mSecWait(100)

                    Do
                        DoEvents()
                        If .DataAvailable And objTcpClient.ReceiveBufferSize >= 20 Then
                            rSize = objTcpClient.ReceiveBufferSize
                            ReDim bytReceiveBuffer(rSize - 1)
                            ReDim mRecvByte(rSize - 1)
                            intDataLength = .Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
                            mRecvByte = bytReceiveBuffer
                            Exit Do
                        End If

                        mSecWait(10) : tCount = tCount + 0.01
                        If tCount > mTimeout Then
                            mErrorComment = "ERROR:MCOMM RCV TIMEOUT SIZE : " & objTcpClient.ReceiveBufferSize.ToString & " byte"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Loop

                End With

            Catch Ex As Exception
                If mErrorComment = "" Then mErrorComment = "PC Link Communication Error"
                Throw New ApplicationException(mErrorComment)

            Finally
                .Close()

            End Try

        End With

        objTcpClient = Nothing

        'Receivetext Data ===========================
        rDisp = ""
        mRecvText = ""
        For i = 12 To intDataLength - 1
            bText = Hex(mRecvByte(i))
            If Len(bText) <= 1 Then bText = "0" & bText
            'rDisp = rDisp & " " & bText
            'mRecvText = mRecvText & bText
            sb.Append(bText)
        Next
        mRecvText = sb.ToString
        'For i = 0 To intDataLength - 1
        '    bText = Hex(mRecvByte(i))
        '    If Len(bText) <= 1 Then bText = "0" & bText
        '    rDisp = rDisp & " " & bText
        '    mRecvText = mRecvText & bText
        'Next
        'Debug.Print("RCV :" & rDisp)
        Debug.Print(mRecvText)

        'Data Size ==================================
        mRecvDataSize = CLng(mRecvByte(16)) * 256 + CLng(mRecvByte(17))

        'Error Code =================================
        mRecvErrCode = CLng(mRecvByte(18)) * 256 + CLng(mRecvByte(19))
        Debug.Print("DATA SIZE : " & mRecvDataSize & " ERR CODE : " & Hex(mRecvErrCode))

        If mRecvErrCode <> 0 Then

            If mRecvErrCode > 32767 Then
                eCode = CInt(mRecvErrCode - 65536)
            Else
                eCode = CInt(mRecvErrCode)
            End If

            mErrorComment = "ERROR:MCOM CODE " & Hex(eCode)
            Throw New ApplicationException(mErrorComment)

        End If

    End Sub

    Private Sub F3MCommandSendRcvH()

        '==========================================================================================
        ' F.Name   : F3MCommandSendRcv
        ' Function : PC Link Command Communication
        ' Data     : Send command
        ' Return   : Recv command
        ' Remarks  :
        '==========================================================================================

        'Dim bytReceiveBuffer(255) As Byte
        'Dim bytSendBuffer As Byte()
        Dim intDataLength As Integer

        'Dim rSize As Integer
        Dim tCount As Double = 0
        'Dim bText As String
        Dim rDisp As String

        Dim sb As New System.Text.StringBuilder()

        If mTimeout = 0 Then mTimeout = 3 'ﾀｲﾑｱｳﾄの設定がなければﾀｲﾑｱｳﾄを3秒とする

        objTcpClient = New System.Net.Sockets.TcpClient
        With objTcpClient

            Try

                objTcpClient.Connect(mIPAdd, mRemPort)
                objNetStream = .GetStream
                With objNetStream
                    bytSendBuffer = mSendByte
                    .Write(bytSendBuffer, 0, bytSendBuffer.Length)

                    'mSecWait(10)

                    Do
                        DoEvents()
                        If .DataAvailable And objTcpClient.ReceiveBufferSize >= 20 Then
                            'rSize = objTcpClient.ReceiveBufferSize
                            'ReDim bytReceiveBuffer(rSize - 1)
                            'ReDim mRecvByte(rSize - 1)
                            intDataLength = .Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
                            mRecvByte = bytReceiveBuffer
                            Exit Do
                        End If

                        mSecWait(10) : tCount = tCount + 0.01
                        If tCount > mTimeout Then
                            mErrorComment = "ERROR:MCOMM RCV TIMEOUT SIZE : " & objTcpClient.ReceiveBufferSize.ToString & " byte"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Loop

                    .Close()
                End With

            Catch Ex As Exception
                If mErrorComment = "" Then mErrorComment = "PC Link Communication Error"
                Throw New ApplicationException(mErrorComment)
            Finally
                .Close()
            End Try
            .Close()
        End With

        objTcpClient = Nothing
        objNetStream = Nothing

        'Receivetext Data ===========================
        rDisp = ""
        mRecvText = ""
        For i = 12 To intDataLength - 1
            'bText = Hex(mRecvByte(i))
            'If Len(bText) <= 1 Then bText = "0" & bText
            'mRecvText &= mRecvByte(i).ToString("X2")
            sb.Append(mRecvByte(i).ToString("X2"))
            'sb.Append(bText)
        Next
        mRecvText = sb.ToString

    End Sub

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
    End Sub

End Class
