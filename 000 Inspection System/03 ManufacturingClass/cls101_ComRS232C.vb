Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : cls110_Barcode_Reader
'Function   : RS232C Communication class 
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'           :
'Issue      : 2008/12/26  Rev.0    T.Yanagawa       New Release
'============================================================================================================================

Imports System.Windows.Forms.Application

Public Class cls101_ComRS232C
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
    Dim mSendByteData As Byte()

    Dim mRecvByte As Byte()

    'Property =====================================================================================
    'Communication Object -------------   
    Public WriteOnly Property RS_COMM() As System.IO.Ports.SerialPort
        Set(ByVal Value As System.IO.Ports.SerialPort)
            mRSComm = Value
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

    'SendByte -----------------------------
    Public Property SendByteData() As Byte()
        Set(ByVal Value() As Byte)
            mSendByteData = Value
        End Set
        Get
            SendByteData = mSendByteData
        End Get
    End Property

    Public WriteOnly Property TIMEOUT() As Integer
        Set(ByVal Value As Integer)
            'mTimeOut = Value * 1000
            mTimeOut = Value
        End Set
    End Property

    'Setting -------------------------
    Public WriteOnly Property SETTING() As String
        Set(ByVal Value As String)
            mSetting = Value
        End Set
    End Property

    'Terminal -------------------------
    Public Property TERMINAL_CHAR() As String
        Set(ByVal Value As String)
            mTerminalChar = Value
        End Set
        Get
            TERMINAL_CHAR = mTerminalChar
        End Get
    End Property

    'Receive ---------------------------
    Public ReadOnly Property RECEIVE_DATA() As String
        Get
            RECEIVE_DATA = mRecvData
        End Get
    End Property

    Public ReadOnly Property RECEIVE_BYTE_DATA() As Byte()
        Get
            RECEIVE_BYTE_DATA = mRecvByte
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

    Public Sub PortSetup()
        '==========================================================================================
        ' Function : COM port initial & open
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim PortSet As String() = Split(mSetting, ",")
        mTimeOut = 0

        Try
            With mRSComm

                .Close()
                .PortName = mPort

                .BaudRate = CInt(PortSet(0))

                Select Case PortSet(1)
                    Case "E" : .Parity = IO.Ports.Parity.Even
                    Case "M" : .Parity = IO.Ports.Parity.Mark
                    Case "N" : .Parity = IO.Ports.Parity.None
                    Case "O" : .Parity = IO.Ports.Parity.Odd
                    Case "S" : .Parity = IO.Ports.Parity.Space
                End Select

                .DataBits = CInt(PortSet(2))

                Select Case PortSet(3)
                    Case "N" : .StopBits = IO.Ports.StopBits.None
                    Case "1" : .StopBits = IO.Ports.StopBits.One
                    Case "1.5" : .StopBits = IO.Ports.StopBits.OnePointFive
                    Case "2" : .StopBits = IO.Ports.StopBits.Two
                End Select

                .RtsEnable = True
                .ReceivedBytesThreshold = 1
                .Handshake = IO.Ports.Handshake.None

                If .IsOpen = True Then
                    .DiscardInBuffer() : .DiscardOutBuffer()
                Else
                    .Open()
                    .DiscardInBuffer() : .DiscardOutBuffer()

                End If

                .ReadTimeout = 10 * 1000    '受信タイムアウト10秒

            End With

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Call PortClose()        'エラーが起きたらポートを閉じる
            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub PortOpen()

        If mRSComm.IsOpen Then
            mRSComm.DiscardInBuffer()
            mRSComm.DiscardOutBuffer()
        Else
            mRSComm.Open()
            mRSComm.DiscardInBuffer()
            mRSComm.DiscardOutBuffer()
        End If

    End Sub

    Public Sub PortClose()
        '==========================================================================================
        ' Function : COM port close
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Try
            Threading.Thread.Sleep(100)
            If Not mRSComm Is Nothing Then mRSComm.Close()

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub FuncRecvCmd()

        '==========================================================================================
        ' Function : Data read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================
        Dim dataLength As Integer
        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        mRecvData = ""

        StartTime = Microsoft.VisualBasic.DateAndTime.Timer()

        ' [1] 먼저 특정 시간(mTimeOut)까지 데이터가 들어오길 대기한다.
        Do
            NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
            ErrTime = NowTime - StartTime
            If ErrTime < 0 Then
                ' Timer()가 24시간마다 0으로 리셋될 수 있으므로 보정
                ErrTime = 24.0# * 3600.0# - StartTime + NowTime
            End If

            DoEvents()

            dataLength = mRSComm.BytesToRead
            If dataLength > 0 Then Exit Do ' 읽을 데이터가 있으면 대기 종료

        Loop While (mTimeOut > ErrTime)

        Try
            ' [2] 실제 데이터 수신 루프
            '    - CR+LF(vbCrLf) 또는 LF(vbLf) 중 하나라도 발견되면 종료
            Do
                ''만약 현재 읽을 수 있는 데이터가 없다면, 타임아웃 체크
                'If mRSComm.BytesToRead = 0 Then
                '    NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
                '    If (NowTime - StartTime) > mTimeOut Then
                '        Throw New TimeoutException("受信がタイムアウトしました。")
                '    End If
                '    DoEvents()
                '    Continue Do
                'End If

                ' 한 글자씩 읽어서 수신 버퍼에 추가
                'mRecvData &= Chr(mRSComm.ReadChar)
                mRecvData += Chr(mRSComm.ReadChar)
                DoEvents()

                ' CR+LF를 찾는 경우
                If InStr(mRecvData, vbCrLf) > 0 Then
                    Exit Do
                End If

                ' LF만 들어왔을 경우
                If InStr(mRecvData, vbLf) > 0 Then
                    Exit Do
                End If

                '' 루프 도중 타임아웃 체크
                'NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
                'If (NowTime - StartTime) > mTimeOut Then
                '    Throw New TimeoutException("受信がタイムアウトしました。")
                'End If

            Loop While True

            ' 수신 데이터에서 CR+LF 또는 LF 제거
            If Right(mRecvData, 2) = vbCrLf Then
                ' CR+LF 제거
                mRecvData = Left(mRecvData, Len(mRecvData) - 2)
            ElseIf Right(mRecvData, 1) = vbLf Then
                ' LF만 제거
                mRecvData = Left(mRecvData, Len(mRecvData) - 1)
            End If

            Debug.Print("Received Data: " & mRecvData)

        Catch ex As TimeoutException
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1)
            If InStr(mErrorComment, vbCr) > 0 Then
                mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            End If

            Debug.Print("Error: " & mErrorComment)

            ' 포트를 닫고 상위(호출부)로 예외 던짐
            PortClose()
            Throw New ApplicationException(mErrorComment)
        End Try

    End Sub

    Public Sub FuncRecvCmd_ori()

        '==========================================================================================
        ' Function : Data read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================
        Dim dataLength As Integer
        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        mRecvData = ""

        StartTime = Microsoft.VisualBasic.DateAndTime.Timer()

        Do
            NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
            ErrTime = NowTime - StartTime
            If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime
            DoEvents()

            dataLength = mRSComm.BytesToRead
            If dataLength <> 0 Then Exit Do '受信すべきデータがあれば抜ける

        Loop While mTimeOut > ErrTime

        Try
            Do While InStr(mRecvData, mTerminalChar) <= 0
                mRecvData += Chr(mRSComm.ReadChar)
                DoEvents()
            Loop

            mRecvData = Left(mRecvData, Len(mRecvData) - Len(mTerminalChar))

            Debug.Print(mRecvData)

            'Catch ex As Exception
        Catch ex As TimeoutException
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Call PortClose()        'エラーが起きたらポートを閉じる
            Throw New ApplicationException(mErrorComment)

        End Try

        'mRecvData = mRSComm.ReadLine

    End Sub

    Public Sub FuncBarRecvCmd()

        '==========================================================================================
        ' Function : Data read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mRecvData = ""
        'If mTimeOut = 0 Then
        '    mRSComm.ReadTimeout = 5 * 1000
        'Else
        '    mRSComm.ReadTimeout = mTimeOut
        'End If

        Try
            Do While InStr(mRecvData, mTerminalChar) <= 0
                mRecvData += Chr(mRSComm.ReadChar)
            Loop

            'mRecvData = Left(mRecvData, Len(mRecvData) - Len(mTerminalChar))

            Debug.Print(mRecvData)

            'Catch ex As Exception
        Catch ex As TimeoutException
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Call PortClose()        'エラーが起きたらポートを閉じる
            Throw New ApplicationException(mErrorComment)

        End Try

        'mRecvData = mRSComm.ReadLine

    End Sub

    Public Sub RecBufIni()

        mRSComm.DiscardInBuffer()

    End Sub

    Public Sub FuncSendCmd()

        Try
            Debug.Print(mSendCmd)
            mRSComm.WriteLine(mSendCmd)
            'mRSComm.Write(mSendCmd & mTerminalChar)
            DoEvents()

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Call PortClose()        'エラーが起きたらポートを閉じる
            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub FuncSendByteData()

        Try
            mRSComm.Write(mSendByteData, 0, mSendByteData.Length)
            DoEvents()

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Call PortClose()        'エラーが起きたらポートを閉じる
            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub FuncRecvByteData()

        '==========================================================================================
        ' Function : Data read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dataLength As Integer
        Dim bytReceiveBuffer() As Byte
        Dim bText As String
        Dim intDataLength As Integer
        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        StartTime = Microsoft.VisualBasic.DateAndTime.Timer()
        mRecvData = ""

        mRSComm.ReadTimeout = 5 * 1000    '受信タイムアウト5秒
        If mTimeOut = 0 Then mTimeOut = 5

        Try
            Do
                NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
                ErrTime = NowTime - StartTime
                If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime

                If mRSComm.BytesToRead > 1 Then
                    Debug.Print(mRSComm.BytesToRead.ToString)
                    dataLength = mRSComm.BytesToRead
                    Exit Do
                End If
            Loop While mTimeOut > ErrTime

            ReDim bytReceiveBuffer(dataLength - 1) : ReDim mRecvByte(dataLength - 1)
            intDataLength = mRSComm.Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
            mRecvByte = bytReceiveBuffer
            'Do While InStr(mRecvData, mTerminalChar) <= 0
            '    mRecvData += Chr(mRSComm.ReadChar)
            '    DoEvents()
            'Loop

            For i = 0 To intDataLength - 1
                bText = Hex(mRecvByte(i))
                If Len(bText) <= 1 Then bText = "0" & bText
                mRecvData &= bText
            Next i

            If InStr(mRecvData, "04") <> 0 Then mRecvData = Left(mRecvData, Len(mRecvData) - Len("04"))

            mRecvData = Chr(Convert.ToInt32(mRecvData, 16))
            Debug.Print(mRecvData)

            'Catch ex As Exception
        Catch ex As TimeoutException
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Call PortClose()        'エラーが起きたらポートを閉じる
            Throw New ApplicationException(mErrorComment)

        End Try

        mRSComm.ReadTimeout = 10 * 1000    '受信タイムアウト10秒

    End Sub

    Public Sub FuncRecvByteData2()

        '==========================================================================================
        ' Function : Data read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dataLength As Integer
        Dim bytReceiveBuffer() As Byte
        'Dim bText As String
        Dim intDataLength As Integer
        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        StartTime = Microsoft.VisualBasic.DateAndTime.Timer()
        mRecvData = ""

        mRSComm.ReadTimeout = 5 * 1000    '受信タイムアウト5秒
        If mTimeOut = 0 Then mTimeOut = 5

        Try
            Do
                NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
                ErrTime = NowTime - StartTime
                If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime

                If mRSComm.BytesToRead > 0 Then
                    Debug.Print(mRSComm.BytesToRead.ToString)
                    dataLength = mRSComm.BytesToRead
                    Exit Do
                End If
            Loop While mTimeOut > ErrTime

            ReDim bytReceiveBuffer(dataLength - 1)
            intDataLength = mRSComm.Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
            mRecvByte = bytReceiveBuffer
            'Do While InStr(mRecvData, mTerminalChar) <= 0
            '    mRecvData += Chr(mRSComm.ReadChar)
            '    DoEvents()
            'Loop

            'For i = 0 To intDataLength - 1
            '    bText = Hex(mRecvByte(i))
            '    If Len(bText) <= 1 Then bText = "0" & bText
            '    mRecvData &= bText
            'Next i

            'If InStr(mRecvData, "04") <> 0 Then mRecvData = Left(mRecvData, Len(mRecvData) - Len("04"))

            'mRecvData = Chr(Convert.ToInt32(mRecvData, 16))
            'Debug.Print(mRecvData)

            'Catch ex As Exception
        Catch ex As TimeoutException
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Call PortClose()        'エラーが起きたらポートを閉じる
            Throw New ApplicationException(mErrorComment)

        End Try

        mRSComm.ReadTimeout = 10 * 1000    '受信タイムアウト10秒

    End Sub

    Public Sub FuncSendRecvCmd()

        Try
            Call PortSetup()
            Call FuncSendCmd()
            Call FuncRecvCmd()
        Catch ex As Exception
            mErrorComment = "RS232C通信ｴﾗｰ"
            Throw New ApplicationException(mErrorComment)
        Finally
            Call PortClose()
        End Try
        
    End Sub

    Public Sub FuncKYOUSEIRecvCmd()

        '==========================================================================================
        ' Function : Data read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dataLength As Integer
        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        mRecvData = ""

        StartTime = Microsoft.VisualBasic.DateAndTime.Timer()

        Do
            NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
            ErrTime = NowTime - StartTime
            If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime
            DoEvents()

            dataLength = mRSComm.BytesToRead
            If dataLength <> 0 Then
                mRecvData += Chr(mRSComm.ReadChar)
                Exit Do '受信すべきデータがあれば抜ける
            End If

        Loop While mTimeOut > ErrTime

        Debug.Print(mRecvData)

    End Sub

    Public Sub FuncLR02OptRecvCmd()

        '==========================================================================================
        ' Function : Data read
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dataLength As Integer
        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double
        Dim mTImeOut As Integer = 5
        Dim CRCounter As Integer = 0

        mRecvData = ""

        StartTime = Microsoft.VisualBasic.DateAndTime.Timer()

        Do
            NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
            ErrTime = NowTime - StartTime
            If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime
            'DoEvents()

            dataLength = mRSComm.BytesToRead
            If dataLength <> 0 Then

                Do While InStr(mRecvData, mTerminalChar) <= 0
                    mRecvData += Chr(mRSComm.ReadChar)
                    'If mTImeOut < ErrTime Then Exit Do
                    If Strings.Right(mRecvData, 1) = vbCr Then
                        CRCounter += 1
                        If CRCounter = 3 Then Exit Do
                    End If
                Loop
                'mRecvData = mRSComm.ReadLine
                Exit Do '受信すべきデータがあれば抜ける
            End If

        Loop While mTImeOut > ErrTime

        Debug.Print(mRecvData)

    End Sub

End Class

