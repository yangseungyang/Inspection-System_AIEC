Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : cls080_F3LinkToLE
'Function   : Communication F3LE etc 
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'           :
'Issue      : 2009/02/  Rev.0    T.Yanagawa       New Release
'============================================================================================================================
'Doevents
Imports System.Windows.Forms.Application

Public Class cls080_F3LinkToLE

    'Local ================================================
    Dim mPort As String
    Dim mSetting As String
    Dim mTerminalChar As String
    Dim mRecvData As String
    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mTimeOut As Integer = 0

    Dim mSendByte() As Byte
    Dim mRecvByte() As Byte

    Dim mIPAdd As String = ""

    Private intPortNo As Integer = 12289
    Private objNetStream As System.Net.Sockets.NetworkStream
    Private objTcpClient As System.Net.Sockets.TcpClient

    'Property =====================================================================================
    'IPAddress -----------------------------
    Public Property IPAdd() As String
        Set(ByVal Value As String)
            mIPAdd = Value
        End Set
        Get
            IPAdd = mIPAdd
        End Get
    End Property

    'Port
    Public WriteOnly Property Port() As Integer
        Set(ByVal Value As Integer)
            intPortNo = Value
        End Set
    End Property

    'Receive ---------------------------
    Public ReadOnly Property RECEIVE_DATA() As String
        Get
            RECEIVE_DATA = mRecvData
        End Get
    End Property

    'TimeOut---------------------------
    Public WriteOnly Property TimeOut() As Integer
        Set(ByVal Value As Integer)
            mTimeOut = Value
        End Set
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

    'SendByte
    Public WriteOnly Property SendByte() As Byte()
        Set(ByVal sValue As Byte())
            mSendByte = sValue
        End Set
    End Property

    Public ReadOnly Property RecvByte() As Byte()
        Get
            RecvByte = mRecvByte
        End Get
    End Property

    Public Sub F3PCLinkCommunication(ByVal SendCommand As String, ByRef RecvCommand As String)
        '==========================================================================================
        ' F.Name   : F3PCLinkCommunication
        ' Function : PC Link Command Communication
        ' Data     : Send command
        ' Return   : Recv command
        ' Remarks  :
        '==========================================================================================

        Dim bytReceiveBuffer(255) As Byte
        Dim bytSendBuffer As Byte()
        Dim intDataLength As Integer
        Dim strReceivedData As String = ""

        objTcpClient = New System.Net.Sockets.TcpClient
        With objTcpClient

            Try
                .Connect(mIPAdd, intPortNo)
                objNetStream = .GetStream
                With objNetStream
                    bytSendBuffer = System.Text.Encoding.Default.GetBytes(SendCommand & vbCrLf)
                    .Write(bytSendBuffer, 0, bytSendBuffer.Length)

                    Do
                        If .DataAvailable Then
                            intDataLength = .Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
                            If intDataLength = 0 Then
                                mSecWait(100)
                            Else
                                strReceivedData &= System.Text.Encoding.Default.GetString(bytReceiveBuffer).Substring(0, intDataLength)
                                'ElseIf Not strReceivedData Is Nothing Then
                                RecvCommand = strReceivedData
                                mRecvData = strReceivedData
                                Exit Do
                            End If

                        End If
                    Loop

                End With

            Catch Ex As Exception
                mErrorComment = "PC Link Communication Error"
                Throw New ApplicationException(mErrorComment)

            Finally
                .Close()

            End Try

        End With

        If InStr(mRecvData, vbCrLf) <> 0 Then
            'mRecvData = Strings.Left(mRecvData, Len(mRecvData) - Len(vbCrLf))
            mRecvData = Replace(mRecvData, vbCrLf, "")
            RecvCommand = mRecvData
        End If
        'Debug.Print(mRecvData)
        objTcpClient = Nothing

    End Sub

    Public Sub F3PCLinkCommunication_RP(ByVal SendCommand As String, ByRef RecvCommand As String)
        '==========================================================================================
        ' F.Name   : F3PCLinkCommunication
        ' Function : PC Link Command Communication
        ' Data     : Send command
        ' Return   : Recv command
        ' Remarks  :
        '==========================================================================================

        Dim bytReceiveBuffer(255) As Byte
        Dim bytSendBuffer As Byte()
        Dim intDataLength As Integer
        Dim strReceivedData As String = ""

        objTcpClient = New System.Net.Sockets.TcpClient
        With objTcpClient

            Try
                .Connect(mIPAdd, 5000)
                objNetStream = .GetStream
                With objNetStream
                    bytSendBuffer = System.Text.Encoding.Default.GetBytes(SendCommand & vbCrLf)
                    .Write(bytSendBuffer, 0, bytSendBuffer.Length)

                    Do
                        If .DataAvailable Then
                            intDataLength = .Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
                            If intDataLength = 0 Then
                                mSecWait(100)
                            Else
                                strReceivedData &= System.Text.Encoding.Default.GetString(bytReceiveBuffer).Substring(0, intDataLength)
                                'ElseIf Not strReceivedData Is Nothing Then
                                RecvCommand = strReceivedData
                                mRecvData = strReceivedData
                                Exit Do
                            End If

                        End If
                    Loop

                End With

            Catch Ex As Exception
                mErrorComment = "PC Link Communication Error"
                Throw New ApplicationException(mErrorComment)

            Finally
                .Close()

            End Try

        End With

        If InStr(mRecvData, vbCrLf) <> 0 Then
            'mRecvData = Strings.Left(mRecvData, Len(mRecvData) - Len(vbCrLf))
            mRecvData = Replace(mRecvData, vbCrLf, "")
            RecvCommand = mRecvData
        End If
        'Debug.Print(mRecvData)
        objTcpClient = Nothing

    End Sub

    Public Sub SocketCommunication(ByVal SendCommand As String, ByRef RecvCommand As String)
        '==========================================================================================
        ' F.Name   : SocketPCLinkCommunication
        ' Function : PC Link Command Communication
        ' Data     : Send command
        ' Return   : Recv command
        ' Remarks  :
        '==========================================================================================

        Dim bytReceiveBuffer(255) As Byte
        Dim bytSendBuffer As Byte()
        Dim intDataLength As Integer
        Dim strReceivedData As String = ""
        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        If mTimeOut = 0 Then mTimeOut = 5

        StartTime = Microsoft.VisualBasic.DateAndTime.Timer()

        objTcpClient = New System.Net.Sockets.TcpClient
        With objTcpClient

            Try

                .Connect(mIPAdd, intPortNo)
                objNetStream = .GetStream
                With objNetStream
                    bytSendBuffer = System.Text.Encoding.Default.GetBytes(SendCommand)
                    .Write(bytSendBuffer, 0, bytSendBuffer.Length)

                    Do

                        System.Threading.Thread.Sleep(10)

                        If .DataAvailable Then
                            intDataLength = .Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
                            If intDataLength = 0 Then
                                mSecWait(100)
                            Else
                                strReceivedData &= System.Text.Encoding.Default.GetString(bytReceiveBuffer).Substring(0, intDataLength)
                                'ElseIf Not strReceivedData Is Nothing Then
                                RecvCommand = strReceivedData
                                mRecvData = strReceivedData
                                Exit Do
                            End If

                        End If

                        NowTime = Microsoft.VisualBasic.DateAndTime.Timer()
                        ErrTime = NowTime - StartTime
                        If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime
                        If mTimeOut < ErrTime Then
                            mErrorComment = "受信ﾀｲﾑｱｳﾄｴﾗｰ"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Loop

                End With

            Catch Ex As Exception
                mErrorComment = "PC Link Communication Error"
                Throw New ApplicationException(mErrorComment)

            Finally
                .Close()

            End Try

        End With

        If InStr(mRecvData, vbCrLf) <> 0 Then
            mRecvData = Strings.Left(mRecvData, Len(mRecvData) - Len(vbCrLf))
        End If

        objTcpClient = Nothing

    End Sub

    Private Sub mWait(ByVal Sec As Double)
        '==========================================================================================
        ' F.Name   : mWait
        ' Function : Sec Timer
        ' Data     : 0.000 sec
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        StartTime = Second(Now)
        Do While Sec > ErrTime
            NowTime = Second(Now)
            ErrTime = NowTime - StartTime
            If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime
            DoEvents()
        Loop

    End Sub

    Private Sub mSecWait(ByVal mSec As Integer)
        '==========================================================================================
        ' F.Name   : mWait
        ' Function : mSec Timer
        ' Data     : msec
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait

    End Sub

End Class
