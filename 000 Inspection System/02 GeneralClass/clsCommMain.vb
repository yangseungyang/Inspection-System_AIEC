'------------------------------------------------------------------------
'   clsCommMain
'   Copyright(c)  2008 Yokogawa Electric Corporation.
'
'   Comment:  <通信 MAIN Class>
'   Modify 
'   > Ver 0.00 | 2008/10/13 | New Release                               | K.Minami
'   > Ver 0.01 | 2008/10/29 | Ethernet対応によるErr.処理追加            | K.Minami
'   > Ver 
'
'------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports CommonCommunicationLibrary
Imports CommonCommunicationLibrary.clsCommBase
Imports System.IO
Imports Microsoft.VisualBasic

Public Class clsCommMain
    '----------------------------------------
    '   Const Define
    '----------------------------------------
    '설정 FILE KeyWord
    Protected Const SetupKeywordCommDevice As String = "DEVICE"
    Protected Const SetupKeywordCommParameter As String = "PARAM"
    Protected Const SetupKeywordCommDeviceTerminater As String = "TERM"
    'Terminater 문자
    Public Const TerminaterCharactors As String = ControlChars.Lf & ControlChars.Cr & ControlChars.CrLf
    'Work用 BufferSize
    Public Const MaxBufferSize As Integer = 4096

    Public Enum DeviceType
        None = -1
        TmGpib = 0
        TmRs232c
        TmUsb
        TmEther
        Rs232c
        TcpClient
        NIFBus
        TmUsbtmc
        TmUsbtmcDL
        NIGpib
        TmVXI11
    End Enum

    '----------------------------------------
    '   EVENT
    '----------------------------------------
    Public Event EventCommLogMessage(ByVal msg As String)

    '----------------------------------------
    '   Local Variable
    '----------------------------------------
    Private myTargetName As String
    Private myComm As clsCommBase
    Private myParam As String
    Private myParamDefault As String

    '----------------------------------------
    '   Constructor ／ Distructor
    '----------------------------------------
    Public Sub New()
        myTargetName = ""
        myParam = ""
        myParamDefault = ""
        myComm = Nothing
    End Sub

    Protected Overrides Sub Finalize()
        If myComm Is Nothing Then
            Return
        End If
        '# Ver0.01ADD[Try-Catch:Err. Skip]
        Try
            If COMM_STATUS_CONNECT = myComm.CommStatus Then
                myComm.Close()
            End If
        Catch ex As System.Exception
            GoTo skipErr
        End Try
skipErr:

    End Sub

    '----------------------------------------
    '    Property
    '----------------------------------------
    Public Property TargetName() As String
        Get
            Return myTargetName
        End Get
        Set(ByVal Value As String)
            myTargetName = Value
        End Set
    End Property

    Public Property Device() As clsCommBase
        Get
            Return myComm
        End Get
        Set(ByVal Value As clsCommBase)
            myComm = Value
        End Set
    End Property
    'Open時 사용되는 Parameter
    Public Property Parameter() As String
        Get
            Return myParam
        End Get
        Set(ByVal Value As String)
            myParam = Value
        End Set
    End Property
    '설정 FILE로 부터 READ된 Parameter
    Public Property ParamDefault() As String
        Get
            Return myParamDefault
        End Get
        Set(ByVal Value As String)
            myParamDefault = Value
        End Set
    End Property

    '----------------------------------------
    '   Method
    '----------------------------------------
    Public Overridable Sub Initialize(ByVal id As Integer, ByVal filename As String)
        Dim dev As String
        Dim term As COMM_TERM

        dev = Me.LoadSetupData(filename, SetupKeywordCommDevice & CStr(id))
        myParamDefault = Me.LoadSetupData(filename, SetupKeywordCommParameter & CStr(id))
        Me.myParam = Me.myParamDefault
        term = CType(CInt(Me.LoadSetupData(filename, SetupKeywordCommDeviceTerminater & CStr(id))), COMM_TERM)
        Select Case CType(dev, DeviceType)
            Case DeviceType.TmGpib
                If -1 = term Then
                    Me.myComm = New clsCommTmGpib
                Else
                    Me.myComm = New clsCommTmGpib(term)
                End If
                Exit Select
            Case DeviceType.TmRs232c
                If -1 = term Then
                    Me.myComm = New clsCommTmRs
                Else
                    Me.myComm = New clsCommTmRs(term)
                End If
                Exit Select
            Case DeviceType.TmEther
                Me.myComm = New clsCommTmEther
                Exit Select
            Case DeviceType.TmUsb
                Me.myComm = New clsCommTmUsb
                Exit Select
            Case DeviceType.Rs232c
                If -1 = term Then
                    Me.myComm = New clsCommRs232c
                Else
                    Me.myComm = New clsCommRs232c(term)
                End If
                Exit Select
            Case DeviceType.TcpClient
                Me.myComm = New clsCommTcpClient
                Exit Select
            Case DeviceType.NIFBus
                '    Me.myComm = New clsCommNiFBUS
                '    Exit Select
                'Case DeviceType.TmUsbtmc
                '    Me.myComm = New clsCommTmUSBTMC
                '    Exit Select
                'Case DeviceType.TmUsbtmcDL
                '    Me.myComm = New clsCommTmUSBTMC_DL
                '    Exit Select
                'Case DeviceType.NIGpib
                '    Me.myComm = New clsCommNiGpib
                '    Exit Select
                'Case DeviceType.TmVXI11
                '    Me.myComm = New clsCommTmVXI11
                Exit Select
            Case Else
                myComm = Nothing
        End Select
    End Sub

    Public Overridable Sub Open(Optional ByVal param As String = Nothing)
        If param Is Nothing Then
            'ini FILE에서 정의한 Parameter
            Me.myParam = Me.myParamDefault
        Else
            'Param.을 임의로 변경 할 경우
            Me.myParam = param
        End If
        Me.myComm.Open(myParam)
        RaiseEvent EventCommLogMessage("Comm Open: " & param)
    End Sub

    Public Overridable Sub Close()
        Me.myComm.Close()
        RaiseEvent EventCommLogMessage("Comm Close:")
    End Sub

    Public Overridable Sub Send(ByVal cmd As String)
        Me.myComm.Write(cmd)
        RaiseEvent EventCommLogMessage("Comm Cmd: " & cmd)
        Return
    End Sub

    Public Overridable Sub SendBinary(ByVal cmd As String)
        Dim data(MaxBufferSize) As Byte
        Dim size As Integer

        size = Me.ConvStr16ToByte(cmd, data)
        Me.myComm.Write(data, size)
        RaiseEvent EventCommLogMessage("Comm Cmd: " & cmd)
        Return
    End Sub

    Public Overridable Sub Receive(ByRef qry As String)
        Dim flag As Boolean
        Dim rLen As Integer

        qry = New String(CChar(" "), MaxBufferSize)
        Me.myComm.Read(qry, qry.Length, rLen, flag)
        If rLen = 0 Then
            qry = ""
        Else
            qry = Microsoft.VisualBasic.Strings.Left(qry, rLen)
        End If
        RaiseEvent EventCommLogMessage("Comm Qry: " & qry)
        Return
    End Sub

    Public Overridable Sub Receive(ByRef qry As String, ByVal len As Integer, ByRef rLen As Integer, ByRef flg As Boolean)
        Dim q As String

        Me.myComm.Read(qry, len, rLen, flg)
        q = qry.Remove(rLen, qry.Length - rLen)
        q = q.TrimEnd(TerminaterCharactors.ToCharArray())
        RaiseEvent EventCommLogMessage("Comm Qry: " & q)
        Return
    End Sub

    Public Overridable Sub ReceiveBinary(ByRef qry As String, ByVal len As Integer, ByRef rLen As Integer, ByRef flg As Boolean)
        Dim data(MaxBufferSize) As Byte

        Me.myComm.Read(data, len, rLen, flg)
        qry = Me.ConvByteToStr16(data, rLen)
        flg = True
        RaiseEvent EventCommLogMessage("Comm Qry: " & qry)
        Return
    End Sub

    Public Overridable Sub SendReceive(ByVal cmd As String, ByRef qry As String, Optional ByVal WaitmS As Integer = 0)
        Me.Send(cmd)
        System.Threading.Thread.Sleep(WaitmS)         'mSec(秒)Wait
        Me.Receive(qry)
    End Sub

    Public Overridable Sub SendReceiveBinary(ByVal cmd As String, ByRef qry As String, ByRef rLen As Integer)
        Throw New System.ApplicationException("송수신、미실장.")
    End Sub
    '설정FILE READ
    Protected Function LoadSetupData(ByVal filename As String, ByVal id As String) As String
        Dim fp As StreamReader
        Dim buff As String
        Dim txt As String
        Dim n As Integer

        On Error GoTo LoadSetupDataError

        fp = New StreamReader(File.Open(filename, FileMode.Open, FileAccess.Read), System.Text.Encoding.Default)
        txt = ""
        Do While fp.Peek > -1
            buff = fp.ReadLine()
            buff = buff.Trim()
            n = buff.IndexOf("<" & id & ">")
            If n > -1 Then
                Do
                    txt = txt & fp.ReadLine()
                    n = txt.IndexOf("\")
                    If -1 = n Then
                        Exit Do
                    End If
                    txt = txt.Remove(n, 1)
                    txt = txt.Trim()
                Loop
                Exit Do
            End If
        Loop
        fp.Close()
        Return txt

LoadSetupDataError:
        MsgBox("통신설정 FILE(" & filename & ")이 없습니다.")
        Return ""
    End Function
    '16진수 문자열 --> Byet배열로 변환
    Public Function ConvStr16ToByte(ByVal val As String, ByRef valB() As Byte) As Integer
        Dim n, cnt As Integer

        On Error Resume Next
        cnt = 0
        For n = 0 To CInt(val.Length - 1) Step 2
            valB(cnt) = CByte("&H" & val.Substring(n, 2))
            cnt = cnt + 1
        Next
        If CBool(val.Length Mod 2) Then
            valB(cnt) = CByte("&H" & val.Substring(n, 2))
            cnt = cnt + 1
        End If
        Return cnt
    End Function
    'Byet배열 --> 16진수 문자열로 변환
    Public Function ConvByteToStr16(ByVal valB() As Byte, ByVal size As Integer) As String
        Dim n As Integer
        Dim str As String = ""

        For n = 0 To size - 1
            str = str & String.Format("{0:X2}", valB(n))
        Next
        Return str
    End Function

End Class
