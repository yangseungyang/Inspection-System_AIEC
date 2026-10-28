'---------------------------------------------------------------
'   clsCommMain
'   Copyright(c)  2006 Yokogawa Electric Corporation.
'
'   Comment:  通信メインクラス
'   Modify : Ver 0.00 2006/08/20 新規作成
'            Ver x.xx xxxx/xx/xx
'
'---------------------------------------------------------------

Option Strict On
Option Explicit On

'Imports c = CommonCommunicationLibrary
Imports CommonCommunicationLibrary
Imports CommonCommunicationLibrary.clsCommBase
Imports System.IO
Imports Microsoft.VisualBasic
'Imports TmctlAPINet                 'TmctlAPINet.dll to your project, and add a reference to your project by reference settings. Please copy tmctl.dll, Ykusb.dll and YKMUSB.dll to the bin director(or Windows System directory).
'Imports NationalInstruments.NI4882  'NationalInstruments.Common.dll and NationalInstruments.NI4882.dll to your project, and add a reference to your project by reference settings.

Public Class cls103_CommMain
    '----------------------------------------
    '   定数宣言
    '----------------------------------------
    '設定フ?イルキ?ワ?ド
    Protected Const SetupKeywordCommDevice As String = "DEVICE"
    Protected Const SetupKeywordCommParameter As String = "PARAM"
    Protected Const SetupKeywordCommDeviceTerminater As String = "TERM"
    '???ネ??文字
    Public Const TerminaterCharactors As String = ControlChars.Lf & ControlChars.Cr & ControlChars.CrLf
    'ワ?ク用バッフ?サイズ
    Public Const MaxBufferSize As Integer = 16   'size of the buffer for receiving    
    'Public Const MaxBufferSize As Integer = 4096   'size of the buffer for receiving    
    'Public Const MaxBufferSize As Integer = 8192   'size of the buffer for receiving    2014/06/25

    'Private myConnectionStatus As Boolean   'False=disconnect, True=connet
    'Private myTMCTL_id As Integer

    Public Enum DeviceType
        None = -1
        TmGpib = 9
        TmRs232c
        TmUsb
        TmEther
        Rs232c
        TcpClient
        NIFBus
        TmUsbtmc
        TmUsbtmcDL
        NIGpib = 0
        TmVXI11
    End Enum

    '----------------------------------------
    '   イベント
    '----------------------------------------
    Public Event EventCommLogMessage(ByVal msg As String)

    '----------------------------------------
    '   ロ?カル変数
    '----------------------------------------
    Private myTargetName As String
    Private myComm As clsCommBase
    Private myParam As String
    Private myParamDefault As String
    'Private myDevice As Integer
    'Private myTMCTL As TMCTL
    '    Private myTerm As Integer


    'Private mGPIB As c.clsCommNiGpib
    '----------------------------------------
    '   コンストラク??デストラク?
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

        If COMM_STATUS_CONNECT = myComm.CommStatus Then
            myComm.Close()
        End If
    End Sub

    '----------------------------------------
    '    プロパティ
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
    'Open時に使用されるパラメ??
    Public Property Parameter() As String
        Get
            Return myParam
        End Get
        Set(ByVal Value As String)
            myParam = Value
        End Set
    End Property
    '設定フ?イルより読み込まれたパラメ??
    Public Property ParamDefault() As String
        Get
            Return myParamDefault
        End Get
        Set(ByVal Value As String)
            myParamDefault = Value
        End Set
    End Property

    '----------------------------------------
    '   メ?ッド
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
                Me.myComm = New clsCommNiFBUS
                Exit Select
            Case DeviceType.TmUsbtmc
                Me.myComm = New clsCommTmUSBTMC
                Exit Select
            Case DeviceType.TmUsbtmcDL
                Me.myComm = New clsCommTmUSBTMC_DL
                Exit Select
            Case DeviceType.NIGpib
                Me.myComm = New clsCommNiGpib
                Exit Select
            Case DeviceType.TmVXI11
                Me.myComm = New clsCommTmVXI11
                Exit Select
            Case Else
                myComm = Nothing
        End Select
    End Sub
    Public Overridable Sub Open(Optional ByVal param As String = Nothing)
        If param Is Nothing Then
            'iniフ?イルで定?したパラメ??
            Me.myParam = Me.myParamDefault
        Else
            'paramを任意で変更したい場合
            Me.myParam = param
        End If
        Me.myComm.Open(myParam)

        'Dim ret As Integer
        'mGPIB = New clsCommNiGpib
        'ret = myComm.Open(myParam)
        RaiseEvent EventCommLogMessage("Comm Open: " & myParam)
    End Sub

    Public Overridable Sub Close()
        Me.myComm.Close()
        RaiseEvent EventCommLogMessage("Comm Close:")
    End Sub

    Public Overridable Sub Send(ByVal cmd As String)
        System.Threading.Thread.Sleep(300)
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

    Public Overridable Sub SendReceive(ByVal cmd As String, ByRef qry As String, Optional ByVal WaitmS As Integer = 300)
        Me.Send(cmd)
        System.Threading.Thread.Sleep(WaitmS)         'mSec秒のWait
        Me.Receive(qry)
    End Sub

    Public Overridable Sub SendReceiveBinary(ByVal cmd As String, ByRef qry As String, ByRef rLen As Integer)
        Throw New System.ApplicationException("送受信、未実装。")
    End Sub
    '設定フ?イルの読み込み
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
        MsgBox("通信設定フ?イル(" & filename & ")がありません。")
        Return ""
    End Function
    '?六進文字列をByet配列へ変換
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
    'Byet配列を?六進文字列へ変換
    Public Function ConvByteToStr16(ByVal valB() As Byte, ByVal size As Integer) As String
        Dim n As Integer
        Dim str As String = ""

        For n = 0 To size - 1
            str = str & String.Format("{0:X2}", valB(n))
        Next
        Return str
    End Function

End Class
