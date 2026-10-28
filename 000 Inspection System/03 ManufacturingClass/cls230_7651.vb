'------------------------------------------------------------------------
'   cls7651
'   Copyright(c)  2012 Yokogawa Electric Corporation.
'
'   Comment:  <7651 Class>
'   Modify 
'   > Ver 0.00 | 2012/5/30 | New Release                               | S.W.BAEK
'   > Ver 
'
'------------------------------------------------------------------------

Option Strict On
Option Explicit On 

Public Class cls230_7651
    Inherits cls103_CommMain
    '----------------------------------------
    '   Const Define
    '----------------------------------------
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    '기기명
    Private Const EquipmentName As String = "7651"
    '기기별 setting FILE Name
    Private SetupFilename As String = mAppPath & "\010 EquipmentData\Drive\Setup7651.ini"

    '----------------------------------------
    '   Local Variable
    '----------------------------------------
    Private myID As Integer                             '식별용ID(複数台 사용 시, 필요)
    Private myRangeAuto As Boolean

    '----------------------------------------
    '   コンストラクタ／デストラクタ
    '----------------------------------------
    Public Sub New(ByVal id As Integer)
        MyBase.New()
        myID = id
        myRangeAuto = True
        Me.TargetName = EquipmentName
        Me.Initialize(id, SetupFilename)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    '----------------------------------------
    '    プロパティ
    '----------------------------------------
    Public ReadOnly Property Model() As String
        Get
            Return EquipmentName
        End Get
    End Property
    Public ReadOnly Property Id() As Integer
        Get
            Return myID
        End Get
    End Property

    '----------------------------------------
    '    Enum
    '----------------------------------------
    '電圧/電流レンジ
    Public Enum Range
        RAuto = 0
        V10m
        V100m
        V1
        V10
        V30
        A1m
        A10m
        A100m
    End Enum

    'ファンクションモード
    Public Enum Mode
        DCV = 0
        DCA
    End Enum

    '----------------------------------------
    '   メソッド
    '----------------------------------------
    'コマンド送信
    Public Overrides Sub Send(ByVal cmd As String)
        Device.Write(cmd)
        Device.Write("E")           '実行トリガー
        Return

    End Sub

    '本体の初期化
    Public Sub Preset()
        Me.Send("RC")
        Me.Send("LA120")                 'Limit 120mA
    End Sub

    'DC電圧出力
    Public Sub SetVoltage(ByVal volt As Double)
        If myRangeAuto = True Then
            Me.Send("SA" & CStr(volt))
        Else
            Me.Send("S" & CStr(volt))
        End If
    End Sub

    'DC電流出力
    Public Sub SetAmpere(ByVal amp As Double)
        If myRangeAuto = True Then
            Me.Send("SA" & CStr(amp))
        Else
            Me.Send("A" & CStr(amp))
        End If
    End Sub

    '出力ON/OFFsetting
    Public Overloads Sub SetOutput(ByVal State As Boolean)
        If State = True Then
            Me.Send("O1")   'Output On
        Else
            Me.Send("O0")   'Output Off
        End If
    End Sub
    Public Overloads Sub SetOutput(ByVal volt As Double, ByVal State As Boolean)
        If myRangeAuto = True Then
            Me.Send("SA" & CStr(volt))
        Else
            Me.Send("S" & CStr(volt))
        End If

        If State = True Then
            Me.Send("O1")   'Output On
        Else
            Me.Send("O0")   'Output Off
        End If
    End Sub

    'ファンクションのsetting
    Public Sub SetRange(ByVal Mode As Mode, ByVal Range As Range)
        Dim cmd As String

        Select Case Mode
            Case Mode.DCV
                Me.Send("F1")
                If Range = Range.RAuto Then
                    myRangeAuto = True
                Else
                    myRangeAuto = False
                    Select Case Range
                        Case Range.V10m : cmd = "R2"
                        Case Range.V100m : cmd = "R3"
                        Case Range.V1 : cmd = "R4"
                        Case Range.V10 : cmd = "R5"
                        Case Range.V30 : cmd = "R6"
                        Case Else
                            Throw New System.ApplicationException(Me.GetErrorHeader & "範囲外のレンジです。")
                    End Select
                    Me.Send(cmd)
                End If
            Case Mode.DCA
                Me.Send("F5")
                If Range = Range.RAuto Then
                    myRangeAuto = True
                Else
                    myRangeAuto = False
                    Select Case Range
                        Case Range.A1m : cmd = "R4"
                        Case Range.A10m : cmd = "R5"
                        Case Range.A100m : cmd = "R6"
                        Case Else
                            Throw New System.ApplicationException(Me.GetErrorHeader & "範囲外のレンジです。")
                    End Select
                    Me.Send(cmd)
                End If
            Case Else
                Throw New System.ApplicationException(Me.GetErrorHeader & "範囲外のモードです。")
        End Select


        Return
    End Sub


    'エラーメッセージヘッダー作成
    Protected Function GetErrorHeader() As String
        Return String.Format("%s_%d:", Me.TargetName, Me.myID)
    End Function

End Class
