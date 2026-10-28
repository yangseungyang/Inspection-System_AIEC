'---------------------------------------------------------------
'   clsADRx
'   Copyright(c)  2006 Yokogawa Electric Corporation.
'
'   Comment:  可変抵抗発生器
'   > Ver 0.00 | 2012/5/30 | New Release                               | S.W.BAEK
'            Ver x.xx xxxx/xx/xx
'
'---------------------------------------------------------------
Option Strict On
Option Explicit On 

Imports VB = Microsoft.VisualBasic
Public NotInheritable Class cls231_ADR_3204
    Inherits cls103_CommMain

    '----------------------------------------
    '   定数宣言
    '----------------------------------------
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    '?器名
    Private Const EquipmentName As String = "ADR_3204"
    '?器別設定フ?イル名
    Private SetupFilename As String = mAppPath & "\010 EquipmentData\Drive\SetupADRx.ini"
    'Preset時の抵抗値
    Private Const MinResistor As String = "0.20"
    'SendReceiveのWiat時間
    Private Const WaitTime As Integer = 500

    '----------------------------------------
    '   ロ?カル変数
    '----------------------------------------
    Private myID As Integer                                 '識別用ID(複数台使用時に必要)
    Private myRange As ResistanceRange

    '----------------------------------------
    '   コンストラク??デストラク?
    '----------------------------------------
    Public Sub New(ByVal id As Integer)
        MyBase.New()

        myID = id
        TargetName = EquipmentName
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
    Public Enum ResistanceRange As Integer
        R1 = 1  '2K Range
        R2 = 2  '20K Range
        R3 = 3  '200K Range
    End Enum

    '----------------------------------------
    '   メ?ッド
    '----------------------------------------
    Public Sub Preset()
        Dim ReceiveString As String = ""
        Try
            Call SendReceive("MM", ReceiveString, WaitTime)
            Call System.Threading.Thread.Sleep(500)
            Call SendReceive("R1", ReceiveString, WaitTime)
            Call System.Threading.Thread.Sleep(100)
            Call SendReceive(MinResistor, ReceiveString, WaitTime)
            Call System.Threading.Thread.Sleep(100)
        Catch ex As System.Exception
            Throw New System.ApplicationException(String.Format("{0} : Preset に失敗しました。", EquipmentName), ex)
        End Try
    End Sub

    Public Overloads Sub SetResistance(ByVal Resistance As Double)
        Dim ReceiveString As String = ""
        Dim cmd As String = ""
        Try
            If Resistance < 2000 Then
                SetResistanceRange(ResistanceRange.R1)
                cmd = VB.Format(Resistance, "0.000")
            ElseIf Resistance < 20000 Then
                SetResistanceRange(ResistanceRange.R2)
                cmd = VB.Format(Resistance, "0.00")
            ElseIf Resistance < 200000 Then
                SetResistanceRange(ResistanceRange.R3)
                cmd = VB.Format(Resistance, "0.0")
            End If

            'なぜか有効桁数+1桁まで定?しないと全桁数が有効にならない
            cmd &= "0"

            Call SendReceive(cmd, ReceiveString, WaitTime)
            Call System.Threading.Thread.Sleep(100)
        Catch ex As System.Exception
            Throw New System.ApplicationException(String.Format("{0} : SetResistance {1}ohm に失敗しました。", EquipmentName, Resistance), ex)
        End Try
    End Sub

    Public Sub SetResistanceRange(ByVal ResistanceRange As ResistanceRange)
        Dim ReceiveString As String = ""
        Try
            If System.Enum.IsDefined(GetType(ResistanceRange), ResistanceRange) = False Then
                Throw New System.ApplicationException(String.Format("ResistanceRange の値 {0} が異常です。", ResistanceRange))
            End If
            Call SendReceive("MM", ReceiveString, WaitTime)
            Call System.Threading.Thread.Sleep(500)
            Call SendReceive(ResistanceRange.ToString, ReceiveString, WaitTime)
            Call System.Threading.Thread.Sleep(100)

            myRange = ResistanceRange
        Catch ex As System.Exception
            Throw New System.ApplicationException(String.Format("{0} : SetResistanceRange {1} に失敗しました。", EquipmentName, ResistanceRange), ex)
        End Try
    End Sub

    Public Sub StatusCheck()
        Throw New System.NotSupportedException
    End Sub

    Public Sub StatusReset()
        Throw New System.NotSupportedException
    End Sub
    Public Sub Wait(ByVal Sec As Double)
        System.Threading.Thread.Sleep(CInt(Sec * 1000)) 'Sec秒のWaits
    End Sub
End Class

