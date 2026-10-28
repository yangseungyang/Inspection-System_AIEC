'---------------------------------------------------------------
'   cls34410A
'
'   Comment:  34410Aクラス
'   > Ver 0.00 | 2015/4/24 | New Release   
'            Ver x.xx xxxx/xx/xx
'
'---------------------------------------------------------------

Option Strict On
Option Explicit On

Public Class cls240_34410A
    Inherits cls103_CommMain
    '----------------------------------------
    '   定数宣言
    '----------------------------------------
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    '?器名
    Private Const EquipmentName As String = "34410A"
    '?器別設定フ?イル名
    Private SetupFilename As String = mAppPath & "\010 EquipmentData\Drive\Setup34410A.ini"
    'コ?ンド送信後の待ち時間
    Private Const WaitTime As Integer = 100                         '送信時の待ち時間  100ms
    'エラ?メッセ?ジ
    Private Const ErrorMessageRange As String = "設定範囲外のレンジです。"

    '----------------------------------------
    '   ロ?カル変数
    '----------------------------------------
    Private myID As Integer                                         '識別用ID(複数台使用時に必要)
    Private myRangeDCVTable(DCVoltageRange.Qty) As String           '直流電圧レンジ設定コ?ンドテ?ブル
    Private myRangeDCATable(DCAmpereRange.Qty) As String            '直流電流レンジ設定コ?ンドテ?ブル
    Private myRangeACVTable(ACVoltageRange.Qty) As String           '交流電圧レンジ設定コ?ンドテ?ブル
    Private myRangeACATable(ACAmpereRange.Qty) As String            '交流電流レンジ設定コ?ンドテ?ブル
    Private myRangeRegistanceTable(RegistanceRange.Qty) As String   '抵抗レンジ設定コ?ンドテ?ブル
    Private myAutoZeroTable(AutoZeroMode.Qty) As String             'オ?ト?ロコ?ンドテ?ブル
    Private myMeasureFunctionTable(MeasureFunction.Qty) As String   '測定FUNCTIONコ?ンドテ?ブル

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
    '   Enum
    '----------------------------------------

    'Measure Function
    Public Enum MeasureFunction
        DCV = 0      'DC電圧
        ACV          'AC電圧
        ACDCV        'AC+DC電圧
        OHM          '抵抗2線
        OHMF         '抵抗4線
        DCI          'DC電流
        ACI          'AC電流
        ACDCI        'AC+DC電流
        FREQ         '周波数
        Qty
    End Enum

    '直流電圧レンジ
    Public Enum DCVoltageRange
        VAuto = 0
        V100mV
        V1V
        V10V
        V100V
        V1000V
        Qty
    End Enum

    '交流電圧レンジ
    Public Enum ACVoltageRange
        VAuto = 0
        V100mV
        V1V
        V10V
        V100V
        V1000V
        Qty
    End Enum

    '抵抗レンジ
    Public Enum RegistanceRange
        RAuto = 0
        R100
        R1K
        R10K
        R100K
        R1M
        R10M
        R100M
        R1G
        Qty
    End Enum

    '直流電流
    Public Enum DCAmpereRange
        AAuto = 0
        A100uA
        A1mA
        A10mA
        A100mA
        A1A
        A3A
        Qty
    End Enum

    '交流電流
    Public Enum ACAmpereRange
        AAuto = 0
        A100uA
        A1mA
        A10mA
        A100mA
        A1A
        A3A
        Qty
    End Enum

    'オ?ト?ロ
    Public Enum AutoZeroMode
        MOff = 0
        MOn
        MOnce
        Qty
    End Enum

    '----------------------------------------
    '   メ?ッド
    '----------------------------------------
    '初期化(?体リセット)
    Public Sub Preset()
        Me.Send("*RST")            'リセット
        Me.Send("*CLS")
        Me.Send(":CONF:VOLT:DC AUTO")
        Me.Send("VOLT:DC:NPLC 10")
        Me.Send("VOLT:DC:IMP:AUTO ON")
        Me.Send("DISP ON")
    End Sub

    'オ?ト?ロ設定
    Public Sub SetAutoZero(ByVal mode As AutoZeroMode)
        Dim cmd As String

        cmd = myAutoZeroTable(mode)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & "設定範囲外のNULLモ?ドです。")
        End If
        Me.Send(cmd)
    End Sub

    'DCV NPLC SET
    Public Function NPLCVSet(ByVal NPLC As Single) As Single
        Me.Send("VOLT:DC:NPLC " & NPLC)
    End Function
    '直流電圧レンジ設定
    Public Sub SetDCVoltageRange(ByVal range As DCVoltageRange)
        Dim r As DCVoltageRange

        r = CType(range, DCVoltageRange)
        Select Case range

            Case DCVoltageRange.VAuto : Me.Send("CONF:VOLT:DC AUTO")
            Case DCVoltageRange.V100mV : Me.Send("CONF:VOLT:DC 0.1")
            Case DCVoltageRange.V1V : Me.Send("CONF:VOLT:DC 1, 1E-5")
            Case DCVoltageRange.V10V : Me.Send("CONF:VOLT:DC 10")
            Case DCVoltageRange.V100V : Me.Send("CONF:VOLT:DC 100")
            Case DCVoltageRange.V1000V : Me.Send("CONF:VOLT:DC 1000")
            Case Else
                Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End Select

    End Sub
    '直流電流レンジ設定
    Public Sub SetDCAmpereRange(ByVal range As DCAmpereRange)
        Dim r As DCAmpereRange

        r = CType(range, DCAmpereRange)
        Select Case range

            Case DCAmpereRange.AAuto : Me.Send("CONF:CURR:DC AUTO")
            Case DCAmpereRange.A100uA : Me.Send("CONF:CURR:DC 0.1")
            Case DCAmpereRange.A1mA : Me.Send("CONF:CURR:DC 1, 1E-5")
            Case DCAmpereRange.A10mA : Me.Send("CONF:CURR:DC 10")
            Case DCAmpereRange.A100mA : Me.Send("CONF:CURR:DC 100")
            Case DCAmpereRange.A1A : Me.Send("CONF:CURR:DC 1000")
            Case DCAmpereRange.A3A : Me.Send("CONF:CURR:DC 3000")
            Case Else
                Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End Select

    End Sub
    'DMM MEASURE
    Public Function GetMeasureDMM(ByVal WT As Double) As Double
        Dim ReceiveString As String = ""

        Wait(WT)

        Try
            Call SendReceive("read?", ReceiveString)
            Return CDbl(ReceiveString)
            Me.Send("DISP ON")
        Catch ex As System.Exception
            Throw New System.ApplicationException(String.Format("{0} : GetDMM に失敗しました。", EquipmentName), ex)
        End Try
    End Function

    Public Sub Wait(ByVal Sec As Double)
        System.Threading.Thread.Sleep(CInt(Sec * 1000)) 'Sec秒のWaits
    End Sub

    '-------------------------------------
    '   プライベ?ト関数
    '-------------------------------------
    'エラ?メッセ?ジヘッ??作成
    Protected Overridable Function GetErrorHeader() As String
        Return String.Format("%s_%d:", Me.TargetName, Me.myID)
    End Function

End Class
