'---------------------------------------------------------------
'   cls3458A
'   Copyright(c)  2006 Yokogawa Electric Corporation.
'
'   Comment:  3458Aクラス
'   Modify : Ver 0.00 2007/02/16 新規作成
'            Ver 0.10 2013/10/28 change SetFunction for Chronus  T.Suda
'            Ver x.xx xxxx/xx/xx
'
'---------------------------------------------------------------

Option Strict On
Option Explicit On 

Imports System.Threading.Thread
'Imports Communication

Public Class cls232_3458A
    Inherits cls103_CommMain
    '----------------------------------------
    '   定数宣言
    '----------------------------------------
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    '機器名
    Private Const EquipmentName As String = "3458A"

    '機器別設定ファイル名
    Private SetupFilename As String = mAppPath & "\010 EquipmentData\Drive\Setup3458A.ini"

    'コマンド送信後の待ち時間
    Private Const WaitTime As Integer = 100                         '送信時の待ち時間  100ms

    'エラーメッセージ
    Private Const ErrorMessageRange As String = "Out of Setting for Range"

    '----------------------------------------
    '   ローカル変数
    '----------------------------------------
    Private myID As Integer                                         '識別用ID(複数台使用時に必要)
    Private myRangeDCVTable(DCVoltageRange.Qty) As String           '直流電圧レンジ設定コマンドテーブル
    Private myRangeDCATable(DCAmpereRange.Qty) As String            '直流電流レンジ設定コマンドテーブル
    Private myRangeACVTable(ACVoltageRange.Qty) As String           '交流電圧レンジ設定コマンドテーブル
    Private myRangeACATable(ACAmpereRange.Qty) As String            '交流電流レンジ設定コマンドテーブル
    Private myRangeRegistanceTable(RegistanceRange.Qty) As String   '抵抗レンジ設定コマンドテーブル
    Private myRangeRegistance4Table(RegistanceRange.Qty) As String     '抵抗4線レンジ設定コマンドテーブ
    Private myAutoZeroTable(AutoZeroMode.Qty) As String             'オートゼロコマンドテーブル
    Private myMeasureFunctionTable(MeasureFunction.Qty) As String   '測定FUNCTIONコマンドテーブル

    '----------------------------------------
    '   コンストラクタ／デストラクタ
    '----------------------------------------
    Public Sub New(ByVal id As Integer)
        MyBase.New()
        myID = id
        TargetName = EquipmentName
        Me.Initialize(id, SetupFilename)
        Me.InitializeCommandTable()
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
        DCV = 1      'DC電圧
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
        V1mV
        V10mV
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
        V10mv
        V100mv
        V1V
        V10V
        V100V
        V1000V
        Qty
    End Enum

    '抵抗レンジ
    Public Enum RegistanceRange
        RAuto = 0
        R10
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
        A100nA
        A1uA
        A10uA
        A100uA
        A1mA
        A10mA
        A100mA
        A1A
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
        Qty
    End Enum

    'オートゼロ
    Public Enum AutoZeroMode
        MOff = 0
        MOn
        MOnce
        Qty
    End Enum

    '----------------------------------------
    '   メソッド
    '----------------------------------------
    '初期化(本体リセット)
    Public Sub Preset()
        'Me.Send("RESET")            'リセット
        'Threading.Thread.Sleep(3000)
        'Me.Send("LFREQ LINE")       'Line Frequency
        'Threading.Thread.Sleep(100)
        'Me.Send("NDIG 8.5")         'Number of Digits
        'Threading.Thread.Sleep(100)
        'Me.Send("TRIG AUTO")        'Trigger
        ''Me.Send("TRIG SYN")        'Trigger要求時出力。連続測定時は"AUTO"
        'Me.Send("NPLC 1")          'Number of Power Line Cycles
        'Threading.Thread.Sleep(100)
        'Me.Send("NRDGS 1,AUTO")    'Number of Readings
        'Me.Send("END ON")           '
        'Threading.Thread.Sleep(100)

        Me.Send("RESET")            'リセット
        Me.Send("LFREQ LINE")       'Line Frequency
        Me.Send("NDIG 8.5")         'Number of Digits
        Me.Send("TRIG AUTO")        'Trigger
        Me.Send("NPLC 10")          'Number of Power Line Cycles
        Me.Send("NRDGS 10,AUTO")    'Number of Readings
        Me.Send("AZERO ON")         '
        Me.Send("END ON")           '
    End Sub

    '直流電圧レンジ設定
    Public Sub SetDCVoltageRange(ByVal range As DCVoltageRange)
        Dim cmd As String

        cmd = myRangeDCVTable(range)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End If
        Me.Send(cmd)
    End Sub

    '交流電圧レンジ設定
    Public Sub SetACVoltageRange(ByVal range As ACVoltageRange)
        Dim cmd As String

        cmd = myRangeACVTable(range)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End If
        Me.Send(cmd)
    End Sub

    '直流電流レンジ設定
    Public Sub SetDCAmpereRange(ByVal range As DCAmpereRange)
        Dim cmd As String

        cmd = myRangeDCATable(range)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End If
        Me.Send(cmd)
    End Sub

    '交流電流レンジ設定
    Public Sub SetACAmpereRange(ByVal range As ACAmpereRange)
        Dim cmd As String

        cmd = myRangeACATable(range)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End If
        Me.Send(cmd)
    End Sub

    '抵抗レンジ設定
    Public Sub SetRegistance2WRange(ByVal range As RegistanceRange)
        Dim cmd As String

        cmd = myRangeRegistanceTable(range)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End If
        Me.Send(cmd)
    End Sub
    '４線抵抗レンジ設定
    Public Sub SetRegistance4WRange(ByVal range As RegistanceRange)
        Dim cmd As String

        'cmd = myRangeRegistance4Table(range)
        cmd = myRangeRegistanceTable(range)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & ErrorMessageRange)
        End If
        Me.Send(cmd)
    End Sub
    'オートゼロ設定
    Public Sub SetAutoZero(ByVal mode As AutoZeroMode)
        Dim cmd As String

        cmd = myAutoZeroTable(mode)
        If "" = cmd Then
            Throw New System.ApplicationException(Me.GetErrorHeader & "Out of Setting for NULL Mode")
        End If
        Me.Send(cmd)
    End Sub

    '測定FUNCTION設定
    'Public Sub SetFunction(ByVal Func As MeasureFunction, ByVal MaxInput As Double, ByVal Resolution As Double)
    '    Dim qry As String = ""
    '    Dim cmd As String

    '    cmd = myMeasureFunctionTable(Func) & "," & CStr(MaxInput) & "," & CStr(Resolution)

    '    Me.Send(cmd)
    'End Sub
    ' change for Chronus
    Public Sub SetFunction(ByVal Func As MeasureFunction, ByVal MaxInput As String, ByVal Resolution As String)
        Dim qry As String = ""
        Dim cmd As String

        cmd = myMeasureFunctionTable(Func) & "," & MaxInput & "," & Resolution

        Me.Send(cmd)
    End Sub

    '測定値取得
    Public Function GetMeasureValue() As Double
        Dim qry As String = ""

        Me.Receive(qry)

        Return CDbl(qry)
    End Function

    '-------------------------------------
    '   プライベート関数
    '-------------------------------------
    'エラーメッセージヘッダー作成
    Protected Overridable Function GetErrorHeader() As String
        Return String.Format("%s_%d:", Me.TargetName, Me.myID)
    End Function

    '電圧/電流コマンドテーブル作成
    Private Sub InitializeCommandTable()
        Dim n As Integer

        'クリア
        For n = 0 To DCVoltageRange.Qty
            Me.myRangeDCVTable(n) = ""
        Next
        For n = 0 To DCAmpereRange.Qty
            Me.myRangeDCATable(n) = ""
        Next
        For n = 0 To ACVoltageRange.Qty
            Me.myRangeACVTable(n) = ""
        Next
        For n = 0 To ACAmpereRange.Qty
            Me.myRangeACATable(n) = ""
        Next
        For n = 0 To RegistanceRange.Qty
            Me.myRangeRegistanceTable(n) = ""
        Next
        For n = 0 To AutoZeroMode.Qty
            Me.myAutoZeroTable(n) = ""
        Next
        For n = 0 To MeasureFunction.Qty
            Me.myMeasureFunctionTable(n) = ""
        Next
        '直流電圧レンジ設定コマンドテーブル初期化
        myRangeDCVTable(DCVoltageRange.VAuto) = "RANGE AUTO"
        myRangeDCVTable(DCVoltageRange.V1mV) = "RANGE 0.001"
        myRangeDCVTable(DCVoltageRange.V10mV) = "RANGE 0.01"
        myRangeDCVTable(DCVoltageRange.V100mV) = "RANGE 0.1"
        myRangeDCVTable(DCVoltageRange.V1V) = "RANGE 1"
        myRangeDCVTable(DCVoltageRange.V10V) = "RANGE 10"
        myRangeDCVTable(DCVoltageRange.V100V) = "RANGE 100"
        myRangeDCVTable(DCVoltageRange.V1000V) = "RANGE 1000"
        '直流電流レンジ設定コマンドテーブル初期化
        myRangeDCATable(DCAmpereRange.AAuto) = "RANGE AUTO"
        myRangeDCATable(DCAmpereRange.A100nA) = "RANGE 0.1E-6"
        myRangeDCATable(DCAmpereRange.A1uA) = "RANGE 1E-6"
        myRangeDCATable(DCAmpereRange.A10uA) = "RANGE 10E-6"
        myRangeDCATable(DCAmpereRange.A100uA) = "RANGE 100E-6"
        myRangeDCATable(DCAmpereRange.A1mA) = "RANGE 1E-3"
        myRangeDCATable(DCAmpereRange.A10mA) = "RANGE 10E-3"
        myRangeDCATable(DCAmpereRange.A100mA) = "RANGE 100E-3"
        myRangeDCATable(DCAmpereRange.A1A) = "RANGE 1"
        '交流電圧レンジ設定コマンドテーブル初期化
        myRangeACVTable(ACVoltageRange.VAuto) = "RANGE AUTO"
        myRangeACVTable(ACVoltageRange.V10mv) = "RANGE 0.01"
        myRangeACVTable(ACVoltageRange.V100mv) = "RANGE 0.1"
        myRangeACVTable(ACVoltageRange.V1V) = "RANGE 1"
        myRangeACVTable(ACVoltageRange.V10V) = "RANGE 10"
        myRangeACVTable(ACVoltageRange.V100V) = "RANGE 100"
        myRangeACVTable(ACVoltageRange.V1000V) = "RANGE 1000"
        '交流電流レンジ設定コマンドテーブル初期化
        myRangeACATable(ACAmpereRange.AAuto) = "RANGE AUTO"
        myRangeACATable(ACAmpereRange.A100uA) = "RANGE 100E-6"
        myRangeACATable(ACAmpereRange.A1mA) = "RANGE 1E-3"
        myRangeACATable(ACAmpereRange.A10mA) = "RANGE 10E-3"
        myRangeACATable(ACAmpereRange.A100mA) = "RANGE 100E-3"
        myRangeACATable(ACAmpereRange.A1A) = "RANGE 1"
        '抵抗レンジ設定コマンドテーブ初期化
        myRangeRegistanceTable(RegistanceRange.RAuto) = "RANGE AUTO"
        myRangeRegistanceTable(RegistanceRange.R10) = "RANGE 10"
        myRangeRegistanceTable(RegistanceRange.R100) = "RANGE 100"
        myRangeRegistanceTable(RegistanceRange.R1K) = "RANGE 1E3"
        myRangeRegistanceTable(RegistanceRange.R10K) = "RANGE 10E3"
        myRangeRegistanceTable(RegistanceRange.R100K) = "RANGE 100E3"
        myRangeRegistanceTable(RegistanceRange.R1M) = "RANGE 1E6"
        myRangeRegistanceTable(RegistanceRange.R10M) = "RANGE 10E6"
        myRangeRegistanceTable(RegistanceRange.R100M) = "RANGE 100E6"
        myRangeRegistanceTable(RegistanceRange.R1G) = "RANGE 1E9"
        'オートゼロモードコマンドテーブル初期化
        myAutoZeroTable(AutoZeroMode.MOff) = "AZERO OFF"
        myAutoZeroTable(AutoZeroMode.MOn) = "AZERO ON"
        myAutoZeroTable(AutoZeroMode.MOnce) = "AZERO ONCE"
        '測定FUNCTIONコマンドテーブル初期化
        myMeasureFunctionTable(MeasureFunction.DCV) = "FUNC DCV"
        myMeasureFunctionTable(MeasureFunction.ACV) = "FUNC ACV"
        myMeasureFunctionTable(MeasureFunction.ACDCV) = "FUNC ACDCV"
        myMeasureFunctionTable(MeasureFunction.OHM) = "FUNC OHM"
        myMeasureFunctionTable(MeasureFunction.OHMF) = "FUNC OHMF"
        myMeasureFunctionTable(MeasureFunction.DCI) = "FUNC DCI"
        myMeasureFunctionTable(MeasureFunction.ACI) = "FUNC ACI"
        myMeasureFunctionTable(MeasureFunction.ACDCI) = "FUNC ACDCI"
        myMeasureFunctionTable(MeasureFunction.FREQ) = "FUNC FREQ"
    End Sub


    '高速測定設定
    Public Sub Hightsetting()
        Dim qry As String = ""
        Me.Send("RESET")             'リセット
        Threading.Thread.Sleep(3000)
        Me.Send("END")               'GPIBのEOI をON
        Threading.Thread.Sleep(100)
        Me.Send("END?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        'Me.Send("DCI 100E-3")       'DCI 測定ﾚﾝｼﾞ 120mA以下の場合は、100mAも使用可能
        Me.Send("APER 20E-3")        'アパーチャー時間 20m秒：NPLC=1と同じ。積分時間：20ms
        Threading.Thread.Sleep(100)
        Me.Send("APER?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("DISP OFF")          '3458Aの表示をOFF。表示ONでは高速処理に追いつかない。
        Threading.Thread.Sleep(100)
        Me.Send("DISP?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("NRDGS 50,TIMER")    'タイマーのタイミングで50回測定
        Threading.Thread.Sleep(100)
        Me.Send("NRDGS?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("TIMER 0.050")           'タイマーは50m秒。アパーチャー時間の2倍以上
        Threading.Thread.Sleep(100)
        Me.Send("TIMER?")
        Me.Receive(qry)
    End Sub

    'SIM Setting of the transmitter power limit measurement  SIM伝送器電源リミット測定設定
    Public Sub Limitsetting()
        Dim qry As String = ""
        Me.Send("RESET")             'Reset
        Threading.Thread.Sleep(3000)
        Me.Send("END")               'Turn on EOI of GPIB
        Threading.Thread.Sleep(100)
        Me.Send("END?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("DCI 100E-3")        'DCI 測定ﾚﾝｼﾞ 120mA以下の場合は、100mAも使用可能
        Threading.Thread.Sleep(100)
        Me.Send("FUNC?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("APER 20E-3")        'アパーチャー時間 20m秒：NPLC=1と同じ。積分時間：20ms
        Threading.Thread.Sleep(100)
        Me.Send("APER?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("DISP OFF")          '3458Aの表示をOFF。表示ONでは高速処理に追いつかない。
        Threading.Thread.Sleep(100)
        Me.Send("DISP?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("NRDGS 90,TIMER")   'タイマーのタイミングで50回測定(2014-10-27)
        Threading.Thread.Sleep(100)
        Me.Send("NRDGS?")
        Me.Receive(qry)
        Threading.Thread.Sleep(100)
        Me.Send("TIMER 0.100")       'タイマーは100mSEC。アパーチャー時間の2倍以上(2014-10-27)
        Threading.Thread.Sleep(100)
        Me.Send("TIMER?")
        Me.Receive(qry)
        'Me.Send("NRDGS 100,TIMER")   'タイマーのタイミングで100回測定
        'Threading.Thread.Sleep(100)
        'Me.Send("TIMER 0.050")       'タイマーは50mSEC。アパーチャー時間の2倍以上
    End Sub
    'DMM MEASURE
    Public Function GetMeasureDMM(ByVal WT As Integer) As Double
        Dim ReceiveString As String = ""

        mSecWait(WT)

        Try
            'Call SendReceive("read?", ReceiveString)
            Call SendReceive("", ReceiveString)
            'Call SendReceive("E", ReceiveString)
            Return CDbl(ReceiveString)
            Me.Send("DISP ON")
        Catch ex As System.Exception
            Throw New System.ApplicationException(String.Format("{0} : GetDMM に失敗しました。", EquipmentName), ex)
        End Try
    End Function
    Private Sub mSecWait(ByVal mSec As Integer)
        '==========================================================================================
        ' Function : mSec Wait
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait

    End Sub
End Class
