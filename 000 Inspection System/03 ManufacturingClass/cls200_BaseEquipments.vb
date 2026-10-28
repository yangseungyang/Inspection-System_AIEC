'---------------------------------------------------------------
'   clsBaseEquipments
'   Copyright(c)  2006 Yokogawa Electric Corporation.
'
'   Comment:  WTシリ?ズ検査装置共通関数
'   Modify : Ver 0.00 2006/11/07 新規作成
'            Ver x.xx xxxx/xx/xx
'
'---------------------------------------------------------------

Option Strict On
Option Explicit On

'Imports WT500.clsCommonMessageBox
Imports VB = Microsoft.VisualBasic

Public Class clsBaseEquipments
    '----------------------------------------
    '   定数宣言
    '----------------------------------------

    '----------------------------------------
    '   ロ?カル変数
    '----------------------------------------
    'Private myGOD As clsWT3000
    'Private myDMM_3 As cls240_34410A
    'Private myDMM_4 As cls240_34410A
    'Private myPCR As cls220_PCRxM
    'Private myACV As cls5700A
    'Private myFG As clsWF1946A
    'Private myGOS_1 As cls230_7651
    'Private myGOS_2 As cls230_7651
    'Private myGuild As clsTM7620
    'Private myDL As clsDLSeries
    Private myDMM As cls240_34410A
    'Private myDMM_2 As cls232_3458A
    'Private myADR_1 As cls231_ADR_3204

    'Private myRelay As clsContecDIO

    '----------------------------------------
    '   コンストラク??デストラク?
    '----------------------------------------
    Public Sub New()
        'myGOD = New clsWT3000(1)
        'myDMM_3 = New cls240_34410A(3)
        'myDMM_4 = New cls240_34410A(4)
        'myPCR = New cls220_PCRxM(0)
        'myACV = New cls5700A(0)
        'myFG = New clsWF1946A(0)
        'myGOS_1 = New cls230_7651(1)
        'myADR_1 = New cls231_ADR_3204(2)
        'myGOS_2 = New cls230_7651(2)
        'myGuild = New clsTM7620(0)
        'myDL = New clsDLSeries(0)
        myDMM = New cls240_34410A(1)
        'myDMM_2 = New cls232_3458A(4)


        'myRelay = New clsContecDIO
    End Sub

    Protected Overrides Sub Finalize()
        'myGOD = Nothing
        'myDMM_3 = Nothing
        'myDMM_4 = Nothing
        'myPCR = Nothing
        'myACV = Nothing
        'myFG = Nothing
        'myGOS_1 = Nothing
        'myGOS_2 = Nothing
        'myGuild = Nothing
        'myDL = Nothing
        myDMM = Nothing
        'myDMM_2 = Nothing
        'myADR_1 = Nothing

        'myRelay = Nothing

        MyBase.Finalize()
    End Sub

    '----------------------------------------
    '   プロパティのGet/Let/Set
    '----------------------------------------
    'Public ReadOnly Property GOD() As clsWT3000
    '    Get
    '        Return myGOD
    '    End Get
    'End Property

    'Public ReadOnly Property DMM_3() As cls240_34410A
    '    Get
    '        Return myDMM_3
    '    End Get
    'End Property
    'Public ReadOnly Property DMM_4() As cls240_34410A
    '    Get
    '        Return myDMM_4
    '    End Get
    'End Property
    'Public ReadOnly Property PCR() As cls220_PCRxM
    '    Get
    '        Return myPCR
    '    End Get
    'End Property
    'Public ReadOnly Property ACV() As cls5700A
    '    Get
    '        Return myACV
    '    End Get
    'End Property
    'Public ReadOnly Property FG() As clsWF1946A
    '    Get
    '        Return myFG
    '    End Get
    'End Property
    'Public ReadOnly Property GOS_1() As cls230_7651
    '    Get
    '        Return myGOS_1
    '    End Get
    'End Property
    Public ReadOnly Property DMM() As cls240_34410A
        Get
            Return myDMM
        End Get
    End Property
    'Public ReadOnly Property DMM_2() As cls232_3458A
    '    Get
    '        Return myDMM_2
    '    End Get
    'End Property
    'Public ReadOnly Property ADR_1() As cls231_ADR_3204
    '    Get
    '        Return myADR_1
    '    End Get
    'End Property

    'Public ReadOnly Property GOS_2() As cls230_7651
    '    Get
    '        Return myGOS_2
    '    End Get
    'End Property
    'Public ReadOnly Property Guild() As clsTM7620
    '    Get
    '        Return myGuild
    '    End Get
    'End Property
    'Public ReadOnly Property DL() As clsDLSeries
    '    Get
    '        Return myDL
    '    End Get
    'End Property
    'Public ReadOnly Property Relay() As clsContecDIO
    '    Get
    '        Return myRelay
    '    End Get
    'End Property

    '----------------------------------------
    '   メ?ッド
    '----------------------------------------
    Public Overridable Sub Open(Optional ByVal param As String = Nothing)
        'myGOD.Open()
        'myDMM_3.Open()
        'myDMM_4.Open()
        'myPCR.Open()
        'myACV.Open()
        'myFG.Open()
        'myGOS_1.Open()
        'myADR_1.Open()           

        'myGOS_2.Open()
        'myGuild.Open()
        ''myDL.Open()        '同期測定でのみ使用
        'myRelay.Open()
        'myDMM_1.Open()
        myDMM.Open()

        'myGOD.Preset()
        'myDMM.Preset()
        'myACV.Preset()
        'myFG.Preset()
        'myGOS.Preset()
        'myGuild.Preset()
        ''myDL.Preset()
        'myRelay.Reset()
    End Sub

    Public Overridable Sub Close()
        'myGOD.Send(":HOLD ON")
        'Me.Waits(0.25)

        'myACV.SetOutput(False)
        'Me.GOS.SetOutput(False)
        'Me.VGeneJIGU_Opr(False)
        'Me.Relay.KDOpr(False)
        'Me.FG.SetChannel(clsWF1946A.Channel.CH1) : Me.FG.SetOutput(False)
        'Me.FG.SetChannel(clsWF1946A.Channel.CH2) : Me.FG.SetOutput(False)

        ''myDMM.Preset()
        ''myACV.Preset()
        ''myFG.Preset()
        ''myGOS.Preset()
        ''myGuild.Preset()
        ''myDL.Preset()
        ''myRelay.Reset()

        'myGOD.Close()
        'myDMM_3.Close()
        'myDMM_4.Close()
        'myPCR.Close()
        'myACV.Close()
        'myFG.Close()
        'myGOS_1.Close()
        'myADR_1.Close()           
        'myGOS_2.Close()
        'myGuild.Close()
        ''myDL.Close()
        myDMM.Close()
        'myDMM_2.Close()


        'myRelay.Close()
    End Sub

    Public Overridable Sub Preset()
        myDMM.Preset()
        'myDMM_2.Preset()
        'myDMM_3.Preset()
        'myDMM_4.Preset()
        'myGOS_1.Preset()
        'myGOS_2.Preset()
        'myADR_1.Preset()       VER 0.03 JDI DEL

    End Sub

    '----------------------------------------
    '   5700A/WF1946A関連
    '----------------------------------------
    Public Overloads Sub SetAcvWF(ByVal volt As Double, ByVal freq As Double, _
                                    Optional ByVal VIPhaseLock As Boolean = False, _
                                    Optional ByVal ACVPhaseLock As Boolean = True)
        ''5700Aの位相LockOff
        'myACV.SetPhaseLock(False)
        'Me.Waits(0.5)

        ''WF1946Aの設定
        'myFG.SetChannel(clsWF1946A.Channel.CH1)

        'If VIPhaseLock = True Then
        '    myFG.SetChanMode(clsWF1946A.ChanMode.Phase)         'CH1,CH2 Phase Lock
        'Else
        '    myFG.SetChanMode(clsWF1946A.ChanMode.Independent)   'CH1,CH2独立
        'End If

        ''5700A出力ON
        'myACV.SetOutputV(True, volt, freq)

        ''ACVPhaseLock = True だったらWF1946Aを出力し、5700A位相Lock
        'If ACVPhaseLock = True Then
        '    'WF1946Aの設定
        '    myFG.SetFunction(clsWF1946A.Waveform.Sinusoid, clsWF1946A.OscMode.Normal)   '正弦波、発信モ?ド=連続
        '    myFG.SetAmplitude(clsWF1946A.AmpMode.Vrms, 5, 0)                            '5Vrms、Offset=0
        '    myFG.SetFrequency(freq)
        '    myFG.SetOutput(True)
        '    Me.Waits(1)

        '    '5700A位相Lock
        '    myACV.SetPhaseLock(True)
        'End If
    End Sub

    Public Function SetAcvDmm(ByVal SetV As Double, ByVal PerLimit As Double, ByRef HoseiCoefficientV As Double) As Boolean
        'ＧＯＳから任意の電圧を発生させ、ＤＭＭで値を追い込む
        'Dim ReadValue As Double
        'Dim i As Integer
        'Dim IgnoreFlag As Boolean = False

        'Me.ACV.SetOutputV(True, SetV, 0)

        'Do
        '    For i = 0 To 9

        '        ReadValue = myDMM.GetMeasureValue

        '        If SetV = 0 Then
        '            HoseiCoefficientV = 1
        '            If ReadValue > (-1 * PerLimit / 100) And ReadValue < (PerLimit / 100) Then
        '                Return True
        '            End If

        '            SetV = -0.8 * ReadValue     '発散防?のため0.8?にしておく
        '            Me.ACV.SetVoltage(SetV)
        '        Else
        '            HoseiCoefficientV = SetV / ReadValue
        '            If HoseiCoefficientV > 1.5 Then
        '                Exit For
        '            End If
        '            If HoseiCoefficientV > (1 - PerLimit / 100) And HoseiCoefficientV < (1 + PerLimit / 100) Then
        '                Return True
        '            End If

        '            SetV = SetV * HoseiCoefficientV
        '            Me.ACV.SetVoltage(SetV)
        '        End If

        '    Next

        '    If IgnoreFlag = True Then Return False

        '    '危険防?のため電圧発生OFF後作業者に確認
        '    Me.ACV.SetOutput(False)
        '    Select Case Me.Msg_AbortRetryIgnore("電圧の追い込みができませんでした")
        '        Case VB.MsgBoxResult.Abort : Throw New System.ApplicationException("電圧の追い込み失敗")
        '        Case VB.MsgBoxResult.Retry
        '        Case VB.MsgBoxResult.Ignore : IgnoreFlag = True '無視の場合1度だけRetryしFalseでも継続
        '        Case Else
        '            Throw New System.ApplicationException
        '    End Select
        'Loop

        'Return False
    End Function

    '----------------------------------------
    '   電圧発生冶具関連
    '----------------------------------------
    Public Sub VGeneJIGU_Opr(ByVal Status As Boolean)
        ' 出力ＯＮ?ＯＦＦ
        ' Status = true ：出力ＯＮ
        ' Status = false：出力ＯＦＦ

        'If Status = True Then
        '    myGOS.SetOutput(True)                   ' 直流電圧発生器７６５１の出力ＯＮ
        '    myFG.SetChannel(clsWF1946A.Channel.CH2) ' ＷＦ１９４６??作?ャンネル設定
        '    myFG.SetOutput(True)                    ' ＷＦ１９４６?出力ＯＮ
        '    myRelay.VGeneOPR("VOPR_ON")             ' 電圧発生冶具出力ＯＮ
        'End If

        'If Status = False Then
        '    myRelay.VGeneOPR("VOPR_OFF")            ' 電圧発生冶具出力ＯＦＦ
        '    myFG.SetChannel(clsWF1946A.Channel.CH2) ' ＷＦ１９４６??作?ャンネル設定
        '    myFG.SetOutput(False)                   ' ＷＦ１９４６?出力ＯＦＦ
        '    myGOS.SetOutput(False)                  ' 直流電圧発生器７６５１の出力ＯＦＦ
        'End If

    End Sub

    Public Function SetVGeneJIGU(ByVal SetV As Double, ByVal SetF As Double, ByVal VIPhaseLock As Boolean) As Double
        '' SetV : 電圧発生冶具?発生電圧
        '' SetF : 電圧発生冶具?発生周波数
        '' VIPhaseLock : V Iの位相ロック設定
        '' SetGosV : GOS設定電圧(戻り値として使用)

        'Dim VGenFreq As String         ' 電圧発生冶具?体の周波数レンジ
        'Dim VGeneRange As String
        'Dim SetGosV As Double

        ''電圧発生冶具?体の電圧レンジ 設定 [Rangeの?30%が発生範囲]
        'Select Case SetV
        '    Case 0.0
        '        VGeneRange = "VRng_Non"
        '        SetGosV = 1.0           '1V固定
        '    Case 0.02 To 0.065          '0.035→0.02に変更2007/03/06
        '        VGeneRange = "50mV"
        '        SetGosV = SetV / 0.05   '1V=50mV
        '    Case 0.07 To 0.13
        '        VGeneRange = "100mV"
        '        SetGosV = SetV / 0.1    '1V=100mV
        '    Case 0.14 To 0.28           '直線性5.5A出力のため0.26→0.28
        '        VGeneRange = "200mV"
        '        SetGosV = SetV / 0.2    '1V=200mV
        '    Case 0.35 To 0.65
        '        VGeneRange = "500mV"
        '        SetGosV = SetV / 0.5    '1V=500mV
        '    Case 0.7 To 1.3
        '        VGeneRange = "1V"
        '        SetGosV = SetV          '1V=1V
        '    Case 1.4 To 2.6
        '        VGeneRange = "2V"
        '        SetGosV = SetV / 2      '1V=2V
        '    Case 3.0
        '        VGeneRange = "5V"
        '        SetGosV = SetV / 5      '1V=5V
        '    Case 3.5 To 6.5
        '        VGeneRange = "5V"
        '        SetGosV = SetV / 5      '1V=5V
        '    Case 7 To 13
        '        VGeneRange = "10V"
        '        SetGosV = SetV / 10     '1V=10V 
        '    Case Else
        '        Throw New System.ApplicationException("電圧発生冶具の出力電圧選択違反(想定外の電圧)")
        'End Select

        '' 電圧発生冶具?体の電圧レンジ設定
        'myRelay.VGeneRng(VGeneRange)

        '' 電圧発生冶具?体の周波数レンジ設定
        'If SetF < 20000 Then            'High Frequency ??しきい値変更（ハ?ド?価結果に?う）
        '    VGenFreq = "under20k"
        'Else
        '    VGenFreq = "above20k"
        'End If
        'myRelay.VGeneFrq(VGenFreq)

        '' ＷＦ１９４６?の設定
        'If VIPhaseLock = True Then
        '    myFG.SetChanMode(clsWF1946A.ChanMode.Phase)         'CH1,CH2 Phase Lock
        'Else
        '    myFG.SetChanMode(clsWF1946A.ChanMode.Independent)   'CH1,CH2独立
        'End If
        'myFG.SetChannel(clsWF1946A.Channel.CH2)                                     'CH2を選択
        'myFG.SetFunction(clsWF1946A.Waveform.Sinusoid, clsWF1946A.OscMode.Normal)   '正弦波、発信モ?ド=連続
        'myFG.SetAmplitude(clsWF1946A.AmpMode.Vrms, 5, 0)                            '5Vrms、Offset=0
        'If VIPhaseLock = False Then                                                 '2Channel独立だったら、改めてCH2の周波数設定
        '    myFG.SetFrequency(SetF)
        'End If

        '' 直流電圧発生器７６５１の設定
        'myGOS.SetVoltage(SetGosV)

        'Return SetGosV

    End Function

    ' 高周波電流発生器の設定
    Public Function SetIGeneJIGU(ByVal SetI As Double, ByVal SetFreq As Double) As Double
        'Dim SetGosV As Double
        'Dim SetRange As String
        'Dim ret As Boolean

        ''高周波電流発生用の経路設定
        'Me.Relay.KDSet()
        'Me.Waits(1)

        ''電圧側で設定するので不要
        ''If VIPhaseLock = True Then
        ''    Me.FG.SetChanMode(clsWF1946A.ChanMode.Phase)         'CH1,CH2 Phase Lock
        ''Else
        ''    Me.FG.SetChanMode(clsWF1946A.ChanMode.Independent)   'CH1,CH2独立
        ''End If

        ''WF1946A設定
        'Me.FG.SetChannel(clsWF1946A.Channel.CH2)
        'Me.FG.SetAmplitude(clsWF1946A.AmpMode.Vrms, 5, 0)   '5 Vrms固定で使用する
        'Me.FG.SetFrequency(SetFreq)

        ''高周波電流発生器?体の周波数レンジ 設定
        'Select Case SetFreq
        '    Case Is <= 300000 : Me.Relay.KDFrq("300KHZ")
        '    Case Is <= 500000 : Me.Relay.KDFrq("500KHZ")
        '    Case Is <= 700000 : Me.Relay.KDFrq("700KHZ")
        '    Case Is <= 1000000 : Me.Relay.KDFrq("1MHZ")
        'End Select

        ''高周波電流発生器?体の電流レンジ 設定 [Rangeの?20%が発生範囲]
        'Select Case SetI
        '    Case 0.004 To 0.006
        '        SetRange = "5MA"
        '        SetGosV = SetI / 0.005  '1V=5mA
        '    Case 0.008 To 0.012
        '        SetRange = "10MA"
        '        SetGosV = SetI / 0.01   '1V=10mA
        '    Case 0.016 To 0.024
        '        SetRange = "20MA"
        '        SetGosV = SetI / 0.02   '1V=20mA
        '    Case 0.04 To 0.06
        '        SetRange = "50MA"
        '        SetGosV = SetI / 0.05   '1V=50mA
        '    Case 0.08 To 0.12
        '        SetRange = "100MA"
        '        SetGosV = SetI / 0.1    '1V=100mA
        '    Case 0.16 To 0.24
        '        SetRange = "200MA"
        '        SetGosV = SetI / 0.2    '1V=200mA
        '    Case 0.4 To 0.6
        '        SetRange = "500MA"
        '        SetGosV = SetI / 0.5    '1V=500mA
        '    Case 0.8 To 1.2
        '        SetRange = "1A"
        '        SetGosV = SetI / 1      '1V=1A
        '    Case 1.6 To 2.4
        '        SetRange = "2A"
        '        SetGosV = SetI / 2      '1V=2A
        '    Case 4 To 6
        '        SetRange = "5A"
        '        SetGosV = SetI / 5      '1V=5A
        '    Case 8 To 12
        '        SetRange = "10A"
        '        SetGosV = SetI / 10     '1V=10A
        '    Case Else
        '        Throw New System.ApplicationException("電圧発生冶具の出力電圧選択違反(想定外の電圧)")
        'End Select

        ''50mA,1MHzは特性が曲がっているため電圧値を変更する2007/03/06
        'If SetI = 0.05 And SetFreq = 1000000 Then
        '    SetGosV = 0.67
        'End If

        'Me.Relay.KDRng(SetRange)

        'Me.FG.SetOutput(True)
        'Me.GOS.SetOutput(SetGosV, True)

        'Return SetGosV

    End Function

    ' GuildLine(TM7620)の設定
    Public Function SetGuildLine(ByVal SetI As Double, ByVal SetFreq As Double) As Double
        'Dim GuildVolt As Double

        'Select Case System.Math.Abs(SetI)
        '    Case Is < 0.2
        '        If SetFreq = 0 Then
        '            Me.Relay.Set7620_DC("7620_2A")
        '        Else
        '            Me.Relay.Set7620_AC("7620_2A")
        '        End If
        '        Guild.SetARange(clsTM7620.ARange.A200m)
        '        GuildVolt = SetI / 0.2  '1V = 0.2A
        '    Case Is <= 2.0
        '        If SetFreq = 0 Then
        '            Me.Relay.Set7620_DC("7620_2A")
        '        Else
        '            Me.Relay.Set7620_AC("7620_2A")
        '        End If
        '        Guild.SetARange(clsTM7620.ARange.A2)
        '        GuildVolt = SetI / 2.0  '1V = 2A
        '    Case Else
        '        If SetFreq = 0 Then
        '            Me.Relay.Set7620_DC("7620_20A")
        '        Else
        '            Me.Relay.Set7620_AC("7620_20A")
        '        End If
        '        Guild.SetARange(clsTM7620.ARange.A20)
        '        GuildVolt = SetI / 20.0  '1V = 20A
        'End Select

        'Guild.SetVRange(clsTM7620.VRange.V1)    '1V入力Range設定

        'Return GuildVolt
    End Function

    '----------------------------------------
    '   その他
    '----------------------------------------
    Public Function DevideString(ByVal InString As String, ByVal KeyName As String, _
                                   ByRef LeftString As String, ByRef RightString As String) As Integer
        'KeyNameの?入位置を戻り値とする。
        'KeyNameの左右のStringをLeftString, RightStringとして切り出す。
        Dim pos As Integer

        pos = InString.IndexOf(KeyName)
        LeftString = InString
        RightString = ""
        If pos > -1 Then
            LeftString = InString.Substring(0, pos).Trim
            RightString = InString.Substring(pos + KeyName.Length).Trim
        End If

        DevideString = pos
    End Function

    Public Sub CheckTestBreak(ByVal TestBreak As Boolean)
        If TestBreak = True Then
            Throw New System.ApplicationException("検査の中断")
        End If

    End Sub

    '---------------------------------------------
    '   メ?ッド?メッセ?ジ関連
    '---------------------------------------------
    Public Function Msg_AbortRetryIgnore(ByVal msg As String) As VB.MsgBoxResult
        Dim msg2 As String

        msg2 = msg & VB.vbCrLf & _
                "終了しますか？" & VB.vbCrLf & VB.vbCrLf & _
                "[Abort]  終了" & VB.vbCrLf & _
                "[Retry]  再実行" & VB.vbCrLf & _
                "[Ignore] 無視して次の検査へ" & VB.vbCrLf
        'SoundPlay(SoundSelect.Err)
        Msg_AbortRetryIgnore = VB.MsgBox(msg2, VB.MsgBoxStyle.AbortRetryIgnore Or VB.MsgBoxStyle.Critical, "Error発生")
        'SoundStop()
    End Function

    Public Function Msg_RetryCancel(ByVal msg As String) As VB.MsgBoxResult
        'SoundPlay(SoundSelect.Err)
        Msg_RetryCancel = VB.MsgBox(msg, VB.MsgBoxStyle.RetryCancel)
        'SoundStop()
    End Function

    Public Sub Msg_Indicate(ByVal Title As String, ByVal msg As String)
        Dim msg2 As String
        msg2 = "------------------------------------------" & VB.vbCrLf
        msg2 &= Title & VB.vbCrLf
        msg2 &= "------------------------------------------" & VB.vbCrLf & VB.vbCrLf
        msg2 &= msg & VB.vbCrLf & VB.vbCrLf
        msg2 &= "設備ができたら､「OK」をクリックして下さい."

        'SoundPlay(SoundSelect.OK)
        VB.MsgBox(msg2, VB.MsgBoxStyle.OKOnly, "作業指示")
        'SoundStop()
    End Sub

    Public Function Msg_Confirm(ByVal Title As String, ByVal msg As String) As VB.MsgBoxResult
        Dim msg2 As String
        msg2 = "------------------------------------------" & VB.vbCrLf
        msg2 &= Title & VB.vbCrLf
        msg2 &= "------------------------------------------" & VB.vbCrLf & VB.vbCrLf
        msg2 &= msg & VB.vbCrLf & VB.vbCrLf
        msg2 &= "動作が正常であれば「はい」を押して下さい。"

        'SoundPlay(SoundSelect.OK)
        Msg_Confirm = VB.MsgBox(msg2, VB.MsgBoxStyle.YesNo Or VB.MsgBoxStyle.Question, "確認画面")
        'SoundStop()
    End Function

    Public Sub Waits(ByVal Sec As Double)
        System.Threading.Thread.Sleep(CInt(Sec * 1000)) 'Sec秒のWaits
    End Sub

    '---------------------------------------------
    '   プライベ?ト関数(汎用)
    '---------------------------------------------
    Private Function GetAppPath() As String
        Return System.IO.Path.GetDirectoryName( _
            System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function


End Class
