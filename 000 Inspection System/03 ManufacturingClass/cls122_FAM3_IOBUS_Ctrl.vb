Option Explicit On
Imports System.Windows.Forms.Application


Public Class cls122_FAM3_IOBUS_Ctrl
    Dim cMCOM As cls120_SP_M_Command

    Dim mPort As String
    Dim mSetting As String
    Dim mTerminalChar As String
    Dim mRecvData As String
    Dim mErrorComment As String
    Dim mErrorCode As String

    Dim mIPAdd As String = ""
    Dim intPortNo As Integer

    Private objNetStream As System.Net.Sockets.NetworkStream
    Private objTcpClient As System.Net.Sockets.TcpClient

    Dim mintUnit As Integer = 0      'DUT検査ユニット番号
    Dim mintSlot As Integer = 6      'DUT検査スロット番号
    Dim mintXD6407 As Integer = 7    'BM128 ＤＩスロット番号
    Dim mintYD6408 As Integer = 8    'BM128 ＤＯスロット番号

    Dim pbstrRcvBuf As String                   'RP45-5P からの 受信データ
    Dim strSndBuf As String

    Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3
    Dim IP_U2 As String = "192.168.1.11"        'B側用LE
    Dim IP_SP As String = "192.168.0.2"         '被検査用SP

    Dim cLnk As cls080_F3LinkToLE               'v0.02

    'SEND DATA ----------------------------
    Public ReadOnly Property SEND_DATA() As String
        Get
            SEND_DATA = strSndBuf
        End Get
    End Property

    'RCV DATA ----------------------------
    Public ReadOnly Property RCV_DATA() As String
        Get
            RCV_DATA = pbstrRcvBuf
        End Get
    End Property


    Public Sub RWIOS_READ(ByVal Unit As Integer, ByVal Slot As Integer, ByVal addr As Integer)
        '(01)   "RWIOS_READ(Unit, Slot, Addr)"
        '       CU,CXで使用
        strSndBuf = "RWIOS_READ(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(addr)) & ")"
        'Call UnitMainCmd(0, strSndBuf)

        Call UnitMainCmd(2, strSndBuf)
    End Sub

    Public Sub RWIOS_WRITE(ByVal Unit As Integer, ByVal Slot As Integer, ByVal addr As Integer, ByVal Buff As Long)
        '(02)   "RWIOS_WRITE(Unit, Slot, Addr, Buff)"
        '       CU,CXで使用
        strSndBuf = "RWIOS_WRITE(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(addr)) & "," & Trim(CStr(Buff)) & ")"
        'Call UnitMainCmd(0, strSndBuf)

        Call UnitMainCmd(2, strSndBuf)
    End Sub

    Public Sub IO_READ_DATA(ByVal Unit As Integer, ByVal Slot As Integer, ByVal Position As Integer, ByVal Number As Integer)
        '(03)   "IO_READ_DATA(Unit, Slot, レジスタ先頭番号, データ数)"
        '       "IO_READ_DATA(Unit, Slot, Position, Number)"
        strSndBuf = "IO_READ_DATA(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(Position)) & "," & Trim(CStr(Number)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IO_WRITE_DATA(ByVal Unit As Integer, ByVal Slot As Integer, ByVal Position As Integer, ByVal Number As Integer, ByVal Data As Integer)
        '(04)   "IO_WRITE_DATA(Unit, Slot, レジスタ先頭番号, データ数, データ)"
        '       "IO_WRITE_DATA(Unit, Slot, Position, Number, Data)"
        strSndBuf = "IO_WRITE_DATA(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(Position)) & "," & Trim(CStr(Number)) & "," & Trim(CStr(Data)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IO_READ_INRELAY(ByVal Unit As Integer, ByVal Slot As Integer, ByVal Position As Integer, ByVal Number As Integer)
        '(05)   "IO_READ_INRELAY(Unit, Slot, リレー先頭番号, ブロック数)"
        '       "IO_READ_INRELAY(Unit, Slot, Position, Number)"
        strSndBuf = "IO_READ_INRELAY(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(Position)) & "," & Trim(CStr(Number)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IO_WRITE_OUTRELAY(ByVal Unit As Integer, ByVal Slot As Integer, ByVal Position As Integer, ByVal Number As Integer, ByVal Data As Integer, ByVal mask As Integer)
        '(06)   "IO_WRITE_OUTRELAY(Unit, Slot, リレー先頭番号, ブロック数, データ, マスク)"
        '       "IO_WRITE_OUTRELAY(Unit, Slot, Position, Number, Data, Mask)"
        strSndBuf = "IO_WRITE_OUTRELAY(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(Position)) & "," & Trim(CStr(Number)) & "," & Trim(CStr(Data)) & "," & Trim(CStr(mask)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IO_READ_INRELAY_POINT(ByVal Unit As Integer, ByVal Slot As Integer, ByVal Position As Integer)
        '(07)   "IO_READ_INRELAY_POINT(Unit, Slot, リレー先頭番号)"
        '       "IO_READ_INRELAY_POINT(Unit, Slot, Position)"
        strSndBuf = "IO_READ_INRELAY_POINT(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(Position)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IO_WRITE_OUTRELAY_POINT(ByVal Unit As Integer, ByVal Slot As Integer, ByVal Position As Integer, ByVal Data As Integer)
        '(08)   "IO_WRITE_OUTRELAY_POINT(Unit, Slot, リレー先頭番号, データ)"
        '       "IO_WRITE_OUTRELAY_POINT(Unit, Slot, Position, Data)"
        strSndBuf = "IO_WRITE_OUTRELAY_POINT(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(Position)) & "," & Trim(CStr(Data)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IO_READ(ByVal Unit As Integer, ByVal Slot As Integer, ByVal AccessArea As Integer, ByVal AccessType As Integer, ByVal AccessAddress As Integer, ByVal Number As Integer)
        '(09)   "IO_READ(Unit, Slot, Memory or G/A, Byte Word Dword, レジスタ先頭番号, データ数)"
        '       "IO_READ(Unit, Slot, AccessArea, AccessType, AccessAddress, Number)"
        strSndBuf = "IO_READ(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(AccessArea)) & "," & Trim(CStr(AccessType)) & "," & _
                    Trim(CStr(AccessAddress)) & "," & Trim(CStr(Number)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IO_WRITE(ByVal Unit As Integer, ByVal Slot As Integer, ByVal AccessArea As Integer, ByVal AccessType As Integer, ByVal AccessAddress As Integer, ByVal Number As Integer, ByVal Data As Integer)
        '(10)   "IO_WRITE(Unit, Slot, Memory or G/A, Byte Word Dword, レジスタ先頭番号, データ数, データ)"
        '       "IO_WRITE(Unit, Slot, AccessArea, AccessType, AccessAddress, Number, Data)"
        strSndBuf = "IO_WRITE(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(AccessArea)) & "," & Trim(CStr(AccessType)) & "," & _
                    Trim(CStr(AccessAddress)) & "," & Trim(CStr(Number)) & "," & Trim(CStr(Data)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub RWIOS_READ_AVERAGE(ByVal Unit As Integer, ByVal Slot As Integer, ByVal addr As Integer, ByVal Count As Integer)
        '(11)   "RWIOS_READ_AVERAGE(Unit, Slot, Addr, Count)"
        '       CU,CX調整時に必要回数測定し平均値を返す処理(0F20,0F24,0F28,10B0,10B4,10B8,1240,1244,1248,13D0,13D4,13D8)
        strSndBuf = "RWIOS_READ_AVERAGE(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(addr)) & "," & Trim(CStr(Count)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub RWIOS_READ_LH(ByVal Unit As Integer, ByVal Slot As Integer, ByVal addr As Integer)
        '(12)   "RWIOS_READ_LH(Unit, Slot, Addr)"
        '       CU,CX検査時に上位下位の整合を取る必要がある場合に使用
        strSndBuf = "RWIOS_READ_LH(" & Trim(CStr(Unit)) & "," & Trim(CStr(Slot)) & "," & _
                    Trim(CStr(addr)) & ")"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IOT_ATEINTRDISABLE()
        '(13)   "IOT_ATEINTRDISABLE()"  割込み無効化
        strSndBuf = "IOT_ATEINTRDISABLE()"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IOT_ATEINTRENABLE()
        '(14)   "IOT_ATEINTRENABLE()"   割込み有効化
        strSndBuf = "IOT_ATEINTRENABLE()"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    Public Sub IOT_ATEIOREFRESH()
        '(15)   "IOT_ATEIOREFRESH()"    IOリフレッシュ
        strSndBuf = "IOT_ATEIOREFRESH()"
        Call UnitMainCmd(0, strSndBuf)
    End Sub

    '送信
    'Index = 0 : RP45-5P
    'Index = 1 : LE11-0T
    'Index = 2 : CPU用LE
    Private Sub UnitMainCmd(ByVal Index As Integer, ByVal strbuf As String)
        'Dim sngStartTime As Single
        Select Case Index
            Case 0
                mIPAdd = "192.168.1.10"
                intPortNo = 5000
            Case 1  'IO用LE
                mIPAdd = "192.168.1.12"
                intPortNo = 12289
            Case 2  'CPU用LE
                mIPAdd = "192.168.1.11"
                intPortNo = 12289

                'frmTestMain.wskUnitMain(Index).RemoteHost = "192.168.1.14"
                'frmTestMain.wskUnitMain(Index).RemotePort = 8500            '画像認識装置は 8500 にする。
            Case Else
        End Select

        pbstrRcvBuf = ""

        On Error GoTo errHandler

        SendRecive(strbuf, pbstrRcvBuf)

        DoEvents()

        Exit Sub
errHandler:
        Resume Next
    End Sub

    Public Sub Slot2CpuReset()
        'SLOT2 CPUリセット
        'Mコマンド使用 バイナリ形式 レスポンスなし

        Dim bytSendDataByte(21) As Byte

        On Error GoTo errHandler

        cMCOM = New cls120_SP_M_Command
        cMCOM.IPAdd = "192.168.1.12"
        cMCOM.RemPort = 12290           'LE バイナリ形式時は 12290 にする。

        cMCOM.SlotNo = 2        'CPU Slot番号
        cMCOM.UnitNo = 0        'ユニット番号

        cMCOM.CPU_RESTERT()
        mWait(3)

        'ダミーラダープログラムを再スタートさせる
        Call UnitMainCmd(1, "02STA")

        cMCOM = Nothing

        Exit Sub
errHandler:
        Resume Next
    End Sub
    Public Sub Slot2CpuReset_SP7x()
        'SLOT2 CPUリセット
        'Mコマンド使用 バイナリ形式 レスポンスなし

        Dim bytSendDataByte(21) As Byte

        On Error GoTo errHandler

        cMCOM = New cls120_SP_M_Command
        cMCOM.IPAdd = "192.168.1.11"
        cMCOM.RemPort = 12289           'LE バイナリ形式時は 12290 にする。

        cMCOM.SlotNo = 2        'CPU Slot番号
        cMCOM.UnitNo = 0        'ユニット番号

        cMCOM.CPU_RESTERT()
        mWait(3)

        'ダミーラダープログラムを再スタートさせる
        Call UnitMainCmd(2, "02STA")

        cMCOM = Nothing

        Exit Sub
errHandler:
        Resume Next
    End Sub

    Public Sub BUSMONITOR_TB04_START()

        Call IOT_ATEINTRDISABLE()

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 2, 1)        'VCC_B ON
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 0)        'XDFAILアサート

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 1)        'ｱﾄﾞﾚｽﾗｯﾁ解除　新ﾊﾞｽ

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 40, 1)       'B側ﾒｲﾝCPUを起こさない

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 1)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 1)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 0)        'ｱﾄﾞﾚｽﾗｯﾁ　新ﾊﾞｽ

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 1)        'XDFAILネゲート

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 0)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 0)

        mWait(2)
        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    Public Sub BUSMONITOR_TB04_START_slot1()

        Call IOT_ATEINTRDISABLE()

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 2, 1)        'VCC_B ON
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 0)        'XDFAILアサート

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 1)        'ｱﾄﾞﾚｽﾗｯﾁ解除　新ﾊﾞｽ

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 40, 1)       'B側ﾒｲﾝCPUを起こさない

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 0)        'ｱﾄﾞﾚｽﾗｯﾁ　新ﾊﾞｽ

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 1)        'XDFAILネゲート

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 0)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 0)

        mWait(2)
        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    '型名確認
    Public Sub AssignTst()

        Call UnitMainCmd(1, "02INF3," & Trim(Str(mintUnit)))

    End Sub

    Public Sub AssignTst_CPU(ByVal slot As Integer)

        Call UnitMainCmd(2, "0" & CStr(slot) & "INF3," & Trim(Str(mintUnit)))

        'I/Oゲートアレイ読み出し
        cMCOM = New cls120_SP_M_Command
        cMCOM.IPAdd = "192.168.1.11"
        cMCOM.RemPort = 12290           'LE バイナリ形式時は 12290 にする。

        cMCOM.SlotNo = slot        'CPU Slot番号
        cMCOM.UnitNo = 0        'ユニット番号

        mWait(1)

        'SP28の時は下で動く
        'Call cMCOM.GATE_ARRAY_READ(0, slot, "0060B200")
        'Form1.TextBox1.Text = Form1.TextBox1.Text & cMCOM.RecvText & vbCrLf
        'Call cMCOM.MEM_READ("0060B200")
        'Form1.TextBox1.Text = Form1.TextBox1.Text & cMCOM.RecvText & vbCrLf

        cMCOM = Nothing

    End Sub

    'TB04 終了
    Public Sub BUSMONITOR_TB04_END()

        Call IOT_ATEINTRDISABLE()

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 1)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 1)
        mSecWait(100)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 0)        'XDFAILアサート
        mSecWait(500)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 2, 0)        'VCC_B OFF
        mWait(2)
        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    'TB04 Slot6 → 11 に変更
    Public Sub BUSMONITOR_TB04_6TO11()

        Call IOT_ATEINTRDISABLE()

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 0)        'XDFAILアサート

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 1)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 1)        'SLOT 設定
        mWait(0.5)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 1)        'XDFAILネゲート
        mWait(2)
        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    'XDFAIL信号
    Public Sub BUSMONITOR_TB04_XDFAIL()

        Call IOT_ATEINTRDISABLE()

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 0)        'XDFAILアサート
        mWait(0.5)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 1)        'XDFAILネゲート
        mWait(2)
        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()
    End Sub

    'XACS検査
    Public Function IOBusTstXACS() As Boolean

        IOBusTstXACS = True

        'Y33->1　ラッチする(信号線が繋がっている場合アクセス後に保持)
        '被検査モジュールに何らかのアクセス(0x7Cxx00-01:アドレス/モジュールIDレジスタ)｡
        'X33->1　ならＯＫ。
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 1)
        Call IO_READ(mintUnit, mintSlot, 1, 2, &H0, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = pbstrRcvBuf & "   :   " & "= " & "1"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If pbstrRcvBuf <> "1" Then
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstXACS = False : Exit Function
        End If

        'XACSラッチクリアの確認
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = pbstrRcvBuf & "   :   " & "= " & "0"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If pbstrRcvBuf <> "0" Then
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstXACS = False : Exit Function
        End If

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)

        'Y33->0　スルーする(信号線が繋がっている場合アクセス後にクリア)
        '被検査モジュールに何らかのアクセス(0x7Cxx00-01:アドレス/モジュールIDレジスタ)｡
        'X33->0　ならＯＫ。
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 0)
        Call IO_READ(mintUnit, mintSlot, 1, 2, &H0, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = pbstrRcvBuf & "   :   " & "= " & "0"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If pbstrRcvBuf <> "0" Then
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstXACS = False : Exit Function
        End If

    End Function

    'XSFAIL検査
    Public Function IOBusTstXSFAIL() As Boolean

        IOBusTstXSFAIL = True

        Call IO_READ(mintUnit, mintSlot, 1, 2, &H24, 1)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 12, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 12)
        Call IO_READ(mintUnit, mintSlot, 1, 2, &H24, 1)

        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = "0x" & Hex(Val(pbstrRcvBuf)) & "   :   " & "= " & "0x1000"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If Hex(Val(pbstrRcvBuf)) <> "1000" Then
            'pbintOKNGFlag = 3 : Exit Sub
            IOBusTstXSFAIL = False : Exit Function
        End If

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 12, 0)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 12)
        Call IO_READ(mintUnit, mintSlot, 1, 2, &H24, 1)

        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = "0x" & Hex(Val(pbstrRcvBuf)) & "   :   " & "= " & "0x0"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If Val(pbstrRcvBuf) <> 0 Then
            'pbintOKNGFlag = 3 : Exit Sub
            IOBusTstXSFAIL = False : Exit Function
        End If

    End Function

    'XDFAILセット(CPU)
    Public Function IOBusTstSetReSetCpuXDFAIL() As Boolean

        Call IOT_ATEINTRDISABLE()

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 1)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 1)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 0)        'XDFAILアサート

        Call mSecWait(10)

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 1)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 0)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 0)

        mWait(2)
        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

        IOBusTstSetReSetCpuXDFAIL = True

    End Function

    'XSFAILセット(CPU)
    Public Function IOBusTstSetCpuXSFAIL() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 12, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 12)

        IOBusTstSetCpuXSFAIL = True
    End Function

    'XDFAILリセット(CPU)
    Public Function IOBusTstReSetCpuXDFAIL() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 1)

        IOBusTstReSetCpuXDFAIL = True
    End Function

    'XSFAILリセット(CPU)
    Public Function IOBusTstReSetCpuXSFAIL() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 13, 0)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 14, 1)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 14, 0)

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 12, 0)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 12)

        IOBusTstReSetCpuXSFAIL = True
    End Function

    'XLFAILセット(CPU)
    Public Function IOBusTstSetCpuXLFAIL() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 11)

        IOBusTstSetCpuXLFAIL = True
    End Function

    'XLFAILリセット(CPU)
    Public Function IOBusTstReSetCpuXLFAIL() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 0)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 11)

        IOBusTstReSetCpuXLFAIL = True
    End Function

    'XPFAILセット(CPU)
    Public Function IOBusTstSetCpuXPFAIL() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 10)

        IOBusTstSetCpuXPFAIL = True
    End Function

    'XPFAILリセット(CPU)
    Public Function IOBusTstReSetCpuXPFAIL() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 0)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 10)

        IOBusTstReSetCpuXPFAIL = True
    End Function

    'XBCLRセット(CPU)
    Public Function IOBusTstSetCpuXBCLR() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 21, 1)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 21)

        IOBusTstSetCpuXBCLR = True
    End Function

    'XBCLRリセット(CPU)
    Public Function IOBusTstReSetCpuXBCLR() As Boolean

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 21, 0)
        Call IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 21)

        IOBusTstReSetCpuXBCLR = True
    End Function

    'XDFAIL検査
    Public Function IOBusTstXDFAIL() As Boolean

        IOBusTstXDFAIL = True

        Call IO_WRITE(mintUnit, mintSlot, 1, 1, &HD, 1, &H80)       'メインユニットCPU実装３つまで
        Call IO_READ(mintUnit, mintSlot, 1, 1, &HD, 1)              'メインユニットCPU実装３つまで
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = "0x" & Hex(Val(pbstrRcvBuf)) & "   :   " & "= " & "0x80"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If Hex(Val(pbstrRcvBuf)) <> "80" Then
            'pbintOKNGFlag = 3 : Exit Sub

        End If
        mWait(0.5)

        Call BUSMONITOR_TB04_XDFAIL()

        Call IO_READ(mintUnit, mintSlot, 1, 1, &HD, 1)              'メインユニットCPU実装３つまで
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = "0x" & Hex(Val(pbstrRcvBuf)) & "   :   " & "= " & "0x0"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If Hex(Val(pbstrRcvBuf)) <> "0" Then
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstXDFAIL = False : Exit Function
        End If

        Call IO_READ(mintUnit, mintSlot, 1, 2, &H0, 1)
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = "0x" & Hex(Val(pbstrRcvBuf)) & "   :   " & "= " & "0x5xxx"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If Len(Hex(Val(pbstrRcvBuf))) = 4 And Microsoft.VisualBasic.Strings.Left(Hex(Val(pbstrRcvBuf)), 1) = "5" Then
        Else
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstXDFAIL = False : Exit Function
        End If

    End Function

    'SLOT検査(SLOT6(5)&SLOT11(A)で確認)
    Public Function IOBusTstSLOT() As Boolean

        IOBusTstSLOT = True

        Call IO_READ(mintUnit, mintSlot, 1, 2, &H0, 1)
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = "0x" & Hex(Val(pbstrRcvBuf)) & "   :   " & "= " & "0x5xxx"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If Len(Hex(Val(pbstrRcvBuf))) = 4 And Microsoft.VisualBasic.Strings.Left(Hex(Val(pbstrRcvBuf)), 1) = "5" Then
        Else
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstSLOT = False : Exit Function
        End If

        Call IO_READ(mintUnit, 11, 1, 2, &H0, 1)
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = pbstrRcvBuf & "   :   " & "= " & "NG"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If InStr(pbstrRcvBuf, "NG") > 0 Then
        Else
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstSLOT = False : Exit Function
        End If

        Call BUSMONITOR_TB04_6TO11()          'TB04 Slot 6 to 11

        Call IO_READ(mintUnit, mintSlot, 1, 2, &H0, 1)
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = pbstrRcvBuf & "   :   " & "= " & "NG"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If InStr(pbstrRcvBuf, "NG") > 0 Then
        Else
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstSLOT = False : Exit Function
        End If

        Call IO_READ(mintUnit, 11, 1, 2, &H0, 1)
        'txtValueChk表示>>>>>
        'pbvrtTestMainValueText = "0x" & Hex(Val(pbstrRcvBuf)) & "   :   " & "= " & "0xAxxx"
        'Call txtValueChkShow(pbvrtTestMainValueText)
        '<<<<<<<<<<<<<<<<<<<<

        If Len(Hex(Val(pbstrRcvBuf))) = 4 And Microsoft.VisualBasic.Strings.Left(Hex(Val(pbstrRcvBuf)), 1) = "A" Then
        Else
            'pbintOKNGFlag = 3 : Exit Function
            IOBusTstSLOT = False : Exit Function
        End If

    End Function

    Public Sub SendRecive(ByVal SendCommand As String, ByRef RecvCommand As String)

        Dim bytReceiveBuffer(255) As Byte
        Dim bytSendBuffer As Byte()
        Dim intDataLength As Integer
        Dim strReceivedData As String = ""

        objTcpClient = New System.Net.Sockets.TcpClient
        objTcpClient.Connect(mIPAdd, intPortNo)
        objNetStream = objTcpClient.GetStream

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

        objTcpClient.Close()
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

    Public Sub BUSMONITOR_TB04_CPU_START(ByVal SLOT As Integer)

        Call IOT_ATEINTRDISABLE()

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 2, 1)        'VCC_B ON

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 1)       '追加
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 1)       '追加

        'Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 9, 0)        'XDFAILアサート
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 0)        'XDFAILアサート
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 1)        'ｱﾄﾞﾚｽﾗｯﾁ解除

        'Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 40, 1)

        If SLOT = 1 Then
            Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 40, 1)
        Else
            Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 40, 0)
        End If

        Select Case SLOT
            Case 1  'SLOT1の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 2  'SLOT2の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 3  'SLOT3の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 4  'SLOT4の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 5
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 6
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定

            Case Else

        End Select

        '新バスモニタ追加   2010/08/05
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 0)        'ｱﾄﾞﾚｽﾗｯﾁ
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 1, 1)        'XDFAILネゲート

        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 10, 0)
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 11, 0)

        mWait(2)
        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    Public Sub BUSMONITOR_TB04_CPU_SLOTChange(ByVal SLOT As Integer)

        Call IOT_ATEINTRDISABLE()
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 1)        'ｱﾄﾞﾚｽﾗｯﾁ解除

        If SLOT = 1 Then
            Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 40, 1)
        Else
            Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 40, 0)
        End If

        Select Case SLOT
            Case 1  'SLOT1の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 2  'SLOT2の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 3  'SLOT3の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 4  'SLOT4の場合
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定
            Case 5
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 1)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 0)        'SLOT 設定

            Case 9
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 3, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 4, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 5, 0)        'SLOT 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 6, 1)        'SLOT 設定
            Case Else

        End Select

        '新バスモニタ追加   2010/08/05
        Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 7, 0)        'ｱﾄﾞﾚｽﾗｯﾁ

        Call mSecWait(100)

        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    Public Sub BUSMONITOR_TB04_CPU_XBRChange(ByVal XBR As Integer)

        Call IOT_ATEINTRDISABLE()

        Select Case XBR
            Case 0
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 17, 1)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 18, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 19, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 20, 0)        'XBR 設定
            Case 1
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 17, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 18, 1)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 19, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 20, 0)        'XBR 設定
            Case 2
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 17, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 18, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 19, 1)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 20, 0)        'XBR 設定
            Case 3
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 17, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 18, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 19, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 20, 1)        'XBR 設定
            Case Else
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 17, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 18, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 19, 0)        'XBR 設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 20, 0)        'XBR 設定
        End Select

        Call mSecWait(100)

        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    Public Sub BUSMONITOR_TB04_CPU_XIRQChange(ByVal XIRQ As Integer)

        Call IOT_ATEINTRDISABLE()

        Select Case XIRQ
            Case 0
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 29, 1)        'XIRQ0設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 30, 0)        'XIRQ0設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 31, 0)        'XIRQ0設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 32, 0)        'XIRQ0設定
            Case 1
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 29, 0)        'XIRQ1設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 30, 1)        'XIRQ1設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 31, 0)        'XIRQ1設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 32, 0)        'XIRQ1設定
            Case 2
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 29, 0)        'XIRQ2設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 30, 0)        'XIRQ2設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 31, 1)        'XIRQ2設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 32, 0)        'XIRQ2設定
            Case 3
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 29, 0)        'XIRQ3設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 30, 0)        'XIRQ3設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 31, 0)        'XIRQ3設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 32, 1)        'XIRQ3設定
            Case Else
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 29, 0)        'XIRQ設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 30, 0)        'XIRQ設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 31, 0)        'XIRQ設定
                Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 32, 0)        'XIRQ設定
        End Select

        Call mSecWait(100)

        Call IOT_ATEIOREFRESH()
        Call IOT_ATEINTRENABLE()

    End Sub

    Public Sub BUS_MONITOR_POWER_OFF()

        Dim scmd As String = ""
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U1

        scmd = "01" & "BWRY" & "901" & ",001,0"
        Try
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cLnk = Nothing
        End Try

        scmd = "01" & "BWRY" & "301" & ",001,0"
        Try
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception

        Finally
            cLnk = Nothing
        End Try

    End Sub

    Public Sub BUS_MONITOR_POWER_ON()

        Dim scmd As String = ""
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U1
        scmd = "01" & "BWRY" & "301" & ",001,1"
        Try
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception

        Finally
            cLnk = Nothing
        End Try

    End Sub

    Public Sub BUS_MONITOR_POWER_ON_IOMODE()

        Dim scmd As String = ""
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U1

        scmd = "01" & "BWRY" & "901" & ",001,1"                 'ﾊﾞｽﾓﾆﾀIOﾓｰﾄﾞ設定
        Try
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cLnk = Nothing
        End Try

        Call mSecWait(500)

        scmd = "01" & "BWRY" & "301" & ",001,1"
        Try
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception

        Finally
            cLnk = Nothing
        End Try

    End Sub

End Class
