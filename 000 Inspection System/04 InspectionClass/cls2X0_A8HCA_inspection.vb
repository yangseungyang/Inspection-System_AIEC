'==============================================================================================================
'rev 0.00 F3CU/CX Remakeによるクラス追加


'==============================================================================================================

Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm020_InspectionStartEnd
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.Net.NetworkInformation
Imports VB = Microsoft.VisualBasic
Imports System.Math
Imports Renci.SshNet
Imports System.IO
Imports Org.BouncyCastle.Crypto
Imports System.Net.Mail
Imports System.Threading.Tasks
Imports System.ComponentModel.Design


Public Class cls2X0_A8HCA_inspection
    Friend myEq As clsBaseEquipments
    'Class ================================================
    Dim cIA8Comm As cls101_ComRS232C
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cLnk As cls080_F3LinkToLE
    Dim cISPComm As cls101_ComRS232C
    Dim cWin As cls001_Window_Control
    Dim cMac As cls031_MacAddress
    Dim cGOS_1 As cls230_7651
    Dim cDMM_1 As cls232_3458A
    Dim cDMM_2 As cls232_3458A
    Dim cADR_1 As cls231_ADR_3204

    Dim cSer As cls030_LineSerial
    Dim cLbl As cls011_Excel_Operation
    Dim cMcm As cls120_SP_M_Command
    Dim cPCSP As cls060_PCNetSPEED
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl
    Dim cTEST As cls121_SP_TEST_MODE
    Dim cKP As cls124_KP_IOBUS_TEST

    Dim cPF As cls125_PF_IOBUS_TEST
    Dim cBUBU As cls123_BUBU_IOBUS_TEST

    Dim dfrm As frm500_MessagePopUp

    Private objSck As New System.Net.Sockets.TcpClient
    Private objStm As System.Net.Sockets.NetworkStream

    Dim bytReceiveBuffer(255) As Byte
    Dim bytSendBuffer As Byte()
    Dim mRecvByte() As Byte
    Dim RecvText As String

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String

    Dim mTestSpec As String
    Dim mTestMeasure As String

    Dim Dmm_Read_Data As Double
    Dim InputSetVoltage As Double

    Dim rs232cRZPort As String = "COM7"

    Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3のLE
    Dim IP_U2 As String = "192.168.1.11"        'ﾒｲﾝCPU SP76
    Dim IP_U3 As String = "192.168.1.12"        'ﾒｲﾝCPU SP28

    Dim mTestComAddress As Integer = &H5000
    Dim mTestInAddress As Integer = &H5004
    Dim mTestCom As Integer = &H55AA
    Dim mTestClearCom As Integer = &H0

    Dim mTestExeAddress As Integer = &H55AA

    Dim mlngDataL As Long     'LONG WORDデータの下位ビット
    Dim mlngDataH As Long     'LONG WORDデータの上位ビット
    Dim mlngDataLH As Long     'LONG WORDデータ
    Dim mlngDataLHMin(2) As Long     'LONG WORDデータ 最小値
    Dim mlngDataLHMax(2) As Long     'LONG WORDデータ 最大値
    Dim mlngDataStd As Single     'データの基準値  
    Dim mlngDataMin As Single     'データの最小値
    Dim mlngDataMax As Single     'データの最大値
    Dim mlngDataSpec As Single    '判定の基準値
    Dim mlngData(16) As Single     '
    Dim mintCntTDNum As Integer          '検査データ番号のカウント用

    Dim mTimeOut As Integer

    Dim iName As String

    'Property =====================================================================================
    'Error message --------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public ReadOnly Property TestSpec() As String
        Get
            TestSpec = mTestSpec
        End Get
    End Property

    Public ReadOnly Property TestMeasure() As String
        Get
            TestMeasure = mTestMeasure
        End Get
    End Property

    Public ReadOnly Property ERR_CODE() As String
        Get
            ERR_CODE = mErrorCode
        End Get
    End Property

    'LED test 용
    Public Structure LedTestInfo
        Public testName As String       ' 예: "XSFAIL", "Main LED", ...
        Public ppt1 As String           ' 예: "LED_XSFAIL_1.ppt", ...
        Public ppt2 As String
        Public sendCmd As String        ' 예: "TS,0A00,0", "TS,0D00,0" ...
        Public checkTitle As String     ' 예: "XSFAILの確認", "Main LEDの確認" ...
        Public checkTarget As String    ' 예: "XSFAIL LED", "Main LED"
        Public checkColor As String     ' 예: "GREEN"
        Public offText As String        ' 예: "OFF"
        Public onText1 As String        ' 예: "ON"
        Public onText2 As String        ' 예: "ON"
        Public errMsg As String         ' 예: "XSFAIL LED 異常"
    End Structure

    Public Sub TestExecute(ByVal iName As String)
        myEq = New clsBaseEquipments

        Select Case Mid(iName, 1, 3)


            Case "I01" : Call I01(iName)            'I01:検査メインプログラムのVersionテスト
            Case "I02" : Call I02(iName)            'I02:Ethernetテスト
            Case "I03" : Call I03(iName)            'I03スイッチテスト
            Case "I04" : Call I04(iName)            'I04:FPGAテスト
            Case "I05" : Call I05(iName)            'I05:電源モジュールテスト(XDFAIL)
            Case "I06" : Call I06(iName)            'I05:電源モジュールテスト(XPFAIL)


            Case "I10" : Call I10(iName)            'I10:電圧確認テスト

            Case "I11" : Call I11(iName)            'I11:MRAM初期化(초기화후  off/on 필요)
            Case "I12" : Call I12(iName)            'I12:MRAMテスト

            Case "I13" : Call I13(iName)            'I13:eMMC信号線テスト
            Case "I14" : Call I14(iName)            'I14:Q-SPI信号線テスト

            Case "I15" : Call I15(iName)            'I15:バススピードテスト(SDメモリカードテスト)
            Case "I16" : Call I16(iName)            'I16:ライトプロテクトテスト(SDメモリカードテスト)
            Case "I17" : Call I17(iName)            'I17短絡保護テスト(SDメモリカードテスト)
            Case "I18" : Call I18(iName)            'I18:ライトプロテクトテスト(SDメモリカードテスト)

            Case "I20" : Call I20(iName)            'I20:バッテリ電圧電流確認テスト
            Case "I21" : Call I21(iName)            'I21:バッテリ電圧低下確認テスト

            Case "I30" : Call I30(iName)            'I30:時刻設定テスト(RTCテスト)
            Case "I31" : Call I31(iName)            'I31:バックアップテスト（/BATのみ実施) (RTCテスト)

            Case "I40" : Call I40(iName)            'I40:LEDテスト

            Case "I89" : Call I89(iName)            'I89:SD挿抜TEST
            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF

            Case "I92" : Call I92(iName)            'I92:検査開始時処理
            Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save    
            Case "I95" : Call I95(iName)            'I95:測定器Preset ' 수정필요=======
            Case "I97" : Call I97(iName)            'I97:ﾕﾆｯﾄ電源OFF&ON
            Case "I98" : Call I98(iName)            'I98:出荷設定確認
            Case "I99" : Call I99(iName)            'I99:最初ﾕﾆｯﾄ電源ON

            Case "IA1" : Call IA1(iName)            'IA1:eMMC Download
            Case "IA2" : Call IA2(iName)            'IA2:eMMC SUM値算出テスト
            Case "IA3" : Call IA3(iName)            'IA3:Q-SPI SUM値算出テスト
            Case "IA4" : Call IA4(iName)            'IA4:Mac Address Port0
            Case "IA5" : Call IA5(iName)            'IA5:Mac Address Port1
            Case "IA6" : Call IA6(iName)            'IA6:Mac Address Port2
            Case "IA7" : Call IA7(iName)            'IA7:Mac Address Port3
            Case "IA8" : Call IA8(iName)            'IA8:Mac Address Port4 (TSN END POINT)
            Case "IA9" : Call IA9(iName)            'IA9:MACｱﾄﾞﾚｽ重複確認

            Case "IB1" : Call IB1(iName)            'IB1:Flash Memoryへの書込み
            Case "IB2" : Call IB2(iName)            'IB2:Flash MemoryへMac Address書込み
            Case "IB3" : Call IB3(iName)            'IB3:Ubuntu Patch
            Case "IB4" : Call IB4(iName)            'IB4:Q-SPI Download

            'Ver 0.01 HSJ ADD STRART  ====
            Case "IC1" : Call IC1(iName)            'IC1:ランタイムライセンス設定
            Case "IC2" : Call IC2(iName)            'IC2:TPMテスト
            Case "IC3" : Call IC3(iName)            'IC3:eFUSE設定
                'Ver 0.01 HSJ ADD END  ====

            Case Else : mErrorComment = "検査プロシージャが定義されていません。"
                MsgBox(mErrorComment)

        End Select


    End Sub

    Public Sub I01(ByVal iName As String)
        '==============================================================================================================
        ' Function : I01:検査メインプログラムのVersionテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As String = ""
        Dim rData() As String = {""}
        Dim dbName As String = mAppPath & "520 Product Firmware\A8HCA\MainPRGRev.mdb"
        Dim rField As String = "Ver"
        Dim Model As String = ModelRead()
        Dim tbName As String = "tblMainPRG"
        Dim fName() As String = {""}
        Dim sSQL As String = ""
        Dim sData() As String = Nothing
        Dim cADO As cls010_MDB_Operation = Nothing
        Dim strSetRev As String = ""
        '==============================================================================================================

        Try
            Try
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName

                sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model LIKE '" & Model & "'"
                cADO.DB_SQL = sSQL
                Call cADO.DB_READ_JET(fName, sData)
                strSetRev = sData(0)

            Catch ex As Exception
                cADO = Nothing
                mErrorComment = "DB読み込み失敗"
                Throw New ApplicationException(mErrorComment)
            End Try



            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            '検査メインプログラムのVersionテスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "検査メインプログラムのVersionテスト　開始" & vbCrLf)


            'Send command ---------------------------
            cIA8Comm.SendCmd = $"TS,0DC0,1,{strSetRev}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = $"OK,1DC0,{strSetRev}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "検査メインプログラムのVersionテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "検査メインプログラムのVersionテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "検査メインプログラムのVersionテストOK")

    End Sub


    Public Sub I02(ByVal iName As String)
        '==============================================================================================================
        ' Function : I02:Ethernetテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}
        Dim retValue As String = ""
        'Dim mMeas As Double
        '==============================================================================================================

        Try
            ShowMessagePopupAndCheck(mAppPath,
                                         "LANLED.ppt",
                                         "Ethernetの確認",
                                         "Ethernetテスト",
                                         "YELLOW", "OFF", "OFF", "OFF",
                                         "Ethernet LED 異常")
            mSecWait(2000)

            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'Ethrnet검사 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "Ethernet検査開始" & vbCrLf)


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0420" & ",1" & ",3E8"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(300)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1420,1,1"
                mTestMeasure = cIA8Comm.RECEIVE_DATA

                Call WriteLog(False, "判定基準 : OK,1420,1,1")
                mErrorComment = "Ethernet検査" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "LAN0-3のLNK, ACT LEDの目視検査 開始") ' Ver 0.02 hsj add
            Do
                ' Retry PPT 추가
                ShowMessagePopupAndCheck_RETRY(mAppPath,
                                     "LANLED_1.ppt",
                                     "Ethernetの確認",
                                     "Ethernetテスト",
                                     "YELLOW",
                                     "OFF",
                                     "OFF",
                                     "OFF",
                                     "Ethernet LED 異常",
                                     retValue)

                Select Case retValue
                    Case "OK"
                        Call WriteLog(False, "LNK, ACT LEDの目視検査 完了")
                        Exit Do

                    Case "NG"
                        Throw New ApplicationException

                    Case "RETRY"
                        ' LED TEST 커맨드 재송신
                        ' SendCmd 
                        Call WriteLog(False, "LNK, ACT LEDの目視検査 RETRY")
                        cIA8Comm.SendCmd = "TS," & "0420" & ",1" & ",3E8"
                        WriteLog(False, "Send : " & cIA8Comm.SendCmd)
                        cIA8Comm.FuncSendCmd()

                        mSecWait(500)

                        ' RecvCmd
                        cIA8Comm.FuncRecvCmd()
                        WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

                        ' Recv Data 파싱
                        Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
                        If Me.ERR_COMMENT <> "" Then
                            mTestSpec = "OK,1420,1,1"
                            mTestMeasure = cIA8Comm.RECEIVE_DATA

                            Call WriteLog(False, "判定基準 : OK,1420,1,1")
                            mErrorComment = "Ethernetテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                            Throw New ApplicationException(mErrorComment)
                        End If

                        ' PPT2 반복 => Do 루프 계속
                    Case Else
                        Throw New ApplicationException("Ethernet TEST: " & retValue)
                End Select
            Loop

            'Try
            '    dfrm = New frm500_MessagePopUp
            '    dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\A8HCA\LANLED.ppt"
            '    dfrm.DISP_TITLENAME = "LAN LED 検査"
            '    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            '    Call dfrm.frmShow()

            '    Do
            '        'wait push button -------------------
            '        If dfrm.DIALOG_RESULT = "NG" Then
            '            mErrorComment = "Ethernet LED 異常"
            '            Throw New ApplicationException(mErrorComment)
            '        End If

            '        If dfrm.DIALOG_RESULT = "OK" Then
            '            Exit Do
            '        End If

            '        DoEvents()

            '        Call mSecWait(200)

            '    Loop
            'Catch ex As Exception
            '    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            '    mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            '    Call WriteLog(False, mErrorComment)
            '    Throw New ApplicationException(mErrorComment)
            'Finally
            '    dfrm.Close()
            '    dfrm = Nothing
            'End Try



            ''LAN LED 준비******
            ''Popup ============================================
            'Call WriteLog(False, "")
            'Call WriteLog(False, "LAN LED 目視検査結果")

            'Try
            '    dfrm = New frm500_MessagePopUp
            '    dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\A8HCA\LANLED_1.ppt"
            '    dfrm.DISP_TITLENAME = "LAN LED 検査"
            '    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            '    Call dfrm.frmShow()

            '    Do
            '        'wait push button -------------------
            '        If dfrm.DIALOG_RESULT = "NG" Then
            '            mErrorComment = "Ether LED 異常"
            '            Throw New ApplicationException(mErrorComment)
            '        End If

            '        If dfrm.DIALOG_RESULT = "OK" Then
            '            Exit Do
            '        End If

            '        DoEvents()

            '        Call mSecWait(200)

            '    Loop

            '    Call WriteLog(False, "LAN LED 目視検査OK")

            'Catch ex As Exception
            '    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            '    mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            '    Call WriteLog(False, mErrorComment)
            '    Throw New ApplicationException(mErrorComment)
            'Finally
            '    dfrm.Close()
            '    dfrm = Nothing
            'End Try

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "Ethernetテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "Ethernetテスト OK")

    End Sub
    Public Sub I03(ByVal iName As String)
        '==============================================================================================================
        ' Function : I03 スイッチテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}
        Dim mFileName As String

        Dim modeArray() As String = {"0", "1"}
        Dim valueArray() As String = {"1", "2", "4", "8"}
        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            '모드 스위치(RSW) 테스트 
            Call WriteLog(False, "")
            Call WriteLog(False, "モードスイッチテスト開始")

            For Each mode As String In modeArray
                For Each val As String In valueArray
                    mFileName = String.Format("RSW_{0}_{1}.ppt", mode, val)

                    ' PPT Send
                    ShowMessagePopupAndCheck(mAppPath,
                                         mFileName,
                                         "Rotary SWの確認",
                                         "Rotary SW TR",
                                         "GREEN", "OFF", "ON", "ON",
                                         "モードスイッチテスト備　異常")
                    Call WriteLog(False, String.Format("Switch {0}:{1} 確認 OK", mode, val)) 'ver 0.02 hsj add

                Next
            Next
            Call WriteLog(False, "モードスイッチテストOK" & vbCrLf)

            'Set 스위치 테스트
            Call WriteLog(False, "SET スイッチテスト開始")

            ' PPT Send
            ShowMessagePopupAndCheck(mAppPath,
                                 "SET_SW.ppt",
                                 "SET SWの確認",
                                 "SET SW PUSH",
                                 "GREEN", "OFF", "ON", "ON",
                                 "SET SW　異常")

            '바로 날리면 고속으로 통신 끊김
            mSecWait(1500)
            ' Send
            cIA8Comm.SendCmd = "TS," & "0D50" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(1500)
            ' Recv
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

            ' Recv Data 파싱
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1D50,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, "判定基準 :OK,1D50,0")

                mErrorComment = "SETスイッチテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "SET　スイッチテストOK" & vbCrLf)

            'Reset 스위치 테스트
            Call WriteLog(False, "RESET スイッチテスト開始")

            ShowMessagePopupAndCheck(mAppPath,
                                 "RESET_SW.ppt",
                                 "RESET SWの確認",
                                 "RESET SW PUSH",
                                 "GREEN", "OFF", "ON", "ON",
                                 "RESET SW　異常")

            Call WriteLog(False, "RESET スイッチテストOK")

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "スイッチテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "スイッチテストOK")
    End Sub

    Public Sub I11(ByVal iName As String)
        '==============================================================================================================
        ' Function : I11:MRAM初期化
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'MRAM初期化***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "MRAM初期化開始")


            'Send command ---------------------------
            'cIA8Comm.SendCmd = "TS," & "0330" & ",0" 'Ver 0.02 HSJ del
            cIA8Comm.SendCmd = "TS,0330,1,0" 'Ver 0.02 HSJ add
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(5000) ' MRAM 초기화에 시간이 소요되지 않음

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1330,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "MRAM初期化" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            mSecWait(500)
            Call WriteLog(False, "MRAM初期化 OK" & vbCrLf)

            'MRAM초기화-재기동 ***************************************************
            ' 전원 OFF
            Call WriteLog(False, "モジュールの再起動開始")
            Me.I91("検査モジュールの電源OFF")
            mSecWait(2000)
            Me.I90("検査モジュールの電源ON")
            'Call WaitStart(20)
            'mSecWait(10000)
            Call WriteLog(False, "モジュールの再起動OK" & vbCrLf)

            'Ver 0.02 HSJ add Start===
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()
            Call WriteLog(False, "")
            Call WriteLog(False, "MRAM LOG 初期化開始")

            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0330,1,1"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(3000)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1330,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "MRAMLOG 初期化" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            mSecWait(500)
            Call WriteLog(False, "MRAM LOG 初期化 OK" & vbCrLf)

            'MRAM초기화-재기동 ***************************************************
            ' 전원 OFF
            Call WriteLog(False, "モジュールの再起動開始")
            Me.I91("検査モジュールの電源OFF")
            mSecWait(2000)
            Me.I90("検査モジュールの電源ON")
            'Call WaitStart(20)
            'mSecWait(10000)
            Call WriteLog(False, "モジュールの再起動OK" & vbCrLf)

            'Ver 0.02 HSJ add End===


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "MRAM初期化 NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "MRAM初期化OK")

    End Sub

    Public Sub I12(ByVal iName As String)
        '==============================================================================================================
        ' Function : I12:MRAMテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'MRAM信号線テスト***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "MRAM信号線テスト開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0300" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1300,100,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "MRAMテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            mSecWait(500)
            Call WriteLog(False, "MRAM信号線テスト OK" & vbCrLf)


            'MRAMテスト-Write ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "MRAMテスト-Write開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0310" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1310,100,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "MRAMテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "MRAMテスト-Write OK" & vbCrLf)

            'MRAMテスト-재기동 ***************************************************
            ' 전원 OFF
            Call WriteLog(False, "モジュールの再起動開始")
            Me.I91("検査モジュールの電源OFF")
            mSecWait(2000)
            Me.I90("検査モジュールの電源ON")
            Call WriteLog(False, "モジュールの再起動OK" & vbCrLf)

            'MRAMテスト-Read ***************************************************
            Call WriteLog(False, "MRAMテスト-Read開始")
            mSecWait(5000)
            'A8 Communication =================================
            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0320" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(500)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1320,100,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "MRAMテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "MRAMテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "MRAMテストOK")

    End Sub

    Public Sub I13(ByVal iName As String)
        '==============================================================================================================
        ' Function : I13:eMMC信号線テスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'eMMC信号線テスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "eMMC信号線テスト開始" & vbCrLf)


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0200" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1200,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "eMMC信号線テスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "eMMC信号線テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "eMMC信号線テストOK")

    End Sub

    Public Sub I14(ByVal iName As String)
        '==============================================================================================================
        ' Function : I14:Q-SPI信号線テスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'eMMC信号線テスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "Q-SPI信号線テスト開始" & vbCrLf)


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0100" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1100,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "Q-SPI信号線テスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "Q-SPI信号線テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "Q-SPI信号線テストOK")

    End Sub
    Public Sub I15(ByVal iName As String)
        '==============================================================================================================
        ' Function : I15:バススピードテスト(SDメモリカードテスト)
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'バススピードテスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "バススピードテスト(SDメモリカードテスト)開始" & vbCrLf)

            'mSecWait(3000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0810" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            WaitStart(40)
            'mSecWait(8000)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1810,53,0,7"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "バススピードテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "バススピードテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "バススピードテスト(SDメモリカードテスト)OK")

    End Sub
    Public Sub I16(ByVal iName As String)
        '==============================================================================================================
        ' Function : I16: ライトプロテクトテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        ' SSH 관련 변수
        Dim host As String = "192.168.0.10"
        Dim user As String = "a8admin"
        Dim pass As String = "a8pass"
        'Dim user As String = "root"
        'Dim pass As String = "root_ykgw"
        Dim result As String = ""

        '==============================================================================================================

        Try
            ' 라이트프로텍트테스트 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "ライトプロテクトテスト開始")


            Me.I91("検査モジュールの電源OFF")


            '  SD카드 교체, LAN 케이블 교체
            ShowMessagePopupAndCheck(mAppPath,
                                 "SD_Test.ppt",
                                 "SD_Testの確認",
                                 "SD_Test",
                                 "GREEN", "OFF", "ON", "ON",
                                 "SD_Test異常")

            ' 전원 ON - eMMC 부트
            Me.I90("検査モジュールの電源ON")


            'result = ssh.ExecuteRemoteCommand(host, user, pass, "cat /sys/kernel/debug/mmc1/ios")
            'Call WriteLog(False, "バススピードテスト: " & result)

            'If result.Contains("DDR50") OrElse result.Contains("SDR50") Then
            '    Call WriteLog(False, "バススピードテスト OK ")
            'Else
            '    Throw New ApplicationException("バススピードテスト NG: DDR50及び SDR50 無し")
            'End If


            ' SD카드 Write Protect 상태 확인
            '    "cat /sys/block/mmcblk1/ro" 결과값: 0 → 쓰기 가능, 1 → 쓰기 금지
            result = ssh.ExecuteRemoteCommand(host, user, pass, "cat /sys/block/mmcblk1/ro")
            Call WriteLog(False, "ライトプロテクトテスト 状態: " & result.Trim())

            If result.Trim() = "0" Then
                ' 쓰기 가능
                mErrorComment = "ライトプロテクト 異常 : Write 可能"
                Throw New ApplicationException(mErrorComment)
            Else
                ' 쓰기 금지
                Call WriteLog(False, "ライトプロテクト OK : Write 不可能")
            End If

            'SD 카드 재교체 및 로타리 스위치 0에서 E로 돌릴 것

            ShowMessagePopupAndCheck(mAppPath,
                                 "SD_Test1.ppt",
                                 "SD_Testの確認",
                                 "SD_Test",
                                 "GREEN", "OFF", "ON", "ON",
                                 "SD_Test異常")


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "ライトプロテクトテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

        End Try

        Call WriteLog(False, "ライトプロテクトテストOK")

    End Sub



    Public Sub I17(ByVal iName As String)
        '==============================================================================================================
        ' Function : I17:短絡保護テスト(SDメモリカードテスト)
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}
        Dim FSR As Double = 6.144     ' ±6.144 V
        Dim ADC_max As Double = 32752 ' ADC 최대값

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()
            ' 短絡保護テスト ***************************************************
            ' SD카드 부트
            Call WriteLog(False, "短絡保護テスト 開始")

            ' SD 카드 교체, 
            ShowMessagePopupAndCheck(mAppPath,
                                 "SD_Short.ppt",
                                 "SD_ShortTestの確認",
                                 "SD_Short",
                                 "GREEN", "OFF", "ON", "ON",
                                 "SD_ShortTest異常")

            ' SD 카드 단락전 전압측정
            'MeasureVoltage(cIA8Comm, "3_3V_SD_電源電圧_短絡前", "0920", "46CC", "408D", FSR, ADC_max)
            MeasureVoltage(cIA8Comm, "3_3V_SD_電源電圧_短絡前", "0920", "46D5", "42AA")
            ' 전원 on - SD 카드 부팅
            'Me.I90("検査モジュールの電源ON")

            ' SD 카드 단락
            PLC_Y("01" & "BWRY" & "306" & ",001," & "1")
            '단락하고 SD 전압 측정
            'MeasureVoltage(cIA8Comm, "3_3V_SD_電源電圧_短絡", "0920", "A69", "0000", FSR, ADC_max)
            MeasureVoltage(cIA8Comm, "3_3V_SD_電源電圧_短絡", "0920", "A69", "0000")

            ' SD 카드 단락 해제
            PLC_Y("01" & "BWRY" & "306" & ",001," & "0")

            ' SD 카드 교체, 
            ShowMessagePopupAndCheck(mAppPath,
                                 "SD_Short_1.ppt",
                                 "SD_ShortTestの確認",
                                 "SD_Short",
                                 "GREEN", "OFF", "ON", "ON",
                                 "SD_ShortTest異常")

            Call WriteLog(False, "短絡保護テスト OK")

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "短絡保護テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
            PLC_Y("01" & "BWRY" & "306" & ",001," & "0")
        End Try

        Call WriteLog(False, "短絡保護テスト OK")

    End Sub

    Public Sub I18(ByVal iName As String)
        '==============================================================================================================
        ' Function : I18:ライトプロテクトテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}


        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()
            ' ライトプロテクトテスト ***************************************************
            ' SD카드 교체 - Write Protect 교체
            Call WriteLog(False, "ライトプロテクトテスト 開始")

            ' SD 카드 교체, 
            'ShowMessagePopupAndCheck(mAppPath,
            '                     "SD_Short.ppt",
            '                     "SD_ShortTestの確認",
            '                     "SD_Short",
            '                     "YELLOW", "OFF", "ON", "ON",
            '                     "SD_ShortTest異常")



            ' 전원 on - SD 카드 부팅
            'Me.I90("検査モジュールの電源ON")
            mSecWait(1000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0800,0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(1000)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            If Not String.IsNullOrEmpty(cIA8Comm.RECEIVE_DATA) Then
                Dim splitted() As String = cIA8Comm.RECEIVE_DATA.Split(","c)
                If splitted.Length >= 3 Then
                    Dim lastValue As String = splitted(splitted.Length - 1).Trim()

                    If lastValue = "0" Then
                        ' 0이면 NG
                        mErrorComment = "ライトプロテクトテスト : NG (" & cIA8Comm.RECEIVE_DATA & ")"
                        Call WriteLog(False, mErrorComment)
                        Throw New ApplicationException(mErrorComment)
                    ElseIf lastValue = "1" Then
                        ' 1이면 PASS
                        Call WriteLog(False, "ライトプロテクトテスト : PASS (" & cIA8Comm.RECEIVE_DATA & ")")
                    Else
                        ' 그 외 예외처리 (필요시)
                        ' mErrorComment = "Unexpected Response: " & cIA8Comm.RECEIVE_DATA
                        ' Throw New ApplicationException(mErrorComment)
                    End If
                End If
            End If


            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "ライトプロテクトテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If

            ' SD 카드 교체, 
            ShowMessagePopupAndCheck(mAppPath,
                                 "SD_Short_2.ppt",
                                 "SDの交換",
                                 "SDの交換",
                                 "YELLOW", "OFF", "OFF", "OFF",
                                 "SD_ShortTest異常")


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "ライトプロテクトテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
            PLC_Y("01" & "BWRY" & "306" & ",001," & "0")
        End Try

        Call WriteLog(False, "ライトプロテクトテストOK")

    End Sub

    Public Sub I20(ByVal iName As String)
        '==============================================================================================================
        ' Function : I20:バッテリ電圧電流確認テスト（/NBT未実施）
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Dim batVoltValue As Double
        Dim batCurrValue As Double
        Dim batCurrValue_uA As Double
        '==============================================================================================================

        Try
            ' 전원 OFF
            Me.I91("検査モジュールの電源OFF")

            PLC_Y("01" & "BWRY" & "301" & ",032," & "00000000000000000000000000000000")

            ' 배터리 전압 테스트
            Call WriteLog(False, "バッテリ電圧テスト開始" & vbCrLf)

            DMMRead(1, "V10V")
            batVoltValue = Dmm_Read_Data

            Call WriteLog(False, "バッテリ電圧 SPEC : 2.8V≦測定値≦3.5V")

            ' 범위 판정(2.8V 이상 ~ 3.2V 이하) -> BAT 전압 spec 변경 max 3.2 -> 3.5V
            If batVoltValue >= 2.8 AndAlso batVoltValue <= 3.5 Then
                Call WriteLog(False, "バッテリ電圧測定値: " & batVoltValue.ToString("F3") & " V")
                Call WriteValueData("bat_volt", CStr(batVoltValue)) 'Ver 0.03 HSJ Add
            Else
                Call WriteLog(False, "バッテリ電圧が規定範囲(2.8～3.5V)外です: " & batVoltValue.ToString("F3") & " V")
                mErrorComment = "バッテリ電圧NG"
                mTestSpec = "2.8≦測定値≦3.5"
                Call WriteLog(False, "spec : " & mTestSpec)
                Throw New ApplicationException(mErrorComment)
            End If


            ' 배터리 전류 테스트
            Call WriteLog(False, "バッテリ電流テスト開始")
            PLC_Y("01" & "BWRY" & "304" & ",002," & "11")
            mSecWait(1500)
            'DMMRead(1, "A100uA") 'Auto로 해야 uA로 측정이 가능
            DMMRead(1, "AAUTO")
            batCurrValue = Dmm_Read_Data
            If batCurrValue < 0 Then
                mErrorComment = "バッテリ電流 NG : ケーブル接続を確認してください " & vbCrLf & "(測定値 : " & batCurrValue.ToString("E") & "A)"
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If

            batCurrValue_uA = batCurrValue * 1000000.0

            ' 범위 판정(1µA 이하)
            If batCurrValue_uA <= 1.0 Then
                Call WriteLog(False, "バッテリ電流測定値: " & batCurrValue_uA.ToString("F6") & " µA (1µA 以下)")
                Call WriteValueData("bat_curr", CStr(batCurrValue_uA)) 'Ver 0.03 HSJ Add
            Else
                mTestSpec = "規定値≦1µA"
                mErrorComment = "バッテリ電流が規定値(1µA)NG : " & batCurrValue_uA.ToString("F6") & " µA"
                Call WriteLog(False, "spec : " & mTestSpec)
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException("バッテリ電流範囲エラー")
            End If
            Me.I90("検査モジュールの電源ON")
        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "バッテリ電圧電流確認テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            PLC_Y("01" & "BWRY" & "304" & ",002," & "00")

        End Try

        Call WriteLog(False, "バッテリ電圧電流確認テストOK")

    End Sub
    Public Sub I21(ByVal iName As String)
        '==============================================================================================================
        ' Function : I21:バッテリ電圧低下確認テスト（/NBT未実施）
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Try
            ' 배터리 전압 저하 테스트
            Call WriteLog(False, "バッテリ電圧低下確認テスト開始" & vbCrLf)

            ' 배터리 A8HCA에 연결 YC08, 4,8
            PLC_Y("01" & "BWRY" & "304" & ",001," & "1")
            PLC_Y("01" & "BWRY" & "308" & ",001," & "1")
            ' 전원 OFF
            Me.I91("検査モジュールの電源OFF")
            WaitStart(3)
            ' 전원 ON
            Me.I90("検査モジュールの電源ON")
            ' 배터리 확인 지시 ppt
            ShowMessagePopupAndCheck(mAppPath,
                                 "BAT_LED_1.ppt",
                                 "BAT LEDの確認",
                                 "BAT LED",
                                 "YELLOW", "OFF", "OFF", "OFF",
                                 "BAT LED準備異常")

            ' 배터리 A8HCA에 연결 해제 YC08, 4, 8
            PLC_Y("01" & "BWRY" & "304" & ",001," & "0")
            PLC_Y("01" & "BWRY" & "308" & ",001," & "0")

            Call WriteLog(False, "バッテリショット開始")
            ' 배터리 A8HCA 쇼트 YC08, 7
            PLC_Y("01" & "BWRY" & "307" & ",001," & "1")

            ' BAT LED 켜지는데 걸리는 시간
            WaitStart(35)
            Call WriteLog(False, "BATLED目視検査開始")
            ' BAT LED 검사 ppt
            ShowMessagePopupAndCheck(mAppPath,
                                 "BAT_LED_2.ppt",
                                 "BAT LEDの確認",
                                 "BAT LED",
                                 "YELLOW", "OFF", "OFF", "OFF",
                                 "BAT LED異常")

            ' 배터리 A8HCA 쇼트 해제 YC08, 7
            PLC_Y("01" & "BWRY" & "307" & ",001," & "0")

            If InStr(Model, "NBT") = 0 Then
                ShowMessagePopupAndCheck(mAppPath,
                     "DustCap2.ppt",
                     "DustCap Check",
                     "DustCap Check",
                     "YELLOW", "OFF", "OFF", "OFF",
                     "DustCap 準備　異常")
            End If

            '' 배터리 A8HCA에 연결 YC08, 4,8
            'PLC_Y("01" & "BWRY" & "304" & ",001," & "1")
            'PLC_Y("01" & "BWRY" & "308" & ",001," & "1")
            'mSecWait(4000)

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "バッテリ電圧低下確認テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally

        End Try

        Call WriteLog(False, "バッテリ電圧低下確認テストOK")

    End Sub


    Public Sub I30(ByVal iName As String)
        'ver 0.02 hsj add start===
        '==============================================================================================================
        ' Function : I30:RTC設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer = 0
        Dim rData() As String = {}

        Dim nWeek As String = ""

        Dim Model As String = ModelRead()
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        ' ────────────────────────────────────── 1. 로그 시작
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        ' ────────────────────────────────────── 2. UTC 확인

        Call WriteLog(False, "RTC TimeZone : UTC")          ' 장치 RTC 기준(UTC) 고정 표기

        ' ────────────────────────────────────── 3. 현지(PC) 시각·요일 계산
        Dim nowPC As DateTime = DateTime.Now                      '  현지 시각 사용
        Dim year4 As String = nowPC.Year.ToString("0000")
        Dim month2 As String = nowPC.Month.ToString()
        Dim day2 As String = nowPC.Day.ToString()
        Dim hour2 As String = nowPC.Hour.ToString()
        Dim min2 As String = nowPC.Minute.ToString()
        Dim sec2 As String = nowPC.Second.ToString()

        ' 로그용 표시
        Dim rtcDisp As String = String.Join(",", {year4, month2, day2, hour2, min2, sec2})
        Call WriteLog(False, "RTC設定時刻：" & rtcDisp)

        ' ────────────────────────────────────── 4. Send 커맨드 조립
        Dim sendCmd As String = $"TS,0610,6,{year4},{month2},{day2},{hour2},{min2},{sec2}"

        ' ────────────────────────────────────── 5. RS-232C 통신
        Try
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                Throw New ApplicationException("RS232SET Read異常")
            End If

            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            ' ── Send
            cIA8Comm.SendCmd = sendCmd
            Call WriteLog(False, "Send : " & sendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(6000)

            ' ── Recv
            Call cIA8Comm.FuncRecvCmd()
            Call RPRecvCmd_1(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            If Me.ERR_COMMENT <> "" Then
                mTestSpec = rtcDisp
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Throw New ApplicationException("Recv : " & Me.ERR_COMMENT & "　異常")
            End If

            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA & "　正常")

            ' ────────────────────────── 6. 보조 로그
            Call WriteLog(False, $"RTCSetDate:{nowPC:yy/MM/dd}")
            Call WriteLog(False, $"RTCSetTime:{nowPC:HH:mm:ss}")

            Call WriteLog(False, iName & " 正常")

        Catch ex As Exception
            Call WriteLog(False, ex.Message)
            Throw
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try
        'ver 0.02 hsj add end===

        'ver 0.02 hsj del start===
        ''==============================================================================================================
        '' Function : I30:RTC設定
        '' Data     : Process name
        '' Return   : Nothing
        '' Remarks  :
        ''==============================================================================================================

        'Dim mSetting As String = ""
        'Dim mDelim As String = ""

        'Dim rRes As String = ""
        'Dim rCode As String = ""
        'Dim rCount As Integer = 0
        'Dim rData() As String = {}

        'Dim nDate As String = ""
        'Dim nWeek As String = ""
        'Dim nTime As String = ""
        'Dim nAll As String = ""
        'Dim wData As String = ""
        'Dim i As Integer = 0

        'Dim dateValue As Date
        'Dim WeekValue As String

        'Dim Model As String = ModelRead()
        'Dim dbName As String
        'Dim tbName As String
        'Dim sSQL As String
        'Dim fName() As String = {""}
        'Dim sData() As String = {""}
        'Dim rField As String
        'Dim kField As String
        'Dim TimeZone As String

        'Call WriteLog(False, "")
        'Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        'Call WriteLog(False, "*")
        ''==========================================================================================


        'Try
        '    '出荷設定Dataの読込**********
        '    cADO = New cls010_MDB_Operation

        '    dbName = mAppPath & "530 Shipment Setting\ShipmentSetting.mdb"

        '    tbName = "tblShipmentSetting"
        '    rField = "RTC" : kField = "MSCODE"
        '    sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & kField & " = '" & Model & "'"

        '    cADO.DB_PATH = dbName
        '    cADO.TB_NAME = tbName
        '    cADO.DB_SQL = sSQL

        '    Call cADO.DB_READ_JET(fName, sData)
        '    cADO = Nothing

        '    TimeZone = sData(0)


        '    If TimeZone = "" Then
        '        mErrorComment = "ShipmentSetting.mdb登録無し。"
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    Call WriteLog(False, "RTC TimeZone : " & TimeZone)

        'Catch ex As Exception
        '    mErrorComment = "ShipmentSetting.mdb 異常"
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End Try
        '' 시간대(Time Zone) 확인
        'If TimeZone = "JST/KST" Then
        '    nAll = System.DateTime.Now.ToString("yy/MM/dd HH:mm:ss")
        '    nWeek = CStr(CInt(Format(DatePart(DateInterval.Weekday, Now))) - 1)
        '    dateValue = Date.Today
        'ElseIf TimeZone = "UTC" Then
        '    nAll = System.DateTime.UtcNow.ToString("yy/MM/dd HH:mm:ss")
        '    nWeek = CStr(CInt(Format(DatePart(DateInterval.Weekday, System.DateTime.UtcNow))) - 1)
        '    dateValue = System.DateTime.UtcNow
        'Else
        '    mErrorComment = "ShipmentSetting.mdb 異常"
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End If


        '' 현재 시간을 토대로 TS커맨드 형식 변환 
        'nDate = Mid(nAll, 1, 8)

        'nTime = Mid(nAll, 10, 8)
        'nAll = Mid(nDate, 1, 2)                                 '年
        'nAll &= (CInt(Mid(nDate, 4, 2))).ToString("00")         '月
        'nAll &= Mid(nDate, 7, 2)                                '日
        'nAll &= nWeek                                           '曜日番号
        'nAll &= Mid(nTime, 1, 2)                                '時
        'nAll &= Mid(nTime, 4, 2)                                '分
        'nAll &= Mid(nTime, 7, 2)                                '秒

        'For i = 1 To Len(nAll)
        '    If i <> 13 Then
        '        wData = wData & Mid(nAll, i, 1) & ","
        '    Else
        '        wData = wData & Mid(nAll, i, 1)
        '    End If
        'Next

        'Call WriteLog(False, "月は0始まり、曜日は日曜が0")
        'Call WriteLog(False, "RTC設定時刻：" & wData)

        'Try

        '    'RP Communication =================================
        '    cIA8Comm = New cls101_ComRS232C
        '    cIA8Comm.RS_COMM = Insfrm.SerialPort1
        '    cIA8Comm.PORT = rs232cRZPort
        '    If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
        '        mErrorComment = "RS232SET Read異常"
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    cIA8Comm.SETTING = mSetting : cIA8Comm.TERMINAL_CHAR = mDelim
        '    cIA8Comm.PortSetup()

        '    'Send command ---------------------------
        '    cIA8Comm.SendCmd = "TS," & "0610" & ",D," & wData
        '    Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
        '    Call cIA8Comm.FuncSendCmd()

        '    'Recv Command  --------------------------
        '    Call cIA8Comm.FuncRecvCmd()
        '    Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

        '    If Me.ERR_COMMENT <> "" Then
        '        mTestSpec = nAll
        '        mTestMeasure = cIA8Comm.RECEIVE_DATA
        '        mErrorComment = "Recv : " & Me.ERR_COMMENT & "　異常"
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA & "　正常")

        '    Call WriteLog(False, "RTCSetDate:" & Mid(nDate, 1, 2) & "/" & (CInt(Mid(nDate, 4, 2))).ToString("00") & "/" & Mid(nDate, 7, 2))
        '    Call WriteLog(False, "RTCSetWeek:" & nWeek)



        '    WeekValue = dateValue.ToString("dddd")
        '    Call WriteLog(False, "RTCSetWeekは曜日 : " & WeekValue)

        '    Call WriteLog(False, "RTCSetTime:" & nTime)

        '    Call WriteLog(False, iName & " 正常")
        '    ' =============================================================

        'Catch ex As Exception
        '    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
        '    mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    If cIA8Comm IsNot Nothing Then
        '        cIA8Comm.PortClose()
        '        cIA8Comm = Nothing
        '    End If
        'End Try
        'ver 0.02 hsj del end===
    End Sub
    Public Sub I31(ByVal iName As String)
        'ver 0.02 hsj add start ===
        '==============================================================================================================
        ' Function : I31:バックアップテスト（/NBT未実施) (RTCテスト)
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : Recv 응답시간이 PC시간보다 9시간 전이어야 함 (10초 오차)
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Dim temp(12) As String
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer = 0
        Dim rData() As String = {}

        Dim RPDate As String = ""
        Dim RPTime As String = ""

        Dim Month As Integer = 0
        Dim Month_str As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        Try
            ' RS232 통신 설정
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort
            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            ' RTC 데이터 요청
            cIA8Comm.SendCmd = "TS," & "0620" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(2000)

            ' 응답 수신 및 PC 시간 읽기 (Recv 직후 바로 PC 시간 획득)
            Call cIA8Comm.FuncRecvCmd()
            Dim pcTimeAtRecv As DateTime = DateTime.Now  ' Recv 받은 시점의 PC 시간

            Call RPRecvCmd_1(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            If Me.ERR_COMMENT <> "" Then
                mErrorComment = "Recv : " & Me.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call WriteLog(False, "PC Time at Recv : " & pcTimeAtRecv.ToString("yy/MM/dd HH:mm:ss"))

            ' 모듈에서 읽은 RTC 데이터 파싱
            For i As Integer = 0 To 12
                temp(i) = Hex(Val("&h" & rData(i)))  ' 16진 → 10진 
            Next i

            ' 월 계산 (temp(2)="0", temp(3)="3" → "03")
            Month = CInt(temp(2) & temp(3))
            Month_str = Month.ToString("00")

            ' 날짜(YY/MM/DD) 구성
            RPDate = temp(0) & temp(1) & "/" & Month_str & "/" & temp(4) & temp(5)
            Call WriteLog(False, "Module Date : " & RPDate)

            ' 시간(HH:mm:ss) 구성
            RPTime = temp(7) & temp(8) & ":" & temp(9) & temp(10) & ":" & temp(11) & temp(12)
            Call WriteLog(False, "Module Time : " & RPTime)

            ' 모듈 DateTime 변환
            Dim moduleDateTimeStr As String = RPDate & " " & RPTime
            Dim moduleDateTime As DateTime

            Try
                moduleDateTime = DateTime.ParseExact(moduleDateTimeStr, "yy/MM/dd HH:mm:ss", Nothing)
            Catch exParse As Exception
                mErrorComment = "Module RTC Data Parsing NG : " & moduleDateTimeStr
                Throw New ApplicationException(mErrorComment)
            End Try

            ' PC 시간에서 9시간을 뺀 예상 시간 계산
            Dim expectedDateTime As DateTime = pcTimeAtRecv.AddHours(-9)

            Call WriteLog(False, "Expected Time (PC-9H) : " & expectedDateTime.ToString("yy/MM/dd HH:mm:ss"))
            Call WriteLog(False, "誤差：10秒以内")
            ' 시간 차이 계산 및 검증 (10초 오차 허용)
            Dim timeDifference As Double = Math.Abs((moduleDateTime - expectedDateTime).TotalSeconds)

            Call WriteLog(False, "Time Difference : " & timeDifference.ToString("F1") & " seconds")

            If timeDifference > 10.0 Then
                mErrorComment = String.Format("RTC Time Sync NG: モジュール={0}, 期待値(PC-9H)={1}, 時差={2:F1}秒",
                                        moduleDateTime.ToString("yy/MM/dd HH:mm:ss"),
                                        expectedDateTime.ToString("yy/MM/dd HH:mm:ss"),
                                        timeDifference)
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "RTC バックアップテスト OK - 時間同期確認済み")

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "RTC バックアップテスト NG"
            End If
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try
        'ver 0.02 hsj add end ===
        'ver 0.02 hsj del start ===
        ''==============================================================================================================
        '' Function : I31:バックアップテスト（/NBT未実施) (RTCテスト)
        '' Data     : Process name
        '' Return   : Nothing
        '' Remarks  :
        ''==============================================================================================================

        'Dim mSetting As String = ""
        'Dim mDelim As String = ""
        'Dim mReadLine As String = ""

        'Dim temp(12) As String
        'Dim rRes As String = ""
        'Dim rCode As String = ""
        'Dim rCount As Integer = 0
        'Dim rData() As String = {}

        'Dim RPDate As String = ""
        'Dim RPTime As String = ""
        'Dim RPWeek As String = ""

        'Dim RPSetDate As String = ""
        'Dim RPSetTime As String = ""
        'Dim RPSetWeek As String = ""

        'Dim Month As Integer = 0
        'Dim Month_str As String = ""

        'Dim highTime As String = ""
        'Dim valueData As Date
        'Dim valueData1 As Date

        'Dim diffSec As Double
        'Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"
        ''로컬 PC(시스템) 현재 날짜/시간/요일


        'Call WriteLog(False, "")
        'Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        'Call WriteLog(False, "*")
        ''==========================================================================================

        'Try
        '    'RTC 설정 시간 Read 
        '    Dim sr As System.IO.StreamReader = New System.IO.StreamReader(lFile)

        '    Do While sr.Peek() >= 0

        '        mReadLine = sr.ReadLine
        '        If InStr(mReadLine, "RTCSetDate:") <> 0 Then
        '            RPSetDate = Mid(mReadLine, Len("RTCSetDate:") + 1, mReadLine.Length - Len("RTCSetDate:"))
        '        ElseIf InStr(mReadLine, "RTCSetWeek:") <> 0 Then
        '            RPSetWeek = Mid(mReadLine, Len("RTCSetWeek:") + 1, mReadLine.Length - Len("RTCSetWeek:"))
        '        ElseIf InStr(mReadLine, "RTCSetTime:") <> 0 Then
        '            RPSetTime = Mid(mReadLine, Len("RTCSetTime:") + 1, mReadLine.Length - Len("RTCSetTime:"))
        '        End If

        '    Loop
        '    sr.Close()
        '    sr = Nothing

        '    cIA8Comm = New cls101_ComRS232C
        '    cIA8Comm.RS_COMM = Insfrm.SerialPort1
        '    cIA8Comm.PORT = rs232cRZPort
        '    If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
        '        mErrorComment = "RS232SET Read異常"
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    cIA8Comm.SETTING = mSetting
        '    cIA8Comm.TERMINAL_CHAR = mDelim
        '    cIA8Comm.PortSetup()

        '    cIA8Comm.SendCmd = "TS," & "0620" & ",0"
        '    Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
        '    Call cIA8Comm.FuncSendCmd()

        '    mSecWait(2000)

        '    Call cIA8Comm.FuncRecvCmd()
        '    Call RPRecvCmd_1(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

        '    Dim now As DateTime = DateTime.Now
        '    Dim nowTime As DateTime = DateTime.ParseExact(now.ToString("HH:mm:ss"), "HH:mm:ss", Nothing)
        '    Dim nowDateStr As String = now.ToString("yy/MM/dd")

        '    Dim nowWeekday As Integer = CInt(now.DayOfWeek)

        '    If Me.ERR_COMMENT <> "" Then
        '        mErrorComment = "Recv : " & Me.ERR_COMMENT
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

        '    ' 모듈에서 읽은 RTC Date/Time/Week ====================================


        '    For i As Integer = 0 To 12
        '        temp(i) = Hex(Val("&h" & rData(i)))  ' 16진 → 10진 
        '    Next i

        '    ' 월 계산 (temp(2)="0", temp(3)="3" → "03")
        '    Month = CInt(temp(2) & temp(3))
        '    Month_str = Month.ToString("00")

        '    ' 날짜(YY/MM/DD) 구성

        '    RPDate = temp(0) & temp(1) & "/" & Month_str & "/" & temp(4) & temp(5)
        '    Call WriteLog(False, "A8HCA Date : " & RPDate)
        '    Call WriteLog(False, "PC Date : " & nowDateStr)

        '    ' 시간(HH:mm:ss) 구성
        '    RPTime = temp(7) & temp(8) & ":" & temp(9) & temp(10) & ":" & temp(11) & temp(12)
        '    Call WriteLog(False, "A8HCA Time : " & RPTime)
        '    Call WriteLog(False, "PC Time : " & nowTime)

        '    ' 요일(0~6) → string
        '    RPWeek = CStr(Val(temp(6)))
        '    Call WriteLog(False, "A8HCA Week : " & RPWeek)
        '    Call WriteLog(False, "PC Week : " & nowWeekday)

        '    '--------------------------------------------------------------------------
        '    ' 시간 비교 방법
        '    ' (1) 날짜 비교: RPDate(YY/MM/dd) vs. PC의 YY/MM/dd
        '    ' (2) 요일 비교: RPWeek(0~6) vs. PC의 nowWeekday(0~6)
        '    ' (3) 시각(초단위) 비교: 모듈 시간과 PC 시간의 오차가 ±5초 이내인지 확인 
        '    '                        TC 문의 결과 : 시간 오차는 확인후 FK1 기입 예정
        '    '--------------------------------------------------------------------------

        '    ' (1) 날짜가 동일한지?
        '    If RPDate <> nowDateStr Then
        '        mErrorComment = String.Format("RTC Date NG : モジュール={0}, PC={1}", RPDate, nowDateStr)
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    ' (2) 요일(0~6)이 동일한지?
        '    If CInt(RPWeek) <> nowWeekday Then
        '        mErrorComment = String.Format("RTC Week NG : モジュール={0}, PC={1}", RPWeek, nowWeekday)
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    ' (3) 시간 비교(오차 5초)
        '    Dim moduleDateTimeStr As String = RPDate & " " & RPTime
        '    Dim moduleDateTime As DateTime

        '    Try
        '        ' moduleDateTimeStr이 "yy/MM/dd HH:mm:ss" 형태라는 가정
        '        moduleDateTime = DateTime.ParseExact(moduleDateTimeStr, "yy/MM/dd HH:mm:ss", Nothing)
        '    Catch exParse As Exception
        '        mErrorComment = "Module RTC Data Parsing NG : " & moduleDateTimeStr
        '        Throw New ApplicationException(mErrorComment)
        '    End Try

        '    ' PC 현재 시간과의 초 차이
        '    'Dim diffSec As Double = Math.Abs((moduleDateTime - now).TotalSeconds)
        '    diffSec = Math.Abs((moduleDateTime - now).TotalSeconds)
        '    If diffSec > 5.0 Then
        '        mErrorComment = String.Format("RTC Time NG: モジュール={0}, PC={1}, 違い={2:F1}秒", RPTime, now.ToString("HH:mm:ss"), diffSec)
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    valueData = CDate(RPSetTime)
        '    valueData1 = DateAdd(DateInterval.Second, 60, valueData)
        '    highTime = valueData1.ToString("HH:mm:ss")
        '    'Date Time judge ================================================
        '    If RPTime > highTime Then
        '        mTestSpec = "D:" & RPSetDate & " W:" & RPSetWeek & " T:" & RPSetTime
        '        mTestMeasure = "D:" & RPDate & " W:" & RPWeek & " T:" & RPTime

        '        mErrorComment = "RTC Comparison NG : 60s Time Over"
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    Call WriteLog(False, "RTC バックアップテスト OK")

        'Catch ex As Exception
        '    If String.IsNullOrEmpty(mErrorComment) Then
        '        mErrorComment = "RTC バックアップテスト NG"
        '    End If
        '    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
        '    mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    If cIA8Comm IsNot Nothing Then
        '        cIA8Comm.PortClose()
        '        cIA8Comm = Nothing
        '    End If
        'End Try
        'ver 0.02 hsj del end ===
    End Sub


    Public Sub I40(ByVal iName As String)
        '==============================================================================================================
        ' Function : I40 LEDテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rData() As String = {""}
        Dim retValue As String = ""
        Dim rCount As String = ""

        Try
            '======================================
            ' A8 Communication 설정
            '======================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If

            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            Dim ledTests() As LedTestInfo = {
        New LedTestInfo With {
            .testName = "XSFAIL LED",
            .ppt1 = "LED_XSFAIL_1.ppt",
            .ppt2 = "LED_XSFAIL_2.ppt",
            .sendCmd = "TS,0A00,0",
            .checkTitle = "XSFAILの確認",
            .checkTarget = "XSFAIL LED",
            .checkColor = "YELLOW",
            .offText = "OFF",
            .onText1 = "OFF",
            .onText2 = "OFF",
            .errMsg = "XSFAIL LED 異常"
        },
        New LedTestInfo With {
            .testName = "RUN/ALM/ERR LED",
            .ppt1 = "LED_MAIN_1.ppt",
            .ppt2 = "LED_MAIN_2.ppt",
            .sendCmd = "TS,0D00,0",
            .checkTitle = "Main LEDの確認",
            .checkTarget = "Main LED",
            .checkColor = "YELLOW",
            .offText = "OFF",
            .onText1 = "OFF",
            .onText2 = "OFF",
            .errMsg = "Main LED 異常"
        },
        New LedTestInfo With {
            .testName = "User, SD, USB LED",
            .ppt1 = "LED_USER_1.ppt",
            .ppt2 = "LED_USER_2.ppt",
            .sendCmd = "TS,0D20,0",
            .checkTitle = "User LEDの確認",
            .checkTarget = "User LED",
            .checkColor = "YELLOW",
            .offText = "OFF",
            .onText1 = "OFF",
            .onText2 = "OFF",
            .errMsg = "User LED 異常"
        },
        New LedTestInfo With {
            .testName = "7 Segment LED",
            .ppt1 = "LED_7SEG_1.ppt",
            .ppt2 = "LED_7SEG_2.ppt",
            .sendCmd = "TS,0D30,0",
            .checkTitle = "7 Segment LEDの確認",
            .checkTarget = "7 Segment LED",
            .checkColor = "YELLOW",
            .offText = "OFF",
            .onText1 = "OFF",
            .onText2 = "OFF",
            .errMsg = "7 Segment LED 異常"}
        }

            ' RDY 확인 작업지시 ppt
            Call WriteLog(False, "RDY LED Test開始")
            ShowMessagePopupAndCheck(mAppPath,
                                 "RDY.ppt",
                                 "RDY LEDの確認",
                                 "RDY LED",
                                 "YELLOW", "OFF", "OFF", "OFF",
                                 "RDY LED異常")
            Call WriteLog(False, "RDY LED Test OK")
            For Each t In ledTests
                ' 테스트 시작 로그
                Call WriteLog(False, "")
                Call WriteLog(False, t.testName & "テスト開始")

                ' PPT 1
                ShowMessagePopupAndCheck(mAppPath,
                                     t.ppt1,
                                     t.checkTitle,
                                     t.checkTarget,
                                     t.checkColor,
                                     t.offText,
                                     t.onText1,
                                     t.onText2,
                                     t.errMsg)
                mSecWait(500)

                ' SendCmd
                cIA8Comm.SendCmd = t.sendCmd
                Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
                Call cIA8Comm.FuncSendCmd()

                mSecWait(2000)

                ' Recv
                Call cIA8Comm.FuncRecvCmd()
                Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

                ' Recv Data 파싱
                Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
                If Me.ERR_COMMENT <> "" Then
                    mTestMeasure = cIA8Comm.RECEIVE_DATA

                    mErrorComment = "FPGAテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                    Throw New ApplicationException(mErrorComment)
                End If

                Do
                    ' Retry PPT 추가
                    ShowMessagePopupAndCheck_RETRY(mAppPath,
                                     t.ppt2,
                                     t.checkTitle,
                                     t.checkTarget,
                                     t.checkColor,
                                     t.offText,
                                     t.onText1,
                                     t.onText2,
                                     t.errMsg,
                                     retValue)

                    Select Case retValue
                        Case "OK"
                            Exit Do

                        Case "NG"
                            Throw New ApplicationException

                        Case "RETRY"
                            ' LED TEST 커맨드 재송신
                            ' SendCmd 
                            cIA8Comm.SendCmd = t.sendCmd
                            WriteLog(False, "Send : " & cIA8Comm.SendCmd)
                            cIA8Comm.FuncSendCmd()

                            mSecWait(500)

                            ' RecvCmd
                            cIA8Comm.FuncRecvCmd()
                            WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

                            ' Recv Data 파싱
                            Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
                            If Me.ERR_COMMENT <> "" Then
                                mTestMeasure = cIA8Comm.RECEIVE_DATA
                                mErrorComment = "FPGAテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                                Throw New ApplicationException(mErrorComment)
                            End If

                            ' PPT2 반복 => Do 루프 계속
                        Case Else
                            Throw New ApplicationException("LED TEST: " & retValue)
                    End Select
                Loop
                Call WriteLog(False, t.testName & "テストOK" & vbCrLf)
            Next

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "LEDテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "LEDテストOK")
    End Sub

    Public Sub I10(ByVal iName As String)
        '==============================================================================================================
        ' Function : I10:電圧確認テスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim FSR As Double = 6.144     ' ±6.144 V
        Dim ADC_max As Double = 32752 ' ADC 최대값

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            MeasureVoltage_Sysmon(cIA8Comm, "3_3V_電源電圧", "0910", "9111", "8444")
            mSecWait(300)
            MeasureVoltage_Sysmon(cIA8Comm, "1_8V_電源電圧1", "0940", "A222", "8E82")
            mSecWait(300)
            MeasureVoltage_Sysmon(cIA8Comm, "1_8V_電源電圧2", "0950", "5111", "4741")
            mSecWait(300)
            MeasureVoltage_Sysmon(cIA8Comm, "1_2V_電源電圧", "0960", "6E14", "5D03")
            mSecWait(300)
            MeasureVoltage_Sysmon(cIA8Comm, "0_85V_電源電圧1", "0980", "5037", "3F26")
            mSecWait(300)
            MeasureVoltage_Sysmon(cIA8Comm, "0_85V_電源電圧2", "0990", "4F5C", "4000")
            mSecWait(300)

            'ver 0.02 HSJ del Start===
            'MeasureVoltage(cIA8Comm, "5_2V_電源電圧", "0900", "7072", "681E", FSR, ADC_max)  'High=7287, Low=5DB4
            'mSecWait(300)
            'MeasureVoltage(cIA8Comm, "3_3V_SD_電源電圧", "0920", "46CC", "408D", FSR, ADC_max)
            'mSecWait(300)
            'MeasureVoltage(cIA8Comm, "2_5V_電源電圧", "0930", "3799", "3085", FSR, ADC_max)
            'mSecWait(300)
            'MeasureVoltage(cIA8Comm, "1_05V_電源電圧", "0970", "18FD", "12BE", FSR, ADC_max)
            'mSecWait(300)
            'MeasureVoltage(cIA8Comm, "0_6V_電源電圧", "09A0", "E29", "AD4", FSR, ADC_max)
            'ver 0.02 HSJ del End===

            'ver 0.02 HSJ add Start===
            MeasureVoltage(cIA8Comm, "5_2V_電源電圧", "0900", "7080", "682A")
            mSecWait(300)
            MeasureVoltage(cIA8Comm, "3_3V_SD_電源電圧", "0920", "46D5", "42AA")
            mSecWait(300)
            MeasureVoltage(cIA8Comm, "2_5V_電源電圧", "0930", "37A0", "308A")
            mSecWait(300)
            MeasureVoltage(cIA8Comm, "1_05V_電源電圧", "0970", "1900", "12C0")
            mSecWait(300)
            MeasureVoltage(cIA8Comm, "0_6V_電源電圧", "09A0", "E2A", "AD5")
            'ver 0.02 HSJ add End===

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "電圧確認テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "")
        Call WriteLog(False, "電圧確認テスト OK")
    End Sub


    Public Sub I04(ByVal iName As String)
        '==============================================================================================================
        ' Function : I04:FPGAテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}
        Dim Model As String = ModelRead()

        Dim rDTkey As String = ""
        Dim tbName As String
        Dim fName() As String = {""}
        Dim strSetFPGA As String = ""
        Dim strSetID As String = ""

        '==============================================================================================================
        'FPGA Versionの読み 
        Dim dbName As String = mAppPath & "520 Product Firmware\A8HCA\FPGAVer.mdb"
        Dim rField As String = "Ver"
        Dim sSQL As String = ""
        Dim sData() As String = Nothing
        Dim cADO As cls010_MDB_Operation = Nothing
        'Dim tableList As String() = {"tblFPGAVer", "tblmouduleID"} 'ver 0.02 hsj del
        Dim tableList As String() = {"tblmouduleID"} 'ver 0.02 hsj add

        For Each tbName In tableList
            Try
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName

                sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model LIKE '" & Model & "'"
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(fName, sData)
                strSetID = sData(0) 'ver 0.02 hsj add

                'ver 0.02 hsj del start===
                'Select Case tbName
                '    Case "tblFPGAVer"
                '        strSetFPGA = sData(0)
                '    Case "tblmouduleID"
                '        strSetID = sData(0)
                'End Select
                'ver 0.02 hsj del end===
                Call WriteLog(False, "モジュールID : " & strSetID) 'ver 0.02 hsj add
            Catch ex As Exception
                cADO = Nothing
                mErrorComment = "DB読み込み失敗"
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
        Next


        Try
            'ver 0.02 hsj del start===
            'Call WriteLog(False, "FPGA Revision テスト開始")
            ''FPGA Revision test ***************************************************

            '' Send
            'cIA8Comm.SendCmd = $"TS,0D80,1,{strSetFPGA}"
            'Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            'Call cIA8Comm.FuncSendCmd()

            '' Recv
            'Call cIA8Comm.FuncRecvCmd()
            'Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

            '' Recv Data 파싱
            'Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
            'If Me.ERR_COMMENT <> "" Then
            '    ' ok면 0이 나옴
            '    mTestSpec = "OK,1D80,0"
            '    mTestMeasure = cIA8Comm.RECEIVE_DATA
            '    mErrorComment = "FPGAテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
            '    Throw New ApplicationException(mErrorComment)
            'End If
            'Call WriteLog(False, "FPGA Revision テスト OK" & vbCrLf)
            'ver 0.02 hsj del end===

            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'FPGA ModuleID test ***************************************************
            Call WriteLog(False, "FPGA ModuleID テスト開始")
            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            ' Send
            cIA8Comm.SendCmd = $"TS,0DB0,1,{strSetID}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            ' Recv
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

            ' Recv Data 파싱
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1DB0,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "FPGAテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "FPGA ModuleID テスト OK" & vbCrLf)


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "FPGAテスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "FPGAテスト OK")

    End Sub

    Public Sub I05(ByVal iName As String)
        '==============================================================================================================
        ' Function : I05:電源モジュールテスト(XDFAIL)
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}


        '==============================================================================================================

        Try


            'XDFAILテスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "XDFAILテスト開始")

            'XDFAIL 쇼트시 모듈 재기동 확인 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "XDFAIL_TEST.ppt",
                     "XDFAILテスト開始",
                     "XDFAILテスト",
                     "YELLOW", "OFF", "ON", "ON",
                     "XDFAILテスト　異常")

            Call WriteLog(False, "XDFAIL Short(JIG Board)") 'ver 0.02 hsj add

            'XDFAIL 쇼트 ---------------------------
            PLC_Y("01" & "BWRY" & "302" & ",001," & "1")
            mSecWait(1000)
            PLC_Y("01" & "BWRY" & "302" & ",001," & "0")
            Call WriteLog(False, "XDFAIL Open(JIG Board)") 'ver 0.02 hsj add
            WaitStart(30)

            'XDFAIL 쇼트시 모듈 재기동 확인 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "XDFAIL_TEST_1.ppt",
                     "XDFAILテスト開始",
                     "XDFAILテスト",
                     "YELLOW", "OFF", "ON", "ON",
                     "XDFAILテスト　異常")


            Call WriteLog(False, "XDFAILテスト OK")

            'XPFAILテスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "XPFAILテスト開始")


            mSecWait(1000)
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'Send command ---------------------------
            Call WriteLog(False, "XPFAIL Open(JIG Board) HIGH 確認") 'ver 0.02 hsj add
            cIA8Comm.SendCmd = "TS,0A10,1,0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1A10,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "電源モジュールテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            mSecWait(2000)
            'XPFAIL 쇼트 ---------------------------
            Call WriteLog(False, "XPFAIL Short(JIG Board)") 'ver 0.02 hsj add
            PLC_Y("01" & "BWRY" & "303" & ",001," & "1")
            WaitStart(5)
            'Send command ---------------------------
            Call WriteLog(False, "XPFAIL Short(JIG Board) LOW 確認") 'ver 0.02 hsj add
            cIA8Comm.SendCmd = "TS,0A10,1,1"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            'WaitStart(90)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1A10,1"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "電源モジュールテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            'XPFAIL 쇼트 ---------------------------
            Call WriteLog(False, "XPFAIL Open (JIG Board)") 'ver 0.02 hsj add
            PLC_Y("01" & "BWRY" & "303" & ",001," & "0")

            Call WriteLog(False, "XPFAILテストOK")

            '電力モニタ信号テスト（I2C）***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "電力モニタ信号テスト（I2C）開始")

            'I2C 24V 측정
            Call MeasureVoltage_I2C(cIA8Comm)

            Call WriteLog(False, "電力モニタ信号テスト（I2C）OK")


        Catch ex As Exception
            PLC_Y("01" & "BWRY" & "302" & ",001," & "0")
            PLC_Y("01" & "BWRY" & "303" & ",001," & "0")
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "電源モジュールテスト NG"
            End If

            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
            PLC_Y("01" & "BWRY" & "302" & ",001," & "0")
            PLC_Y("01" & "BWRY" & "303" & ",001," & "0")
        End Try

        Call WriteLog(False, "電源モジュールテストOK")

    End Sub
    Public Sub I06(ByVal iName As String)
        '==============================================================================================================
        ' Function : I06:電源モジュールテスト(XPFAIL)
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}


        '==============================================================================================================

        Try


            'XDFAILテスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "XDFAILテスト開始")

            'XDFAIL 쇼트시 모듈 재기동 확인 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "XDFAIL_TEST.ppt",
                     "XDFAILテスト開始",
                     "XDFAILテスト",
                     "YELLOW", "OFF", "ON", "ON",
                     "XDFAILテスト　異常")

            'XDFAIL 쇼트 ---------------------------
            PLC_Y("01" & "BWRY" & "302" & ",001," & "1")
            mSecWait(1000)
            PLC_Y("01" & "BWRY" & "302" & ",001," & "0")

            WaitStart(26)

            'XDFAIL 쇼트시 모듈 재기동 확인 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "XDFAIL_TEST_1.ppt",
                     "XDFAILテスト開始",
                     "XDFAILテスト",
                     "YELLOW", "OFF", "ON", "ON",
                     "XDFAILテスト　異常")


            Call WriteLog(False, "XDFAILテスト OK")

            'XPFAILテスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "XPFAILテスト開始")


            mSecWait(1000)
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0A10,1,0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1A10,0"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "電源モジュールテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            mSecWait(2000)
            'XPFAIL 쇼트 ---------------------------
            PLC_Y("01" & "BWRY" & "303" & ",001," & "1")
            mSecWait(5000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0A10,1,1"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(2000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = "OK,1A10,1"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "電源モジュールテスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            'XPFAIL 쇼트 ---------------------------
            PLC_Y("01" & "BWRY" & "303" & ",001," & "0")

            Call WriteLog(False, "XPFAILテストOK")

            '電力モニタ信号テスト（I2C）***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "電力モニタ信号テスト（I2C）開始")

            'I2C 24V 측정
            Call MeasureVoltage_I2C(cIA8Comm)

            Call WriteLog(False, "電力モニタ信号テスト（I2C）OK")


        Catch ex As Exception
            PLC_Y("01" & "BWRY" & "302" & ",001," & "0")
            PLC_Y("01" & "BWRY" & "303" & ",001," & "0")
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "電源モジュールテスト NG"
            End If

            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
            PLC_Y("01" & "BWRY" & "302" & ",001," & "0")
            PLC_Y("01" & "BWRY" & "303" & ",001," & "0")
        End Try

        Call WriteLog(False, "電源モジュールテストOK")

    End Sub

    Public Sub IA1(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA1:eMMC Download
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'eMMC Download ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "eMMC Download開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0F30" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'DL 대기
            Call WriteLog(False, "Download 180s Wait")
            Call WaitStart(180)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "eMMC Download" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "eMMC Download NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "eMMC Download OK")

    End Sub

    Public Sub IB3(ByVal iName As String)
        '==============================================================================================================
        ' Function : IB3:Ubuntu Patch 
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'eMMC Download ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "Ubuntu Patch開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0F40" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'DL 대기
            'Call WriteLog(False, "Download 420s Wait")
            Call WaitStart(60)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "Ubuntu Patch" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "Ubuntu Patch NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "Ubuntu Patch OK")

    End Sub
    'ver 0.02 HSJ add start ===
    Public Sub IB4(ByVal iName As String)
        '==============================================================================================================
        ' Function : IB4:Q-SPI Download
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'Q-SPI Download ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "Q-SPI Download開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0F70" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'DL 대기
            Call WriteLog(False, "Download 180s Wait")
            Call WaitStart(180)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "Q-SPI Download" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "Q-SPI Download NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "Q-SPI Download OK")

    End Sub
    'ver 0.02 HSJ add End ===

    'Ver 0.01 HSJ Add Start ====
    Public Sub IC1(ByVal iName As String)
        '==============================================================================================================
        ' Function : IC1:ランタイムライセンス書込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}
        Dim Model As String = ModelRead()

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            Call WriteLog(False, "ランタイムライセンス設定 開始")

            'ランタイムライセンス初期化 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "ランタイムライセンス初期化")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS," & "0FA0" & ",0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "ランタイムライセンス初期化" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If

            '/STN, /ADV 사양만 등록실시
            If InStr(Model, "/STN") > 0 Or InStr(Model, "/ADV") > 0 Then

                'ランタイムライセンス登録 ***************************************************
                Call WriteLog(False, "")
                Call WriteLog(False, "ランタイムライセンス登録")


                'Send command ---------------------------
                If InStr(Model, "/STN") > 0 Then
                    cIA8Comm.SendCmd = "TS,0FB0,6,EtherCatMain,1,Motion,8,Pid,16"
                ElseIf InStr(Model, "/ADV") > 0 Then
                    cIA8Comm.SendCmd = "TS,0FB0,6,EtherCatMain,1,Motion,32,Pid,64"
                End If

                Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
                Call cIA8Comm.FuncSendCmd()


                'Recv Command  --------------------------
                Call cIA8Comm.FuncRecvCmd()
                Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
                Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

                'Recv가 NG면 ERR 처리
                If Me.ERR_COMMENT <> "" Then
                    mTestMeasure = cIA8Comm.RECEIVE_DATA
                    Call WriteLog(False, mTestMeasure)
                    mErrorComment = "ランタイムライセンス登録" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                    Throw New ApplicationException(mErrorComment)
                End If
            End If 'ver 0.02 HSJ add
            'ランタイムライセンス有効化 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "ランタイムライセンス有効化")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0FC0,0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(5000)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "ランタイムライセンス有効化" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            'End If 'Ver 0.02 HSJ del
        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "ランタイムライセンス有効化 異常"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "ランタイムライセンス設定 OK")
    End Sub

    Public Sub IC2(ByVal iName As String)
        '==============================================================================================================
        ' Function : IC2:TPMテスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()


            'TPM信号線テスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "TPM信号線テスト 開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0F90,0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "TPM信号線テスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If

            'TPM初期化 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "TPM初期化 開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0FD0,0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            WaitStart(15)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "TPM初期化" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "TPMテスト 異常"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "TPMテスト OK")

    End Sub

    Public Sub IC3(ByVal iName As String)
        '==============================================================================================================
        ' Function : IC3:eFUSE設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        '==============================================================================================================

        Try
            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()


            'eFUSE設定 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "eFUSE設定 開始")

            'Send command ---------------------------
            cIA8Comm.SendCmd = "TS,0F80,0"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            'DL 대기 
            Call WriteLog(False, "eFUSE設定 Wait")
            Call WaitStart(3)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "eFUSE設定" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "eFUSE設定 異常"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "eFUSE設定 OK")

    End Sub
    'Ver 0.01 HSJ Add End ====

    Public Sub IA3(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA3:Q-SPI SUM値算出テスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        ' A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}
        Dim dbName As String = mAppPath & "520 Product Firmware\A8HCA\FirmName.mdb"
        Dim Model As String = ModelRead()
        Dim tbName As String = "tblQSPI"
        Dim fName() As String = {""}
        Dim sSQL As String = ""
        Dim sData() As String = Nothing
        Dim cADO As cls010_MDB_Operation = Nothing

        Dim Ops(3) As String
        'Dim rField As String = "mtd0, mtd1, mtd2, mtd3" 'VER 0.04 HSJ DEL
        Dim rField As String = "mtd0, mtd1" 'VER 0.04 HSJ ADD
        Dim Op1 As String = ""
        Dim Op2 As String = ""
        'Dim Op3 As String = "" 'VER 0.04 HSJ DEL
        'Dim Op4 As String = "" 'VER 0.04 HSJ DEL
        '==============================================================================================================

        Try
            Try
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName

                sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model LIKE '" & Model & "'"
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(fName, sData)

                'If sData IsNot Nothing AndAlso sData.Length >= 4 Then'VER 0.04 HSJ DEL
                If sData IsNot Nothing AndAlso sData.Length >= 2 Then 'VER 0.04 HSJ ADD
                    Op1 = sData(0)
                    Op2 = sData(1)
                    'Op3 = sData(2) 'VER 0.04 HSJ DEL
                    'Op4 = sData(3) 'VER 0.04 HSJ DEL
                End If

            Catch ex As Exception
                cADO = Nothing
                mErrorComment = "DB読み込み失敗"
                Throw New ApplicationException(mErrorComment)
            End Try

            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'Q-SPI SUM値算出テスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "Q-SPI SUM値算出テスト開始" & vbCrLf)


            'Send command ---------------------------
            'cIA8Comm.SendCmd = $"TS,0110,4,{Op1},{Op2},{Op3},{Op4}" 'VER 0.04 HSJ DEL
            cIA8Comm.SendCmd = $"TS,0110,2,{Op1},{Op2}" 'VER 0.04 HSJ ADD
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            WaitStart(10)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                Call WriteLog(False, mTestMeasure)
                mErrorComment = "Q-SPI SUM値算出テスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "Q-SPI SUM値算出テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "Q-SPI SUM値算出テスト OK")

    End Sub

    Public Sub IA2(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA2:eMMC SUM値算出テスト
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        ' A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}
        Dim dbName As String = mAppPath & "520 Product Firmware\A8HCA\FirmName.mdb"
        Dim Model As String = ModelRead()
        Dim tbName As String = "tblEMMC"
        Dim fName() As String = {""}
        Dim sSQL As String = ""
        Dim sData() As String = Nothing
        Dim cADO As cls010_MDB_Operation = Nothing

        Dim Ops(3) As String
        Dim rField As String = "mmcblk0p1, mmcblk0p2, mmcblk0p3, mmcblk0p4"
        Dim fieldList() As String = rField.Split(","c)
        Dim Op1 As String
        Dim Op2 As String
        Dim Op3 As String
        Dim Op4 As String
        '==============================================================================================================

        Try
            Try
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName

                sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model LIKE '" & Model & "'"
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(fName, sData)

                Op1 = sData(0)
                Op2 = sData(1)
                Op3 = sData(2)
                Op4 = sData(3)

            Catch ex As Exception
                cADO = Nothing
                mErrorComment = "DB読み込み失敗"
                Throw New ApplicationException(mErrorComment)
            End Try



            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'eMMC SUM値算出テスト ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "eMMC SUM値算出テスト開始")


            'Send command ---------------------------
            cIA8Comm.SendCmd = $"TS,0210,4,{Op1},{Op2},{Op3},{Op4}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()

            WaitStart(100)

            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RPRecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "eMMC SUM値算出テスト" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If


        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "eMMC SUM値算出テスト NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "eMMC SUM値算出テスト OK")

    End Sub

    Public Sub IA4(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA4:Mac Address Port0
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Dim mcAddress As String
        Dim Serial As String = SerialRead()
        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()
        Dim FolderPath As String = "A8HCA"

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'PLASMAｼﾘｱﾙ番号でMACｱﾄﾞﾚｽ登録があるかどうかﾁｪｯｸ
        Dim dbName As String = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "PastMAC"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblPastMAC"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Serial ='" & Serial & "' AND LAN_Port='" & "LAN0" & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT <> 0 Then
                Call WriteLog(False, "MacAddress.mdbに" & Serial & "で登録したAddress再使用")
                Call WriteLog(False, "MacAddress1 : " & sData(0))
                Exit Sub
            End If

        Catch ex As Exception
            mErrorComment = "MacAddress.mdbのRead　異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        Try
            'Get MacAddress -----------------------------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            Call cMac.IssueMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If
            mcAddress = cMac.MacAddress

            cMac = Nothing

            Call WriteLog(False, "MacAddress1 : " & mcAddress)


            'Next Mac Address ==============================================
            Call WriteLog(False, "Next MAC ADDRESS ･･･")

            'Issue Next MacAddress -------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            cMac.Serial = Serial
            cMac.LAN_Port = "LAN0"

            Call cMac.IssueNextMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If

            cMac = Nothing

            'LOG Disp -----------------------------------
            Call WriteLog(False, "Next MAC ADDRESS OK!")

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMac = Nothing
        End Try

    End Sub

    Public Sub IA5(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA5:Mac Address Port1
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Dim mcAddress As String
        Dim Serial As String = SerialRead()
        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()
        Dim FolderPath As String = "A8HCA"

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'PLASMAｼﾘｱﾙ番号でMACｱﾄﾞﾚｽ登録があるかどうかﾁｪｯｸ
        Dim dbName As String = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "PastMAC"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblPastMAC"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Serial ='" & Serial & "' AND LAN_Port='" & "LAN1" & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT <> 0 Then
                Call WriteLog(False, "MacAddress.mdbに" & Serial & "で登録したAddress再使用")
                Call WriteLog(False, "MacAddress2 : " & sData(0))
                Exit Sub
            End If

        Catch ex As Exception
            mErrorComment = "MacAddress.mdbのRead　異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        Try
            'Get MacAddress -----------------------------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            Call cMac.IssueMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If
            mcAddress = cMac.MacAddress

            cMac = Nothing

            Call WriteLog(False, "MacAddress2 : " & mcAddress)


            'Next Mac Address ==============================================
            Call WriteLog(False, "Next MAC ADDRESS ･･･")

            'Issue Next MacAddress -------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            cMac.Serial = Serial
            cMac.LAN_Port = "LAN1"

            Call cMac.IssueNextMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If

            cMac = Nothing

            'LOG Disp -----------------------------------
            Call WriteLog(False, "Next MAC ADDRESS OK!")

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMac = Nothing
        End Try

    End Sub
    Public Sub IA6(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA6:Mac Address Port2
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Dim mcAddress As String
        Dim Serial As String = SerialRead()
        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()
        Dim FolderPath As String = "A8HCA"

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'PLASMAｼﾘｱﾙ番号でMACｱﾄﾞﾚｽ登録があるかどうかﾁｪｯｸ
        Dim dbName As String = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "PastMAC"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblPastMAC"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Serial ='" & Serial & "' AND LAN_Port='" & "LAN2" & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT <> 0 Then
                Call WriteLog(False, "MacAddress.mdbに" & Serial & "で登録したAddress再使用")
                Call WriteLog(False, "MacAddress3 : " & sData(0))
                Exit Sub
            End If

        Catch ex As Exception
            mErrorComment = "MacAddress.mdbのRead　異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        Try
            'Get MacAddress -----------------------------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            Call cMac.IssueMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If
            mcAddress = cMac.MacAddress

            cMac = Nothing

            Call WriteLog(False, "MacAddress3 : " & mcAddress)


            'Next Mac Address ==============================================
            Call WriteLog(False, "Next MAC ADDRESS ･･･")

            'Issue Next MacAddress -------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            cMac.Serial = Serial
            cMac.LAN_Port = "LAN2"

            Call cMac.IssueNextMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If

            cMac = Nothing

            'LOG Disp -----------------------------------
            Call WriteLog(False, "Next MAC ADDRESS OK!")

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMac = Nothing
        End Try

    End Sub
    Public Sub IA7(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA7:Mac Address Port3
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Dim mcAddress As String
        Dim Serial As String = SerialRead()
        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()
        Dim FolderPath As String = "A8HCA"

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'PLASMAｼﾘｱﾙ番号でMACｱﾄﾞﾚｽ登録があるかどうかﾁｪｯｸ
        Dim dbName As String = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "PastMAC"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblPastMAC"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Serial ='" & Serial & "' AND LAN_Port='" & "LAN3" & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT <> 0 Then
                Call WriteLog(False, "MacAddress.mdbに" & Serial & "で登録したAddress再使用")
                Call WriteLog(False, "MacAddress4 : " & sData(0))
                Exit Sub
            End If

        Catch ex As Exception
            mErrorComment = "MacAddress.mdbのRead　異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        Try
            'Get MacAddress -----------------------------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            Call cMac.IssueMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If
            mcAddress = cMac.MacAddress

            cMac = Nothing

            Call WriteLog(False, "MacAddress4 : " & mcAddress)


            'Next Mac Address ==============================================
            Call WriteLog(False, "Next MAC ADDRESS ･･･")

            'Issue Next MacAddress -------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            cMac.Serial = Serial
            cMac.LAN_Port = "LAN3"

            Call cMac.IssueNextMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If

            cMac = Nothing

            'LOG Disp -----------------------------------
            Call WriteLog(False, "Next MAC ADDRESS OK!")

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMac = Nothing
        End Try

    End Sub
    Public Sub IA8(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA8:Mac Address Port4(TSN END POINT)
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Dim mcAddress As String
        Dim Serial As String = SerialRead()
        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()
        Dim FolderPath As String = "A8HCA"

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'PLASMAｼﾘｱﾙ番号でMACｱﾄﾞﾚｽ登録があるかどうかﾁｪｯｸ
        Dim dbName As String = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "PastMAC"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblPastMAC"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Serial ='" & Serial & "' AND LAN_Port='" & "LAN4" & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT <> 0 Then
                Call WriteLog(False, "MacAddress.mdbに" & Serial & "で登録したAddress再使用")
                Call WriteLog(False, "MacAddress5 : " & sData(0))
                Exit Sub
            End If

        Catch ex As Exception
            mErrorComment = "MacAddress.mdbのRead　異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        Try
            'Get MacAddress -----------------------------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            Call cMac.IssueMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If
            mcAddress = cMac.MacAddress

            cMac = Nothing

            Call WriteLog(False, "MacAddress5 : " & mcAddress)


            'Next Mac Address ==============================================
            Call WriteLog(False, "Next MAC ADDRESS ･･･")

            'Issue Next MacAddress -------
            cMac = New cls031_MacAddress
            cMac.DBFileName = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"

            cMac.Serial = Serial
            cMac.LAN_Port = "LAN4"

            Call cMac.IssueNextMacAddress()
            If cMac.ERR_MSG <> "" Then
                mErrorComment = cMac.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If

            cMac = Nothing

            'LOG Disp -----------------------------------
            Call WriteLog(False, "Next MAC ADDRESS OK!")

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMac = Nothing
        End Try

    End Sub
    Public Sub IB1(ByVal iName As String)
        '==============================================================================================================
        ' Function : IB1:Flash Memoryへの書込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : Macaddress, Style Number, Release Number, Product name
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        ' A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As String = ""
        Dim rData() As String = {""}
        Dim dbName As String = mAppPath & "\500 Model Information\ModelInformation.mdb"
        Dim Model As String = ModelRead()
        Dim Serial As String = SerialRead()
        Dim tbName As String = "tblModelInformation"
        Dim fName() As String = {""}
        Dim sSQL As String = ""
        Dim sData() As String = Nothing
        Dim cADO As cls010_MDB_Operation = Nothing

        Dim rField As String = "StyleNum, ReleaseNum"
        Dim fieldList() As String = rField.Split(","c)
        Dim styleNum As String
        Dim releaseNum As String

        '==============================================================================================================

        Try
            ' Stvle Number, Release Number를 Read
            Try
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName

                sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model LIKE '" & Model & "'"
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(fName, sData)
                styleNum = sData(0)
                releaseNum = sData(1)
            Catch ex As Exception
                cADO = Nothing
                mErrorComment = "DB読み込み失敗"
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'A8 Communication =================================
            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort

            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read 異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            'モジュール形名設定 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "モジュール形名書込み開始")
            mSecWait(1000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = $"TS,0130,2,1,{Model}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "モジュール形名書込み" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            ' Write값과 Read값이 다르면 ERR 처리
            If Model = rCount Then
                Call WriteLog(False, $"書き込み値={Model}, 読み込み値={rCount})")
            Else
                Call WriteLog(False, $"書き込み値={Model}, 読み込み値={rCount})")

                mTestSpec = $"OK,1130,{Model}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = $"モジュール形名比較異常 : 書き込み値={Model}, 読み込み値={rCount}"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "モジュール形名書込み OK" & vbCrLf)


            'シリアル番号設定 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "シリアル番号設定開始")
            mSecWait(1000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = $"TS,0130,2,2,{Serial}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "シリアル番号書込み" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            ' Write값과 Read값이 다르면 ERR 처리
            If Serial = rCount Then
                Call WriteLog(False, $"書き込み値={Serial}, 読み込み値={rCount})")
            Else
                Call WriteLog(False, $"書き込み値={Serial}, 読み込み値={rCount})")

                mTestSpec = $"OK,1130,{Serial}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = $"シリアル番号比較異常 : 書き込み値={Serial}, 読み込み値={rCount}"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "シリアル番号書込み OK" & vbCrLf)




            'リリースナンバ設定 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "リリースナンバ書込み開始")
            mSecWait(1000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = $"TS,0130,2,3,{releaseNum}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = $"OK,1130,{releaseNum}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "リリースナンバ書込み" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            ' Write값과 Read값이 다르면 ERR 처리
            If releaseNum = rCount Then
                Call WriteLog(False, $"書き込み値={releaseNum}, 読み込み値={rCount})")
            Else
                Call WriteLog(False, $"書き込み値={releaseNum}, 読み込み値={rCount})")

                mTestSpec = $"OK,1130,{releaseNum}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = $"リリースナンバ比較異常 : 書き込み値={releaseNum}, 読み込み値={rCount}"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "リリースナンバ書込み OK" & vbCrLf)

            'カレントナンバ設定 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "カレントナンバ書込み開始")
            mSecWait(1000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = $"TS,0130,2,4,{releaseNum}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = $"OK,1130,{releaseNum}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "カレントナンバ書込み" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            ' Write값과 Read값이 다르면 ERR 처리
            If releaseNum = rCount Then
                Call WriteLog(False, $"書き込み値={releaseNum}, 読み込み値={rCount})")
            Else
                Call WriteLog(False, $"書き込み値={releaseNum}, 読み込み値={rCount})")

                mTestSpec = $"OK,1130,{releaseNum}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = $"カレントナンバ比較異常 : 書き込み値={releaseNum}, 読み込み値={rCount}"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "カレントナンバ書込み OK" & vbCrLf)

            'スタイルナンバ・リリースナンバ設定 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "スタイルナンバ書込み開始")
            mSecWait(1000)
            'Send command ---------------------------
            cIA8Comm.SendCmd = $"TS,0130,2,5,{styleNum}"
            Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)
            Call cIA8Comm.FuncSendCmd()
            mSecWait(1000)
            'Recv Command  --------------------------
            Call cIA8Comm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)
            Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
            'Recv가 NG면 ERR 처리
            If Me.ERR_COMMENT <> "" Then
                mTestSpec = $"OK,1130,{styleNum}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = "スタイルナンバ書込み" & " : " & CStr(cIA8Comm.RECEIVE_DATA) & "  異常"
                Throw New ApplicationException(mErrorComment)
            End If
            ' Write값과 Read값이 다르면 ERR 처리
            If styleNum = rCount Then
                Call WriteLog(False, $"書き込み値={styleNum}, 読み込み値={rCount})")
            Else
                Call WriteLog(False, $"書き込み値={styleNum}, 読み込み値={rCount})")

                mTestSpec = $"OK,1130,{styleNum}"
                mTestMeasure = cIA8Comm.RECEIVE_DATA
                mErrorComment = $"スタイルナンバ比較異常 : 書き込み値={styleNum}, 読み込み値={rCount}"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "スタイルナンバ書込み OK" & vbCrLf)




        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "Flash Memoryへの書込み NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try

        Call WriteLog(False, "Flash Memoryへの書込み OK")

    End Sub

    Public Sub IB2(ByVal iName As String)
        '==============================================================================================================
        ' Function : IB2:Flash MemoryへMac Address書込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        '==========================================================================================
        Dim m_Count As Integer = 0

        Dim mMACAddress1 As String = ""
        Dim mMACAddress2 As String = ""
        Dim mMACAddress3 As String = ""
        Dim mMACAddress4 As String = ""
        Dim mMACAddress5 As String = ""

        Dim mMACAddressCon1 As String = ""
        Dim mMACAddressCon2 As String = ""
        Dim mMACAddressCon3 As String = ""
        Dim mMACAddressCon4 As String = ""
        Dim mMACAddressCon5 As String = ""

        Dim Model As String = ModelRead()
        Dim Model_Len As Integer = 0
        Dim Model_Char() As String = {""}
        Dim Model_Asc() As Integer = {0}
        Dim Model_Hex() As String = {""}
        Dim Model_Hex_Full As String = ""
        Dim Obj_add As String = ""
        Dim j_Hex As String = ""
        Dim k_Hex As String = ""
        Dim Model_Mod As Integer = 0   '0=偶数, 1=奇数
        Dim j As Integer = 0
        Dim k As Integer = 0

        Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"

        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As String = ""
        Dim rData() As String = {""}

        Dim i As Integer

        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mReadLine As String = ""

        Try
            ' Temporary.txt에서 MAC Address 읽어오기
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(lFile)
            Do While sr.Peek() >= 0
                mReadLine = sr.ReadLine
                If InStr(mReadLine, "MacAddress1 : ") <> 0 Then
                    mMACAddress1 = Mid(mReadLine, Len("MacAddress1 : ") + 1, mReadLine.Length - Len("MacAddress1 : "))
                ElseIf InStr(mReadLine, "MacAddress2 : ") <> 0 Then
                    mMACAddress2 = Mid(mReadLine, Len("MacAddress2 : ") + 1, mReadLine.Length - Len("MacAddress2 : "))
                ElseIf InStr(mReadLine, "MacAddress3 : ") <> 0 Then
                    mMACAddress3 = Mid(mReadLine, Len("MacAddress3 : ") + 1, mReadLine.Length - Len("MacAddress3 : "))
                ElseIf InStr(mReadLine, "MacAddress4 : ") <> 0 Then
                    mMACAddress4 = Mid(mReadLine, Len("MacAddress4 : ") + 1, mReadLine.Length - Len("MacAddress4 : "))
                ElseIf InStr(mReadLine, "MacAddress5 : ") <> 0 Then
                    mMACAddress5 = Mid(mReadLine, Len("MacAddress5 : ") + 1, mReadLine.Length - Len("MacAddress5 : "))
                End If
            Loop
            sr.Close()
            sr = Nothing

            '  MAC 주소가 모두 정상적으로 읽혔는지 확인
            If mMACAddress1 = "" Or mMACAddress2 = "" Or mMACAddress3 = "" Or mMACAddress4 = "" Or mMACAddress5 = "" Then
                mErrorComment = "アドレスが正常に発行されていない可能性があります。確認してください。"
                Throw New ApplicationException(mErrorComment)
            End If


            ' MAC 주소를 XX:XX:XX:XX:XX:XX 형태로 변환
            'mMACAddressCon1 = ConvertToMacFormat(mMACAddress1)
            'mMACAddressCon2 = ConvertToMacFormat(mMACAddress2)
            'mMACAddressCon3 = ConvertToMacFormat(mMACAddress3)
            'mMACAddressCon4 = ConvertToMacFormat(mMACAddress4)
            'mMACAddressCon5 = ConvertToMacFormat(mMACAddress5)

            ' 변환된 어드레스 로그 확인용
            Call WriteLog(False, "MAC Address1:" & mMACAddress1)
            Call WriteLog(False, "MAC Address2:" & mMACAddress2)
            Call WriteLog(False, "MAC Address3:" & mMACAddress3)
            Call WriteLog(False, "MAC Address4:" & mMACAddress4)
            Call WriteLog(False, "MAC Address5:" & mMACAddress5)

            ' 5개의 MAC 주소를 배열로 묶기
            Dim macList As String() = {mMACAddress1, mMACAddress2, mMACAddress3, mMACAddress4, mMACAddress5}

            cIA8Comm = New cls101_ComRS232C
            cIA8Comm.RS_COMM = Insfrm.SerialPort1
            cIA8Comm.PORT = rs232cRZPort
            If RS232CSettingRead(rs232cRZPort, mSetting, mDelim) = False Then
                mErrorComment = "RS232SET Read異常"
                Throw New ApplicationException(mErrorComment)
            End If
            cIA8Comm.SETTING = mSetting
            cIA8Comm.TERMINAL_CHAR = mDelim
            cIA8Comm.PortSetup()

            ' 5개의 MAC 주소를 순차적으로 Flash 메모리에 쓰기
            ' ※ idx=0 → writeAddress=6 >> 임시 FX3 TS 커맨드 
            '    idx=1 → writeAddress=7
            '    idx=2 → writeAddress=8
            '    idx=3 → writeAddress=9
            '    idx=4 → writeAddress=A (TSN eND POINT)
            For idx As Integer = 0 To 4
                Call WriteLog(False, "")
                Call WriteLog(False, "MAC" & (idx + 1).ToString() & "書込み ============" & " " & Format(DateTime.Now, "HH:mm:ss"))

                Dim writeAddress As String
                If idx < 4 Then
                    writeAddress = (6 + idx).ToString()
                Else
                    ' 5번째는 "A"로 처리 > TSN END POINT 카키코미
                    writeAddress = "A"
                End If

                cIA8Comm.SendCmd = $"TS,0130,2,{writeAddress},{macList(idx)}"
                Call WriteLog(False, "Send : " & cIA8Comm.SendCmd)

                ' Send / Recv
                mSecWait(1000)
                Call cIA8Comm.FuncSendCmd()
                mSecWait(1000)
                Call cIA8Comm.FuncRecvCmd()

                Call WriteLog(False, "Recv : " & cIA8Comm.RECEIVE_DATA)

                ' 응답 분석, RecvCmd로 rCount가 string임
                Call RecvCmd(cIA8Comm.RECEIVE_DATA, rRes, rCode, rCount, rData)
                If Me.ERR_COMMENT <> "" Then
                    mTestMeasure = cIA8Comm.RECEIVE_DATA
                    mErrorComment = "MAC" & (idx + 1).ToString() & "書込み　異常"
                    Throw New ApplicationException(mErrorComment)
                End If
            Next
            ' Temporary.mdb에 맥어드레스 저장
            Call WriteData("MacAddress", mMACAddress1)
            Call WriteData("MacAddress1", mMACAddress2)
            Call WriteData("MacAddress2", mMACAddress3)
            Call WriteData("MacAddress3", mMACAddress4)
            Call WriteData("MacAddress4", mMACAddress5)

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "Mac Address書込み NG"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cIA8Comm IsNot Nothing Then
                cIA8Comm.PortClose()
                cIA8Comm = Nothing
            End If
        End Try
    End Sub

    Public Sub IA9(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA9: MACアドレス 重複確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        Dim cSQL As cls010_MDB_Operation = Nothing
        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""
        Dim rField As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()
        Dim currentMacAddress As String = ""
        Dim macFieldNames As String() =
        {
            "MacAddress",
            "MacAddress1",
            "MacAddress2",
            "MacAddress3",
            "MacAddress4"
        }
        Try

            ' 중복 검사
            For i As Integer = 0 To macFieldNames.Length - 1

                '  Temporary.mdb에서 해당 MacAddress 읽어오기

                rField = macFieldNames(i)

                cSQL = New cls010_MDB_Operation
                dbName = mAppPath & "210 Result Data\Temporary.mdb"
                tbName = "tblResult"
                sSQL = "SELECT " & rField & " FROM " & tbName

                cSQL.DB_PATH = dbName
                cSQL.TB_NAME = tbName
                cSQL.DB_SQL = sSQL

                Call cSQL.DB_READ_JET(fName, sData)
                currentMacAddress = sData(0)  ' 읽은 MAC 주소

                cSQL = Nothing

                If String.IsNullOrEmpty(currentMacAddress) Then
                    mErrorComment = $"{macFieldNames(i)} が空です。MACアドレスが正しく読み取れているか確認してください。"
                    Throw New ApplicationException(mErrorComment)
                End If

                ' MacAddress.mdb (서버)에서 현재 MAC 주소가 1개인지 확인
                cSQL = New cls010_MDB_Operation
                dbName = ServerPath & "\MACaddress\A8HCA\MacAddress.mdb"
                tbName = "tblPastMAC"

                sSQL = "SELECT PastMAC FROM " & tbName & " WHERE PastMAC = '" & currentMacAddress & "'"
                cSQL.DB_PATH = dbName
                cSQL.TB_NAME = tbName
                cSQL.DB_SQL = sSQL

                Call cSQL.DB_READ_JET(fName, sData)

                ' DB_DATA_COUNT 값으로 중복 또는 미등록 판정
                Select Case cSQL.DB_DATA_COUNT
                    Case 0
                        ' 0개라면 등록되지 않음
                        mErrorComment = $"{macFieldNames(i)} が登録されていません！"
                        Throw New ApplicationException(mErrorComment)

                    Case 1
                        ' 1개라면 정상
                        Call WriteLog(False, $"{macFieldNames(i)} は1つ登録確認OK")

                    Case Else
                        ' 2개 이상이면 중복
                        mErrorComment = $"{macFieldNames(i)} が1つ以上存在しています！"
                        Throw New ApplicationException(mErrorComment)
                End Select

                cSQL = Nothing
            Next

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            If cSQL IsNot Nothing Then
                cSQL = Nothing
            End If
        End Try
    End Sub

    Public Sub IA10(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA10: MACアドレス 重複確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        Dim wTime As Double = 0
        Dim mcAddress As String
        Dim Serial As String = SerialRead()
        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()
        Dim FolderPath As String = "A8HCA"

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        Dim dbName As String = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
        Dim tbName As String = "tblPastMAC"
        Dim rField As String = "PastMAC"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        '-----------------------------------------------------------------
        '  LAN0 ~ LAN4까지의 각 포트별로 이미 사용한(등록된) 맥주소가 있는지 확인 후
        '  등록되어 있지 않다면 신규로 MAC을 발번한 뒤 NextMac 발번 로직까지 수행
        '-----------------------------------------------------------------
        For i As Integer = 0 To 4
            Dim currentLanPort As String = "LAN" & i

            Try
                ' DB에서 해당 시리얼/포트로 과거 발번된 맥어드레스가 있는지 확인
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName

                sSQL = "SELECT " & rField &
                   " FROM " & tbName &
                   " WHERE Serial = '" & Serial & "'" &
                   "   AND LAN_Port = '" & currentLanPort & "'"

                cADO.DB_SQL = sSQL
                Call cADO.DB_READ_JET(fName, sData)

                ' 이미 등록된 맥주소가 있다면 재사용 후 로그만 출력
                If cADO.DB_DATA_COUNT <> 0 Then
                    Call WriteLog(False, "MacAddress.mdb에 " & Serial & " / " & currentLanPort & " 로 등록된 Address 재사용")
                    Call WriteLog(False, "MacAddress(" & currentLanPort & ") : " & sData(0))
                    GoTo SkipMacIssue
                End If

            Catch ex As Exception
                mErrorComment = "MacAddress.mdb(" & currentLanPort & ") Read 에러: " & ex.Message
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                cADO = Nothing
            End Try

            '-----------------------------------------------------------------
            '  신규 맥주소 발번 로직
            '-----------------------------------------------------------------
            Try
                ' 1) 현재 맥주소 발급
                cMac = New cls031_MacAddress
                cMac.DBFileName = dbName

                Call cMac.IssueMacAddress()
                If cMac.ERR_MSG <> "" Then
                    mErrorComment = cMac.ERR_MSG
                    Throw New ApplicationException(mErrorComment)
                End If
                mcAddress = cMac.MacAddress
                Call WriteLog(False, "MacAddress(" & currentLanPort & ") : " & mcAddress)

                ' 2) Next MAC ADDRESS 발급
                Call WriteLog(False, "Next MAC ADDRESS ･･･")

                cMac.Serial = Serial
                cMac.LAN_Port = currentLanPort

                Call cMac.IssueNextMacAddress()
                If cMac.ERR_MSG <> "" Then
                    mErrorComment = cMac.ERR_MSG
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, "Next MAC ADDRESS OK!")

            Catch ex As Exception
                ' 예외 발생 시 메시지 후 Throw
                mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1)
                mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                cMac = Nothing
            End Try

SkipMacIssue:
            ' 다음 LAN 포트로 진행
        Next

    End Sub


    Public Sub I89(ByVal iName As String)
        '==============================================================================================================
        ' Function : I89:SD挿抜TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'SD挿抜TESTメッセージ表示
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\A8HCA\SDLubricationTest.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "SD挿抜TEST 異常 "
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "SD挿抜TEST 正常" & " " & Format(DateTime.Now, "HH:mm:ss"))

    End Sub
    Public Sub I90(ByVal iName As String)
        '==============================================================================================================
        ' Function : I90:ﾕﾆｯﾄ電源ON
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim slot As Integer = 2

        Dim rCmd As String = ""

        If InStr(iName, "I90") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================

        Try
            Double.Parse(Strings.Right(iName, 1))
            slot = CInt(Strings.Right(iName, 1))        '文字列の一番右の数字がｽﾛｯﾄ番号
        Catch
            '文字列の一番右が数字でなければｽﾛｯﾄ2（ｱﾄﾞｵﾝCPU）で起動
        End Try

        'Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入：ｽﾛｯﾄ" & CStr(slot))
        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入")

        Try
            PLC_Y("01" & "BWRY" & "301" & ",001," & "1")
            WaitStart(32)
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

        End Try

        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入OK")

    End Sub
    Public Sub I99(ByVal iName As String)
        '==============================================================================================================
        ' Function : I99:最初ﾕﾆｯﾄ電源ON
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim slot As Integer = 2

        Dim rCmd As String = ""

        If InStr(iName, "I90") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================

        Try
            Double.Parse(Strings.Right(iName, 1))
            slot = CInt(Strings.Right(iName, 1))        '文字列の一番右の数字がｽﾛｯﾄ番号
        Catch
            '文字列の一番右が数字でなければｽﾛｯﾄ2（ｱﾄﾞｵﾝCPU）で起動
        End Try

        'Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入：ｽﾛｯﾄ" & CStr(slot))
        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入")

        Try
            '7segment,DP(점)LED 테스트 ***************************************************
            Call WriteLog(False, "")
            Call WriteLog(False, "DP LEDテスト開始")

            'DP 확인 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "DP_LED.ppt",
                     "DP LEDテスト開始",
                     "DP LEDテスト",
                     "YELLOW", "OFF", "ON", "ON",
                     "DP LEDテスト　異常")

            PLC_Y("01" & "BWRY" & "301" & ",001," & "1")
            WaitStart(27)
            'DP 확인 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "DP_LED1.ppt",
                     "DP LEDテスト開始",
                     "DP LEDテスト",
                     "YELLOW", "OFF", "ON", "ON",
                     "DP LEDテスト　異常")

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

        End Try

        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入OK")

    End Sub
    Public Sub I91(ByVal iName As String)
        '==============================================================================================================
        ' Function : I91:ﾕﾆｯﾄ電源OFF
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        If InStr(iName, "I91") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================
        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源遮断")

        Try
            PLC_Y("01" & "BWRY" & "301" & ",001," & "0")
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally

        End Try

        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源遮断OK")

    End Sub

    Public Sub I92(ByVal iName As String)
        '==============================================================================================================
        ' Function : I92:検査開始準備
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Dim mFileName As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        '==========================================================================================

        Try
            If InStr(Model, "NBT") > 0 Then
                mFileName = "NBT.ppt"
            Else
                mFileName = "NORMAL.ppt"
            End If

            ShowMessagePopupAndCheck(mAppPath,
                     mFileName,
                     "PWR FAIL 점등확인",
                     "검사준비확인",
                     "YELLOW", "OFF", "OFF", "OFF",
                     "検査準備　異常")

            ShowMessagePopupAndCheck(mAppPath,
                     "SetUp.ppt",
                     "PWR FAIL 점등확인",
                     "검사준비확인",
                     "YELLOW", "OFF", "OFF", "OFF",
                     "検査準備　異常")

            If InStr(Model, "NBT") = 0 Then
                ShowMessagePopupAndCheck(mAppPath,
                     "SetUp1.ppt",
                     "PWR FAIL 점등확인",
                     "검사준비확인",
                     "YELLOW", "OFF", "OFF", "OFF",
                     "検査準備　異常")
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally


        End Try

        Call WriteLog(False, "検査準備OK")
    End Sub

    Public Sub I93(ByVal iName As String)
        '==============================================================================================================
        ' Function : I93:検査終了時メッセージ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        'Dim scmd As String
        Dim rCmd As String = ""

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Dim mReadLine As String = ""

        Dim temp(12) As String
        Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"

        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rData() As String = {}

        Dim Model As String = ModelRead()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================


        Try
            ' 방진캡 부착
            'ShowMessagePopupAndCheck(mAppPath,
            '         "DustCap.ppt",
            '         "DustCap Check",
            '         "DustCap Check",
            '         "YELLOW", "OFF", "OFF", "OFF",
            '         "DustCap 準備　異常")


            'If InStr(Model, "NBT") = 0 Then
            '    ShowMessagePopupAndCheck(mAppPath,
            '         "DustCap2.ppt",
            '         "DustCap Check",
            '         "DustCap Check",
            '         "YELLOW", "OFF", "OFF", "OFF",
            '         "DustCap 準備　異常")
            'End If

            '검사완료 - > 여기에 방진캡 부착 집어넣을 에정임
            ShowMessagePopupAndCheck(mAppPath,
                     "Fin.ppt",
                     "DustCap Check",
                     "DustCap Check",
                     "GREEN", "OFF", "OFF", "OFF",
                     "出荷準備　異常")

        Catch ex As Exception

            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

            dfrm = Nothing

        End Try

        '最終確認
        Call WriteLog(False, "*")

        Call WriteLog(False, "●　検査合格時の作業  ======================================")
        Call WriteLog(False, "・主銘板に印字された内容を確認して、製品に貼り付ける。")
        Call WriteLog(False, "-以上、お疲れ様でした。(^o^) -")

        Insfrm.btnRePrintLabel.Enabled = True

    End Sub

    Public Sub I94(ByVal iName As String)
        '==============================================================================================================
        ' Function : I94:Result Log Save
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 検査合格の場合だけ作成
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Dim res As String() = {}
        'Make Folder ========================================================
        Dim SearchArea As String = "\\10.36.14.211\aiec\result_LOG"
        Dim SearchYear As String = Format(DateTime.Now, "yyyy")
        Dim SearchMonth As String = Format(DateTime.Now, "yyyy-MM")
        Dim CalYear As String
        Dim CalMonth As String
        Dim CalTemp As String
        'Search date folder -----------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea
            cFop.FOLDER_NAME = SearchYear

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cFop = Nothing

        End Try

        'Empty check ------------------------------------------------
        Dim sFolder As String = String.Empty
        For Each stFilePath As String In res
            sFolder &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then make folder -------------------------------------
        If sFolder = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.FOLDER_PATH = SearchArea & "\" & SearchYear

                Call cFop.FOLDER_MAKE()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try

            If mErrorComment <> "" Then Exit Sub
        End If

        'Make File ========================================================
        'Search month date file ------------------------------------------- 
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\" & SearchYear
            cFop.FILE_NAME = SearchMonth & ".mdb"

            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        Dim sFile = String.Empty
        sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then copy file ---------------------------------------
        If sFile = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = SearchArea & "\" & "ResultMaster.mdb"
                cFop.FOLDER_PATH = SearchArea & "\" & SearchYear & "\"     '& SearchFolder & "\"
                cFop.FILE_NAME = SearchMonth & ".mdb"

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = "ResultMaster.mdbがありません。"
                Throw New ApplicationException(mErrorComment)
            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Result Dataの読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String = "tblResult"
        ' StyleNum, ReleaseNum 필드 추가
        Dim rField As String = "StartDateTime, MSCODE, SerialNo, EquipmentNo, RunMode, StyleNum, ReleaseNum"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim fNameDel() As String = {""}
        Dim sDataDel() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing
        End Try

        'Search and delete redundant data=======================================================
        If sData(4) = "MAN" Then
            Call WriteLog(False, "Manual検査時には Reuslt Log Saveしません。")
            Exit Sub
        ElseIf sData(4) = "AUT" Then
            For i = 0 To 1
                dbName = SearchArea & "\" & SearchYear & "\" & SearchMonth & ".mdb"

                Select Case i 'yyyy/mm 再設定
                    Case 0 '現月の場合
                    Case 1 '先月の場合
                        If Strings.Right(SearchMonth, 2) = "01" Then ' たとえば、2018/01なら 2017/12で検索 

                            CalYear = CStr(CInt(SearchYear) - 1)
                            CalMonth = "12"
                            dbName = SearchArea & "\" & CalYear & "\" & CalYear & "-" & CalMonth & ".mdb"

                        Else ' 2月～12月なら -1ヵ月

                            CalTemp = CStr(CInt(Strings.Right(SearchMonth, 2)) - 1)

                            If CInt(CalTemp) >= 10 Then
                                CalMonth = SearchYear & "-" & CalTemp
                            Else
                                CalMonth = SearchYear & "-0" & CalTemp
                            End If

                            dbName = SearchArea & "\" & SearchYear & "\" & CalMonth & ".mdb"

                        End If
                End Select

                Try
                    cADO = New cls010_MDB_Operation
                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName

                    sSQL = "SELECT * FROM " & tbName & " WHERE MSCODE = '" & sData(1) & "' AND SerialNo = '" & sData(2) & "'"
                    cADO.DB_SQL = sSQL

                    Call cADO.DB_READ_JET(fNameDel, sDataDel)

                    If cADO.DB_DATA_COUNT > 0 Then
                        sSQL = "DELETE FROM " & tbName & " WHERE MSCODE = '" & sDataDel(0) & "' AND SerialNo = '" & sDataDel(1) & "'"
                        cADO.DB_PATH = dbName
                        cADO.TB_NAME = tbName
                        cADO.DB_SQL = sSQL

                        Call cADO.DB_EXEC_JET()
                    End If
                Catch ex As Exception
                    If mErrorComment = "" Then mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                Finally
                    cADO = Nothing
                End Try
            Next
        Else
            mErrorComment = "Run Mode 異常"
            Throw New ApplicationException(mErrorComment)

        End If

        'Write of Data ---------------------------------------------
        dbName = SearchArea & "\" & SearchYear & "\" & SearchMonth & ".mdb"
        tbName = "tblResult"

        ' StyleNo, ReleaseNo 컬럼 추가하여 INSERT 구문 수정
        sSQL = "INSERT INTO " & tbName & " ("
        sSQL = sSQL & fName(0) & ", "
        sSQL = sSQL & fName(1) & ", "
        sSQL = sSQL & fName(2) & ", "
        sSQL = sSQL & fName(3) & ", "
        'sSQL = sSQL & fName(4) & ", "
        sSQL = sSQL & "StyleNo, "
        sSQL = sSQL & "ReleaseNo"
        sSQL = sSQL & ") VALUES("
        sSQL = sSQL & "'" & sData(0) & "', "
        sSQL = sSQL & "'" & sData(1) & "', "
        sSQL = sSQL & "'" & sData(2) & "', "
        sSQL = sSQL & "'" & sData(3) & "', "
        'sSQL = sSQL & "'" & sData(4) & "', "
        sSQL = sSQL & "'" & sData(5) & "', "  ' StyleNum 값을 StyleNo 컬럼에 저장
        sSQL = sSQL & "'" & sData(6) & "'"    ' ReleaseNum 값을 ReleaseNo 컬럼에 저장
        sSQL = sSQL & ")"

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing
        End Try

    End Sub

    'origin
    'Public Sub I94(ByVal iName As String)
    '    '==============================================================================================================
    '    ' Function : I94:Result Log Save
    '    ' Data     : Process name
    '    ' Return   : Nothing
    '    ' Remarks  : 検査合格の場合だけ作成
    '    '==============================================================================================================
    '    Call WriteLog(False, "")
    '    Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
    '    Call WriteLog(False, "*")
    '    '==============================================================================================================

    '    Dim res As String() = {}
    '    'Make Folder ========================================================
    '    Dim SearchArea As String = "\\10.36.14.211\aiec\result_LOG"
    '    Dim SearchYear As String = Format(DateTime.Now, "yyyy")
    '    Dim SearchMonth As String = Format(DateTime.Now, "yyyy-MM")
    '    Dim CalYear As String
    '    Dim CalMonth As String
    '    Dim CalTemp As String
    '    'Search date folder -----------------------------------------
    '    Try
    '        cFop = New cls012_File_Operation

    '        cFop.ROOT_FOLDER = SearchArea
    '        cFop.FOLDER_NAME = SearchYear

    '        Call cFop.FOLDER_SEARCH(res)

    '    Catch ex As Exception
    '        mErrorComment = cFop.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cFop = Nothing

    '    End Try

    '    'Empty check ------------------------------------------------
    '    Dim sFolder As String = String.Empty
    '    For Each stFilePath As String In res
    '        sFolder &= stFilePath & System.Environment.NewLine
    '    Next stFilePath

    '    'Empty then make folder -------------------------------------
    '    If sFolder = String.Empty Then
    '        Try
    '            cFop = New cls012_File_Operation

    '            cFop.FOLDER_PATH = SearchArea & "\" & SearchYear

    '            Call cFop.FOLDER_MAKE()

    '        Catch ex As Exception
    '            mErrorComment = cFop.ERR_COMMENT
    '            If mErrorComment = "" Then mErrorComment = ex.ToString

    '        Finally
    '            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

    '        End Try

    '        If mErrorComment <> "" Then Exit Sub
    '    End If

    '    'Make File ========================================================
    '    'Search month date file ------------------------------------------- 
    '    Try
    '        cFop = New cls012_File_Operation

    '        cFop.ROOT_FOLDER = SearchArea & "\" & SearchYear
    '        cFop.FILE_NAME = SearchMonth & ".mdb"

    '        Call cFop.FILE_SEARCH(res)

    '    Catch ex As Exception
    '        mErrorComment = cFop.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString
    '    Finally
    '        mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    'Empty check ------------------------------------------------
    '    Dim sFile = String.Empty
    '    sFile = String.Empty
    '    For Each stFilePath As String In res
    '        sFile &= stFilePath & System.Environment.NewLine
    '    Next stFilePath

    '    'Empty then copy file ---------------------------------------
    '    If sFile = String.Empty Then
    '        Try
    '            cFop = New cls012_File_Operation

    '            cFop.ORIGINAL_PATH = SearchArea & "\" & "ResultMaster.mdb"
    '            cFop.FOLDER_PATH = SearchArea & "\" & SearchYear & "\"     '& SearchFolder & "\"
    '            cFop.FILE_NAME = SearchMonth & ".mdb"

    '            Call cFop.FILE_COPY()

    '        Catch ex As Exception
    '            mErrorComment = "ResultMaster.mdbがありません。"
    '            Throw New ApplicationException(mErrorComment)
    '        Finally
    '            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

    '        End Try
    '        If mErrorComment <> "" Then Exit Sub

    '    End If

    '    'Result Dataの読込み
    '    Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
    '    Dim rDTkey As String = ""
    '    Dim tbName As String = "tblResult"
    '    Dim rField As String = "StartDateTime, MSCODE, SerialNo, EquipmentNo, RunMode"
    '    Dim sSQL As String
    '    Dim fName() As String = {""}
    '    Dim sData() As String = {""}
    '    Dim fNameDel() As String = {""}
    '    Dim sDataDel() As String = {""}

    '    Try
    '        cADO = New cls010_MDB_Operation
    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName

    '        sSQL = "SELECT " & rField & " FROM " & tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '    Catch ex As Exception

    '    Finally
    '        cADO = Nothing
    '    End Try

    '    'Search and delete redundant data=======================================================
    '    If sData(4) = "MAN" Then
    '        Call WriteLog(False, "Manual検査時には Reuslt Log Saveしません。")
    '        Exit Sub
    '    ElseIf sData(4) = "AUT" Then
    '        For i = 0 To 1
    '            dbName = SearchArea & "\" & SearchYear & "\" & SearchMonth & ".mdb"

    '            Select Case i 'yyyy/mm 再設定
    '                Case 0 '現月の場合
    '                Case 1 '先月の場合
    '                    If Strings.Right(SearchMonth, 2) = "01" Then ' たとえば、2018/01なら 2017/12で検索 

    '                        CalYear = CStr(CInt(SearchYear) - 1)
    '                        CalMonth = "12"
    '                        dbName = SearchArea & "\" & CalYear & "\" & CalYear & "-" & CalMonth & ".mdb"

    '                    Else ' 2月～12月なら -1ヵ月

    '                        CalTemp = CStr(CInt(Strings.Right(SearchMonth, 2)) - 1)

    '                        If CInt(CalTemp) >= 10 Then
    '                            CalMonth = SearchYear & "-" & CalTemp
    '                        Else
    '                            CalMonth = SearchYear & "-0" & CalTemp
    '                        End If

    '                        dbName = SearchArea & "\" & SearchYear & "\" & CalMonth & ".mdb"

    '                    End If
    '            End Select

    '            Try
    '                cADO = New cls010_MDB_Operation
    '                cADO.DB_PATH = dbName
    '                cADO.TB_NAME = tbName

    '                sSQL = "SELECT * FROM " & tbName & " WHERE MSCODE = '" & sData(1) & "' AND SerialNo = '" & sData(2) & "'"
    '                cADO.DB_SQL = sSQL

    '                Call cADO.DB_READ_JET(fNameDel, sDataDel)

    '                If cADO.DB_DATA_COUNT > 0 Then
    '                    sSQL = "DELETE FROM " & tbName & " WHERE MSCODE = '" & sDataDel(0) & "' AND SerialNo = '" & sDataDel(1) & "'"
    '                    cADO.DB_PATH = dbName
    '                    cADO.TB_NAME = tbName
    '                    cADO.DB_SQL = sSQL

    '                    Call cADO.DB_EXEC_JET()
    '                End If
    '            Catch ex As Exception
    '                If mErrorComment = "" Then mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
    '                Call WriteLog(False, mErrorComment)
    '                Throw New ApplicationException(mErrorComment)
    '            Finally
    '                cADO = Nothing
    '            End Try
    '        Next
    '    Else
    '        mErrorComment = "Run Mode 異常"
    '        Throw New ApplicationException(mErrorComment)

    '    End If
    '    'Write of Data ---------------------------------------------
    '    dbName = SearchArea & "\" & SearchYear & "\" & SearchMonth & ".mdb"
    '    tbName = "tblResult"

    '    sSQL = "INSERT INTO " & tbName & " ("
    '    sSQL = sSQL & fName(0) & ", "
    '    sSQL = sSQL & fName(1) & ", "
    '    sSQL = sSQL & fName(2) & ", "
    '    sSQL = sSQL & fName(3)
    '    sSQL = sSQL & ") VALUES("
    '    sSQL = sSQL & "'" & sData(0) & "', "
    '    sSQL = sSQL & "'" & sData(1) & "', "
    '    sSQL = sSQL & "'" & sData(2) & "', "
    '    sSQL = sSQL & "'" & sData(3) & "'"
    '    sSQL = sSQL & ")"

    '    Try
    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_EXEC_JET()

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing
    '    End Try

    'End Sub
    Public Sub I95(ByVal iName As String)
        '==============================================================================================================
        ' Function : I95:測定器Preset
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        'If InStr(iName, "I95") <> 0 Then
        '    Call WriteLog(False, "")
        '    Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        '    Call WriteLog(False, "*")
        'End If
        '==========================================================================================

        Try
            '測定器Reset
            myEq.Open()
            myEq.Preset()
            myEq.Close()

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

        End Try

        'Call WriteLog(False, "GOS, DMM, ADR Preset")

    End Sub

    Public Sub I97(ByVal iName As String)
        '==============================================================================================================
        ' Function : I97:ﾕﾆｯﾄ電源OFF&ON
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Try
            'Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入")

            PLC_Y("01" & "BWRY" & "301" & ",001," & "0")
            Call WriteLog(False, "検査ユニット電源遮断OK")

            Call WriteLog(False, "約5秒 Wait")
            Call WaitStart(5)

            PLC_Y("01" & "BWRY" & "301" & ",001," & "1")
            Call WriteLog(False, "検査ユニットOK")
            Call WriteLog(False, "モジュール起動まで約22秒 Wait")
            Call WaitStart(32)
            Call WriteLog(False, "検査ユニット電源投入OK")

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "ﾕﾆｯﾄ電源OFF&ON NG"
            End If
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally

        End Try
    End Sub

    Public Sub I98(ByVal iName As String)
        '==============================================================================================================
        ' Function : I98:出荷設定確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'A8 Comm 준비
        Dim cIA8Comm As cls101_ComRS232C = Nothing
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer
        Dim rData() As String = {""}

        Try
            Call WriteLog(False, "出荷設定確認開始")
            Call WriteLog(False, "出荷設定:Mode SW 0,0確認")　'ver 0.02 hsj add
            Call WriteLog(False, "出荷設定:RDY ON 確認")　'ver 0.02 hsj add
            Call WriteLog(False, "出荷設定:BAT LED OFF 確認")　'ver 0.02 hsj add

            '출하설정 1/3 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "Shipment_Set_1.ppt",
                     "出荷設定",
                     "出荷設定",
                     "YELLOW", "OFF", "OFF", "OFF",
                     "出荷設定　異常")

            ' 전원 OFF
            Call WriteLog(False, "モジュールの再起動開始")
            Me.I91("検査モジュールの電源OFF")
            mSecWait(1000)
            Me.I90("検査モジュールの電源ON")
            Call WriteLog(False, "モジュールの再起動OK" & vbCrLf)

            '출하설정 2/3 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "Shipment_Set_2.ppt",
                     "出荷設定",
                     "出荷設定",
                     "YELLOW", "OFF", "OFF", "OFF",
                     "出荷設定　異常")

            '출하설정 3/3 ppt ---------------------------
            ShowMessagePopupAndCheck(mAppPath,
                     "Shipment_Set_3.ppt",
                     "出荷設定",
                     "出荷設定",
                     "YELLOW", "OFF", "OFF", "OFF",
                     "出荷設定　異常")

            Call WriteLog(False, "出荷設定確認OK")
        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "出荷設定 NG"
            End If
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally

        End Try
    End Sub

    Private Function FlagCheck() As Boolean

        Dim cMcm1 As cls120_SP_M_Command
        Dim t, t1 As Integer
        Dim TimeOut As Integer
        Dim TotalWait As Double

        cMcm1 = New cls120_SP_M_Command
        cMcm1.IPAdd = IP_U2                      'CPU用LE経由
        cMcm1.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm1.SlotNo = 1                         'CPU Slot番号
        cMcm1.UnitNo = 0                         'ユニット番号

        If mTimeOut = 0 Then

            TimeOut = 10

        Else
            TimeOut = mTimeOut
        End If

        t = System.Environment.TickCount
        Do
            t1 = System.Environment.TickCount
            TotalWait = CInt((t1 - t) / 1000)
            If TotalWait >= TimeOut Then
                mErrorComment = "timeout" : Exit Function
            End If

            Call cMcm1.IO_READ(0, 2, mTestExeAddress.ToString("X8"))
            DoEvents()
            Select Case Strings.Right(cMcm1.RecvText, 4)
                Case "0200" : Exit Do
                Case Else
            End Select
            System.Threading.Thread.Sleep(10)
        Loop

        FlagCheck = True
        cMcm1 = Nothing

    End Function
    Private Function FullModelRead() As String

        '検査型式の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "MSCODE"            'InspectionModel  MSCODE
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        FullModelRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            FullModelRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function
    Private Function ModelRead() As String

        '検査型式の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "InspectionModel"            'InspectionModel  MSCODE
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        ModelRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            ModelRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function

    Private Function SerialRead() As String

        'PLASMA Serial番号の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "SerialNo"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        SerialRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            SerialRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function

    Private Function ServerPathRead() As String

        'serverパスの読込み
        Dim dbName As String = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "BackupDrivePath"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        ServerPathRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblEquipmentSetting"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            ServerPathRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function

    Public Sub WriteLog(ByVal TmStmp As Boolean, ByVal msg As String)
        '==========================================================================================
        ' Function : Write Log(Temporary.txt)
        ' Data     : Time stamp Y/N , message
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"
        Try
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = lFile

            If msg <> "*" Then
                cTxt.WRITE_TEXT_FILE(TmStmp, msg)
            Else
                cTxt.WRITE_TEXT_LINE(TmStmp, "*", 50)
            End If

        Catch ex As Exception
            mErrorComment = cTxt.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cTxt = Nothing

        End Try
        DoEvents()

        Call Insfrm.LogBrowserNavi()

    End Sub

    Private Function RS232CSettingRead(ByVal mPort As String, ByRef mSetting As String, ByRef mTerminalChar As String) As Boolean

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim SETDATA() As String

        If mPort = "" Then
            RS232CSettingRead = False
            Exit Function
        End If


        cADO = New cls010_MDB_Operation
        dbName = mAppPath & "010 EquipmentData\機器構成一覧.xls"
        tbName = "[RS232C$]"
        sSQL = "SELECT 設定 , デリミタ FROM " & tbName & " WHERE PORT LIKE '" & mPort & "'"

        cADO.DB_PATH = dbName
        cADO.TB_NAME = tbName
        cADO.DB_SQL = sSQL

        cADO.DB_READ_EXCEL_JET(fName, sData)

        If cADO.DB_DATA_COUNT = 0 Then
            RS232CSettingRead = False
            cADO = Nothing
            Exit Function
        Else
            SETDATA = Split(sData(0), "-")
            mSetting = SETDATA(0) & "," & SETDATA(1) & "," & SETDATA(2) & "," & SETDATA(3)
            Select Case sData(1)
                Case "vbcrlf" : mTerminalChar = vbCrLf
                Case "vbcr" : mTerminalChar = vbCr
                Case "vblf" : mTerminalChar = vbLf
            End Select
            RS232CSettingRead = True
        End If

        cADO = Nothing

    End Function

    Private Sub RPRecvCmd(ByVal mRecvCmd As String, ByRef rRes As String, ByRef rCode As String, ByRef rCount As Integer, ByRef rData() As String)
        '==========================================================================================
        ' F.Name   : RPRecvCmd
        ' Function : RS_Communication
        ' Data     : mRecvCmd
        ' Return   : rRes,rCode,rCount,rData()
        ' Remarks  :
        '==========================================================================================

        Dim TotalWait As Single
        Dim RcvData() As String
        Dim RcvCount() As Integer
        Dim i As Integer

        rRes = ""
        rCode = ""
        rCount = 0
        mErrorComment = ""

        On Error GoTo ErrorTrap

        'Result,CODE,DATA COUNT,DATA1,DATA2,...,DATAF(MAX) --
        RcvData = Split(mRecvCmd, ",")

        rRes = RcvData(0)
        rCode = RcvData(1)

        rCount = Convert.ToInt32(RcvData(2), 16)  'rCount에 split한 값을 저장해야 전원확인 테스트에서 확인 가능

        'Result check --------------------
        If rRes <> "OK" Then
            mErrorComment = rRes & " " & rCode & " " & rCount & " "
            For i = 0 To (rCount - 1)

                If i = (rCount - 1) Then
                    mErrorComment = mErrorComment & rData(i)
                Else
                    mErrorComment = mErrorComment & rData(i) & " "
                End If

            Next

            GoTo ErrorTrap

        End If

        Exit Sub

ErrorTrap:

        If mErrorComment = "" Then mErrorComment = "ERROR:RP COM : " & mRecvCmd

    End Sub
    Private Sub RecvCmd(ByVal mRecvCmd As String, ByRef rRes As String, ByRef rCode As String, ByRef rCount As String, ByRef rData() As String)
        '==========================================================================================
        ' F.Name   : RPRecvCmd
        ' Function : RS_Communication
        ' Data     : mRecvCmd
        ' Return   : rRes,rCode,rCount,rData()
        ' Remarks  :
        '==========================================================================================

        Dim TotalWait As Single
        Dim RcvData() As String
        Dim RcvCount() As Integer
        Dim i As Integer

        rRes = ""
        rCode = ""
        rCount = ""
        mErrorComment = ""

        On Error GoTo ErrorTrap

        'Result,CODE,DATA COUNT,DATA1,DATA2,...,DATAF(MAX) --
        RcvData = Split(mRecvCmd, ",")

        rRes = RcvData(0)
        rCode = RcvData(1)
        rCount = RcvData(2)  'rCount에 split한 값을 저장해야 전원확인 테스트에서 확인 가능

        'Result check --------------------
        If rRes <> "OK" Then
            mErrorComment = rRes & " " & rCode & " " & rCount & " "
            GoTo ErrorTrap

        End If

        Exit Sub

ErrorTrap:

        If mErrorComment = "" Then mErrorComment = "ERROR:RP COM : " & mRecvCmd

    End Sub
    Private Sub RPRecvCmd_1(ByVal mRecvCmd As String,
                        ByRef rRes As String, ByRef rCode As String,
                        ByRef rCount As Integer, ByRef rData() As String)

        Dim RcvData() As String
        Dim i As Integer

        ' 초기화
        rRes = ""
        rCode = ""
        rCount = 0
        mErrorComment = ""

        On Error GoTo ErrorTrap

        ' [1] 응답 문자열 Split
        RcvData = Split(mRecvCmd, ",")
        ' 예: "OK,1620,4,8,0,0,0,5,D,0,1,0,0,4,3" → 
        '      RcvData = { "OK", "1620", "4", "8", "0", "0", "0", "5", "D", "0", "1", "0", "0", "4", "3" }

        ' [2] 최소 토큰 체크 (OK, <명령어>, 데이터1… 은 적어도 3개 이상이어야 함)
        If RcvData.Length < 3 Then
            mErrorComment = mRecvCmd
            GoTo ErrorTrap
        End If

        ' [3] 첫 번째 토큰(OK) 확인
        If RcvData(0).ToUpper() <> "OK" Then
            mErrorComment = mRecvCmd
            GoTo ErrorTrap
        End If

        ' [4] rRes, rCode, rCount 설정
        rRes = RcvData(0)    ' 예: "OK"
        rCode = RcvData(1)   ' 예: "1620"
        '   신형 형식 가정: 나머지( index 2 ~ 끝 ) 전부가 rData
        rCount = RcvData.Length - 2

        If rCount <= 0 Then
            mErrorComment = "Data Err"
            GoTo ErrorTrap
        End If

        ' [5] rData 배열에 복사
        ReDim rData(rCount - 1)

        For i = 0 To rCount - 1
            rData(i) = RcvData(i + 2)
        Next

        ' (오류 없이 끝까지 왔으면 정상 종료)
        Exit Sub

ErrorTrap:
        ' 에러가 발생할 경우
        If mErrorComment = "" Then mErrorComment = "ERROR:RP COM : " & mRecvCmd
    End Sub

    Private Sub WriteValueData(ByVal fName As String, ByVal sData As String)
        '==========================================================================================
        ' Function : Write Value data(F3RP7XData\xxxx.mdb)
        ' Data     : Field name , data
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""
        Dim rsData() As String = {""}
        Dim rdData() As String = {""}
        Dim rfName() As String = {""}
        Dim rdName() As String = {""}

        Dim DataCount As Integer = 0

        Dim res As String() = {}
        Dim SearchArea As String = mAppPath & "210 Result Data\A8HCAData"
        Dim SearchFile As String = Format(DateTime.Now, "yyyy")


        dbName = mAppPath & "210 Result Data\Temporary.mdb"
        tbName = "tblResult"
        sSQL = "SELECT StartDateTime,MSCODE,SerialNo FROM " & tbName

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(rfName, rsData)

            DataCount = cADO.DB_DATA_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        If DataCount = 0 Then
            mErrorComment = "Temporary.mdbにData無し"
            Throw New ApplicationException(mErrorComment)
        End If

        'Search month date file****************************************
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\"
            cFop.FILE_NAME = SearchFile & ".mdb"

            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Throw New ApplicationException(mErrorComment)
        Finally
            cFop = Nothing
        End Try


        'Empty check***************************************************
        Dim sFile = String.Empty
        sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then copy file******************************************
        If sFile = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = SearchArea & "\" & "A8HCADataMaster.mdb"
                cFop.FOLDER_PATH = SearchArea & "\"
                cFop.FILE_NAME = SearchFile & ".mdb"

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString
                Throw New ApplicationException(mErrorComment)
            Finally
                cFop = Nothing
            End Try
        End If



        dbName = SearchArea & "\" & SearchFile & ".mdb"
        tbName = "tblResult"

        'データがあるかどうかチェック
        sSQL = "SELECT * FROM " & tbName & " "
        sSQL &= "WHERE 検査日時='" & rsData(0) & "' AND MSCODE='" & rsData(1) & "' AND SerialNo='" & rsData(2) & "'"
        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(rdName, rdData)
            DataCount = cADO.DB_DATA_COUNT
        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        'Write of data ----------------------------------------------
        If DataCount = 0 Then
            sSQL = "INSERT INTO " & tbName & " (検査日時,MSCODE,SerialNo," & fName & ") "
            sSQL &= "VALUES ('" & rsData(0) & "','" & rsData(1) & "','" & rsData(2) & "','" & sData & "')"
        Else
            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= fName & " = '" & sData & "' "
            sSQL &= "WHERE 検査日時='" & rsData(0) & "' AND MSCODE='" & rsData(1) & "' AND SerialNo='" & rsData(2) & "'"
        End If

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_EXEC_JET()
        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

    End Sub
    Private Sub WriteData(ByVal fName As String, ByVal sData As String)
        '==========================================================================================
        ' Function : Write data(Temporary.mdb)
        ' Data     : Field name , data
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""

        dbName = mAppPath & "210 Result Data\Temporary.mdb"
        tbName = "tblResult"

        'Write of data ----------------------------------------------
        sSQL = "UPDATE " & tbName & " SET "
        sSQL &= fName & " = '" & sData & "'"

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try

    End Sub

    Private Sub CountStart(ByVal sVal As Double)
        '==========================================================================================
        ' Function : Count Start
        ' Data     : Setting time
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Call Insfrm.CountStart(sVal)

    End Sub

    Private Sub CountStop()
        '==========================================================================================
        ' Function : Count Stop
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Call Insfrm.CountStop()

    End Sub

    Private Function PLC_Y(ByVal scmd As String) As Long
        Dim cLnk As New cls080_F3LinkToLE
        Dim rCmd As String = ""
        Dim mErrorComment As String

        cLnk.IPAdd = IP_U1
        Try
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Or InStr(rCmd, "ERR") > 0 Then
                If cLnk.ERR_COMMENT = "" Then
                    mErrorComment = "YD Relay CONT ERR"
                Else
                    mErrorComment = cLnk.ERR_COMMENT

                End If

                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cLnk = Nothing
        End Try

    End Function
    Public Sub rBox_Clear_Ter()
        Dim cLnk As New cls080_F3LinkToLE
        Dim rCmd As String = ""
        Dim mErrorComment As String
        Dim mSend As String

        cLnk.IPAdd = IP_U1
        Try
            For i = 4 To 5
                mSend = "01" & "BWRY" & i & "01" & ",032," & "00000000000000000000000000000000"

                Call cLnk.F3PCLinkCommunication(mSend, rCmd)
                If cLnk.ERR_COMMENT <> "" Then
                    mErrorComment = cLnk.ERR_COMMENT
                    Throw New ApplicationException(mErrorComment)
                End If
            Next

        Catch ex As Exception
            cLnk = Nothing
        End Try

    End Sub

    Public Sub rBox_Clear()
        Dim cLnk As New cls080_F3LinkToLE
        Dim rCmd As String = ""
        Dim mErrorComment As String
        Dim mSend As String

        cLnk.IPAdd = IP_U1
        Try
            For i = 4 To 6
                mSend = "01" & "BWRY" & i & "01" & ",032," & "00000000000000000000000000000000"

                Call cLnk.F3PCLinkCommunication(mSend, rCmd)
                If cLnk.ERR_COMMENT <> "" Then
                    mErrorComment = cLnk.ERR_COMMENT
                    Throw New ApplicationException(mErrorComment)
                End If
            Next

        Catch ex As Exception
            cLnk = Nothing
        End Try

    End Sub

    Private Function GOSSend(ByVal GOSAdd As Integer, ByVal Range_Set As String, ByVal InputSetVoltage As Double, ByVal OutSet As Boolean) As Long

        Dim cGOS = New cls230_7651(GOSAdd)
        cGOS = New cls230_7651(GOSAdd)

        Try
            cGOS.Open()

            Select Case Range_Set
                Case "AUTO" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.RAuto)
                Case "V10m" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V10m)
                Case "V100m" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V100m)
                Case "V1" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V1)
                Case "V10" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V10)
                Case "V30" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V30)
                Case "A1m" : cGOS.SetRange(cls230_7651.Mode.DCA, cls230_7651.Range.A1m)
                Case "A10m" : cGOS.SetRange(cls230_7651.Mode.DCA, cls230_7651.Range.A10m)
                Case "A100m" : cGOS.SetRange(cls230_7651.Mode.DCA, cls230_7651.Range.A100m)
            End Select
            cGOS.SetVoltage(InputSetVoltage)
            cGOS.SetOutput(OutSet)

            Call mSecWait(200)

        Catch ex As Exception
            mErrorComment = "Gos 出力設定異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            cGOS.Close()
            cGOS = Nothing
        End Try

    End Function
    Private Function GOSSend_Cal(ByVal GOSAdd As Integer, ByVal Range_Set As String, ByVal InputSetVoltage As Single, ByVal OutSet As Boolean, ByVal dAdd As Integer, ByVal Spec As Single) As Long
        Dim Range_Data As String = ""
        Dim Spec_Min As Single
        Dim Spec_Max As Single
        'Dim Spec As Single = 0.0007
        Dim InputSetVoltage2 As Single = 0.0
        Dim fInputVolt As Single

        Dim dbName As String = mAppPath & "210 Result Data\Output_Adjust.mdb"
        Dim tbname As String = "tblVolt" & dAdd
        Dim sSQL As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim SearchV As String = ""

        Dim cGOS = New cls230_7651(GOSAdd)
        cGOS = New cls230_7651(GOSAdd)
        cADO = New cls010_MDB_Operation

        Try
            SearchV = "V" & Replace(Replace(CStr(InputSetVoltage), ".", ""), "-", "m")

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbname

            sSQL = "SELECT " & SearchV & " FROM " & tbname
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)

            fInputVolt = CSng(sData(0))

            Spec_Max = InputSetVoltage + Spec
            Spec_Min = InputSetVoltage - Spec

            cGOS.Open()

            Select Case Range_Set
                Case "AUTO" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.RAuto)
                Case "V10m" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V10m)
                Case "V100m" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V100m)
                Case "V1" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V1)
                Case "V10" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V10)
                Case "V30" : cGOS.SetRange(cls230_7651.Mode.DCV, cls230_7651.Range.V30)
                Case "A1m" : cGOS.SetRange(cls230_7651.Mode.DCA, cls230_7651.Range.A1m)
                Case "A10m" : cGOS.SetRange(cls230_7651.Mode.DCA, cls230_7651.Range.A10m)
                Case "A100m" : cGOS.SetRange(cls230_7651.Mode.DCA, cls230_7651.Range.A100m)
            End Select

            '調整用PLC Relay設定
            PLC_Y("01" & "BWRY" & "429,002," & "00")

            If dAdd = 3 Then
                PLC_Y("01" & "BWRY" & "429,001," & "1")
            ElseIf dAdd = 4 Then
                PLC_Y("01" & "BWRY" & "430,001," & "1")
            End If

            cGOS.SetVoltage(fInputVolt)
            cGOS.SetOutput(OutSet)

            'Call mSecWait(1000)

            '抵抗設定値測定
            Call DMMRead(dAdd, Range_Set)

            Call WriteLog(True, "入力値：" & fInputVolt)
            Call WriteLog(True, "測定値：" & Dmm_Read_Data)

            '出力値調整
            For i = 1 To 6
                If Dmm_Read_Data < Spec_Max And Dmm_Read_Data > Spec_Min Then
                    If i > 1 Then
                        sSQL = "UPDATE " & tbname & " SET "
                        sSQL &= SearchV & " = '" & CStr(InputSetVoltage2) & "'"

                        cADO.DB_SQL = sSQL

                        Call cADO.DB_EXEC_JET()
                    End If

                    Exit For
                End If
                If i > 5 Then
                    Call WriteLog(True, "調整５回失敗")
                    Throw New ApplicationException("GOS設定異常,Specに合わない")
                End If

                If Dmm_Read_Data > Spec_Max Or Dmm_Read_Data < Spec_Min Then
                    Call WriteLog(True, i & "回、再調整")

                    If i = 1 Then
                        InputSetVoltage2 = fInputVolt + (InputSetVoltage - CSng(Dmm_Read_Data))

                    ElseIf InputSetVoltage - CSng(Dmm_Read_Data) > 0 Then
                        InputSetVoltage2 += CSng(0.000001)
                    ElseIf InputSetVoltage - CSng(Dmm_Read_Data) < 0 Then
                        InputSetVoltage2 -= CSng(0.000001)

                    End If

                    cGOS.SetVoltage(InputSetVoltage2)

                    'Call mSecWait(1000)
                    DMMRead(dAdd, Range_Set)

                    Call WriteLog(True, "入力値：" & InputSetVoltage2)
                    Call WriteLog(True, "測定値：" & Dmm_Read_Data)

                End If
            Next i

        Catch ex As Exception
            mErrorComment = "Gos 出力設定異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            PLC_Y("01" & "BWRY" & "429,002," & "00")
            cGOS.Close()
            cGOS = Nothing
            cADO = Nothing
        End Try

    End Function

    Private Function DMMRead(ByVal DMMAdd As Integer, ByVal Range_Set As String) As Long

        Dim cDMM = New cls240_34410A(DMMAdd)
        cDMM = New cls240_34410A(DMMAdd)

        Try
            cDMM.Open()

            Select Case Range_Set
                Case "VAUTO" : cDMM.SetDCVoltageRange(cls240_34410A.DCVoltageRange.VAuto)
                Case "V100mV" : cDMM.SetDCVoltageRange(cls240_34410A.DCVoltageRange.V100mV)
                Case "V1V" : cDMM.SetDCVoltageRange(cls240_34410A.DCVoltageRange.V1V)
                Case "V10V" : cDMM.SetDCVoltageRange(cls240_34410A.DCVoltageRange.V10V)
                Case "V100V" : cDMM.SetDCVoltageRange(cls240_34410A.DCVoltageRange.V100V)
                Case "V1000V" : cDMM.SetDCVoltageRange(cls240_34410A.DCVoltageRange.V1000V)

                Case "AAUTO" : cDMM.SetDCAmpereRange(cls240_34410A.DCAmpereRange.AAuto)
                Case "A100uA" : cDMM.SetDCAmpereRange(cls240_34410A.DCAmpereRange.A100uA)
                Case "A1mA" : cDMM.SetDCAmpereRange(cls240_34410A.DCAmpereRange.A1mA)
                Case "A10mA" : cDMM.SetDCAmpereRange(cls240_34410A.DCAmpereRange.A10mA)
                Case "A100mA" : cDMM.SetDCAmpereRange(cls240_34410A.DCAmpereRange.A100mA)
                Case "A1A" : cDMM.SetDCAmpereRange(cls240_34410A.DCAmpereRange.A1A)
                Case "A3A" : cDMM.SetDCAmpereRange(cls240_34410A.DCAmpereRange.A3A)
            End Select

            Call mSecWait(200)

            Dmm_Read_Data = CDbl(cDMM.GetMeasureDMM(0))

        Catch ex As Exception
            mErrorComment = "DMM設定異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            cDMM.Close()
            cDMM = Nothing
        End Try

    End Function

    Private Function ADR_Outl(ByVal InputSetReg As Single, ByVal dAdd As Integer) As Long
        cADR_1 = New cls231_ADR_3204(2)

        Try
            PLC_Y("01" & "BWRY" & "401,010," & "0000000000")

            cADR_1.Open()
            cADR_1.SetResistance(InputSetReg)

        Catch ex As Exception
            mErrorComment = "ADR設定異常"
            Throw New ApplicationException(mErrorComment)
        Finally

            cADR_1.Close()
            cADR_1 = Nothing
        End Try

    End Function
    '------------------------------- Ver 0.01 JDI ADD Start ------------------------------------
    Private Function ADRNoCall(ByVal InputSetReg As Single, ByVal dAdd As Integer) As Long
        Dim fInputReg As Single

        Dim dbName As String = mAppPath & "210 Result Data\Output_Adjust.mdb"
        Dim tbname As String = "tblReg" & dAdd
        Dim sSQL As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim SearchR As String = ""

        cADR_1 = New cls231_ADR_3204(2)
        cADO = New cls010_MDB_Operation

        Try
            SearchR = "R" & Replace(CStr(InputSetReg), ".", "")

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbname

            sSQL = "SELECT " & SearchR & " FROM " & tbname & " WHERE DATE = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)

            '---------------------- Ver 0.01 JDi ADD Start -----------------------------
            If sData(0) = "" Then
                mErrorComment = "補正情報無し、RTD、RJCの始業点検が必要"
                Call WriteLog(True, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            '----------------------- Ver 0.01 JDi ADD End ------------------------------

            fInputReg = CSng(sData(0))

            cADR_1.Open()
            cADR_1.SetResistance(fInputReg)

            Call WriteLog(True, "抵抗値：" & fInputReg)

        Catch ex As Exception
            mErrorComment = "ADR設定異常"
            Throw New ApplicationException(mErrorComment)
        Finally

            cADR_1.Close()
            cADR_1 = Nothing
        End Try

    End Function
    '-------------------------------- Ver 0.01 JDI ADD End -------------------------------------
    Private Function ADRCall(ByVal InputSetReg As Single, ByVal Range_Set As String, ByVal dAdd As Integer, ByVal spec As Single, ByVal tWait As Integer) As Long
        Dim Spec_Min As Double
        Dim Spec_Max As Double
        Dim InputSetReg2 As Single = 0.0
        Dim fInputReg As Single

        Dim dbName As String = mAppPath & "210 Result Data\Output_Adjust.mdb"
        Dim tbname As String = "tblReg" & dAdd
        Dim sSQL As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim SearchR As String = ""

        cADR_1 = New cls231_ADR_3204(2)
        cADO = New cls010_MDB_Operation

        Try
            PLC_Y("01" & "BWRY" & "401,010," & "0000000000")

            If dAdd = 3 Then
                PLC_Y("01" & "BWRY" & "414,001," & "1")
                PLC_Y("01" & "BWRY" & "419,001," & "1")
            Else
                PLC_Y("01" & "BWRY" & "415,001," & "1")
                PLC_Y("01" & "BWRY" & "420,001," & "1")
            End If

            SearchR = "R" & Replace(CStr(InputSetReg), ".", "")

            cADO.DB_PATH = dbName

            If dAdd = 4 Then
                cADO.TB_NAME = tbname & "S"
            Else
                cADO.TB_NAME = tbname
            End If

            sSQL = "SELECT " & SearchR & " FROM " & tbname
            If dAdd = 4 Then sSQL = sSQL & "S"

            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)

            fInputReg = CSng(sData(0))

            '調整用PLC Relay設定
            Spec_Max = InputSetReg + spec
            Spec_Min = InputSetReg - spec

            cADR_1.Open()
            cADR_1.SetResistance(fInputReg)

            Call WaitStart(tWait)

            '抵抗設定値測定
            DMMRead(dAdd, Range_Set)

            Call WriteLog(True, "抵抗値：" & fInputReg)
            Call WriteLog(True, "測定値：" & Dmm_Read_Data)

            '出力値調整
            For i = 1 To 6
                If Dmm_Read_Data < Spec_Max And Dmm_Read_Data > Spec_Min Then
                    If i > 1 Then
                        If dAdd = 4 Then
                            sSQL = "UPDATE " & tbname & "S SET "
                        Else
                            sSQL = "UPDATE " & tbname & " SET "
                        End If

                        sSQL &= SearchR & " = '" & CStr(InputSetReg2) & "'"

                        cADO.DB_SQL = sSQL

                        Call cADO.DB_EXEC_JET()
                    Else
                        InputSetReg2 = fInputReg
                    End If

                    Exit For
                End If

                If i > 5 Then
                    Call WriteLog(True, "調整５回失敗")
                    Throw New ApplicationException("GOS設定異常,Specに合わない")
                End If

                If Dmm_Read_Data > Spec_Max Or Dmm_Read_Data < Spec_Min Then
                    Call WriteLog(True, i & "回、再調整")

                    If i = 1 Then
                        InputSetReg2 = fInputReg + (InputSetReg - CSng(Dmm_Read_Data))

                    ElseIf InputSetReg - CSng(Dmm_Read_Data) < 0 Then
                        InputSetReg2 -= CSng(0.001)
                    ElseIf InputSetReg - CSng(Dmm_Read_Data) > 0 Then
                        InputSetReg2 += CSng(0.001)
                    End If

                    cADR_1.SetResistance(InputSetReg2)

                    Call WaitStart(3)
                    DMMRead(dAdd, Range_Set)

                    Call WriteLog(True, "入力値：" & InputSetReg2)
                    Call WriteLog(True, "測定値：" & Dmm_Read_Data)

                End If
            Next i

            '--------------------------- Ver 0.01 JDI ADD Start ----------------------
            If dAdd = 4 Then
                sSQL = "UPDATE " & tbname & " SET "
                sSQL &= SearchR & " = '" & CStr(InputSetReg2) & "'"
                sSQL &= " WHERE DATE = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"

                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbname

                cADO.DB_SQL = sSQL

                Call cADO.DB_EXEC_JET()
            End If
            '---------------------------- Ver 0.01 JDI ADD End -----------------------

        Catch ex As Exception
            mErrorComment = "ADR設定異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            PLC_Y("01" & "BWRY" & "414,002," & "00")
            PLC_Y("01" & "BWRY" & "419,002," & "00")

            cADR_1.Close()
            cADR_1 = Nothing
        End Try

    End Function

    Private Function F3PCLinkCom(ByVal command As String, ByVal comment As String) As String

        F3PCLinkCom = ""

        Dim RecvDataLE As String = ""
        Dim cLnk1 As cls080_F3LinkToLE

        cLnk1 = New cls080_F3LinkToLE
        cLnk1.IPAdd = IP_U1
        cLnk1.F3PCLinkCommunication(command, RecvDataLE)
        If InStr(RecvDataLE, vbCrLf) <> 0 Then RecvDataLE = Replace(RecvDataLE, vbCrLf, "")
        F3PCLinkCom = RecvDataLE

        cLnk1 = Nothing

    End Function
    Private Sub prvRNGSL(ByVal RangeNum As Int16, ByVal SMP As Int32)
        Dim intCh As Integer
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0001", rCmd)  '操作コマンド有効
        mSecWait(200)                                                             '100ms以上ないとNGになる

        cLnk.F3PCLinkCommunication("01SWW005," & "0081" & ",01," & "0001", rCmd)  '電源周波数, 0 : 50Hz, 1 ; 60Hz
        mSecWait(50)
        cLnk.F3PCLinkCommunication("01SWW005," & "0082" & ",01," & "000" & SMP, rCmd)  '入力サンプリング周期変更, 0 : 100ms, 1 : 200ms, 3 : 10ms
        mSecWait(50)

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0002", rCmd)  '操作指示
        mSecWait(100)

        For intCh = 1 To 4
            cLnk.F3PCLinkCommunication("01SWW005,0" & CStr(142 + 200 * (intCh - 1)) & ",01," & RangeNum.ToString("X4"), rCmd)
            mSecWait(50)
        Next intCh

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0004", rCmd)  '操作指示

        If prvDUTX08Check() = False Then
            Throw New ApplicationException(mErrorComment)
        End If

        mSecWait(100)

        'Burn Out
        For intCh = 1 To 4
            cLnk.F3PCLinkCommunication("01SWW005,0" & 150 + 200 * (intCh - 1) & ",01,0000", rCmd)
            mSecWait(50)
        Next intCh

        'RJC
        For intCh = 1 To 4
            cLnk.F3PCLinkCommunication("01SWW005,0" & 149 + 200 * (intCh - 1) & ",01,0001", rCmd)
            mSecWait(50)
        Next intCh

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0008", rCmd)  '操作指示

        If prvDUTX08Check() = False Then
            Throw New ApplicationException(mErrorComment)
        End If

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0000", rCmd)  '操作コマンド無効
        mSecWait(50)

    End Sub

    Private Sub prvRNGBNRJC(ByVal RangeNum As Integer, ScaleNum As Integer)
        Dim intCh As Integer
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0001", rCmd)  '操作コマンド有効
        mSecWait(200)                                        '100ms以上ないとNGになる

        For intCh = 1 To 4
            cLnk.F3PCLinkCommunication("01SWW005,0" & CStr(142 + 200 * (intCh - 1)) & ",01," & RangeNum.ToString("X4"), rCmd)
            mSecWait(50)
        Next intCh

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0004", rCmd)  '操作指示

        If prvDUTX08Check() = False Then Throw New ApplicationException(mErrorComment)

        mSecWait(100)

        Select Case RangeNum
            Case 1          'バーンアウト
                For intCh = 1 To 4
                    cLnk.F3PCLinkCommunication("01SWW005,0" & 150 + 200 * (intCh - 1) & ",01," & ScaleNum.ToString("X4"), rCmd)
                    mSecWait(50)
                Next intCh
            Case 3          'RJC
                cLnk.F3PCLinkCommunication("01SWW005," & "0081" & ",01," & "0001", rCmd)  '電源周波数, 0 : 50Hz, 1 ; 60Hz
                mSecWait(50)
                cLnk.F3PCLinkCommunication("01SWW005," & "0082" & ",01," & "0002", rCmd)  '入力サンプリング周期変更, 0 : 100ms, 1 : 200ms, 2 : 10ms
                mSecWait(50)

                cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0002", rCmd)  '操作指示
                mSecWait(100)

                For intCh = 1 To 4
                    cLnk.F3PCLinkCommunication("01SWW005,0" & 149 + 200 * (intCh - 1) & ",01," & ScaleNum.ToString("X4"), rCmd)
                    mSecWait(50)
                Next intCh
            Case Else
        End Select

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0008", rCmd)  '操作指示

        If prvDUTX08Check() = False Then Throw New ApplicationException(mErrorComment)

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0000", rCmd)    '操作コマンド無効
        mSecWait(100)
    End Sub
    Private Sub prvOUTREG()
        Dim intCh As Integer
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0001", rCmd)  '操作コマンド有効
        mSecWait(200)                                                             '100ms以上ないとNGになる

        cLnk.F3PCLinkCommunication("01SWW005," & "0081" & ",01," & "0001", rCmd)  '電源周波数, 0 : 50Hz, 1 ; 60Hz
        mSecWait(50)

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0002", rCmd)  '操作指示
        mSecWait(100)

        cLnk.F3PCLinkCommunication("01SWW005," & "0087" & ",01," & "00FF", rCmd)  '連続出力設定(OUT1～8)

        For intCh = 1 To 8
            cLnk.F3PCLinkCommunication("01SWW005,00" & 91 + intCh - 1 & ",01," & "0001", rCmd)   '出力設定値を出力(OUT1～8)
            mSecWait(50)
        Next intCh

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0002", rCmd)  '操作指示

        If prvDUTX08Check() = False Then Throw New ApplicationException(mErrorComment)

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0000", rCmd)    '操作コマンド無効
        mSecWait(1000)

    End Sub
    Private Sub prvOUTTR()
        Dim intCh As Integer
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0001", rCmd)  '操作コマンド有効
        mSecWait(200)                                                             '100ms以上ないとNGになる

        cLnk.F3PCLinkCommunication("01SWW005," & "0081" & ",01," & "0001", rCmd)  '電源周波数, 0 : 50Hz, 1 ; 60Hz
        mSecWait(50)

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0002", rCmd)  '操作指示
        mSecWait(100)

        cLnk.F3PCLinkCommunication("01SWW005," & "0087" & ",01," & "0000", rCmd)  '連続出力設定(OUT1～8)

        For intCh = 1 To 8
            cLnk.F3PCLinkCommunication("01SWW005,00" & 91 + intCh - 1 & ",01," & "0001", rCmd)   '出力設定値を出力(OUT1～8)
            mSecWait(50)
        Next intCh

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0002", rCmd)  '操作指示

        If prvDUTX08Check() = False Then Throw New ApplicationException(mErrorComment)

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0000", rCmd)    '操作コマンド無効
        mSecWait(1000)

    End Sub

    Private Sub prvPRST()
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0001", rCmd)  '操作コマンド有効
        mSecWait(200)                                                             '100ms以上ないとNGになる

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0001", rCmd)  '操作指示(全パラメータ初期化)

        Call prvDUTX08Check()

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0000", rCmd)    '操作コマンド無効
        mSecWait(1000)
    End Sub
    '入力のレンジ切換え(校正用)
    'Sub INRNG@
    Private Sub prvINRNG(ByVal RangeNum As Integer)
        Dim intCh As Integer
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0001", rCmd)  '操作コマンド有効
        mSecWait(200)                                                             '100ms以上ないとNGになる 

        cLnk.F3PCLinkCommunication("01SWW005," & "0081" & ",01," & "0001", rCmd)  '電源周波数, 0 : 50Hz, 1 ; 60Hz
        mSecWait(50)

        cLnk.F3PCLinkCommunication("01SWW005," & "0082" & ",01," & "0001", rCmd)  '入力サンプリング周期変更, 0 : 100ms, 1 : 200ms, 2 : 10ms
        mSecWait(50)

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0002", rCmd)  '操作指示
        mSecWait(100)

        For intCh = 1 To 4
            cLnk.F3PCLinkCommunication("01SWW005,0" & 142 + 200 * (intCh - 1) & ",01," & RangeNum.ToString("X4"), rCmd)
            mSecWait(50)
        Next intCh

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0004", rCmd)  '操作指示

        Call prvDUTX08Check()

        mSecWait(200)

        For intCh = 1 To 4
            cLnk.F3PCLinkCommunication("01SWW005,0" & 150 + 200 * (intCh - 1) & ",01," & "0000", rCmd)
            mSecWait(50)
        Next intCh

        cLnk.F3PCLinkCommunication("01SWW005," & "0072" & ",01," & "0008", rCmd)  '操作指示

        Call prvDUTX08Check()

        cLnk.F3PCLinkCommunication("01SWW005," & "0071" & ",01," & "0000", rCmd)  '操作コマンド無効
        mSecWait(1000)

    End Sub


    ' 以下、呼出されるFunction
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '変更時のステータスチェック X08 OFF(0)→ON(1)
    Private Function prvDUTX08Check() As Boolean
        Dim X_08 As Integer = &H4008
        Dim S_Chk As Integer = &H4073
        Dim t, t1 As Integer
        Dim TimeOut As Integer
        Dim TotalWait As Double
        Dim result As Long
        Dim rCmd As String = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2

        If mTimeOut = 0 Then
            TimeOut = 10
        Else
            TimeOut = mTimeOut
        End If

        t = System.Environment.TickCount

        Do
            t1 = System.Environment.TickCount
            TotalWait = CInt((t1 - t) / 1000)
            If TotalWait >= TimeOut Then
                prvDUTX08Check = False
                mErrorComment = "timeout" : Exit Function
            End If

            cLnk.F3PCLinkCommunication("01BRDX508,001", rCmd)

            DoEvents()
            Select Case Mid(rCmd, 5, 1)
                Case "1" : Exit Do

                Case Else
            End Select
        Loop

        cLnk.F3PCLinkCommunication("01BRDX537,001", rCmd)

        If Strings.Right(rCmd, 1) <> "0" Then
            prvDUTX08Check = False
            mErrorComment = "ステータスのチェックNG:" & Convert.ToString(result)
            WriteLog(True, mErrorComment)
            Exit Function
        End If

        prvDUTX08Check = True

    End Function

    '16進数形式 → IEEE754(Float)浮動小数点形式変換
    Public Function HEXtoFloat(ByVal dblDataDEC As String) As Single
        Dim uintRead As UInteger
        Dim ByRead() As Byte

        uintRead = UInteger.Parse(dblDataDEC, System.Globalization.NumberStyles.AllowHexSpecifier)
        ByRead = BitConverter.GetBytes(uintRead)
        HEXtoFloat = BitConverter.ToSingle(ByRead, 0)
        Console.WriteLine("float convert = {0}", HEXtoFloat)

    End Function
    'IEEE754(Float)浮動小数点形式 → 16進数(HEX)形式変換
    Public Shared Function SingleToHex(ByVal Value As Single) As String
        Dim bytValue As Byte() = BitConverter.GetBytes(Value)
        Dim strHex As String = ""

        ' リトルエンディアンならバイト配列を逆転させる
        If BitConverter.IsLittleEndian Then Array.Reverse(bytValue)

        ' バイト配列を文字列化
        For Each b As Byte In bytValue : strHex += b.ToString("X2") : Next
        Return strHex
    End Function

    Private Function Dec2Bin(ByVal iVal As Double) As String
        Dec2Bin = ""
        Do
            If Len(Dec2Bin) >= 16 Then Exit Do

            If (iVal Mod 2) = 0 Then
                Dec2Bin = "0" + Dec2Bin
            Else
                Dec2Bin = "1" + Dec2Bin
            End If

            iVal = CLng(iVal) \ 2
        Loop
    End Function

    Private Sub WaitStart(ByVal sVal As Double)
        '==========================================================================================
        ' Function : Wait Start
        ' Data     : Setting time
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Call Insfrm.WaitStart(sVal)

    End Sub

    Private Sub mSecWait(ByVal mSec As Integer)
        '==========================================================================================
        ' Function : mSec Wait
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait

    End Sub

    Private Sub ShowMessagePopupAndCheck(ByVal appPath As String,
                                 ByVal pptFileName As String,
                                 ByVal funcName As String,
                                 ByVal dispMessage As String,
                                 ByVal ledColor As String,
                                 ByVal buzz1 As String,
                                 ByVal buzz2 As String,
                                 ByVal flash As String,
                                 ByVal mErrorComment As String)
        ' 메시지 폼 생성
        Dim dfrm As New frm500_MessagePopUp

        ' 폼에 표시할 정보 설정
        dfrm.DISP_FILENAME = appPath & "900 Picture\ppt\A8HCA\" & pptFileName
        dfrm.DISP_TITLENAME = funcName
        dfrm.DISP_MESSAGE = dispMessage
        dfrm.LED_COLOR = ledColor
        dfrm.BUZZ1 = buzz1
        dfrm.BUZZ2 = buzz2
        dfrm.FLASH = flash

        ' 폼 표시
        dfrm.frmShow()

        ' 사용자가 버튼 누를 때까지 대기
        Do
            If dfrm.DIALOG_RESULT = "NG" Then
                dfrm.Close()
                dfrm = Nothing

                Throw New ApplicationException(mErrorComment)  ' NG 클릭 시 에러 발생
            End If

            If dfrm.DIALOG_RESULT = "OK" Then
                Exit Do   ' OK 클릭 시 루프 탈출
            End If

            ' CPU 점유 방지
            DoEvents()

        Loop

        ' 정상 종료
        dfrm.Close()
        dfrm = Nothing
    End Sub

    Private Sub ShowMessagePopupAndCheck_RETRY(ByVal appPath As String,
                                 ByVal pptFileName As String,
                                 ByVal funcName As String,
                                 ByVal dispMessage As String,
                                 ByVal ledColor As String,
                                 ByVal buzz1 As String,
                                 ByVal buzz2 As String,
                                 ByVal flash As String,
                                 ByVal mErrorComment As String,
                                 ByRef result As String)
        Dim dfrm As New frm510_MessagePopUp

        dfrm.DISP_FILENAME = appPath & "900 Picture\ppt\A8HCA\" & pptFileName
        dfrm.DISP_TITLENAME = funcName
        dfrm.DISP_MESSAGE = dispMessage
        dfrm.LED_COLOR = ledColor
        dfrm.BUZZ1 = buzz1
        dfrm.BUZZ2 = buzz2
        dfrm.FLASH = flash

        dfrm.frmShow()


        Do
            Select Case dfrm.DIALOG_RESULT
                Case "NG"
                    dfrm.Close()
                    dfrm = Nothing
                    result = "NG"
                    Throw New ApplicationException(mErrorComment)
                Case "OK"
                    result = "OK"
                    Exit Do
                Case "RETRY"
                    result = "RETRY"
                    Exit Do
                Case Else
                    DoEvents()
            End Select
        Loop


        dfrm.Close()
        dfrm = Nothing
    End Sub
    Private Sub MeasureVoltage(ByVal cComm As cls101_ComRS232C,
                          ByVal voltageName As String,
                          ByVal channelCode As String,
                          ByVal hexHighLimit As String,
                          ByVal hexLowLimit As String)

        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer = 0
        Dim rData() As String = {""}
        Dim mMeas As Double = 0.0
        Dim sPec As String = ""

        ' LSB 상수 정의
        Const LSB As Double = 0.003

        Try
            '채널 별 로깅
            Call WriteLog(False, "")
            Call WriteLog(False, $"{voltageName} 測定 =====")

            'Send command ---------------------------
            cComm.SendCmd = $"TS,{channelCode},2,{hexLowLimit},{hexHighLimit}"
            Call WriteLog(False, "Send : " & cComm.SendCmd)
            Call cComm.FuncSendCmd()

            mSecWait(500)

            'Recv Command --------------------------
            Call cComm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cComm.RECEIVE_DATA)
            Call RPRecvCmd(cComm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            If rRes <> "" Then
                ' mMeas 계산을 새로운 공식으로 변경 (소수점 3째자리에서 반올림)
                mMeas = Math.Round((rCount / 16.0) * LSB, 3)
            End If

            ' voltageName에 따라 고정된 스펙 출력
            Select Case voltageName
                Case "3_3V_SD_電源電圧_短絡前"
                    sPec = "3.100V ≦ 計測値 ≦ 3.400V"
                Case "3_3V_SD_電源電圧_短絡"
                    sPec = "0.000V ≦ 計測値 ≦ 0.500V"
                Case "5_2V_電源電圧"
                    sPec = "5.000V ≦ 計測値 ≦ 5.400V"
                Case "3_3V_SD_電源電圧"
                    sPec = "3.100V ≦ 計測値 ≦ 3.400V"
                Case "2_5V_電源電圧"
                    sPec = "2.330V ≦ 計測値 ≦ 2.670V"
                Case "1_05V_電源電圧"
                    sPec = "0.900V ≦ 計測値 ≦ 1.200V"
                Case "0_6V_電源電圧"
                    sPec = "0.520V ≦ 計測値 ≦ 0.680V"
                Case Else
                    sPec = "스펙 미정의"
            End Select

            Call WriteLog(False, "SPEC : " & sPec)
            Call WriteLog(False, "測定直   : " & mMeas & "V")

            If Me.ERR_COMMENT <> "" Then
                mTestSpec = sPec
                mTestMeasure = $"{voltageName} : {mMeas} V"
                mErrorComment = $"{voltageName} : {mMeas} V 異常"
                Call WriteLog(False, "判定基準 : " & sPec)
                Call WriteValueData(voltageName, CStr(mMeas))
                Throw New ApplicationException(mErrorComment)
            Else
                ' 실제 판정은 고정된 스펙 값으로 수행
                Dim specLowVolt As Double = 0.0
                Dim specHighVolt As Double = 0.0

                Select Case voltageName
                    Case "3_3V_SD_電源電圧_短絡前"
                        specLowVolt = 3.1 : specHighVolt = 3.4
                    Case "3_3V_SD_電源電圧_短絡"
                        specLowVolt = 0.000 : specHighVolt = 0.5
                    Case "5_2V_電源電圧"
                        specLowVolt = 5.0 : specHighVolt = 5.4
                    Case "3_3V_SD_電源電圧"
                        specLowVolt = 3.1 : specHighVolt = 3.4
                    Case "2_5V_電源電圧"
                        specLowVolt = 2.33 : specHighVolt = 2.67
                    Case "1_05V_電源電圧"
                        specLowVolt = 0.9 : specHighVolt = 1.2
                    Case "0_6V_電源電圧"
                        specLowVolt = 0.52 : specHighVolt = 0.68
                    Case Else
                        mErrorComment = $"{voltageName} : 스펙 미정의"
                        Throw New ApplicationException(mErrorComment)
                End Select

                If (mMeas < specLowVolt) OrElse (mMeas > specHighVolt) Then
                    mErrorComment = $"{voltageName} : {mMeas} V 異常"
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, $"{voltageName} : {mMeas} V 正常")
                Call WriteValueData(voltageName, CStr(mMeas) & vbCrLf)
            End If

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "TSCommand Recv ERR"
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

    End Sub


    'Private Sub MeasureVoltage(ByVal cComm As cls101_ComRS232C,
    '                       ByVal voltageName As String,
    '                       ByVal channelCode As String,
    '                       ByVal hexHighLimit As String,
    '                       ByVal hexLowLimit As String,
    '                       ByVal FSR As Double,
    '                       ByVal ADC_max As Double)

    '    Dim rRes As String = ""
    '    Dim rCode As String = ""
    '    Dim rCount As Integer = 0
    '    Dim rData() As String = {""}
    '    Dim mMeas As Double = 0.0
    '    Dim sPec As String = ""
    '    Dim lowVolt As Double = 0.0
    '    Dim highVolt As Double = 0.0

    '    Try
    '        '채널 별 로깅
    '        Call WriteLog(False, "")
    '        Call WriteLog(False, $"{voltageName} 測定 =========")

    '        'Send command ---------------------------
    '        cComm.SendCmd = $"TS,{channelCode},2,{hexLowLimit},{hexHighLimit}"
    '        Call WriteLog(False, "Send : " & cComm.SendCmd)
    '        Call cComm.FuncSendCmd()

    '        mSecWait(500)

    '        'Recv Command --------------------------
    '        Call cComm.FuncRecvCmd()
    '        Call WriteLog(False, "Recv : " & cComm.RECEIVE_DATA)
    '        Call RPRecvCmd(cComm.RECEIVE_DATA, rRes, rCode, rCount, rData)


    '        If rRes <> "" Then
    '            mMeas = Math.Round((rCount * FSR) / ADC_max, 3)
    '            Dim raw10bit As Integer = rCount >> 6
    '            'mMeas = Math.Floor((raw10bit * FSR) / ADC_max)
    '        End If

    '        Dim rawLow16 As Integer = Convert.ToInt32(hexLowLimit, 16)
    '        Dim rawHigh16 As Integer = Convert.ToInt32(hexHighLimit, 16)

    '        Dim adcLow10 As Integer = rawLow16 >> 6  ' 상위 10비트
    '        Dim adcHigh10 As Integer = rawHigh16 >> 6

    '        '스펙 범위 (V 단위로 변환)
    '        lowVolt = (Convert.ToInt32(hexLowLimit, 16) * FSR) / ADC_max
    '        highVolt = (Convert.ToInt32(hexHighLimit, 16) * FSR) / ADC_max
    '        'lowVolt = (adcLow10 * FSR) / ADC_max
    '        'highVolt = (adcHigh10 * FSR) / ADC_max
    '        sPec = $"{lowVolt:0.000} V ≦ 測定直 ≦ {highVolt:0.000} V"

    '        Call WriteLog(False, "SPEC : " & sPec)
    '        Call WriteLog(False, "測定直   : " & mMeas & "V")


    '        If Me.ERR_COMMENT <> "" Then
    '            mTestSpec = sPec
    '            mTestMeasure = $"{voltageName} : {mMeas} V"
    '            mErrorComment = $"{voltageName} : {mMeas} V 異常"
    '            Call WriteLog(False, "判定基準 : " & sPec)
    '            Call WriteValueData(voltageName, CStr(mMeas))
    '            Throw New ApplicationException(mErrorComment)
    '        Else
    '            'mMeas < lowVolt or mMeas > highVolt

    '            If (mMeas < lowVolt) OrElse (mMeas > highVolt) Then
    '                mErrorComment = $"{voltageName} : {mMeas} V 異常"
    '                Throw New ApplicationException(mErrorComment)
    '            End If

    '            Call WriteLog(False, $"{voltageName} : {mMeas} V 正常")
    '            Call WriteValueData(voltageName, CStr(mMeas) & vbCrLf)
    '        End If

    '    Catch ex As Exception
    '        If String.IsNullOrEmpty(mErrorComment) Then
    '            mErrorComment = "TSCommand Recv ERR"
    '        End If
    '        Call WriteLog(False, mErrorComment)
    '        Throw New ApplicationException(mErrorComment)
    '    End Try
    'End Sub

    Private Sub MeasureVoltage_Sysmon(ByVal cComm As cls101_ComRS232C,
                                  ByVal voltageName As String,
                                  ByVal channelCode As String,
                                  ByVal hexHighCode As String,
                                  ByVal hexLowCode As String)

        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer = 0
        Dim rData() As String = {""}
        Dim mMeasLow As Double = 0.0
        Dim mMeasHigh As Double = 0.0
        Dim sSpec As String = ""
        Dim mMeas As Double = 0.0

        Try

            WriteLog(False, "")
            WriteLog(False, $"{voltageName} [Sysmon] 測定 =========")


            ' TS 명령 전송 (Low, High 코드 순서)
            cComm.SendCmd = $"TS,{channelCode},2,{hexLowCode},{hexHighCode}"
            WriteLog(False, "Send : " & cComm.SendCmd)
            cComm.FuncSendCmd()

            mSecWait(500)
            cComm.FuncRecvCmd()
            WriteLog(False, "Recv : " & cComm.RECEIVE_DATA)

            RPRecvCmd(cComm.RECEIVE_DATA, rRes, rCode, rCount, rData)

            ' 주소별로 내부 레인지 설정 

            Dim rangeVal As Double = 0.0
            Select Case channelCode
                Case "0910"
                    rangeVal = 6.0
                Case "0940"
                    rangeVal = 3.0
                Case "0950"
                    rangeVal = 6.0
                Case "0960"
                    rangeVal = 3.0
                Case "0980"
                    rangeVal = 3.0
                Case "0990"
                    rangeVal = 3.0

            End Select
            If rRes <> "" Then
                ' 전압 Vin = (rCount * rangeVolt) / 1024
                Dim raw10bit As Integer = rCount >> 6  ' (16bit -> 10bit)
                mMeas = Math.Round((raw10bit * rangeVal) / 1024.0, 2) ' 소수점 3째자리까지 반올림
                'mMeas = Math.Floor((raw10bit * rangeVal) / 1024.0) ' 소수점 3째자리까지 반올림
            End If
            '실제 ADC 코드(16진) → 10비트로 변환 → 전압 계산

            ' Low, High 각각 16진수 → 10진수 → >> 6 → (ADC×range)/1024
            Dim rawLow16 As Integer = Convert.ToInt32(hexLowCode, 16)
            Dim rawHigh16 As Integer = Convert.ToInt32(hexHighCode, 16)

            Dim adcLow10 As Integer = rawLow16 >> 6  ' 상위 10비트
            Dim adcHigh10 As Integer = rawHigh16 >> 6

            mMeasLow = Math.Round((adcLow10 * rangeVal) / 1024.0, 2)
            mMeasHigh = Math.Round((adcHigh10 * rangeVal) / 1024.0, 2)
            'mMeasLow = Math.Floor((adcLow10 * rangeVal) / 1024.0)
            'mMeasHigh = Math.Floor((adcHigh10 * rangeVal) / 1024.0)


            sSpec = $"{mMeasLow:0.000} V ≦ 測定直 ≦ {mMeasHigh:0.000} V"

            'WriteLog(False, "Range : " & rangeVal & " V")
            Call WriteLog(False, "SPEC : " & sSpec)
            Call WriteLog(False, "測定直   : " & mMeas & "V")


            ' =======================

            If Me.ERR_COMMENT <> "" Then
                mTestSpec = sSpec
                mTestMeasure = $"{voltageName} : {mMeas} V"
                mErrorComment = $"{voltageName} : {mMeas} V 異常"
                Call WriteLog(False, "判定基準 : " & sSpec)
                Call WriteValueData(voltageName, CStr(mMeas))
                Throw New ApplicationException(mErrorComment)
            Else
                'mMeas < lowVolt or mMeas > highVolt

                If (mMeas < mMeasLow) OrElse (mMeas > mMeasHigh) Then
                    mErrorComment = $"{voltageName} : {mMeas} V 異常"
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, $"{voltageName} : {mMeas} V 正常")
                Call WriteValueData(voltageName, CStr(mMeas) & vbCrLf)
            End If

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "TSCommand Recv ERR"
            End If
            WriteLog(False, mErrorComment & " : " & ex.Message)
            Throw New ApplicationException(mErrorComment)
        End Try
    End Sub
    'c
    'Private Sub MeasureVoltage_Sysmon(ByVal cComm As cls101_ComRS232C,
    '                              ByVal voltageName As String,
    '                              ByVal channelCode As String,
    '                              ByVal hexHighCode As String,
    '                              ByVal hexLowCode As String)

    '    Dim rRes As String = ""
    '    Dim rCode As String = ""
    '    Dim rCount As Integer = 0
    '    Dim rData() As String = {""}
    '    Dim mMeasLow As Double = 0.0
    '    Dim mMeasHigh As Double = 0.0
    '    Dim sSpec As String = ""
    '    Dim mMeas As Double = 0.0

    '    Try

    '        WriteLog(False, "")
    '        WriteLog(False, $"{voltageName} [Sysmon] 測定 =========")

    '        '----------------------------------------------------------------------
    '        ' 1) TS 명령 전송 (Low, High 코드 순서)
    '        '    → 내부적으로 측정 후 ADC값(16비트)을 수신하게 됨
    '        '----------------------------------------------------------------------
    '        cComm.SendCmd = $"TS,{channelCode},2,{hexLowCode},{hexHighCode}"
    '        WriteLog(False, "Send : " & cComm.SendCmd)
    '        cComm.FuncSendCmd()

    '        mSecWait(500)
    '        cComm.FuncRecvCmd()
    '        WriteLog(False, "Recv : " & cComm.RECEIVE_DATA)

    '        '----------------------------------------------------------------------
    '        ' 2) 수신 데이터 파싱하여 rCount (16비트 ADC Code) 구하기
    '        '----------------------------------------------------------------------
    '        RPRecvCmd(cComm.RECEIVE_DATA, rRes, rCode, rCount, rData)

    '        '----------------------------------------------------------------------
    '        ' 3) 채널별 내부 Range 설정
    '        '----------------------------------------------------------------------
    '        Dim rangeVal As Double = 0.0
    '        Select Case channelCode
    '            Case "0910"
    '                rangeVal = 6.0
    '            Case "0940"
    '                rangeVal = 3.0
    '            Case "0950"
    '                rangeVal = 6.0
    '            Case "0960"
    '                rangeVal = 3.0
    '            Case "0980"
    '                rangeVal = 3.0
    '            Case "0990"
    '                rangeVal = 3.0
    '        End Select

    '        '----------------------------------------------------------------------
    '        ' 4) 실제 측정값 계산
    '        '    - 16비트로 들어온 ADC Code를 10비트로 복원하기 위해 64.0으로 나눈다.
    '        '    - Vin = (10비트 ADC코드 × rangeVal) / 1024
    '        '----------------------------------------------------------------------
    '        If rRes <> "" Then
    '            Dim raw10bitDouble As Double = rCount / 64.0  ' (16bit -> 10bit 복원: 64로 나눈다)
    '            mMeas = Math.Round((raw10bitDouble * rangeVal) / 1024.0, 3)
    '        End If

    '        '----------------------------------------------------------------------
    '        ' [참고] Low/High HexCode 각각도 동일 계산
    '        '       → "hexLowCode"와 "hexHighCode"를 별도 측정 기준(스펙 범위)으로 사용
    '        '----------------------------------------------------------------------
    '        Dim rawLow16 As Integer = Convert.ToInt32(hexLowCode, 16)   ' 16진수 -> 10진수
    '        Dim rawHigh16 As Integer = Convert.ToInt32(hexHighCode, 16) ' 16진수 -> 10진수

    '        ' 10비트 복원
    '        Dim adcLow10Double As Double = rawLow16 / 64.0
    '        Dim adcHigh10Double As Double = rawHigh16 / 64.0

    '        ' 각각 전압 계산 (소수점 셋째 자리까지)
    '        mMeasLow = Math.Round((adcLow10Double * rangeVal) / 1024.0, 3)
    '        mMeasHigh = Math.Round((adcHigh10Double * rangeVal) / 1024.0, 3)

    '        '----------------------------------------------------------------------
    '        ' 스펙 문자열
    '        '----------------------------------------------------------------------
    '        sSpec = $"{mMeasLow:0.000} V ≦ 測定直 ≦ {mMeasHigh:0.000} V"
    '        Call WriteLog(False, "SPEC : " & sSpec)
    '        Call WriteLog(False, $"測定直   : {mMeas:F3} V")

    '        ' ======================= 판정 및 예외 처리 ==========================
    '        If Me.ERR_COMMENT <> "" Then
    '            ' 별도 오류가 이미 기록된 경우
    '            mTestSpec = sSpec
    '            mTestMeasure = $"{voltageName} : {mMeas} V"
    '            mErrorComment = $"{voltageName} : {mMeas} V 異常"
    '            Call WriteLog(False, "判定基準 : " & sSpec)
    '            Call WriteValueData(voltageName, CStr(mMeas))
    '            Throw New ApplicationException(mErrorComment)
    '        Else
    '            ' 측정값이 스펙 범위를 벗어나는지 검사
    '            If (mMeas < mMeasLow) OrElse (mMeas > mMeasHigh) Then
    '                mErrorComment = $"{voltageName} : {mMeas} V 異常"
    '                Throw New ApplicationException(mErrorComment)
    '            End If

    '            ' 정상 범위 내
    '            Call WriteLog(False, $"{voltageName} : {mMeas} V 正常")
    '            Call WriteValueData(voltageName, CStr(mMeas) & vbCrLf)
    '        End If

    '    Catch ex As Exception
    '        If String.IsNullOrEmpty(mErrorComment) Then
    '            mErrorComment = "TSCommand Recv ERR"
    '        End If
    '        WriteLog(False, mErrorComment & " : " & ex.Message)
    '        Throw New ApplicationException(mErrorComment)
    '    End Try
    'End Sub


    Private Sub MeasureVoltage_I2C(ByVal cComm As cls101_ComRS232C)


        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rCount As Integer = 0
        Dim rData() As String = {""}
        Dim mMeas As Double = 0.0
        Dim sSpec As String = ""
        Dim mErrorComment As String = ""

        ' 변환 계수
        Dim convertFactor As Double = 0.00125

        Dim channelCode As String = "09B0"
        Dim hexLowLimit As String = "47E0"    ' 16진수(0x47e0=18400) => *0.00125=23.0
        Dim hexHighLimit As String = "4E20"   ' 16진수(0x4e20=20000) => *0.00125=25.0

        Try

            cComm.SendCmd = $"TS,{channelCode},2,{hexLowLimit},{hexHighLimit}"
            Call WriteLog(False, "Send : " & cComm.SendCmd)
            Call cComm.FuncSendCmd()

            mSecWait(500)
            Call cComm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cComm.RECEIVE_DATA)


            RPRecvCmd(cComm.RECEIVE_DATA, rRes, rCode, rCount, rData)


            ' rCount * 0.00125 = 측정 전압(V)
            If rRes <> "" Then
                mMeas = Math.Round(rCount * convertFactor, 3)
            End If

            Dim lowVolt As Double = Math.Round(Convert.ToInt32(hexLowLimit, 16) * convertFactor, 3)   ' 23.0
            Dim highVolt As Double = Math.Round(Convert.ToInt32(hexHighLimit, 16) * convertFactor, 3) ' 25.0

            sSpec = $"{lowVolt:0.000} V ≦ 測定直 ≦ {highVolt:0.000} V"

            Call WriteLog(False, "SPEC      : " & sSpec)
            Call WriteLog(False, "測定値    : " & mMeas & " V")


            If Me.ERR_COMMENT <> "" Then
                mTestSpec = sSpec
                mTestMeasure = $"24V_I2C : {mMeas} V"
                mErrorComment = $"24V_I2C : {mMeas} V 異常"

                Call WriteLog(False, "判定基準 : " & sSpec)
                Call WriteValueData("24V_I2C", CStr(mMeas))
                Throw New ApplicationException(mErrorComment)

            Else
                If (mMeas < lowVolt) OrElse (mMeas > highVolt) Then
                    mErrorComment = $"測定直 {mMeas} V (範囲外)"
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, $"測定直 {mMeas} V (範囲内)")
                Call WriteValueData("24V_I2C", $"{mMeas} V" & vbCrLf)
            End If

        Catch ex As Exception
            If String.IsNullOrEmpty(mErrorComment) Then
                mErrorComment = "TSCommand Recv ERR"
            End If
            Call WriteLog(False, mErrorComment & " : " & ex.Message)

            Throw New ApplicationException(mErrorComment, ex)
        End Try

    End Sub
    Private Function ConvertToMacFormat(ByVal rawMac As String) As String
        If rawMac.Length <> 12 Then
            Return rawMac
        End If

        Return rawMac.Substring(0, 2) & ":" &
               rawMac.Substring(2, 2) & ":" &
               rawMac.Substring(4, 2) & ":" &
               rawMac.Substring(6, 2) & ":" &
               rawMac.Substring(8, 2) & ":" &
               rawMac.Substring(10, 2)
    End Function
    'Private Sub MeasureVoltage_Sysmon(ByVal cComm As cls101_ComRS232C,
    '                              ByVal voltageName As String,
    '                              ByVal channelCode As String,
    '                              ByVal hexLimitHigh As String,
    '                              ByVal hexLimitLow As String)

    '    Dim rRes As String = ""    ' 수신결과
    '    Dim rCode As String = ""   ' 응답 코드
    '    Dim rCount As Integer = 0  ' 실제 ADC 카운트값(0~1023 기대)
    '    Dim rData() As String = {""}
    '    Dim mMeas As Double = 0.0  ' 최종 측정 전압
    '    Dim rangeVolt As Double = 3.0  ' 기본값 3V

    '    Try
    '        '────────────────────────────────────────────────────────
    '        ' 1) 로그 출력 (항목명 등)
    '        '────────────────────────────────────────────────────────
    '        WriteLog(False, "")
    '        WriteLog(False, $"{voltageName} [Sysmon] 測定 =========")

    '        '────────────────────────────────────────────────────────
    '        ' 2) A8에 TS 명령 전송
    '        '    (Low/High Limit를 같이 보내지만, 여기서는
    '        '     rCount 1개만 읽어 전압 변환한다고 가정)
    '        '────────────────────────────────────────────────────────
    '        cComm.SendCmd = $"TS,{channelCode},2,{hexLimitLow},{hexLimitHigh}"
    '        WriteLog(False, "Send : " & cComm.SendCmd)
    '        cComm.FuncSendCmd()

    '        '────────────────────────────────────────────────────────
    '        ' 3) 응답 수신 + 파싱
    '        '    RPRecvCmd 내부에서 rRes, rCode, rCount, rData를 채움
    '        '────────────────────────────────────────────────────────
    '        cComm.FuncRecvCmd()
    '        WriteLog(False, "Recv : " & cComm.RECEIVE_DATA)
    '        RPRecvCmd(cComm.RECEIVE_DATA, rRes, rCode, rCount, rData)

    '        '────────────────────────────────────────────────────────
    '        ' 4) 주소(채널)별로 측정 범위(3V 또는 6V) 결정
    '        '    - 예: 3.3V(0910)은 내부에서 6V 스케일로 측정
    '        '────────────────────────────────────────────────────────
    '        Select Case channelCode
    '            Case "0910"  ' 3.3V
    '                rangeVolt = 6.0
    '            Case "0940"  ' 1.8V
    '                rangeVolt = 3.0
    '            Case "0950"  ' 1.8V, 다른 설정이면 6.0
    '                rangeVolt = 6.0
    '            Case "0960"  ' 1.2V
    '                rangeVolt = 3.0
    '            Case "0980"  ' 0.85V
    '                rangeVolt = 3.0
    '            Case "0990"  ' 0.85V
    '                rangeVolt = 3.0
    '            Case Else
    '                rangeVolt = 3.0
    '        End Select

    '        '────────────────────────────────────────────────────────
    '        ' 5) 실제 전압 계산
    '        '    If rRes <> "" Then ... (질문에 주신 형태)
    '        '────────────────────────────────────────────────────────
    '        If rRes <> "" Then
    '            ' rCount가 0~1023 범위라 가정.
    '            ' 전압 V = (rCount * rangeVolt) / 1024
    '            mMeas = Math.Round((rCount * rangeVolt) / 1024.0, 3)
    '        End If

    '        '────────────────────────────────────────────────────────
    '        ' 6) 로그 출력 및 판정
    '        '────────────────────────────────────────────────────────
    '        WriteLog(False, $"Range : {rangeVolt}V")
    '        WriteLog(False, $"測定ADC : {rCount}  (0~1023 기대)")
    '        WriteLog(False, $"測定電圧 : {mMeas} V")

    '        ' 간단 예시로 스펙 체크
    '        '   - Low/High Limit(0xNNN)도 10비트 값이라 가정
    '        Dim lowDec As Integer = Convert.ToInt32(hexLimitLow, 16)
    '        Dim highDec As Integer = Convert.ToInt32(hexLimitHigh, 16)
    '        Dim lowVolt As Double = Math.Round((lowDec * rangeVolt) / 1024.0, 3)
    '        Dim highVolt As Double = Math.Round((highDec * rangeVolt) / 1024.0, 3)

    '        Dim specMsg As String = $"[{lowVolt} V ~ {highVolt} V]"
    '        WriteLog(False, $"SPEC : {specMsg}")

    '        mMeasLow = Math.Round((adcLow10 * rangeVal) / 1024.0, 3)

    '        If (mMeas < lowVolt) OrElse (mMeas > highVolt) Then
    '            mErrorComment = $"{voltageName}: {mMeas} V 異常"
    '            WriteLog(False, mErrorComment)
    '            Throw New ApplicationException(mErrorComment)
    '        Else
    '            WriteLog(False, $"{voltageName}: {mMeas} V OK")
    '        End If

    '    Catch ex As Exception
    '        If String.IsNullOrEmpty(mErrorComment) Then
    '            mErrorComment = "TSCommand Recv ERR"
    '        End If
    '        WriteLog(False, mErrorComment & " : " & ex.Message)
    '        Throw New ApplicationException(mErrorComment)
    '    End Try
    'End Sub

End Class