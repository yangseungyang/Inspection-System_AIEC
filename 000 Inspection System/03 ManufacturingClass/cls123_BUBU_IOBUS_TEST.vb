Option Strict On : Option Explicit On


'Doevents
Imports System.Windows.Forms.Application

Public Class cls123_BUBU_IOBUS_TEST

    Dim cMcm As cls120_SP_M_Command
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl

    Dim mSlotAddress As Integer = 0
    Dim mLOG As String = ""
    Dim mErrorComment As String
    Dim mIPAdd As String = ""

    'IPAddress -----------------------------
    Public Property IPAdd() As String
        Set(ByVal Value As String)
            mIPAdd = Value
        End Set
        Get
            IPAdd = mIPAdd
        End Get
    End Property

    '検査ｽﾛｯﾄｱﾄﾞﾚｽ設定 ====================================
    Public WriteOnly Property TESTSlotAddress() As Integer
        Set(ByVal rp As Integer)
            mSlotAddress = rp
        End Set
    End Property

    'LOG Text ==============================
    Public ReadOnly Property LOG() As String
        Get
            LOG = mLOG
        End Get
    End Property

    'Error message ===========================
    Public ReadOnly Property ERR_MSG() As String
        Get
            ERR_MSG = mErrorComment
        End Get
    End Property

    Public Function TestPreparation() As Boolean

        'ﾒｲﾝCPU（ｽﾛｯﾄ1：SP7x）→BUBUをﾓﾆﾀする場合の前準備
        'BUBUがｱｸｾｽできるようIOﾊﾞｽﾒﾓﾘ空間にﾏｯﾋﾟﾝｸﾞしておく（？）
        'CPU検査時はｽﾛｯﾄ5が基本！

        TestPreparation = False

        Dim slot As Integer
        Dim WriteAddress As String = ""
        Dim WriteDATA As String = ""
        Dim mMessage As String = ""

        If mSlotAddress = 0 Then
            slot = 5
        Else
            slot = mSlotAddress
        End If

        '検査ｽﾛｯﾄによって書き込むｱﾄﾞﾚｽが異なる為下記処理をする
        Dim Adres1 As String = Convert.ToString(CByte("&H78") + (slot - 1) * CByte("&H80"), 16)
        If Adres1.Length = 3 Then
            Adres1 = "00000" & Adres1
        ElseIf Adres1.Length = 2 Then
            Adres1 = "000000" & Adres1
        End If
        Dim Adres2 As String = Convert.ToString(CByte("&H1B") + (slot - 1) * CByte("&H80"), 16)
        If Adres2.Length = 3 Then
            Adres2 = "00000" & Adres2
        ElseIf Adres1.Length = 2 Then
            Adres2 = "000000" & Adres2
        End If
        Dim Adres3 As String = Convert.ToString(CByte("&H12") + (slot - 1) * CByte("&H80"), 16)
        If Adres3.Length = 3 Then
            Adres3 = "00000" & Adres3
        ElseIf Adres1.Length = 2 Then
            Adres3 = "000000" & Adres3
        End If
        Dim Adres4 As String = Convert.ToString(CByte("&H34") + (slot - 1) * CByte("&H80"), 16)
        If Adres4.Length = 3 Then
            Adres4 = "00000" & Adres4
        ElseIf Adres1.Length = 2 Then
            Adres4 = "000000" & Adres4
        End If

        '==========================================================================================
        Call WriteLog("IOﾊﾞｽﾃｽﾄ前準備")

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        Try
            For i = 0 To 2
                Select Case i
                    Case 0
                        WriteAddress = Adres1 : WriteDATA = "0000" : mMessage = "ﾊﾟﾗﾒｰﾀ終了"
                    Case 1
                        WriteAddress = Adres2 : WriteDATA = "0080" : mMessage = "ｲﾆｼｬﾗｲｽﾞ終了"
                    Case 2
                        WriteAddress = Adres3 : WriteDATA = "01D8" : mMessage = "ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定"
                End Select
                Call cMcm.IO_WRITE(1, 2, WriteAddress, WriteDATA)
                Call WriteLog(mMessage)
                Call WriteLog("ｱﾄﾞﾚｽ:" & WriteAddress & ",書き込みﾃﾞｰﾀ:" & WriteDATA)
            Next

            Call mSecWait(100)
            Call cMcm.IO_READ(0, 4, Adres4)               'BUBUｱﾄﾞﾚｽ0xDを読む
            Call WriteLog("Send : " & "BUBUｱﾄﾞﾚｽ0xD（ｽﾛｯﾄ" & slot.ToString & "：" & Adres4 & "）を読む")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "58" Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽ検査前準備ｴﾗｰ"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

        'cBUS = New cls122_FAM3_IOBUS_Ctrl
        'Call cBUS.IOBusTstReSetCpuXSFAIL()
        'cBUS = Nothing

        TestPreparation = True

    End Function

    Public Function TestPreparation_ForBicSlave() As Boolean

        'ﾒｲﾝCPU（ｽﾛｯﾄ1：SP7x）→BUBUをﾓﾆﾀする場合の前準備
        'BUBUがｱｸｾｽできるようIOﾊﾞｽﾒﾓﾘ空間にﾏｯﾋﾟﾝｸﾞしておく（？）
        'CPU検査時はｽﾛｯﾄ5が基本！
        Dim WriteAddress As String = ""
        Dim WriteDATA As String = ""
        Dim mMessage As String = ""

        TestPreparation_ForBicSlave = False

        Dim slot As Integer
        If mSlotAddress = 0 Then
            slot = 5
        Else
            slot = mSlotAddress
        End If

        '検査ｽﾛｯﾄによって書き込むｱﾄﾞﾚｽが異なる為下記処理をする
        Dim Adres1 As String = Convert.ToString(CByte("&H78") + (slot - 1) * CByte("&H80"), 16)
        If Adres1.Length = 3 Then
            Adres1 = "00000" & Adres1
        ElseIf Adres1.Length = 2 Then
            Adres1 = "000000" & Adres1
        End If
        Dim Adres2 As String = Convert.ToString(CByte("&H1B") + (slot - 1) * CByte("&H80"), 16)
        If Adres2.Length = 3 Then
            Adres2 = "00000" & Adres2
        ElseIf Adres1.Length = 2 Then
            Adres2 = "000000" & Adres2
        End If
        Dim Adres3 As String = Convert.ToString(CByte("&H12") + (slot - 1) * CByte("&H80"), 16)
        If Adres3.Length = 3 Then
            Adres3 = "00000" & Adres3
        ElseIf Adres1.Length = 2 Then
            Adres3 = "000000" & Adres3
        End If
        Dim Adres4 As String = Convert.ToString(CByte("&H34") + (slot - 1) * CByte("&H80"), 16)
        If Adres4.Length = 3 Then
            Adres4 = "00000" & Adres4
        ElseIf Adres1.Length = 2 Then
            Adres4 = "000000" & Adres4
        End If

        '==========================================================================================
        Call WriteLog("IOﾊﾞｽﾃｽﾄ前準備")

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        Try
            For i = 0 To 2
                Select Case i
                    Case 0
                        WriteAddress = Adres1 : WriteDATA = "0000" : mMessage = "ﾊﾟﾗﾒｰﾀ終了"
                    Case 1
                        WriteAddress = Adres2 : WriteDATA = "0080" : mMessage = "ｲﾆｼｬﾗｲｽﾞ終了"
                    Case 2
                        WriteAddress = Adres3 : WriteDATA = "0038" : mMessage = "ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定"
                End Select
                Call cMcm.IO_WRITE(1, 2, WriteAddress, WriteDATA)
                Call WriteLog(mMessage)
                Call WriteLog("ｱﾄﾞﾚｽ:" & WriteAddress & ",書き込みﾃﾞｰﾀ:" & WriteDATA)
            Next

            Call mSecWait(100)
            Call cMcm.IO_READ(0, 4, Adres4)               'BUBUｱﾄﾞﾚｽ0xDを読む
            Call WriteLog("Send : " & "BUBUｱﾄﾞﾚｽ0xD（ｽﾛｯﾄ" & slot.ToString & "：" & Adres4 & "）を読む")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "58" Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽ検査前準備ｴﾗｰ"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

        TestPreparation_ForBicSlave = True

    End Function

    Public Function SlotAddressTest() As Boolean

        'ｽﾛｯﾄ5の状態から検査をする
        'ｽﾛｯﾄ2,3,5,9の状態にして$9のIOﾊﾞｽﾚｼﾞｽﾀﾓﾆﾀでﾓﾆﾀできること

        SlotAddressTest = False

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀ（$200）確認（ｽﾛｯﾄ5）")
        Call cMcm.IO_READ(1, 2, "00000200")
        Call WriteLog(cMcm.RecvText)

        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
            mErrorComment = "XDFAILｽﾛｯﾄ5認識異常" : cMcm = Nothing
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        'SLOTを確認
        Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽ検査")

        cBUS = New cls122_FAM3_IOBUS_Ctrl
        Try
            For i As Integer = 2 To 9

                Call cBUS.BUSMONITOR_TB04_CPU_SLOTChange(i)
                Call WriteLog("SLOT" & CStr(i) & "ｱﾄﾞﾚｽ確認")
                Call mSecWait(100)
                Call cMcm.IO_READ(0, 4, "00000224")
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x224ﾘｰﾄﾞ（$9のｽﾛｯﾄ5）")
                Call WriteLog("Recv : " & cMcm.RecvText)
                If Strings.Right(cMcm.RecvText, 1) <> CStr(i - 1) Then
                    mErrorComment = "SLOTｱﾄﾞﾚｽ認識異常：SLOT" & i.ToString
                    Throw New ApplicationException(mErrorComment)
                End If
                Call cMcm.IO_READ(0, 4, "00000210")
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")
                Call WriteLog("Recv : " & cMcm.RecvText)
                If Strings.Right(cMcm.RecvText, 2) <> "F7" Then
                    mErrorComment = "SLOTｱﾄﾞﾚｽにその他ｼｮｰﾄの疑いあり！"
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog("SLOT" & i.ToString & " OK")

                If i = 3 Then i = i + 1 '次はｽﾛｯﾄ5を検査
                If i = 5 Then i = i + 3 '次はｽﾛｯﾄ9を検査

            Next

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.BUSMONITOR_TB04_CPU_SLOTChange(5)
            cBUS = Nothing
            cMcm = Nothing
        End Try

        cBUS = Nothing
        cMcm = Nothing

        SlotAddressTest = True

    End Function

    Public Function XACSTest() As Boolean
        'XACS検査(CPU)

        XACSTest = False

        Dim mintUnit As Integer = 0      'DUT検査ユニット番号
        Dim mintSlot As Integer = 6      'DUT検査スロット番号
        Dim mintXD6407 As Integer = 7    'BM128 ＤＩスロット番号
        Dim mintYD6408 As Integer = 8    'BM128 ＤＯスロット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl
        'IOバスアクセス
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd
        cMcm.RemPort = 12290                           'LE バイナリ形式時は 12290 にする。

        cMcm.SlotNo = 1        'CPU Slot番号
        cMcm.UnitNo = 0        'ユニット番号

        Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        'Y33->0　スルーする。XACSされない状態で
        'X33->0　XASCネゲートならＯＫ。
        'ラッチ解除
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 1)
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)

        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 0)
        Call mSecWait(1000)

        Dim startime As Double = System.Environment.TickCount
        Do
            Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)
            If (System.Environment.TickCount - startime) / 1000 > 1 Then
                Exit Do '1秒間XACSが0であること      v0.01
            End If
        Loop Until cBUS.RCV_DATA <> "0"

        If cBUS.RCV_DATA <> "0" Then
            cBUS = Nothing : cMcm = Nothing : Exit Function
        End If

        Call WriteLog("1秒間XACSﾈｹﾞｰﾄ確認OK")

        'Y33->1　ラッチする。
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 1)
        Call WriteLog("XACSﾓﾆﾀﾗｯﾁﾓｰﾄﾞｾｯﾄ")

        Call cMcm.IO_READ(0, 4, "00000228")      'ｽﾛｯﾄ5にMｺﾏﾝﾄﾞでIOﾊﾞｽｱｸｾｽ
        Call WriteLog("ｽﾛｯﾄ5にIOﾊﾞｽｱｸｾｽ実行")

        cMcm = Nothing

        'X33->1　XASCアサートならＯＫ。
        Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        If cBUS.RCV_DATA <> "1" Then
            cBUS = Nothing : cMcm = Nothing : Exit Function
        End If

        Call WriteLog("XACSｱｻｰﾄ確認OK")

        'Y33->0　スルーする(信号線が繋がっている場合アクセス後にクリア)
        '被検査モジュールに何らかのアクセス
        'X33->0　ならＯＫ。
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 0)
        Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        'ラッチ解除
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 1)
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)

        XACSTest = True
        cBUS = Nothing : cMcm = Nothing

    End Function

    Public Function XBRTest() As Boolean
        'ﾒｲﾝCPU（SP7x）にてXBR0～3を順次上げ下げして
        '$9のIOﾊﾞｽﾚｼﾞｽﾀにてXBRの変化をﾓﾆﾀできること

        XBRTest = False

        'XBRﾃｽﾄ
        Call WriteLog("XBR検査")
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                     'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Try
            For i As Integer = 3 To 0 Step -1
                Call WriteLog("XBR" & CStr(i) & "ｾｯﾄ")

                Select Case i
                    Case 3 : Call cMcm.MEM_WRITE(1, "3880070B", "08")
                    Case 2 : Call cMcm.MEM_WRITE(1, "3880070B", "04")
                    Case 1 : Call cMcm.MEM_WRITE(1, "3880070B", "02")
                    Case 0 : Call cMcm.MEM_WRITE(1, "3880070B", "01")
                End Select
                Call mSecWait(100)

                Call WriteLog("XBR" & CStr(i) & "確認")
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x224ﾘｰﾄﾞ（$9のｽﾛｯﾄ5）")
                Call cMcm.IO_READ(0, 4, "00000224")
                Call WriteLog("Recv : " & cMcm.RecvText)
                Select Case Mid(Strings.Right(cMcm.RecvText, 2), 1, 1)
                    Case "7"
                        If i <> 3 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case "B"
                        If i <> 2 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case "D"
                        If i <> 1 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case "E"
                        If i <> 0 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case Else
                        mErrorComment = "XBR認識異常：XBR"
                        Throw New ApplicationException(mErrorComment)
                End Select

                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")
                Call cMcm.IO_READ(0, 4, "00000210")
                Call WriteLog("Recv : " & cMcm.RecvText)
                If Strings.Right(cMcm.RecvText, 2) <> "F7" Then
                    mErrorComment = "XBR認識異常：XBR" & i.ToString
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog("XBR" & i.ToString & " OK")

            Next i

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cMcm.MEM_WRITE(1, "3880070B", "08")
            'Call cBUS.BUSMONITOR_TB04_CPU_XBRChange(5)
            cBUS = Nothing
            cMcm = Nothing
        End Try

        cBUS = Nothing
        cMcm = Nothing

        XBRTest = True

    End Function

    Public Function XDFAILTest() As Boolean
        'XDFAIL検査
        'ｽﾛｯﾄｱﾄﾞﾚｽを5→3に変更してXDFAILをｱｻｰﾄしてｽﾛｯﾄｱﾄﾞﾚｽが3に変更されていることを確認する

        XDFAILTest = False

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        '現在のｽﾛｯﾄｱﾄﾞﾚｽをread
        Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀ（$200）確認（ｽﾛｯﾄ5：0x200）")

        Call cMcm.IO_READ(1, 2, "00000200")
        Call WriteLog(cMcm.RecvText)

        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
            mErrorComment = "XDFAILｽﾛｯﾄ5認識異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀ（$200）確認（ｽﾛｯﾄ5）OK")

        Try

            Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽ3に変更")
            Call cBUS.BUSMONITOR_TB04_CPU_SLOTChange(3)
            Call cBUS.IOBusTstSetReSetCpuXDFAIL()

            Call WaitTimer(7)

            Call WriteLog("IOﾊﾞｽ検査ｽﾛｯﾄｱﾄﾞﾚｽ3前準備")
            Call cMcm.IO_WRITE(1, 2, "00000178", "0000")      'IOREGｱﾄﾞﾚｽ0x78に0x0000をﾗｲﾄ
            Call WriteLog("Send : " & "IOREGｱﾄﾞﾚｽ0x78（ｽﾛｯﾄ3：0x178）に0x0000をﾗｲﾄ")
            Call WriteLog("Recv : " & cMcm.RecvText)
            Call cMcm.IO_WRITE(1, 2, "0000011A", "0080")      'IOREGｱﾄﾞﾚｽ0x1Bに0x80をﾗｲﾄ
            Call WriteLog("Send : " & "IOREGｱﾄﾞﾚｽ0x1B（ｽﾛｯﾄ3：0x11A）に0x80をﾗｲﾄ")
            Call WriteLog("Recv : " & cMcm.RecvText)
            Call cMcm.IO_WRITE(1, 2, "00000112", "01D8")      'IOREGｱﾄﾞﾚｽ0x12をBUBUが見える状態(0x01D8)にする
            Call WriteLog("Send : " & "IOREGｱﾄﾞﾚｽ0x12（ｽﾛｯﾄ3：0x112）をBUBUが見える状態(0x01D8)にする")
            Call WriteLog("Recv : " & cMcm.RecvText)
            Call cMcm.IO_READ(0, 4, "00000134")               'BUBUｱﾄﾞﾚｽ0xDを読む
            Call WriteLog("Send : " & "BUBUｱﾄﾞﾚｽ0xD（ｽﾛｯﾄ3：0x134）を読む")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "58" Then
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀ（$200）確認（ｽﾛｯﾄ3：0x100）")

            Call cMcm.IO_READ(1, 2, "00000100")
            Call WriteLog(cMcm.RecvText)

            If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "20" Then
                mErrorComment = "XDFAILｽﾛｯﾄ3認識異常"
                Call WriteLog(mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XDFAIL異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXSFAIL()                  'XSFAILがｾｯﾄされるのでﾘｾｯﾄをする
            cBUS = Nothing
            cMcm = Nothing
        End Try

        XDFAILTest = True

    End Function

    Public Function XIRQTest() As Boolean

        'ﾒｲﾝCPU（SP7x）のXIRQをﾏｽｸしてﾊﾞｽﾓﾆﾀ側からXIRQ0～3を上げ下げして
        '$4のIOﾊﾞｽﾚｼﾞｽﾀにてXIRQの変化がﾓﾆﾀできること
        XIRQTest = False

        Dim mResStr As String = ""

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'ﾒｲﾝCPUのXIRQﾏｽｸ
        Call WriteLog("ﾒｲﾝCPUのXIRQﾏｽｸ")
        Call cMcm.MEM_WRITE(1, "3880021E", "00")

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Try
            For i As Integer = 3 To 0 Step -1
                Call cBUS.BUSMONITOR_TB04_CPU_XIRQChange(i)
                Call WriteLog("XIRQ" & CStr(i) & "確認")
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")      '"ｱﾄﾞﾚｽ0x90ﾘｰﾄﾞ（$4のｽﾛｯﾄ2）"
                Call cMcm.IO_READ(0, 4, "00000210")
                Call WriteLog("Recv : " & cMcm.RecvText)

                Select Case i
                    Case 0 : mResStr = "7"
                    Case 1 : mResStr = "B"
                    Case 2 : mResStr = "D"
                    Case 3 : mResStr = "E"
                End Select

                If Mid(Strings.Right(cMcm.RecvText, 2), 1, 1) <> mResStr Then
                    mErrorComment = "XIRQ異常" & i.ToString
                    Throw New ApplicationException(mErrorComment)
                End If

                Call cMcm.IO_READ(0, 4, "00000224")
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x224ﾘｰﾄﾞ（$9のｽﾛｯﾄ5）")      '"ｱﾄﾞﾚｽ0xA4ﾘｰﾄﾞ（$9のｽﾛｯﾄ2）")  
                Call WriteLog("Recv : " & cMcm.RecvText)
                If Strings.Right(cMcm.RecvText, 2) <> "74" Then
                    mErrorComment = "XIRQその他とｼｮｰﾄの恐れあり！"
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog("XIRQ" & i.ToString & " OK")

            Next
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XIRQ異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.BUSMONITOR_TB04_CPU_XIRQChange(5)
            'ﾒｲﾝCPUのXIRQﾏｽｸ解除
            Call WriteLog("ﾒｲﾝCPUのXIRQﾏｽｸ解除")
            Call cMcm.MEM_WRITE(1, "3880021E", "88")
            cBUS = Nothing
            cMcm = Nothing
        End Try

        XIRQTest = True

    End Function

    Public Function XLFAILTest() As Boolean

        XLFAILTest = False

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call cMcm.IO_READ(0, 4, "00000210")
        Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")
        Call WriteLog("Recv : " & cMcm.RecvText)
        If Strings.Right(cMcm.RecvText, 1) <> "7" Then
            mErrorComment = "XLFAIL異常"
            Throw New ApplicationException(mErrorComment)
        End If

        'XLFAILｾｯﾄ
        If cBUS.IOBusTstSetCpuXLFAIL() = False Then
            Call cBUS.IOBusTstReSetCpuXLFAIL()
            mErrorComment = "XLFAILｾｯﾄ異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog("XLFAILｾｯﾄ")

        Try
            Call cMcm.IO_READ(0, 4, "00000210")
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 1) <> "6" Then
                mErrorComment = "XLFAIL認識異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call cMcm.IO_READ(0, 4, "00000224")
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x224ﾘｰﾄﾞ（$9のｽﾛｯﾄ5）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "74" Then
                mErrorComment = "XLFAILその他とｼｮｰﾄの恐れあり！"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXLFAIL()
            Call WriteLog("XLFAILﾘｾｯﾄOK")
            cMcm = Nothing
            cBUS = Nothing
        End Try

        Call WriteLog("XLFAIL：OK")
        XLFAILTest = True

    End Function

    Public Function XPFAILTest() As Boolean

        XPFAILTest = False

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call WriteLog("ﾒｲﾝCPUのXPFAILﾏｽｸ")
        '事前準備1
        Call cMcm.MEM_WRITE(1, "38800013", "00")
        '事前準備2
        Call cMcm.MEM_WRITE(1, "38800011", "00")

        Try
            Call cMcm.IO_READ(0, 4, "00000210")
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")      '"ｱﾄﾞﾚｽ0x90ﾘｰﾄﾞ（$4のｽﾛｯﾄ2）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 1) <> "7" Then
                mErrorComment = "XPFAIL異常"
                Throw New ApplicationException(mErrorComment)
            End If

            'XPFAILｾｯﾄ
            If cBUS.IOBusTstSetCpuXPFAIL() = False Then
                Call cBUS.IOBusTstReSetCpuXPFAIL()
                mErrorComment = "XPFAILｾｯﾄ異常"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Call cMcm.MEM_WRITE(1, "38800011", "80")
            cMcm = Nothing
            cBUS = Nothing
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog("XPFAILｾｯﾄ")

        Try
            Call cMcm.IO_READ(0, 4, "00000210")
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 1) <> "5" Then
                mErrorComment = "XPFAIL認識異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call cMcm.IO_READ(0, 4, "00000224")
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x224ﾘｰﾄﾞ（$9のｽﾛｯﾄ5）")      '"ｱﾄﾞﾚｽ0xA4ﾘｰﾄﾞ（$9のｽﾛｯﾄ2）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "74" Then
                mErrorComment = "XPFAILその他とｼｮｰﾄの恐れあり！"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XPFAIL認識異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXPFAIL()
            Call WriteLog("XPFAILﾘｾｯﾄOK")
            Call WriteLog("ﾒｲﾝCPUのXPFAILﾏｽｸ解除")
            Call cMcm.MEM_WRITE(1, "38800011", "80")
            cMcm = Nothing
            cBUS = Nothing
        End Try

        XPFAILTest = True

    End Function

    Public Function XSFAILTest() As Boolean

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                     'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call cBUS.IOBusTstReSetCpuXSFAIL()

        Call cMcm.IO_READ(0, 4, "00000210")
        Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")      '"ｱﾄﾞﾚｽ0x90ﾘｰﾄﾞ（$4のｽﾛｯﾄ2）")
        Call WriteLog("Recv : " & cMcm.RecvText)
        If Strings.Right(cMcm.RecvText, 1) <> "7" Then
            mErrorComment = "XSFAIL異常"
            Throw New ApplicationException(mErrorComment)
        End If

        'XSFAILｾｯﾄ
        If cBUS.IOBusTstSetCpuXSFAIL() = False Then
            Call cBUS.IOBusTstReSetCpuXSFAIL()
            mErrorComment = "XSFAILｾｯﾄ異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog("XSFAILｾｯﾄ")

        Try
            Call cMcm.IO_READ(0, 4, "00000210")
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x210ﾘｰﾄﾞ（$4のｽﾛｯﾄ5）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 1) <> "3" Then
                mErrorComment = "XSFAIL認識異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call cMcm.IO_READ(0, 4, "00000224")
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ0x224ﾘｰﾄﾞ（$9のｽﾛｯﾄ5）")      '"ｱﾄﾞﾚｽ0xA4ﾘｰﾄﾞ（$9のｽﾛｯﾄ2）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "74" Then
                mErrorComment = "XSFAILその他とｼｮｰﾄの恐れあり！"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXSFAIL()
            Call WriteLog("XSFAILﾘｾｯﾄOK")
            cMcm = Nothing
            cBUS = Nothing
        End Try

        XSFAILTest = True

    End Function

    Private Sub WriteLog(ByVal mComment As String)
        mLOG &= mComment & vbCrLf
    End Sub

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
    End Sub

    Private Sub WaitTimer(ByVal tm As Double)
        '==========================================================================================
        ' Function : Wait timer
        ' Data     : tm : Wait time(Sec)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim sTime As Double
        Dim nTime As Double
        Dim eTime As Double

        sTime = Microsoft.VisualBasic.DateAndTime.Timer()

        Do While tm > eTime
            nTime = Microsoft.VisualBasic.DateAndTime.Timer()
            eTime = nTime - sTime
            If eTime < 0 Then eTime = 24.0# * 3600.0# - sTime + nTime
            DoEvents()
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

End Class
