'======================================================================================================================================
'
'タイトル:FAM3_IO_PFM_CLS
'説明    :IOプラットフォーム対応したFA-M3のIOモジュールのROM書き込みをするクラス
'         FAM3_IO_PF_CLS.vbに存在しているクラスはFAM3_IO_PFM_CLSとBinFileClsの二つ
'         BinFileClsはバイナリファイル管理用のクラスで読み出したファイルの値を保持する。
'         BinFileClsはSwap、SUM計算、2KByteバウンダリパディング機能をもっている。
'         FAM3_IO_PFM_CLSからBinFileClsが呼ばれる。
'更新履歴:RevUp時は _iIoPfmClsRevを必ず更新すること。この更新履歴ではFAM3_IO_PFM_CLS.vbのRevのみ管理する。
'         Rev0 2019.07.22 T.maenishi
'              新規
'
'
'======================================================================================================================================


Imports System.Runtime.InteropServices  '共用体用

Public Class BinFileCls
    '==================================================================
    'フィールド
    '==================================================================
    '----Long Word Size 共用体---
    <StructLayout(LayoutKind.Explicit)> Private Structure LongWordUnion
        <FieldOffset(0)> Public s32_0 As Int32
        <FieldOffset(0)> Public u32_0 As UInt32
        <FieldOffset(0)> Public s16_0 As Int16
        <FieldOffset(2)> Public s16_1 As Int16
        <FieldOffset(0)> Public u16_0 As UInt16
        <FieldOffset(2)> Public u16_1 As UInt16
        <FieldOffset(0)> Public s8_0 As SByte
        <FieldOffset(1)> Public s8_1 As SByte
        <FieldOffset(2)> Public s8_2 As SByte
        <FieldOffset(3)> Public s8_3 As SByte
        <FieldOffset(0)> Public u8_0 As Byte
        <FieldOffset(1)> Public u8_1 As Byte
        <FieldOffset(2)> Public u8_2 As Byte
        <FieldOffset(3)> Public u8_3 As Byte
    End Structure


    '********外部に公開しない変数など********
    '読みだしたバイナリファイル保持
    Private _ByteFile() As Byte
    Private _Err As Integer

    '==================================================================
    'プロパティ
    '==================================================================
    '******************************************************************
    'プロパティ：ByteFile
    '説明：読みだしたバイナリファイルをByte型配列で返す
    '******************************************************************
    Public ReadOnly Property ByteFile() As Byte()
        Get
            Return _ByteFile
        End Get
    End Property

    '******************************************************************
    'プロパティ
    'ByteSwapFile
    '説明：読みだしたバイナリファイルをSwapしてByte型配列で返す
    '******************************************************************
    Public ReadOnly Property ByteSwapFile() As Byte()
        Get
            '****変数宣言****
            Dim i, j As Integer
            Dim _ByteSwapFile() As Byte

            '---配列要素数変更
            ReDim _ByteSwapFile(_ByteFile.Length - 1)

            '---ByteSwapFileへSwapコピー
            For i = 4 To _ByteFile.Length Step 4
                For j = 0 To 3 Step 1
                    _ByteSwapFile(i - 4 + j) = _ByteFile(i - 1 - j)
                Next
            Next

            Return _ByteSwapFile
        End Get
    End Property

    '******************************************************************
    'プロパティ
    'WordFile
    '説明：読みだしたバイナリファイルをint16型配列で返す
    '******************************************************************
    Public ReadOnly Property WordFile() As Int16()
        Get
            '****変数宣言****
            Dim iDivRem As Integer
            Dim iNum As Integer
            Dim _WordFile() As Int16

            '---配列要素数変更
            iDivRem = _ByteFile.Length Mod 2I
            iNum = CInt(_ByteFile.Length \ 2I)
            If iDivRem <> 0 Then iNum = iNum + 1
            ReDim _WordFile(iNum - 1)

            '---_ByteFileをLwFileへBinaryコピー(Byte型からUint32への型変換)
            Buffer.BlockCopy(_ByteFile, 0, _WordFile, 0, _ByteFile.Length)
            Return _WordFile
        End Get
    End Property

    '******************************************************************
    'プロパティ
    'WordSwapFile
    '説明：'読みだしたバイナリファイルをSwapしてint16型配列で返す
    '******************************************************************
    Public ReadOnly Property WordSwapFile() As Int16()
        Get
            '****変数宣言****
            Dim iDivRem As Integer
            Dim iNum As Integer
            Dim _WordSwapFile() As Int16

            '---配列要素数変更
            iDivRem = ByteSwapFile.Length Mod 2I
            iNum = CInt(ByteSwapFile.Length \ 2I)
            If iDivRem <> 0 Then iNum = iNum + 1
            ReDim _WordSwapFile(iNum - 1)

            '---_ByteFileをLwFileへBinaryコピー(Byte型からUint32への型変換)
            Buffer.BlockCopy(ByteSwapFile, 0, _WordSwapFile, 0, ByteSwapFile.Length)
            Return _WordSwapFile
        End Get
    End Property

    '******************************************************************
    'プロパティ
    'LwFile
    '説明：読みだしたバイナリファイルをint32型配列で返す
    '******************************************************************
    Public ReadOnly Property LwFile() As Int32()
        Get
            '****変数宣言****
            Dim iDivRem As Integer
            Dim iNum As Integer
            Dim _LwFile() As Int32

            '---配列要素数変更
            iDivRem = _ByteFile.Length Mod 4I
            iNum = CInt(_ByteFile.Length \ 4I)
            If iDivRem <> 0 Then iNum = iNum + 1
            ReDim _LwFile(iNum - 1)

            '---_ByteFileをLwFileへBinaryコピー(Byte型からUint32への型変換)
            Buffer.BlockCopy(_ByteFile, 0, _LwFile, 0, _ByteFile.Length)
            Return _LwFile
        End Get
    End Property

    '******************************************************************
    'プロパティ
    'LwSwapFile
    '説明：'読みだしたバイナリファイルをSwapしてint32型配列で返す
    '******************************************************************
    Public ReadOnly Property LwSwapFile() As Int32()
        Get
            '****変数宣言****
            Dim iDivRem As Integer
            Dim iNum As Integer
            Dim _LwSwapFile() As Int32

            '---配列要素数変更
            iDivRem = ByteSwapFile.Length Mod 4I
            iNum = CInt(ByteSwapFile.Length \ 4I)
            If iDivRem <> 0 Then iNum = iNum + 1
            ReDim _LwSwapFile(iNum - 1)

            '---ByteSwapFileを_LwSwapFileへBinaryコピー(Byte型からUint32への型変換)
            Buffer.BlockCopy(ByteSwapFile, 0, _LwSwapFile, 0, ByteSwapFile.Length)
            Return _LwSwapFile
        End Get
    End Property

    '******************************************************************
    'プロパティ
    'SUM
    '説明：読みだしたバイナリファイルのSUM値を計算して返す
    '******************************************************************
    Public ReadOnly Property Sum() As Integer
        Get
            '****変数宣言****
            Dim i As Integer
            Dim _Sum As LongWordUnion

            '---SUM値の計算
            _Sum.s32_0 = 0
            For i = 0 To _ByteFile.Length - 1
                _Sum.s32_0 = _Sum.s32_0 + CInt(_ByteFile(i))
            Next

            Return _Sum.s8_0
        End Get
    End Property

    '******************************************************************
    'プロパティ
    'U16_SUM
    '説明：読みだしたバイナリファイルのSUM値を計算して返す
    '      符号なし16bit
    '******************************************************************
    Public ReadOnly Property u16_Sum() As Integer
        Get
            '****変数宣言****
            Dim i As Integer
            Dim _Sum As LongWordUnion

            '---SUM値の計算
            _Sum.u32_0 = 0
            For i = 0 To _ByteFile.Length - 1
                _Sum.u32_0 = _Sum.u32_0 + CUInt(_ByteFile(i))
            Next

            Return _Sum.u16_0
        End Get
    End Property


    '******************************************************************
    'プロパティ: Err
    '返り値： Err = 0  ファイル読み出し成功
    '         Err = 1  ファイル読み出し失敗
    '******************************************************************
    Public ReadOnly Property Err() As Integer
        Get
            Return _Err
        End Get
    End Property

    '==================================================================
    'コンストラクタ
    '==================================================================
    Public Sub New(ByVal strFileName As String)
        '****変数宣言****

        '---変数初期化---
        _Err = 0

        '---ファイルを開く---
        Try
            Dim fs As New System.IO.FileStream(strFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)


            '---配列要素数変更
            ReDim _ByteFile(CInt(fs.Length - 1))

            '---ファイルの内容をすべて読み込む---
            fs.Read(_ByteFile, 0, _ByteFile.Length)
            fs.Close()        '閉じる

        Catch ex As Exception
            _Err = 1
        End Try
    End Sub


    '==================================================================
    'メソッド：Set2KByteBoundary
    '説明：_ByteFileの配列の要素数を2Kの倍数に切り上げる
    '==================================================================
    Public Sub Set2KByteBoundary()
        '****変数宣言****
        Dim iNum As Integer
        Dim iDivRem As Integer
        Dim bBuf() As Byte
        Dim iLen As Integer

        '---一時保存---
        iLen = ByteFile.Length
        ReDim bBuf(iLen - 1)
        Buffer.BlockCopy(ByteFile, 0, bBuf, 0, iLen)

        '---配列要素数を2Kbyteの倍数に切り上げる
        iDivRem = iLen Mod 2048
        iNum = iLen \ 2048
        If iDivRem <> 0 Then iNum = iNum + 1
        iNum = iNum * 2048
        ReDim _ByteFile(iNum - 1)
        For i = 0 To iNum - 1 '0クリア
            _ByteFile(i) = 0
        Next

        '---一時保存から戻す---
        Buffer.BlockCopy(bBuf, 0, _ByteFile, 0, iLen)


    End Sub

End Class


Public Class FAM3_IO_PFM_CLS
    '===================================================================================================================================
    '=
    '=共用体定義
    '=
    '===================================================================================================================================
    <StructLayout(LayoutKind.Explicit)> Private Structure LongWordUnion
        <FieldOffset(0)> Public s32_0 As Integer
        <FieldOffset(0)> Public u32_0 As UInteger
        <FieldOffset(0)> Public s16_0 As Short
        <FieldOffset(2)> Public s16_1 As Short
        <FieldOffset(0)> Public u16_0 As UShort
        <FieldOffset(2)> Public u16_1 As UShort
        <FieldOffset(0)> Public u8_0 As Byte
        <FieldOffset(1)> Public u8_1 As Byte
        <FieldOffset(2)> Public u8_2 As Byte
        <FieldOffset(3)> Public u8_3 As Byte
        <FieldOffset(0)> Public s8_0 As SByte
        <FieldOffset(1)> Public s8_1 As SByte
        <FieldOffset(2)> Public s8_2 As SByte
        <FieldOffset(3)> Public s8_3 As SByte
    End Structure
    <StructLayout(LayoutKind.Explicit)> Private Structure WordUnion
        <FieldOffset(0)> Public s16_0 As Short
        <FieldOffset(0)> Public u16_0 As UShort
        <FieldOffset(0)> Public u8_0 As Byte
        <FieldOffset(1)> Public u8_1 As Byte
        <FieldOffset(0)> Public s8_0 As SByte
        <FieldOffset(1)> Public s8_1 As SByte
    End Structure

    '===================================================================================================================================
    '=
    '=定数
    '=
    '===================================================================================================================================
    'FAM3_IO_PFM_CLS Revision
    Private Const _iIoPfmClsRev As Integer = 2

    '===BicSlvレジスタ
    '---モディファイドモジュールアドレスレジスタ関連定数
    '                                                   先頭	740000
    'レジスタ領域	モディファイ[hex]	data[hex]	    モディファイ領域[hex]
    'Unit0	slot1	7C000E	            8017			-	     -
    'Unit0	slot2	7C008E	            8217			760800	760FFF
    'Unit0	slot3	7C010E	            8417			761000	7617FF
    'Unit0	slot4	7C018E	            8617			761800	761FFF
    'Unit0	slot5	7C020E	            8817			762000	7627FF
    '
    '
    'Unit0	slot15	7C070E	            9C17			767000	7677FF
    'Unit0	slot16	7C078E	            9E17			767800	767FFF
    Private Const I_MODIFY_MDL_ADDR_REG_BASE As Integer = (&H7C000EI)       'モディファイドモジュールアドレスREGの先頭番地
    Private Const I_MODIFY_MDL_ADDR_REG_OFFSET As Integer = (&H80I)         'モディファイドモジュールアドレスREGのオフセットアドレス 0x80×(Slot番号-1)
    Private Const US_MODIFY_MDL_ADDR_REG_DATA As UShort = (&H8017US)        'モディファイドモジュールアドレスREGのデータ  A17=1, アドレス設定終了, サイズ=2KByte
    Private Const US_MODIFY_MDL_ADDR_REG_DATA_OFFSET As UShort = (&H200US)  'モディファイドモジュールアドレスREGのデータオフセット  0x200×(Slot番号-1)

    '---ウインドウアドレスレジスタ関連定数
    Private Const I_WINDOW_ADDR_REG_BASE As Integer = (&H7C0010I)       'ウインドウアドレスレジスタ先頭番地
    Private Const I_WINDOW_ADDR_REG_OFFSET As Integer = (&H80I)         'ウインドウアドレスレジスタのオフセットアドレス 0x80×(Slot番号-1)

    '---モディファイメモリ領域関連定数
    Private Const I_MODIFY_MEM_ADDR_BASE As Integer = (&H760000I)       'モディファイメモリ領域先頭番地
    Private Const I_MODIFY_MEM_ADDR_OFFSET As Integer = (&H800I)        'モディファイメモリ領域のオフセットアドレス 0x800×(Slot番号-1)


    '===================================================================================================================================
    'ComCtrlEx関連のクラス、変数
    '===================================================================================================================================
    Private _oComMng As COMCTRLEXLib.ComMng
    Private _oInfo As COMCTRLEXLib.F3OpenInfo2
    Private _sPort As Short
    Private _sSync As Short

    '===================================================================================================================================
    'FPGAデータ、FirmwareData
    '===================================================================================================================================
    'Bin File Class
    Private AppImageRPD As BinFileCls        'CFM1/2に書き込みたいRPDファイル
    Private FactImageNiosFirm As BinFileCls  'Factory ImageのNiosIIを動かすためのFirmware
    Private MCUFlashWrFirm As BinFileCls     'MCU内蔵FLASHを書き込むためのFirmware (MCUの内蔵RAMに配置する)
    Private MCUFirm As BinFileCls            'MCU内蔵FLASHに書き込みたいFirmware


    '=========================================================================================
    '=
    '=初期化・初期設定をここに記載する
    '=
    '=========================================================================================
    '******************************************************************************
    '*コンストラクタ：New
    '*説明： 使用する変数・クラスの初期化
    '******************************************************************************
    Public Sub New()
        Me._oComMng = New COMCTRLEXLib.ComMng
        Me._oInfo = New COMCTRLEXLib.F3OpenInfo2
        Me._sPort = 0
        Me._sSync = 1

    End Sub


    '=========================================================================================
    '=
    '=FA-M3との通信設定等
    '=
    '=========================================================================================
    '******************************************************************************
    '*プロパティ:iPort
    '*説明：      通信ポート番号
    '******************************************************************************
    Public Property sPort() As Integer
        Get
            Return Me._sPort
        End Get
        Set(ByVal value As Integer)
            Me._sPort = CShort(value)
        End Set
    End Property
    '******************************************************************************
    '*プロパティ：iSync
    '*説明： 1:同期, 2:非同期
    '*       同期指定時は，通信設定アプリケーションの終了をメソッド内で待ちます。
    '******************************************************************************
    Public Property sSync() As Integer
        Get
            Return Me._sSync
        End Get
        Set(ByVal value As Integer)
            Me._sSync = CShort(value)
        End Set
    End Property

    '******************************************************************************
    '*メソッド：PortSetWakeUp
    '*説明： ポート設定のダイアログを表示(ComMng.PortSetWakeUpそのまま)
    '*
    '******************************************************************************
    Public Function PortSetWakeUp() As Integer
        Try
            _oComMng.PortSetWakeUp(_sPort, _sSync)
        Catch ex As Exception
            Return 1
        End Try

        Return 0
    End Function

    Public Function GetPortStatus() As Short
        Dim sSts As Short

        Try
            _oComMng.GetPortStatus(_sPort, sSts)
        Catch ex As Exception
            Return -1
        End Try

        Return sSts
    End Function

    '******************************************************************************
    '*メソッド：ComOpn
    '*説明： FA-M3と通信接続する(ComMng.ComOpnそのまま)
    '*
    '******************************************************************************
    Public Function ComOpn() As Integer
        Try
            _oComMng.Comopn2(_sPort, _oInfo, _sSync)
        Catch ex As Exception
            Return 1
        End Try

        Return 0
    End Function
    '******************************************************************************
    '*メソッド:ComCls
    '*説明： FA-M3の通信接続を切断する(ComMng.ComClsそのまま)
    '*
    '******************************************************************************
    Public Function ComCls() As Integer
        Try
            _oComMng.Comcls(_sPort)
        Catch ex As Exception
            'MessageBox.Show("例外が発生しました。通信ケーブル等の状態を確認してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 1
        End Try

        Return 0
    End Function

    '******************************************************************************
    '*メソッド:Sysrst
    '*説明： CPUモジュールにリセットをかける
    '*
    '******************************************************************************
    Public Function Sysrst() As Integer
        Try
            _oComMng.Sysrst(_sPort, 0)
        Catch ex As Exception
            'MessageBox.Show("例外が発生しました。通信ケーブル等の状態を確認してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 1
        End Try

        Return 0
    End Function


    '=========================================================================================
    '=
    '=メソッド： IOバスゲートアレイ操作関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド：SetMdfyMdlAddrReg
    '*説明：モディファイモジュールアドレスレジスタを設定する
    '*      設定値 0x8017  + 0x0200×(slot番号-1)
    '*      アドレスサイズは2KB固定
    '*      Slot1  0x00760000
    '*      Slot2  0x00760800
    '*      Slot3  0x00761000
    '*      ・・・
    '*      Slot16 0x00767800
    '*引数：ByVal iSlot As Integer 設定するスロット番号を指定
    '*返値：Errあり1, Errなし0
    '*
    '******************************************************************************
    Private Function SetMdfyMdlAddrReg(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim iMdfyAddr As Int32
        Dim usMdfyData As UInt16

        '===========SLOT番号からIOモジュールのModifyレジスタのアドレス・設定値を計算する===========
        iMdfyAddr = I_MODIFY_MDL_ADDR_REG_BASE + I_MODIFY_MDL_ADDR_REG_OFFSET * (iSlot - 1)   '
        usMdfyData = CUShort(US_MODIFY_MDL_ADDR_REG_DATA + US_MODIFY_MDL_ADDR_REG_DATA_OFFSET * (iSlot - 1))  '

        '===========MODIFY REGへの設定===========
        Return Me.MemWriteWord2(iMdfyAddr, usMdfyData)

    End Function

    '******************************************************************************
    '*メソッド：SetWndAddrReg
    '*説明：ウインドウアドレスレジスタを設定する
    '*引数：ByVal iSlot As Integer 設定するスロット番号を指定
    '*      ByVal iOffsetAddr As Integer オフセットアドレスを設定する
    '*返値：Errあり1, Errなし0
    '*
    '*
    '*
    '******************************************************************************
    Private Function SetWndAddrReg(ByVal iSlot As Integer, ByVal iOffsetAddr As Integer) As Integer
        '*******変数宣言*******
        Dim iWndAddr As Int32
        Dim usUpperWndAdrReg As UInt16
        Dim usLowerWndAdrReg As UInt16
        Dim iRetval1 As Integer
        Dim iRetval2 As Integer

        '===========SLOT番号からIOモジュールのウインドウアドレスレジスタのアドレス・設定値を計算する===========
        iWndAddr = I_WINDOW_ADDR_REG_BASE + I_WINDOW_ADDR_REG_OFFSET * (iSlot - 1)
        usUpperWndAdrReg = CUShort((iOffsetAddr >> 16) And (&HFFFC))
        usLowerWndAdrReg = CUShort(CUShort((iOffsetAddr And (&H3FF80I)) >> 2) Or &H10)

        '====後で桁あふれがないか確認する 最上位bitが1の場合エラーになりそう・・


        '===========WindowAddress REGへの設定===========
        iRetval1 = Me.MemWriteWord2(iWndAddr, usUpperWndAdrReg)
        iRetval2 = Me.MemWriteWord2(iWndAddr + 2, usLowerWndAdrReg)

        If iRetval1 = 1 Then
            Return 1
        ElseIf iRetval2 = 1 Then
            Return 1
        Else
            Return 0
        End If

    End Function

    '******************************************************************************
    '*メソッド：GetMdfyMemBaseAddr
    '*説明：指定したスロットのモディファイ空間の先頭アドレスを返す
    '*引数：ByVal iSlot As Integer 設定するスロット番号を指定
    '*      ByRef iMdfyMemBaseAddr As Integer指定したスロットのモディファイ空間の先頭アドレス
    '******************************************************************************
    Private Sub GetMdfyMemBaseAddr(ByVal iSlot As Integer, ByRef iMdfyMemBaseAddr As Integer)
        iMdfyMemBaseAddr = I_MODIFY_MEM_ADDR_BASE + I_MODIFY_MEM_ADDR_OFFSET * (iSlot - 1)
    End Sub



    '=========================================================================================
    '=
    '=FA-M3のIO空間のリードライト関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド：GetMemDumpWord
    '*説明：FA-M3からWord単位でデータを読み出す。結果は文字列型で取得。
    '*引数：ByVal iMemAddr As Integer   読み出したいアドレスの先頭
    '*      ByVal sNum As Short         読み出すデータ数(word数)
    '*      ByRef sReadData() As Short  読み出したデータを格納する配列 
    '******************************************************************************
    Private Function MemReadWord(ByVal iMemAddr As Int32, ByVal sNum As Short, ByRef sReadData() As Short) As Integer
        '*******変数宣言*******
        Dim i As Integer

        'Word単位ででCPUモジュールから読み出し
        Try
            For i = 0 To (sNum - 1)
                _oComMng.MemredWord(_sPort, 1, iMemAddr + (i * 2), sReadData(i))
            Next
        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemReadWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try


        Return 0
    End Function


    '******************************************************************************
    '*メソッド名：MemWriteByte
    '*説明：FA-M3へByte単位でデータを書き込む。
    '*引数：ByVal iMemAddr As Integer   書き込みたいアドレス
    '*      ByVal bWrData As Byte      書き込みデータ
    '******************************************************************************
    Private Function MemWriteByte(ByVal iMemAddr As Integer, ByVal bWrData As Byte) As Integer
        Try
            _oComMng.Memwrt(_sPort, 1, 1, iMemAddr, CInt(bWrData))
        Catch ex As Exception
            'MessageBox.Show("例外が発生しました。通信状態・不正なアドレスにアクセスしていないか等を確認してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End Try
        Return 1
    End Function

    '******************************************************************************
    '*メソッド名：MemWriteWord
    '*説明：FA-M3へWord単位でデータを書き込む。
    '*引数：ByVal iMemAddr As UInt32   書き込みたいアドレス
    '*      ByVal sWrData As UInt16    書き込みデータ
    '******************************************************************************
    Private Function MemWriteWord(ByVal iMemAddr As Integer, ByVal sWrData As Short) As Integer
        Dim wData As WordUnion
        Dim bData As Byte


        'ByteSwap
        wData.s16_0 = sWrData
        bData = wData.u8_0
        wData.u8_0 = wData.u8_1
        wData.u8_1 = bData


        '書込み
        Try
            _oComMng.Memwrt(_sPort, 2, 1, iMemAddr, CInt(wData.s16_0))
        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemWriteWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try
        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：MemWriteWord
    '*説明：FA-M3へWord単位でデータを書き込む。
    '*引数：ByVal iMemAddr As UInt32   書き込みたいアドレス
    '*      ByVal sWrData As UInt16    書き込みデータ
    '******************************************************************************
    Private Function MemWriteWord2(ByVal iMemAddr As Integer, ByVal usWrData As UInt16) As Integer
        Dim uiByte1 As UInt16
        Dim uiByte2 As UInt16
        Dim iWrData As Integer

        'ByteSwap
        uiByte1 = ((usWrData >> 8) And &HFFUS)    ' 
        uiByte2 = ((usWrData << 8) And &HFF00US)  ' 
        usWrData = ((uiByte1 + uiByte2))          ' 
        iWrData = CInt(usWrData)                  '

        '書込み
        Try
            _oComMng.Memwrt(_sPort, 2, 1, iMemAddr, iWrData)
        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemWriteWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try
        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：MemWriteWord 1KB
    '*説明：FA-M3へWord単位でデータを書き込む。
    '*引数：ByVal iMemAddr As UInt32   書き込みたいアドレス
    '*      ByVal sWrData As UInt16    書き込みデータ
    '******************************************************************************
    Private Function MemWriteWord_1KB(ByVal iMemAddr As Integer, ByRef sWrData() As Integer, ByVal iStartAddr As Integer) As Integer

        '書込み
        Try
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr, sWrData(iStartAddr))             '256word (512Byte) Write
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr + 512, sWrData(iStartAddr + 128)) '256word (512Byte) Write   (128 LongWord)
        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemWriteWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try
        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：MemWriteWord 2KB
    '*説明：FA-M3へWord単位でデータを書き込む。512Byteを4回ライトして2KByte
    '*引数：ByVal iMemAddr As UInt32   書き込みたいアドレス
    '*      ByVal sWrData As UInt16    書き込みデータ
    '******************************************************************************
    Private Function MemWriteWord_2KB(ByVal iMemAddr As Integer, ByRef sWrData() As Integer, ByVal iStartAddr As Integer) As Integer

        '書込み
        Try
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr, sWrData(iStartAddr))                 '256word (512Byte) Write
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr + &H200, sWrData(iStartAddr + &H80))  '256word (512Byte) Write
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr + &H400, sWrData(iStartAddr + &H100)) '256word (512Byte) Write
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr + &H600, sWrData(iStartAddr + &H180)) '256word (512Byte) Write

        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemWriteWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try
        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：MemWriteWord 1KB
    '*説明：FA-M3へWord単位でデータを書き込む。
    '*引数：ByVal iMemAddr As UInt32   書き込みたいアドレス
    '*      ByVal sWrData As UInt16    書き込みデータ
    '******************************************************************************
    Private Function MemWriteWord_512B_upper(ByVal iMemAddr As Integer, ByRef sWrData() As Integer, ByVal start_addr As Integer) As Integer

        '書込み
        Try
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr, sWrData(start_addr))
        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemWriteWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try
        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：MemWriteWord 1KB
    '*説明：FA-M3へWord単位でデータを書き込む。
    '*引数：ByVal iMemAddr As UInt32   書き込みたいアドレス
    '*      ByVal sWrData As UInt16    書き込みデータ
    '******************************************************************************
    Private Function MemWriteWord_512B_lower(ByVal iMemAddr As Integer, ByRef sWrData() As Integer, ByVal start_addr As Integer) As Integer

        '書込み
        Try
            _oComMng.Memwrt(_sPort, 2, 256, iMemAddr, sWrData(start_addr + 128))
        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemWriteWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try
        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：MemWriteWord valuable
    '*説明：FA-M3へWord単位でデータを書き込む。
    '*引数：ByVal iMemAddr As UInt32   書き込みたいアドレス
    '*      ByVal sWrData As UInt16    書き込みデータ
    '******************************************************************************
    Private Function MemWriteWord_val(ByVal bytenum As Integer, ByVal iMemAddr As Integer, ByRef sWrData As Integer) As Integer

        '書込み
        Try
            _oComMng.Memwrt(_sPort, 2, CShort(bytenum), iMemAddr, sWrData)
        Catch ex As Exception
            'MsgBox("Dump_FAM3クラス MemWriteWordメソッド内で例外が発生！" & vbCrLf & vbCrLf & ex.Message)
            Return 1
        End Try
        Return 0
    End Function



    '=========================================================================================
    '=
    '=FPGAの再コンフィグレーション関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド： ReConfigFPGA
    '*説明：  指定したSLOTのモジュールをCFM0 or CFM1/2へ再コンフィグレーション
    '*引数：  ByVal iSlot As Integer 　　スロット番号を指定する
    '         ByVal iSwapEn As Integer   ライトするデータのワードスワップ
    '                                    0：Swapなし
    '                                    1：Swapあり
    '                                    0,1以外：設定禁止
    '         ByVal iConfSel As Integer  再コンフィグレーションするImageを選ぶ
    '                                    0：FactoryImage
    '                                    1：ApplicationImage
    '                                    0,1以外：設定禁止
    '*返値：  Err発生:0以外, Errなし:0
    '******************************************************************************
    Private Function ReConfigFPGA(ByVal iSlot As Integer, ByVal iSwapEn As Integer, ByVal iConfsel As Integer) As Integer
        '*******変数宣言*******
        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMdfyMemBaseAddr As Integer
        Dim usConfSelData As UInt16 = 1

        'エラー判定 iSwapEnが0 or 1以外のときエラー終了
        If (iSwapEn < 0) Then Return 1
        If (iSwapEn > 1) Then Return 2

        'エラー判定と値代入 iConfselが0 or 1以外のときエラー終了
        If (iConfsel < 0) Then
            Return 3
        ElseIf (iConfsel > 1) Then
            Return 4
        ElseIf (iConfsel = 0) Then
            usConfSelData = 1
        Else
            usConfSelData = 3
        End If

        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 5
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H20000I) ' + offset 0x0002_0000

        '===========CONF_SEL OVERWRIT CONF_SEL=0(CFM0) ===========
        If iSwapEn = 1 Then
            '指定したSLOT(iSlot)のDula Configuration IP レジスタをライト
            Me.MemWriteWord2(iMdfyMemBaseAddr + 6, usConfSelData)
        Else
            '指定したSLOT(iSlot)のDula Configuration IP レジスタをライト
            Me.MemWriteWord2(iMdfyMemBaseAddr + 4, usConfSelData)

        End If

        ''===========RECONFIGURAION Triger===========
        If iSwapEn = 1 Then
            '指定したSLOT(iSlot)のDula Configuration IP レジスタをライト
            Me.MemWriteWord2(iMdfyMemBaseAddr + 2, usConfSelData)
        Else
            '指定したSLOT(iSlot)のDula Configuration IP レジスタをライト
            Me.MemWriteWord2(iMdfyMemBaseAddr, usConfSelData)

        End If

        '===========100ms待つr===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        Return 0

    End Function

    '******************************************************************************
    '*メソッド： ReConfToFactImage
    '*説明：  指定したSLOTのモジュールをCFM0(FactoryImage)で再コンフィグレーション
    '*引数：  ByVal iSlot As Integer 　　スロット番号を指定する
    '         ByVal iSwapEn As Integer   ライトするデータのワードスワップ
    '                                    0：Swapなし
    '                                    1：Swapあり
    '*返値：  Err発生：0以外, Errなし:0
    '******************************************************************************
    Public Function ReConfToFactImage(ByVal iSlot As Integer, ByVal iSwapEn As Integer) As Integer
        Return ReConfigFPGA(iSlot, iSwapEn, 0)
    End Function
    'オーバーロード(iSwapEn=0固定)
    Public Function ReConfToFactImage(ByVal iSlot As Integer) As Integer
        Return ReConfToFactImage(iSlot, 0)
    End Function

    '******************************************************************************
    '*メソッド： ReConfToAppImage
    '*説明：  指定したSLOTのモジュールをCFM1/2(AppImage)で再コンフィグレーション
    '*引数：  ByVal iSlot As Integer 　　スロット番号を指定する
    '         ByVal iSwapEn As Integer   ライトするデータのワードスワップ
    '                                    0：Swapなし
    '                                    1：Swapあり
    '*返値：  Err発生：0以外, Errなし:0
    '******************************************************************************
    Public Function ReConfToAppImage(ByVal iSlot As Integer, ByVal iSwapEn As Integer) As Integer
        Return ReConfigFPGA(iSlot, iSwapEn, 1)
    End Function
    'オーバーロード(iSwapEn=0固定)
    Public Function ReConfToAppImage(ByVal iSlot As Integer) As Integer
        Return ReConfToAppImage(iSlot, 0)
    End Function

    '=========================================================================================
    '=
    '=FPGA Revision Registerの読み出し
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド： ReadFpgaRevReg
    '*説明：  指定したSLOTのモジュールのFPGA Revision Registerを読み出す
    '*引数：  ByVal iSlot As Integer 　　 スロット番号を指定する
    '         ByRef uiRevReg() As UInt32  読み出したFPGA Rev Regの値がこれに入る
    '*返値：  Err発生1, Errなし0
    '******************************************************************************
    Public Function ReadFpgaRevReg(ByVal iSlot As Integer, ByRef uiRevReg() As UInt32) As Integer
        '*******変数宣言*******
        Dim i As Integer    'カウンタ変数
        Dim lwuData As LongWordUnion
        Dim sData(12 - 1) As Short

        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMdfyMemBaseAddr As Integer

        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H20000I) ' + offset 0x0002_0000

        '===========FPGA Rev REGの読み出し===========
        Me.MemReadWord(iMdfyMemBaseAddr + &H100, 12, sData) '0x0002_0100番地がFPGA RevRegなので &H100を加算

        '===========読み出した結果を配列に保存===========
        For i = 0 To CInt(12 / 2 - 1)
            lwuData.s16_1 = sData(i * 2 + 0)
            lwuData.s16_0 = sData(i * 2 + 1)
            uiRevReg(i) = lwuData.u32_0
        Next

        '===========100ms待つ===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        Return 0

    End Function



    '=========================================================================================
    '=
    '=FPGA内蔵FLASHへのAppImage書込み関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド： Get_FPGA_ROM_SUM
    '*説明：  FPGAからUFM/CFMのSUM値を取得する
    '*引数：  ByVal iSlot As Integer  AppImageを書き込むモジュールのSLOT番号を指定する
    '******************************************************************************
    Public Function Get_FPGA_ROM_SUM(ByVal iSlot As Integer, ByRef usSum() As UInt16) As Integer
        '*******変数宣言*******
        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMdfyMemBaseAddr As Integer
        '---FA-M3のIO空間のアドレスと書き込む変数、読み出した値を格納する変数---
        Dim sReadData(10) As Int16
        '型変換用
        Dim wuData As WordUnion

        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)


        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000

        '===========0. FPGA 内蔵ROM SUM値計算開始 ===========
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H452, 1) '0x0001  ROM_SUM_CHK_RUNレジスタセット

        '===========1. WAIT UNTIL ROM_SUM_CHK_RUN Register Clear===========
        '---ROM_SUM_CHK_RUN Register  Check---

        Dim i As Integer = 0
        While True '---無限ループ
            '---ROM_SUM_CHK_RUN Register ---
            If Me.MemReadWord(iMdfyMemBaseAddr + &H452, 1, sReadData) = 1 Then
                Return 1
            End If

            If sReadData(0) = &H0S Then '0x0000
                Exit While
            End If
            i = i + 1
        End While

        '===========2. FPGA 内蔵ROM SUM値読み出し ===========
        Me.MemReadWord(iMdfyMemBaseAddr + &H454, 3, sReadData)  ' 0x454：UFM0_1_SUM, 0x456：CFM1_2_SUM, 0x458：CFM0_SUM
        wuData.s16_0 = sReadData(0)
        usSum(0) = wuData.u16_0
        wuData.s16_0 = sReadData(1)
        usSum(1) = wuData.u16_0
        wuData.s16_0 = sReadData(2)
        usSum(2) = wuData.u16_0

        '===========100ms待つr===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        Return 0

    End Function

    '=========================================================================================
    '=
    '=FPGA内蔵FLASHへのFactory Image書込み関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド： FacImageWrite
    '*説明：  FPGAの内蔵CFM0へFactory Imageを書き込む
    '*引数：  ByVal iSlot As Integer  Factory Imageを書き込むモジュールのSLOT番号を指定する
    '******************************************************************************
    Public Function FacImageWrite(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMdfyMemBaseAddr As Integer
        '---FA-M3のIO空間のアドレスと書き込む変数、読み出した値を格納する変数---
        Dim sReadData(10) As Int16
        '---タイムアップカウンタ---
        Dim iTimUpCnt As Integer = 0  'カウントアップ(COUNT=1000)までにNiosIIがFlagを落とさない場合エラーで終了
        '---カウンタ変数---
        Dim i As Integer


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)


        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000

        '===========0. SET CFM0 書き込み STATUS Register active_flag===========
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H464, 1) '0x0001  CSTATR1_DATA


        '---1KBずつ 140kBまで共有メモリにRPDデータを送信する
        For i = 0 To 139 Step 1 '1KB씩 Window address변경' ~140KB
            'RPD 1KB を書き込む 
            Me.MemWriteWord_1KB(iMdfyMemBaseAddr + &H50, AppImageRPD.LwFile, (256 * i))

            '===========1. SET 00760450 "32'b0000_0000_0000_0011===========
            Me.MemWriteWord2(iMdfyMemBaseAddr + &H464, &H3)

            '===========2. WAIT UNTIL 00760450 0x0001===========
            '---Flag Check---
            While True '---無限ループ
                '---READ NIOS_PROCESS_FLAG_WORD---
                If Me.MemReadWord(iMdfyMemBaseAddr + &H464, 1, sReadData) = 1 Then
                    Return 1
                End If

                If sReadData(0) = &H1S Then     '0x0001
                    Exit While
                ElseIf sReadData(0) = &H0S Then '0x0000
                    Exit While
                End If

                '---タイムアップでエラー終了---
                If iTimUpCnt = 10000 Then
                    Return 1
                End If
                iTimUpCnt = iTimUpCnt + 1

            End While
            iTimUpCnt = 0
        Next i

        '===========100ms待つr===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        Return 0

    End Function

    '=========================================================================================
    '=
    '=FPGA内蔵FLASHへのAppImage書込み関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド： AppImageWrite
    '*説明：  FPGAの内蔵CFM1/2へAppImageを書き込む
    '*引数：  ByVal iSlot As Integer  AppImageを書き込むモジュールのSLOT番号を指定する
    '******************************************************************************
    Public Function AppImageWrite(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMdfyMemBaseAddr As Integer
        '---FA-M3のIO空間のアドレスと書き込む変数、読み出した値を格納する変数---
        Dim sReadData(10) As Int16
        '---タイムアップカウンタ---
        Dim iTimUpCnt As Integer = 0  'カウントアップ(COUNT=1000)までにNiosIIがFlagを落とさない場合エラーで終了
        '---カウンタ変数---
        Dim i As Integer


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)


        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000

        '===========0. SET CFM1/2 書き込み STATUS Register active_flag===========
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H450, 1) '0x0001  CSTATR_DATA


        '---1KBずつ 140kBまで共有メモリにRPDデータを送信する
        For i = 0 To 139 Step 1 '1KB씩 Window address변경' ~140KB
            'RPD 1KB を書き込む 
            Me.MemWriteWord_1KB(iMdfyMemBaseAddr + &H50, AppImageRPD.LwFile, (256 * i))

            '===========1. SET 00760450 "32'b0000_0000_0000_0011===========
            Me.MemWriteWord2(iMdfyMemBaseAddr + &H450, &H3)

            '===========2. WAIT UNTIL 00760450 0x0001===========
            '---Flag Check---
            While True '---無限ループ
                '---READ NIOS_PROCESS_FLAG_WORD---
                If Me.MemReadWord(iMdfyMemBaseAddr + &H450, 1, sReadData) = 1 Then
                    Return 1
                End If

                If sReadData(0) = &H1S Then     '0x0001
                    Exit While
                ElseIf sReadData(0) = &H0S Then '0x0000
                    Exit While
                End If

                '---タイムアップでエラー終了---
                If iTimUpCnt = 10000 Then
                    Return 1
                End If
                iTimUpCnt = iTimUpCnt + 1

            End While
            iTimUpCnt = 0
        Next i

        '===========100ms待つr===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        Return 0

    End Function


    '=========================================================================================
    '=
    '=FPGA内蔵MCU(NiosII)のファームウェアデータ転送関連
    '=
    '=========================================================================================
    '******************************************************************************
    Public Function NiosFirmDataWrite(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMemAddr As Integer

        Dim iRetVal As Integer = 0

        '===========FirmData書き込み前にNiosIIにリセットをかけてとめる===========
        iRetVal = nios_reset_low(iSlot)
        If iRetVal <> 0 Then Return iRetVal

        '===========MODIFY REGESTER SETTING===========
        SetMdfyMdlAddrReg(iSlot)

        'DUAL MEMORY 16KBYTE WRITE
        GetMdfyMemBaseAddr(iSlot, iMemAddr)
        For i = 0 To (FactImageNiosFirm.LwSwapFile.Length * 4 / 1024 - 1) Step 2 '2KB stepでライトする 
            '===========window regester setting===========
            SetWndAddrReg(iSlot, CInt((&H8000I) + (&H400) * i)) ' + offset 0x0000_8000 + 0x0000_0400*i  (0x400=1KByte)

            '===========NiosII Program Write to Memory===========
            Me.MemWriteWord_2KB(iMemAddr, FactImageNiosFirm.LwSwapFile, CInt((256 * i)))   '(256 LongWord = 1KByte)
        Next

        '===========FirmData書き込み後NiosIIのリセット解除=========== Hi->Low->Hiの理由は???
        iRetVal = nios_reset_high(iSlot)
        If iRetVal <> 0 Then Return iRetVal
        iRetVal = nios_reset_low(iSlot)
        If iRetVal <> 0 Then Return iRetVal
        iRetVal = nios_reset_high(iSlot)
        If iRetVal <> 0 Then Return iRetVal

        '===========100ms待つ===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        Return 0

    End Function

    '******************************************************************************
    '*メソッド名：nios_reset_low
    '*説明：FPGA内部バス(avalon)のNios側のリセットをLowにする
    '*引数：ByVal iSlot As Integer リセットLowにするモジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー   
    '******************************************************************************
    Private Function nios_reset_low(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMemAddr As Integer
        Dim usWdata As UInt16

        Dim iRetVal As Integer

        '---wait 1ms---
        System.Threading.Thread.Sleep(10)


        '===========MODIFY REGESTER SETTING===========後でけす
        iRetVal = SetMdfyMdlAddrReg(iSlot)
        If iRetVal <> 0 Then Return iRetVal

        '---wait 1ms---
        System.Threading.Thread.Sleep(10)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H10000I) ' + offset 0x0001_0000


        '===========NiosII RESET High===========  Lowの関数なのに、Highしてよいのか???また論理正しい?????
        'GetMdfyMemBaseAddr(iSlot, iMemAddr)

        'wordサイズで書込み
        'usWdata = 0      '0x000
        'me.MemWriteWord2(iMemAddr + &H10, usWdata)


        'wait 1ms
        'System.Threading.Thread.Sleep(10)

        '===========NiosII RESET Low===========
        GetMdfyMemBaseAddr(iSlot, iMemAddr)

        'wordサイズで書込み
        usWdata = &H1      '0x001
        Me.MemWriteWord2(iMemAddr + &H10, usWdata)


        Return 0

    End Function

    '******************************************************************************
    '*メソッド名：nios_reset_high
    '*説明：FPGA内部バス(avalon)のNios側のリセットをHighにする
    '*引数：ByVal iSlot As Integer リセットHighにするモジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Private Function nios_reset_high(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        '---FA-M3のIO空間のアドレスと書き込む変数---
        Dim iMemAddr As Integer
        Dim usWdata As UInt16

        Dim iRetVal As Integer

        '---wait 1ms---
        System.Threading.Thread.Sleep(10)

        '===========MODIFY REGESTER SETTING===========後でけす
        iRetVal = SetMdfyMdlAddrReg(iSlot)
        If iRetVal <> 0 Then Return iRetVal


        '---wait 1ms---
        System.Threading.Thread.Sleep(10)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H10000I) ' + offset 0x0001_0000


        '===========NiosII RESET High===========  Lowの関数なのに、Highしてよいのか???また論理正しい?????
        GetMdfyMemBaseAddr(iSlot, iMemAddr)
        'wordサイズで書込み
        usWdata = 0      '0x000
        Me.MemWriteWord2(iMemAddr + &H10, usWdata)

        Return 0

    End Function



    '=========================================================================================
    '=
    '=MCU内蔵FLASHへのファームウェアデータ書込み関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド名：MCUSerialWriterModeRUN
    '*説明：Cypress FM4 Family MCU (S6E2H46G0AGV2)をシリアルライタモードで起動する
    '       FPGAのPIO端子を使ってmode端子設定とResetをかける
    '*引数：ByVal iSlot As Integer リセットHighにするモジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Public Function MCUSerialWriterModeRUN(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim iMdfyMemBaseAddr As Integer         '---FA-M3のIO空間のアドレスと書き込む変数、読み出した値を格納する変数---


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1

        '===========wait 3ms'===========
        System.Threading.Thread.Sleep(30)

        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H4000I) ' + offset 0x0000_4000   'GPIO REG BASE Address

        '===========MODE Set And Reset(GPIO5,GPIO6,GPIO7) OUTPUT===========
        Me.MemWriteByte(iMdfyMemBaseAddr + &H4, &HF0)

        'md0 high=========================================================================================================
        Me.MemWriteByte(iMdfyMemBaseAddr, &H80)
        '=========================================================================================================

        'delay 1ms
        System.Threading.Thread.Sleep(10)

        'XRESETO high=========================================================================================================
        Me.MemWriteByte(iMdfyMemBaseAddr, &HA0)
        ''=========================================================================================================

        '===========1000ms待つ===========
        System.Threading.Thread.Sleep(1000)       '100msの根拠は不明(実機確認ベース)


        '*****正常終了*****
        Return 0
    End Function

    '=========================================================================================
    '=
    '=MCU内蔵FLASHへのファームウェアデータ書込み関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド名：MCUSerialWriterModeRUN
    '*説明：Cypress FM4 Family MCU (S6E2H46G0AGV2)をシリアルライタモードで起動する
    '       FPGAのPIO端子を使ってmode端子設定とResetをかける
    '*引数：ByVal iSlot As Integer リセットHighにするモジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Public Function MCUSerialWriterModeRUN1(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim iMdfyMemBaseAddr As Integer         '---FA-M3のIO空間のアドレスと書き込む変数、読み出した値を格納する変数---


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1

        '===========wait 3ms'===========
        System.Threading.Thread.Sleep(30)

        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H4000I) ' + offset 0x0000_4000   'GPIO REG BASE Address

        '===========MODE Set And Reset(GPIO10,GPIO11) OUTPUT=========== REV4
        Me.MemWriteByte(iMdfyMemBaseAddr + &H5, &HC)

        'md0 high=========================================================================================================
        Me.MemWriteByte(iMdfyMemBaseAddr + &H1, &H4)
        '=========================================================================================================

        'delay 1ms
        System.Threading.Thread.Sleep(10)

        'XRESETO high=========================================================================================================
        Me.MemWriteByte(iMdfyMemBaseAddr + &H1, &HC)
        ''=========================================================================================================

        '===========1000ms待つ===========
        System.Threading.Thread.Sleep(1000)       '100msの根拠は不明(実機確認ベース)


        '*****正常終了*****
        Return 0
    End Function

    ' Make By YMFK 製品技術JIK
    '=========================================================================================
    '=
    '=MCU内蔵FLASHへのファームウェアデータ書込み関連
    '=
    '=========================================================================================
    '******************************************************************************
    '*メソッド名：MCUSerialWriterModeRUN
    '*説明：Cypress H8S MCU (HD64F2378BVFQ35WV)をシリアルライタモードで起動する
    '       FPGAのPIO端子を使ってmode端子設定とResetをかける
    '*引数：ByVal iSlot As Integer リセットHighにするモジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Public Function MCUSerialWriterModeRUN_H8S(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim iMdfyMemBaseAddr As Integer         '---FA-M3のIO空間のアドレスと書き込む変数、読み出した値を格納する変数---


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1

        '===========wait 3ms'===========
        System.Threading.Thread.Sleep(30)

        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H4000I) ' + offset 0x0000_4000   'GPIO REG BASE Address

        '===========MODE Set And Reset(GPIO5,GPIO6,GPIO7) OUTPUT===========
        Me.MemWriteByte(iMdfyMemBaseAddr + &H4, &HE0)

        'iIoPfmClsRev(REV2)追加
        'H8S モード設定md01 high md2 low==========================================================================
        Me.MemWriteByte(iMdfyMemBaseAddr, &H80)
        '=========================================================================================================

        'iIoPfmClsRev(REV2)追加
        'delay 1ms
        System.Threading.Thread.Sleep(10)

        'md0 high=========================================================================================================
        Me.MemWriteByte(iMdfyMemBaseAddr, &HA0)
        '=========================================================================================================

        'delay 1ms
        System.Threading.Thread.Sleep(10)

        '***********************************************************************************************************************
        '使用するUARTポートの選択
        'UARTポートの有効化
        '***********************************************************************************************************************
        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H10000I) ' + offset 0x0001_0000  'UART MUX BASE Address

        '===========UART PORT1 Select===========
        Me.MemWriteByte(iMdfyMemBaseAddr + 4, &H1)

        '===========UART OUTPUT ENABLE===========
        Me.MemWriteByte(iMdfyMemBaseAddr, &H1)

        '===========1000ms待つ===========
        System.Threading.Thread.Sleep(1000)       '100msの根拠は不明(実機確認ベース)


        '*****正常終了*****
        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：MCUFlashWrFirmWrite
    '*説明：MCUの内蔵RAMにFLASH書き込み用のファームウェアをダウンロードする
    '*引数：ByVal iSlot As Integer リセットHighにするモジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Public Function MCUFlashWrFirmWrite(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim i As Integer                'カウンタ変数
        Dim iMdfyMemBaseAddr As Integer 'モディファイドモジュールアドレス先頭番地----
        Dim sReadData1(0) As Short      'リード値
        Dim sReadData2(0) As Short      'リード値
        Dim sReadData3(0) As Short      'リード値
        Dim TimOutCnt As Integer = 0    'Time Out カウンタ
        Dim frame_DOWN_CMD(9) As Byte   'CYPRESS MCU Serial Writer Mode DownLoad Command 10Byte
        Dim lwDlAddr As LongWordUnion   'CYPRESS MCU Serial Writer Mode 時MCU内蔵RAMの先頭アドレス
        Dim lwMCUFlashWrFirmLength As LongWordUnion 'CYPRESS MCU Serial Writer Mode 時 送信するFlashWriteFirmのカウント数
        Dim lwDlCmdSum As LongWordUnion  'CYPRESS MCU Serial Writer Mode 時 送信するコマンドのSUM値


        '===========シリアルライタモードへ移行===========
        Me.MCUSerialWriterModeRUN(iSlot)

        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1

        '===========wait 1ms'===========
        System.Threading.Thread.Sleep(10)

        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)


        '***********************************************************************************************************************
        'CYPRESS MCU Serial Writer Mode 時送信するデータを作る
        '(ダウンロードコマンドの送信フレームの作成)
        '***********************************************************************************************************************
        '---CYPRESS MCU Serial Writer Mode 時MCU内蔵RAMの先頭アドレス---
        lwDlAddr.u32_0 = &H2003C000UL

        '---CYPRESS MCU Serial Writer Mode 時 送信するFlashWriteFirmのカウント数 ---
        lwMCUFlashWrFirmLength.s32_0 = MCUFlashWrFirm.ByteFile.Length

        '---DownLoad CommandのSUM値計算(8bit単位単純加算、オーバーフロー無視)---
        lwDlCmdSum.u32_0 = 0
        lwDlCmdSum.u32_0 = CUInt(&H0 _
                         + lwDlAddr.u8_0 _
                         + lwDlAddr.u8_1 _
                         + lwDlAddr.u8_2 _
                         + lwDlAddr.u8_3 _
                         + lwMCUFlashWrFirmLength.u8_0 _
                         + lwMCUFlashWrFirmLength.u8_1 _
                         + lwMCUFlashWrFirmLength.u8_2 _
                         + lwMCUFlashWrFirmLength.u8_3)

        '---CYPRESS MCU Serial Writer Mode DownLoad Command'---
        frame_DOWN_CMD(0) = &H0                         '0x00 ダウンロードコマンド
        frame_DOWN_CMD(1) = lwDlAddr.u8_0               'ダウンロード開始アドレス A7～A0
        frame_DOWN_CMD(2) = lwDlAddr.u8_1               'ダウンロード開始アドレス A15～A8
        frame_DOWN_CMD(3) = lwDlAddr.u8_2               'ダウンロード開始アドレス A23～A16
        frame_DOWN_CMD(4) = lwDlAddr.u8_3               'ダウンロード開始アドレス A31～A24
        frame_DOWN_CMD(5) = lwMCUFlashWrFirmLength.u8_0 'ダウンロードカウント数 BC7～BC0
        frame_DOWN_CMD(6) = lwMCUFlashWrFirmLength.u8_1 'ダウンロードカウント数 BC15～BC8
        frame_DOWN_CMD(7) = lwMCUFlashWrFirmLength.u8_2 'ダウンロードカウント数 BC23～BC16
        frame_DOWN_CMD(8) = lwMCUFlashWrFirmLength.u8_3 'ダウンロードカウント数 BC31～BC24
        frame_DOWN_CMD(9) = lwDlCmdSum.u8_0             'frame_DOWN_CMD(0)～(9)のSUM値


        '***********************************************************************************************************************
        '使用するUARTポートの選択
        'UARTポートの有効化
        '***********************************************************************************************************************
        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H10000I) ' + offset 0x0001_0000  'UART MUX BASE Address

        '===========UART PORT1 Select===========
        Me.MemWriteByte(iMdfyMemBaseAddr + 4, &H1)

        '===========UART OUTPUT ENABLE===========
        Me.MemWriteByte(iMdfyMemBaseAddr, &H1)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000  'Shared Memory BASE Address


        '***********************************************************************************************************************
        'Cypress-M4 Core FM4 Family Serial Writer Mode
        'Reset Command Process '
        '***********************************************************************************************************************
        'STEP 0. SET UART WRITE START FLAG (0x00000460番地へ0xCCCCをライト)   
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H460, &HCCCC)   'MCUのProgaramRAM領域へファームダウンロード開始

        'STEP 1. SET RESET START FLAG (0x00000470番地へ0x0001をライト) 
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H470, &H1)      'reset_flag=ON MCUのProgram RAM領域のリセット実行

        'STEP 2. WAIT UNTIL RAM RESET STOP REGISTER = ON  (0x00000472番地が0x0001になるまで待つ)
        '   TIME OUT発生時エラー終了       (0x00000468番地が0x0000以外のときエラー)
        TimOutCnt = 0
        While True '無限ループ
            '---READ NIOS_PROCESS_FLAG_WORD---
            If Me.MemReadWord(iMdfyMemBaseAddr + &H472, 1, sReadData1) = 1 Then Return 1 'RAM RESET STOP REGISTER 0x472
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

            'RAM RESET STOP REGISTER がONしたら無限ループから抜ける
            If sReadData1(0) = &H1S Then Exit While '0x0001 'RAM RESET STOP REGISTER 

            'NiosIIファーム側がタイムアウトしたらエラー終了
            If sReadData2(0) <> &H0S Then Return 1 'TIME OUT STATUS REGISTER

            '---タイムアップでエラー終了---
            If TimOutCnt = 10000 Then Return 1

            '---タイムアウトカウンタインクリメント---
            TimOutCnt = TimOutCnt + 1
        End While


        'STEP 3. CLEAR TIME OUT STATUS REG  (0x00000468番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)    ' TIME OUT STATUS REG 0クリア

        'STEP 4. CLEAR RAM RESET START REG  (0x00000470番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H470, &H0)    ' RAM RESET START REG  0クリア
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H472, &H0)    ' RAM RESET STOP REG  0クリア


        '***********************************************************************************************************************
        'Cypress-M4 Core FM4 Family Serial Writer Mode
        'Program Download Process 
        '***********************************************************************************************************************
        'STEP 1. SET PROGRAM EXECUTE START REG  (0x0000_0474番地へ0x0001をライト)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H474, &H1)    ' SET PROGRAM EXECUTE START REG 0x0001ライト


        'STEP 2. WAIT UNTIL  PROGRAM EXECUTE STOP FLAG == 1 (0x0000_0476番地  0x0001  )
        '　      WAIT UNTIL  PROGRAM EXECUTE STOP FLAG == 1 (0x0000_0474番地  0x0001  )
        Dim step_cnt As Integer = 0
        While True
            'PROGRAM EXECUTE 4KBYTE WRITE PROCESS
            If step_cnt = 0 Then   '←このコードの理由が分からない。後で調査する
                step_cnt = 1
                'SHARED MEMORYへMCU内蔵Flash書き込み用ファームウェアを転送する(1KByte × 4回)
                For i = 0 To 3 Step 1
                    Me.MemWriteWord2(iMdfyMemBaseAddr + &H462, &H1)    ' MEMORY_1KB_WRITE REG 0x0001ライト
                    Me.MemWriteWord_512B_upper(iMdfyMemBaseAddr + &H50, MCUFlashWrFirm.LwFile, (256 * i))
                    Me.MemWriteWord_512B_lower(iMdfyMemBaseAddr + &H250, MCUFlashWrFirm.LwFile, (256 * i))

                    'SHARED MEMORY WRITE FINISH CHECK 
                    While True
                        If Me.MemReadWord(iMdfyMemBaseAddr + &H462, 1, sReadData1) = 1 Then Return 1 'MEMORY_1KB_WRITE REG
                        If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STAUS REG
                        If sReadData1(0) = &H0 Then Exit While
                        If sReadData2(0) <> &H0 Then
                            Return 1
                        End If
                    End While

                    'SET 0076046A "0000" timeout1 flag
                    Me.MemWriteWord2(iMdfyMemBaseAddr + &H46A, &H0)    ' 0x46Aは存在しているのか???????????
                Next i
            End If

            '---READ NIOS_PROCESS_FLAG_WORD---
            If Me.MemReadWord(iMdfyMemBaseAddr + &H476, 1, sReadData1) = 1 Then Return 1 'RAM WRITE EXECTUTE STOP REG (MCU内蔵RAM転送確認)
            If Me.MemReadWord(iMdfyMemBaseAddr + &H46C, 1, sReadData2) = 1 Then Return 1 'COPY STATUS REG
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData3) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

            'RAM WRITE EXECTUTE STOP REG (MCU内蔵RAM転送確認)
            If sReadData1(0) = &H1 Then
                Exit While

                'COPY STATUS REG 確認
            ElseIf sReadData2(0) = &H1 Then
                'FLAME DOWN DATA CMD, DATA BYTE LENGTH, DATA WORD LENGTH SEND
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500, frame_DOWN_CMD(0))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H1, frame_DOWN_CMD(1))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H2, frame_DOWN_CMD(2))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H3, frame_DOWN_CMD(3))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H4, frame_DOWN_CMD(4))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H5, frame_DOWN_CMD(5))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H6, frame_DOWN_CMD(6))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H7, frame_DOWN_CMD(7))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H8, frame_DOWN_CMD(8))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H9, frame_DOWN_CMD(9))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &HA, CByte(MCUFlashWrFirm.Sum))  'SUM値設定
                Me.MemWriteWord(iMdfyMemBaseAddr + &H500 + &HC, CShort(MCUFlashWrFirm.ByteFile.Length))
                Me.MemWriteWord(iMdfyMemBaseAddr + &H500 + &HE, CShort(MCUFlashWrFirm.WordFile.Length))

                'CLEAR COPY STATUS REG ((0x0000_046C番地へ0x0000をライト)
                Me.MemWriteWord2(iMdfyMemBaseAddr + &H46C, &H0)

                'TIME OUT STATUS REGISTER 確認
            ElseIf sReadData3(0) <> 0 Then
                Return 1
            End If
        End While


        '=====MCUFlashWrFirmの書き込み完了確認====
        While True '無限ループ
            '---EXECUTE FINISH CHECK---
            If Me.MemReadWord(iMdfyMemBaseAddr + &H460, 1, sReadData1) = 1 Then Return 1 'SET UART WRITE START FLAG (書き込み完了確認) 
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER  (タイムアウト検出)
            If sReadData1(0) = 0 Then Exit While '書き込み完了でループを抜ける
            If sReadData2(0) <> 0 Then Return 1 'NiosIIファームのタイムアウト検出でエラー終了
        End While

        'STEP 3. CLEAR TIME OUT STATUS REG  (0x00000468番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)    ' TIME OUT STATUS REG 0クリア

        'STEP 4. OK, CLEAR RAM WRITE EXECTUTE STOP REG 
        Me.MemWriteWord(iMdfyMemBaseAddr + &H476, &H0)

        '===========100ms待つ===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        '*****正常終了*****
        Return 0

    End Function

    '******************************************************************************
    '*メソッド名：MCUFlashWrFirmWrite
    '*説明：MCUの内蔵RAMにFLASH書き込み用のファームウェアをダウンロードする
    '*引数：ByVal iSlot As Integer リセットHighにするモジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Public Function MCUFlashWrFirmWrite1(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim i As Integer                'カウンタ変数
        Dim iMdfyMemBaseAddr As Integer 'モディファイドモジュールアドレス先頭番地----
        Dim sReadData1(0) As Short      'リード値
        Dim sReadData2(0) As Short      'リード値
        Dim sReadData3(0) As Short      'リード値
        Dim TimOutCnt As Integer = 0    'Time Out カウンタ
        Dim frame_DOWN_CMD(9) As Byte   'CYPRESS MCU Serial Writer Mode DownLoad Command 10Byte
        Dim lwDlAddr As LongWordUnion   'CYPRESS MCU Serial Writer Mode 時MCU内蔵RAMの先頭アドレス
        Dim lwMCUFlashWrFirmLength As LongWordUnion 'CYPRESS MCU Serial Writer Mode 時 送信するFlashWriteFirmのカウント数
        Dim lwDlCmdSum As LongWordUnion  'CYPRESS MCU Serial Writer Mode 時 送信するコマンドのSUM値


        '===========シリアルライタモードへ移行===========
        Me.MCUSerialWriterModeRUN1(iSlot)

        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1

        '===========wait 1ms'===========
        System.Threading.Thread.Sleep(10)

        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)


        '***********************************************************************************************************************
        'CYPRESS MCU Serial Writer Mode 時送信するデータを作る
        '(ダウンロードコマンドの送信フレームの作成)
        '***********************************************************************************************************************
        '---CYPRESS MCU Serial Writer Mode 時MCU内蔵RAMの先頭アドレス---
        lwDlAddr.u32_0 = &H2003C000UL

        '---CYPRESS MCU Serial Writer Mode 時 送信するFlashWriteFirmのカウント数 ---
        lwMCUFlashWrFirmLength.s32_0 = MCUFlashWrFirm.ByteFile.Length

        '---DownLoad CommandのSUM値計算(8bit単位単純加算、オーバーフロー無視)---
        lwDlCmdSum.u32_0 = 0
        lwDlCmdSum.u32_0 = CUInt(&H0 _
                         + lwDlAddr.u8_0 _
                         + lwDlAddr.u8_1 _
                         + lwDlAddr.u8_2 _
                         + lwDlAddr.u8_3 _
                         + lwMCUFlashWrFirmLength.u8_0 _
                         + lwMCUFlashWrFirmLength.u8_1 _
                         + lwMCUFlashWrFirmLength.u8_2 _
                         + lwMCUFlashWrFirmLength.u8_3)

        '---CYPRESS MCU Serial Writer Mode DownLoad Command'---
        frame_DOWN_CMD(0) = &H0                         '0x00 ダウンロードコマンド
        frame_DOWN_CMD(1) = lwDlAddr.u8_0               'ダウンロード開始アドレス A7～A0
        frame_DOWN_CMD(2) = lwDlAddr.u8_1               'ダウンロード開始アドレス A15～A8
        frame_DOWN_CMD(3) = lwDlAddr.u8_2               'ダウンロード開始アドレス A23～A16
        frame_DOWN_CMD(4) = lwDlAddr.u8_3               'ダウンロード開始アドレス A31～A24
        frame_DOWN_CMD(5) = lwMCUFlashWrFirmLength.u8_0 'ダウンロードカウント数 BC7～BC0
        frame_DOWN_CMD(6) = lwMCUFlashWrFirmLength.u8_1 'ダウンロードカウント数 BC15～BC8
        frame_DOWN_CMD(7) = lwMCUFlashWrFirmLength.u8_2 'ダウンロードカウント数 BC23～BC16
        frame_DOWN_CMD(8) = lwMCUFlashWrFirmLength.u8_3 'ダウンロードカウント数 BC31～BC24
        frame_DOWN_CMD(9) = lwDlCmdSum.u8_0             'frame_DOWN_CMD(0)～(9)のSUM値


        '***********************************************************************************************************************
        '使用するUARTポートの選択
        'UARTポートの有効化
        '***********************************************************************************************************************
        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H10000I) ' + offset 0x0001_0000  'UART MUX BASE Address

        '===========UART PORT2 Select===========
        Me.MemWriteByte(iMdfyMemBaseAddr, &H2)

        '===========UART OUTPUT ENABLE===========
        Me.MemWriteByte(iMdfyMemBaseAddr + 4, &H1)

        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000  'Shared Memory BASE Address


        '***********************************************************************************************************************
        'Cypress-M4 Core FM4 Family Serial Writer Mode
        'Reset Command Process '
        '***********************************************************************************************************************
        'STEP 0. SET UART WRITE START FLAG (0x00000460番地へ0xCCCCをライト)   
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H460, &HCCCC)   'MCUのProgaramRAM領域へファームダウンロード開始

        'STEP 1. SET RESET START FLAG (0x00000470番地へ0x0001をライト) 
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H470, &H1)      'reset_flag=ON MCUのProgram RAM領域のリセット実行

        'STEP 2. WAIT UNTIL RAM RESET STOP REGISTER = ON  (0x00000472番地が0x0001になるまで待つ)
        '   TIME OUT発生時エラー終了       (0x00000468番地が0x0000以外のときエラー)
        TimOutCnt = 0
        While True '無限ループ
            '---READ NIOS_PROCESS_FLAG_WORD---
            If Me.MemReadWord(iMdfyMemBaseAddr + &H472, 1, sReadData1) = 1 Then Return 1 'RAM RESET STOP REGISTER 0x472
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

            'RAM RESET STOP REGISTER がONしたら無限ループから抜ける
            If sReadData1(0) = &H1S Then Exit While '0x0001 'RAM RESET STOP REGISTER 

            'NiosIIファーム側がタイムアウトしたらエラー終了
            If sReadData2(0) <> &H0S Then Return 1 'TIME OUT STATUS REGISTER

            '---タイムアップでエラー終了---
            If TimOutCnt = 10000 Then Return 1

            '---タイムアウトカウンタインクリメント---
            TimOutCnt = TimOutCnt + 1
        End While


        'STEP 3. CLEAR TIME OUT STATUS REG  (0x00000468番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)    ' TIME OUT STATUS REG 0クリア

        'STEP 4. CLEAR RAM RESET START REG  (0x00000470番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H470, &H0)    ' RAM RESET START REG  0クリア
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H472, &H0)    ' RAM RESET STOP REG  0クリア


        '***********************************************************************************************************************
        'Cypress-M4 Core FM4 Family Serial Writer Mode
        'Program Download Process 
        '***********************************************************************************************************************
        'STEP 1. SET PROGRAM EXECUTE START REG  (0x0000_0474番地へ0x0001をライト)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H474, &H1)    ' SET PROGRAM EXECUTE START REG 0x0001ライト


        'STEP 2. WAIT UNTIL  PROGRAM EXECUTE STOP FLAG == 1 (0x0000_0476番地  0x0001  )
        '　      WAIT UNTIL  PROGRAM EXECUTE STOP FLAG == 1 (0x0000_0474番地  0x0001  )
        Dim step_cnt As Integer = 0
        While True
            'PROGRAM EXECUTE 4KBYTE WRITE PROCESS
            If step_cnt = 0 Then   '←このコードの理由が分からない。後で調査する
                step_cnt = 1
                'SHARED MEMORYへMCU内蔵Flash書き込み用ファームウェアを転送する(1KByte × 4回)
                For i = 0 To 3 Step 1
                    Me.MemWriteWord2(iMdfyMemBaseAddr + &H462, &H1)    ' MEMORY_1KB_WRITE REG 0x0001ライト
                    Me.MemWriteWord_512B_upper(iMdfyMemBaseAddr + &H50, MCUFlashWrFirm.LwFile, (256 * i))
                    Me.MemWriteWord_512B_lower(iMdfyMemBaseAddr + &H250, MCUFlashWrFirm.LwFile, (256 * i))

                    'SHARED MEMORY WRITE FINISH CHECK 
                    While True
                        If Me.MemReadWord(iMdfyMemBaseAddr + &H462, 1, sReadData1) = 1 Then Return 1 'MEMORY_1KB_WRITE REG
                        If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STAUS REG
                        If sReadData1(0) = &H0 Then Exit While
                        If sReadData2(0) <> &H0 Then
                            Return 1
                        End If
                    End While

                    'SET 0076046A "0000" timeout1 flag
                    Me.MemWriteWord2(iMdfyMemBaseAddr + &H46A, &H0)    ' 0x46Aは存在しているのか???????????
                Next i
            End If

            '---READ NIOS_PROCESS_FLAG_WORD---
            If Me.MemReadWord(iMdfyMemBaseAddr + &H476, 1, sReadData1) = 1 Then Return 1 'RAM WRITE EXECTUTE STOP REG (MCU内蔵RAM転送確認)
            If Me.MemReadWord(iMdfyMemBaseAddr + &H46C, 1, sReadData2) = 1 Then Return 1 'COPY STATUS REG
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData3) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

            'RAM WRITE EXECTUTE STOP REG (MCU内蔵RAM転送確認)
            If sReadData1(0) = &H1 Then
                Exit While

                'COPY STATUS REG 確認
            ElseIf sReadData2(0) = &H1 Then
                'FLAME DOWN DATA CMD, DATA BYTE LENGTH, DATA WORD LENGTH SEND
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500, frame_DOWN_CMD(0))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H1, frame_DOWN_CMD(1))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H2, frame_DOWN_CMD(2))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H3, frame_DOWN_CMD(3))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H4, frame_DOWN_CMD(4))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H5, frame_DOWN_CMD(5))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H6, frame_DOWN_CMD(6))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H7, frame_DOWN_CMD(7))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H8, frame_DOWN_CMD(8))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &H9, frame_DOWN_CMD(9))
                Me.MemWriteByte(iMdfyMemBaseAddr + &H500 + &HA, CByte(MCUFlashWrFirm.Sum))  'SUM値設定
                Me.MemWriteWord(iMdfyMemBaseAddr + &H500 + &HC, CShort(MCUFlashWrFirm.ByteFile.Length))
                Me.MemWriteWord(iMdfyMemBaseAddr + &H500 + &HE, CShort(MCUFlashWrFirm.WordFile.Length))

                'CLEAR COPY STATUS REG ((0x0000_046C番地へ0x0000をライト)
                Me.MemWriteWord2(iMdfyMemBaseAddr + &H46C, &H0)

                'TIME OUT STATUS REGISTER 確認
            ElseIf sReadData3(0) <> 0 Then
                Return 1
            End If
        End While


        '=====MCUFlashWrFirmの書き込み完了確認====
        While True '無限ループ
            '---EXECUTE FINISH CHECK---
            If Me.MemReadWord(iMdfyMemBaseAddr + &H460, 1, sReadData1) = 1 Then Return 1 'SET UART WRITE START FLAG (書き込み完了確認) 
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER  (タイムアウト検出)
            If sReadData1(0) = 0 Then Exit While '書き込み完了でループを抜ける
            If sReadData2(0) <> 0 Then Return 1 'NiosIIファームのタイムアウト検出でエラー終了
        End While

        'STEP 3. CLEAR TIME OUT STATUS REG  (0x00000468番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)    ' TIME OUT STATUS REG 0クリア

        'STEP 4. OK, CLEAR RAM WRITE EXECTUTE STOP REG 
        Me.MemWriteWord(iMdfyMemBaseAddr + &H476, &H0)

        '===========100ms待つ===========
        System.Threading.Thread.Sleep(100)       '100msの根拠は不明(実機確認ベース)

        '*****正常終了*****
        Return 0

    End Function

    '******************************************************************************
    '*メソッド名：MCUFlashErase
    '*説明：MCU内蔵FLASH-ROMをERASEする
    '*引数：ByVal iSlot As Integer モジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Public Function MCUFlashErase(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim iMdfyMemBaseAddr As Integer 'モディファイドモジュールアドレス先頭番地----
        Dim sReadData1(0) As Short      'リード値
        Dim sReadData2(0) As Short      'リード値


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        '===========wait 1ms'===========
        System.Threading.Thread.Sleep(10)
        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)


        '***********************************************************************************************************************
        'Windows Adress REG Setting
        'offset Adress 0x00000000
        '***********************************************************************************************************************
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000  '


        '***********************************************************************************************************************
        'MCU内蔵ROM ERASE Process 
        '***********************************************************************************************************************
        'STEP 1. MCU内蔵ROM Erase Start Register (0x0000_0478番地に0x0001をライト)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H478, &H1) 'ROM ERASE START

        'STEP2. WAIT UNTIL 0076047A "0001" ERASE STOP FLAG
        While True
            'READ ERASE STOP FLAG
            If Me.MemReadWord(iMdfyMemBaseAddr + &H47A, 1, sReadData1) = 1 Then Return 1 'ROM_ERASE_STOP
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468
            If sReadData1(0) = &H1 Then Exit While
            If sReadData2(0) <> &H0 Then
                Return 1
            End If
        End While

        'STEP 3. CLEAR TIME OUT STATUS REG  (0x00000468番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)    ' TIME OUT STATUS REG 0クリア

        'STEP 4. OK, CLEAR ERASE STOP REG 
        Me.MemWriteWord(iMdfyMemBaseAddr + &H47A, &H0)

        '*****正常終了*****
        Return 0

    End Function

    Public Function MCUFlashWrite(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim i As Integer                'カウンタ変数
        Dim iMdfyMemBaseAddr As Integer 'モディファイドモジュールアドレス先頭番地----
        Dim sReadData1(0) As Short      'リード値
        Dim sReadData2(0) As Short      'リード値


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        '===========wait 1ms'===========
        System.Threading.Thread.Sleep(10)
        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)
        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000  'Shared Memory BASE Address


        '1. SET 0076047C "0001" WRITE START FLAG
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H47C, &H1)

        '2. WRITE MCU BYTE LENGTH
        Dim lwLen As LongWordUnion
        lwLen.s32_0 = MCUFirm.ByteFile.Length
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H51C, CUShort(lwLen.s16_0))
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H51E, CUShort(lwLen.s16_1))

        '3. WAIT UNTIL 0076047E "0001" WRITE STOP FLAG

        Dim DUMP_1K As Integer = 0
        Dim DUMP_256B As Integer = 0
        Dim remainder As Integer = 0
        DUMP_1K = (MCUFirm.ByteFile.Length - 1) \ 1024
        DUMP_256B = (MCUFirm.ByteFile.Length - 1 - (1024 * DUMP_1K)) \ 256


        'SHARED MEMORY 256BYTE WRITE
        For i = 0 To DUMP_1K Step 1 '1KB씩 Window address변경'139
            Application.DoEvents()
            remainder = MCUFirm.LwFile.Length - (256 * i)

            If remainder < 128 Then
                Me.MemWriteWord_val(remainder, iMdfyMemBaseAddr, MCUFirm.LwFile(256 * i))
            ElseIf remainder > 128 And remainder < 256 Then
                Me.MemWriteWord_512B_upper(iMdfyMemBaseAddr, MCUFirm.LwFile, (256 * i))
                Me.MemWriteWord_val(remainder - 128, iMdfyMemBaseAddr, MCUFirm.LwFile(256 * i))

            Else
                Me.MemWriteWord_512B_upper(iMdfyMemBaseAddr + &H50, MCUFirm.LwFile, (256 * i))
                Me.MemWriteWord_512B_lower(iMdfyMemBaseAddr + &H250, MCUFirm.LwFile, (256 * i))
            End If


            'SET 00760462 "0001" 256BTYE SHARED MEMORY FLAG
            Me.MemWriteWord2(iMdfyMemBaseAddr + &H462, &H1)


            While True
                'SHARED MEMORY WRITE FINISH CHECK 
                If Me.MemReadWord(iMdfyMemBaseAddr + &H462, 1, sReadData1) = 1 Then Return 1 '
                If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

                '   無限ループから抜ける
                If sReadData1(0) = &H0S Then Exit While

                'NiosIIファーム側がタイムアウトしたらエラー終了
                If sReadData2(0) <> &H0S Then Return 1 'TIME OUT STATUS REGISTER
            End While
            'SET 0076046A "0000" timeout1 flag
            Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)
            Application.DoEvents()
        Next i



        While True
            'SHARED MEMORY WRITE FINISH CHECK 
            If Me.MemReadWord(iMdfyMemBaseAddr + &H47E, 1, sReadData1) = 1 Then Return 1 '
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

            ' がONしたら無限ループから抜ける
            If sReadData1(0) = &H1S Then Exit While

            'NiosIIファーム側がタイムアウトしたらエラー終了
            If sReadData2(0) <> &H0S Then Return 1 'TIME OUT STATUS REGISTER
        End While

        '4. SET 00760468 "0000" timeout flag
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)

        '5. OK, THEN 0076047E "0000" SET WRITE STOP FLAG 
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H47E, &H0)


        Return 0
    End Function
    '---------------------------------- Ver 0.35 JDI End ------------------------------------
    ' Make By YMFK 製品技術 JIK

    Public Function H8SFlashWrite(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim i As Integer                'カウンタ変数
        Dim iMdfyMemBaseAddr As Integer 'モディファイドモジュールアドレス先頭番地----
        Dim sReadData1(0) As Short      'リード値
        Dim sReadData2(0) As Short      'リード値


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        '===========wait 1ms'===========
        System.Threading.Thread.Sleep(10)
        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)
        '===========window regester setting===========
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000  'Shared Memory BASE Address


        '1. WRITE MCU BYTE LENGTH
        Dim lwLen As LongWordUnion
        lwLen.s32_0 = MCUFirm.ByteFile.Length
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H600, CUShort(lwLen.s16_0))
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H602, CUShort(lwLen.s16_1))


        '2. SET 00760488 "0001" WRITE START FLAG
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H488, &H1)


        '3. WAIT UNTIL 0076047E "0001" WRITE STOP FLAG

        Dim DUMP_1K As Integer = 0
        Dim DUMP_256B As Integer = 0
        Dim remainder As Integer = 0
        DUMP_1K = (MCUFirm.ByteFile.Length - 1) \ 1024
        DUMP_256B = (MCUFirm.ByteFile.Length - 1 - (1024 * DUMP_1K)) \ 256


        'SHARED MEMORY 256BYTE WRITE
        For i = 0 To DUMP_1K Step 1 '1KB씩 Window address변경'139
            Application.DoEvents()
            remainder = MCUFirm.LwFile.Length - (256 * i)

            If remainder < 128 Then
                Me.MemWriteWord_val(remainder, iMdfyMemBaseAddr, MCUFirm.LwSwapFile(256 * i))
            ElseIf remainder > 128 And remainder < 256 Then
                Me.MemWriteWord_512B_upper(iMdfyMemBaseAddr, MCUFirm.LwSwapFile, (256 * i))
                Me.MemWriteWord_val(remainder - 128, iMdfyMemBaseAddr, MCUFirm.LwSwapFile(256 * i))

            Else
                Me.MemWriteWord_512B_upper(iMdfyMemBaseAddr + &H50, MCUFirm.LwSwapFile, (256 * i))
                Me.MemWriteWord_512B_lower(iMdfyMemBaseAddr + &H250, MCUFirm.LwSwapFile, (256 * i))
            End If


            'SET 00760462 "0001" 256BTYE SHARED MEMORY FLAG
            Me.MemWriteWord2(iMdfyMemBaseAddr + &H490, &H1)


            While True
                'SHARED MEMORY WRITE FINISH CHECK 
                If Me.MemReadWord(iMdfyMemBaseAddr + &H490, 1, sReadData1) = 1 Then Return 1 '
                If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

                '   無限ループから抜ける
                If sReadData1(0) = &H0S Then Exit While

                'NiosIIファーム側がタイムアウトしたらエラー終了
                If sReadData2(0) <> &H0S Then Return 1 'TIME OUT STATUS REGISTER
            End While
            'SET 0076046A "0000" timeout1 flag
            Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)
            Application.DoEvents()
        Next i



        While True
            'SHARED MEMORY WRITE FINISH CHECK 
            If Me.MemReadWord(iMdfyMemBaseAddr + &H48C, 1, sReadData1) = 1 Then Return 1 '
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468

            ' がONしたら無限ループから抜ける
            If sReadData1(0) = &H1S Then Exit While

            'NiosIIファーム側がタイムアウトしたらエラー終了
            If sReadData2(0) <> &H0S Then Return 1 'TIME OUT STATUS REGISTER
        End While

        '4. SET 00760468 "0000" timeout flag
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)

        '5. OK, THEN 0076047E "0000" SET WRITE STOP FLAG 
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H48C, &H0)


        Return 0
    End Function

    '******************************************************************************
    '*メソッド名：H8S(MCU) Flash Initial setting
    '*説明：MCU内蔵FLASH-ROMのBootMode Initial settingする
    '*引数：ByVal iSlot As Integer モジュールのスロット番号を指定
    '*返り値: 0      エラーなし
    '*        0以外  エラー
    '******************************************************************************
    Public Function H8SFlashInitial(ByVal iSlot As Integer) As Integer
        '*******変数宣言*******
        Dim iMdfyMemBaseAddr As Integer 'モディファイドモジュールアドレス先頭番地----
        Dim sReadData1(0) As Short      'リード値
        Dim sReadData2(0) As Short      'リード値


        '***********************************************************************************************************************
        'MODIFY RESGSTER SETTING
        'GET MODIFIED MEMORY AREA BASE ADDRESS 
        '***********************************************************************************************************************
        '===========MODIFY REGESTER SETTING===========
        If SetMdfyMdlAddrReg(iSlot) = 1 Then Return 1
        '===========wait 1ms'===========
        System.Threading.Thread.Sleep(10)
        '===========モディファイ空間の先頭アドレス取得===========
        GetMdfyMemBaseAddr(iSlot, iMdfyMemBaseAddr)


        '***********************************************************************************************************************
        'Windows Adress REG Setting
        'offset Adress 0x00000000
        '***********************************************************************************************************************
        SetWndAddrReg(iSlot, &H0I) ' + offset 0x0000_0000  '


        '***********************************************************************************************************************
        'MCU内蔵ROM Boot Mode Initial Process 
        '***********************************************************************************************************************
        'STEP 1. MCU内蔵ROM Initial setting Start Register (0x0000_0478番地に0x0001をライト)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H480, &H1) 'ROM ERASE START

        'STEP2. WAIT UNTIL 00760484 "0001" Initial setting STOP FLAG
        While True
            'READ ERASE STOP FLAG
            If Me.MemReadWord(iMdfyMemBaseAddr + &H484, 1, sReadData1) = 1 Then Return 1 'ROM_Initial setting_STOP
            If Me.MemReadWord(iMdfyMemBaseAddr + &H468, 1, sReadData2) = 1 Then Return 1 'TIME OUT STATUS REGISTER 0x468
            If sReadData1(0) = &H1 Then Exit While
            If sReadData2(0) <> &H0 Then
                Return 1
            End If
        End While

        'STEP 3. CLEAR TIME OUT STATUS REG  (0x00000468番地0クリア)
        Me.MemWriteWord2(iMdfyMemBaseAddr + &H468, &H0)    ' TIME OUT STATUS REG 0クリア

        'STEP 4. OK, CLEAR ERASE STOP REG 
        Me.MemWriteWord(iMdfyMemBaseAddr + &H484, &H0)

        '*****正常終了*****
        Return 0

    End Function
    '---------------------------------- Ver 0.35 JDI End ------------------------------------

    '=========================================================================================
    '=
    '=ファームデータ管理
    '=
    '=========================================================================================
    Public Sub ReadFactImageNiosFirm(ByVal strFileName As String)
        '****ファームウェアファイルの読み出し****
        FactImageNiosFirm = New BinFileCls(strFileName)
        If FactImageNiosFirm.Err = 0 Then FactImageNiosFirm.Set2KByteBoundary() '配列の要素数を2Kの倍数にする(FPGA内部RAMへ転送する単位が2Kbyteなのでそれにあわせておく)---
    End Sub


    Public Sub ReadMCUFlashWrFirm(ByVal strFileName As String)
        MCUFlashWrFirm = New BinFileCls(strFileName)
    End Sub

    Public Sub ReadAppImageRPD(ByVal strFileName As String)
        AppImageRPD = New BinFileCls(strFileName)
    End Sub

    Public Sub ReadMCUFirm(ByVal strFileName As String)
        MCUFirm = New BinFileCls(strFileName)
    End Sub


    '******************************************************************************
    '*プロパティ:
    '*説明： ファーム関連
    '******************************************************************************
    Public ReadOnly Property FactImageNiosFirmErr() As Integer
        Get
            Return FactImageNiosFirm.Err
        End Get
    End Property
    Public ReadOnly Property FactImageNiosFirm_u16Sum() As UInt16
        Get
            Return CUShort(FactImageNiosFirm.u16_Sum)
        End Get
    End Property


    Public ReadOnly Property MCUFlashWrFirmErr() As Integer
        Get
            Return MCUFlashWrFirm.Err
        End Get
    End Property

    Public ReadOnly Property MCUFlashWrFirm_u16Sum() As UInt16
        Get
            Return CUShort(MCUFlashWrFirm.u16_Sum)
        End Get
    End Property
    Public ReadOnly Property AppImageRPDErr() As Integer
        Get
            Return AppImageRPD.Err
        End Get
    End Property

    Public ReadOnly Property AppImage_u16Sum() As UInt16
        Get
            Return CUShort(AppImageRPD.u16_Sum)
        End Get
    End Property
    Public ReadOnly Property MCUFirmErr() As Integer
        Get
            Return MCUFirm.Err
        End Get
    End Property
    Public ReadOnly Property MCUFirm_u16Sum() As UInt16
        Get
            Return CUShort(MCUFirm.u16_Sum)
        End Get
    End Property

    '=========================================================================================
    '=
    '=プロパティ： FAM3_IO_PFM_CLS Revision
    '=
    '=========================================================================================
    Public ReadOnly Property iIoPfmClsRev() As Integer
        Get
            Return _iIoPfmClsRev
        End Get
    End Property

End Class
