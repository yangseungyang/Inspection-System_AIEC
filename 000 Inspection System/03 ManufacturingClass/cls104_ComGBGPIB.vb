Option Strict On : Option Explicit On

'Imports c = CommonCommunicationLibrary



Public Class cls104_ComGBGPIB
    '----------------------------------------
    '   定数宣言
    '----------------------------------------
    'Dim cADO As cls010_MDB_Operation
    'Dim cLnk As cls080_F3LinkToLE

    Dim IP_U2 As String = "192.168.1.2"
    Dim mGBslot As Integer = 12          '4->12

    Private COMM_OK As Integer = 0
    Private COMM_ERR As Integer = 1

    '----------------------------------------
    '   ローカル変数
    '----------------------------------------
    Private mADDRESS As String
    Private mTerminalChar As String
    Private mErrorComment As String

    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim StartTime As DateTime

    Private myErrFlag As Boolean

    Public Property ADDRESS() As String
        Set(ByVal Value As String)
            mADDRESS = Value
        End Set
        Get
            ADDRESS = mADDRESS
        End Get
    End Property

    'Public Property PORT() As String
    '    Set(ByVal Value As String)
    '        mPort = Value
    '    End Set
    '    Get
    '        PORT = mPort
    '    End Get
    'End Property

    'Public Property SETTING() As String
    '    Set(ByVal Value As String)
    '        mSetting = Value
    '    End Set
    '    Get
    '        SETTING = mSetting
    '    End Get
    'End Property

    Public Property TERMINAL_CHAR() As String
        Set(ByVal Value As String)
            mTerminalChar = Value
        End Set
        Get
            TERMINAL_CHAR = mTerminalChar
        End Get
    End Property

    'Error ----------------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public Function SettingRead() As Boolean

        'Dim dbName As String = ""
        'Dim tbName As String = ""
        'Dim fName() As String = {""}
        'Dim sData() As String = {""}
        'Dim sSQL As String = ""
        'Dim fCount As Integer = 0
        'Dim rCount As Integer = 0

        'If mADDRESS = "" And mPort = "" Then
        '    SettingRead = False
        '    Exit Function
        'End If

        'cADO = New cls010_MDB_Operation
        'dbName = mAppPath & "010 EquipmentData\機器構成一覧.xls"
        'If mADDRESS <> "" Then
        '    tbName = "[GPIB$]"
        '    sSQL = "SELECT 設定 , デリミタ FROM " & tbName & " WHERE アドレス LIKE '" & mADDRESS & "'"
        'ElseIf mPort <> "" Then
        '    tbName = "[RS232C$]"
        '    sSQL = "SELECT 設定 , デリミタ FROM " & tbName & " WHERE PORT LIKE '" & mPort & "'"
        'End If

        'cADO.DB_PATH = dbName
        'cADO.TB_NAME = tbName
        'cADO.DB_SQL = sSQL

        'cADO.DB_READ_EXCEL_JET(fName, sData)

        'If cADO.DB_DATA_COUNT = 0 Then
        '    SettingRead = False
        '    cADO = Nothing
        '    mErrorComment = "機器構成一覧に設定が登録されていません。"
        '    Exit Function
        'Else
        '    mSetting = sData(0)
        '    mTerminalChar = sData(1)
        '    SettingRead = True
        'End If

        'cADO = Nothing

    End Function

    Public Sub Open()

        GpibIFMessageCom("IFC")
        GpibIFMessageCom("REN")

    End Sub

    Public Sub Close()

    End Sub

    Public Sub Send(ByVal cmd As String)

        Dim RecvDataLE As String = ""
        Dim cs() As Char = cmd.ToCharArray
        Array.Resize(Of Char)(cs, cmd.Length + 1)
        Dim j As Integer = 0
        'Dim i As Integer

        Dim cLnk1 As New cls080_F3LinkToLE
        cLnk1.IPAdd = IP_U2      'SP7x

        'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0101,01,0001", RecvDataLE)                                    '受信機器ｱﾄﾞﾚｽ数の設定
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0102,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)        '受信機器のGP-IBｱﾄﾞﾚｽ設定

        'Char型の配列に変換する
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0132,01," & cmd.Length.ToString("X4"), RecvDataLE)            '送信ﾃｷｽﾄﾃﾞｰﾀのﾊﾞｲﾄ数
        Debug.Print("01SWW" & mGBslot.ToString("D3") & ",0132,01," & cmd.Length.ToString("X4"))
        For i = 1 To Math.Ceiling(cmd.Length / 2)
            Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",01" & (32 + i).ToString & ",01," _
                                            & Asc(cs(j)).ToString("X2") & Asc(cs(j + 1)).ToString("X2"), RecvDataLE)                        '送信ﾃﾞｰﾀ
            Debug.Print("01SWW" & mGBslot.ToString("D3") & ",01" & (32 + i).ToString & ",01," _
                                            & Asc(cs(j)).ToString("X2") & Asc(cs(j + 1)).ToString("X2"))
            j += 2
        Next
        
        Try
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString("D3") & "35,001,1", RecvDataLE)                                    'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信要求
            StartTime = Now
            Do
                If (Now - StartTime).Seconds > 5 Then
                    mErrorComment = "ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認ﾀｲﾑｱｳﾄ"
                    Throw New ApplicationException(mErrorComment)
                Else
                    'System.Threading.Thread.Sleep(50)
                    My.Application.DoEvents()
                End If
                Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString("D3") & "03,001", RecvDataLE)                                  'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認。完了すると1
            Loop Until Mid(RecvDataLE, 5, 1) = "1"

            Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString("D3") & "07,001", RecvDataLE)                                      'ｴﾗｰ発生確認
            If Mid(RecvDataLE, 5, 1) = "1" Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信要求異常"
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString("D3") & "35,001,0", RecvDataLE)            'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信要求解除
        End Try

    End Sub

    Public Sub Receive(ByRef Buff As String)

        Dim RecvDataLE As String = ""
        Dim j As Integer = 0
        Dim sb As New System.Text.StringBuilder

        Dim cLnk1 As New cls080_F3LinkToLE
        cLnk1.IPAdd = IP_U2      'SP7x

        'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ受信
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0651,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)  '送信機器ｱﾄﾞﾚｽの設定
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0652,01,0001", RecvDataLE)        '受信機器数の設定
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0653,01,0000", RecvDataLE)        '受信機器GP-IBｱﾄﾞﾚｽの設定
        
        Try
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString("D3") & "36,001,1", RecvDataLE)
            StartTime = Now
            Do
                If (Now - StartTime).Seconds > 5 Then
                    mErrorComment = "ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認ﾀｲﾑｱｳﾄ"
                    Throw New ApplicationException(mErrorComment)
                Else
                    'System.Threading.Thread.Sleep(50)
                    My.Application.DoEvents()
                End If
                Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "04,001", RecvDataLE)
            Loop Until Mid(RecvDataLE, 5, 1) = "1"

            Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "07,001", RecvDataLE)          'ｴﾗｰ発生確認
            If Mid(RecvDataLE, 5, 1) = "1" Then Throw New ApplicationException(mErrorComment)
            Call cLnk1.F3PCLinkCommunication("01SWR" & mGBslot.ToString("D3") & ",0683,01", RecvDataLE)             '受信ﾃﾞｰﾀﾊﾞｲﾄ数
            Dim recByte As Integer = Convert.ToInt32(Mid(RecvDataLE, 5, 4), 16)
            'Dim recText As String = ""
            For i = 1 To Math.Ceiling(recByte / 2)
                Call cLnk1.F3PCLinkCommunication("01SWR" & mGBslot.ToString("D3") & ",06" & (83 + i).ToString & ",01", RecvDataLE)
                'recText &= Chr(Convert.ToInt32(Mid(RecvDataLE, 5, 2), 16)) & Chr(Convert.ToInt32(Mid(RecvDataLE, 7, 2), 16))
                sb.Append(Chr(Convert.ToInt32(Mid(RecvDataLE, 5, 2), 16))) : sb.Append(Chr(Convert.ToInt32(Mid(RecvDataLE, 7, 2), 16)))
            Next
            Buff = sb.ToString   ' recText
        Catch ex As Exception
            mErrorComment = "ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ受信要求異常"
            Throw New ApplicationException(mErrorComment)
        Finally
            sb = Nothing
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString & "36,001,0", RecvDataLE)
        End Try

        cLnk1 = Nothing

    End Sub

    Public Sub SendReceive(ByVal cmd As String, ByRef Buff As String)

        Me.Send(cmd)
        Me.Receive(Buff)

    End Sub


    Private Function GpibIFMessageCom(ByVal mSendmessage As String) As Boolean

        Dim RecvDataLE As String = ""

        Dim cLnk1 As New cls080_F3LinkToLE
        cLnk1.IPAdd = IP_U2      'SP7x

        'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ送信（IFC）
        Select Case mSendmessage
            Case "IFC"
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0051,01,0001", RecvDataLE)        'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ番号設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0052,01,0000", RecvDataLE)        '受信機器数の設定
                '終端文字の設定
                Select Case mTerminalChar
                    Case vbCrLf, String.Empty
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0001,01,0000", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0004,01,0000", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0002,01,0000", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0005,01,0000", RecvDataLE)
                    Case vbCr           'CR
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0001,01,0001", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0004,01,0001", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0002,01,0000", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0005,01,0000", RecvDataLE)
                    Case Chr(3)         'ETX
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0001,01,0002", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0004,01,0002", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0002,01,0000", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0005,01,0000", RecvDataLE)
                    Case Else           '終端文字指定
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0001,01,0003", RecvDataLE)
                        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0004,01,0003", RecvDataLE)

                End Select

            Case "REN"
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0051,01,0004", RecvDataLE)        'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ番号設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0052,01,0001", RecvDataLE)        '受信機器数の設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0053,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)        '受信機器のGP-IBｱﾄﾞﾚｽ設定。
            Case "DCL"
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0051,01,0001", RecvDataLE)        'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ番号設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0052,01,0001", RecvDataLE)        '受信機器数の設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0053,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)        '受信機器のGP-IBｱﾄﾞﾚｽ設定。
            Case "GET"
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0051,01,0003", RecvDataLE)        'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ番号設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0052,01,0001", RecvDataLE)        '受信機器数の設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0053,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)        '受信機器のGP-IBｱﾄﾞﾚｽ設定。
            Case "SPL"
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0051,01,0007", RecvDataLE)        'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ番号設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0052,01,0001", RecvDataLE)        '受信機器数の設定
                Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0053,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)        '受信機器のGP-IBｱﾄﾞﾚｽ設定。
        End Select
        'Call WriteLog(False, "ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ" & mSendmessage & "送信")
        Try
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString & "34,001,1", RecvDataLE)              'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ送信要求
            StartTime = Now
            Do
                If (Now - StartTime).Seconds > 5 Then
                    mErrorComment = "ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認ﾀｲﾑｱｳﾄ"
                    Throw New ApplicationException(mErrorComment)
                Else
                    'System.Threading.Thread.Sleep(50)
                    My.Application.DoEvents()
                End If
                Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "02,001", RecvDataLE)            'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認。完了すると1
            Loop Until Mid(RecvDataLE, 5, 1) = "1"

            Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "07,001", RecvDataLE)                'ｴﾗｰ発生確認
            If Mid(RecvDataLE, 5, 1) = "1" Then Throw New ApplicationException(mErrorComment)
        Catch ex As Exception
            mErrorComment = "ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ" & mSendmessage & "要求異常"
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString & "34,001,0", RecvDataLE)            'ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ送信要求解除
        End Try
        
        'Call WriteLog(False, "ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ" & mSendmessage & "送信OK")
        GpibIFMessageCom = True

        cLnk1 = Nothing

    End Function

    Private Function GpibDevMessageCom(ByVal mSendmessage As String) As String

        Dim RecvDataLE As String = ""
        Dim cs() As Char = mSendmessage.ToCharArray
        Array.Resize(Of Char)(cs, mSendmessage.Length + 1)
        Dim j As Integer = 0
        Dim sb As New System.Text.StringBuilder

        GpibDevMessageCom = ""

        Dim cLnk1 As New cls080_F3LinkToLE
        cLnk1.IPAdd = IP_U2      'SP7x

        'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0101,01,0001", RecvDataLE)               '受信機器ｱﾄﾞﾚｽ数の設定
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0102,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)                '受信機器のGP-IBｱﾄﾞﾚｽ設定

        'Char型の配列に変換する
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0132,01," & mSendmessage.Length.ToString("X4"), RecvDataLE)           '送信ﾃｷｽﾄﾃﾞｰﾀのﾊﾞｲﾄ数
        Debug.Print("01SWW" & mGBslot.ToString("D3") & ",0132,01," & mSendmessage.Length.ToString("X4"))
        For i = 1 To Math.Ceiling(mSendmessage.Length / 2)
            Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",01" & (32 + i).ToString & ",01," _
                                            & Asc(cs(j)).ToString("X2") & Asc(cs(j + 1)).ToString("X2"), RecvDataLE)                                '送信ﾃﾞｰﾀ
            Debug.Print("01SWW" & mGBslot.ToString("D3") & ",01" & (32 + i).ToString & ",01," _
                                            & Asc(cs(j)).ToString("X2") & Asc(cs(j + 1)).ToString("X2"))
            j += 2
        Next
        
        Try
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString & "35,001,1", RecvDataLE)                  'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信要求
            StartTime = Now
            Do
                If (Now - StartTime).Seconds > 5 Then
                    mErrorComment = "ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認ﾀｲﾑｱｳﾄ"
                    Throw New ApplicationException(mErrorComment)
                Else
                    'System.Threading.Thread.Sleep(50)
                    My.Application.DoEvents()
                End If
                Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "03,001", RecvDataLE)                'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認。完了すると1
            Loop Until Mid(RecvDataLE, 5, 1) = "1"
            Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "07,001", RecvDataLE)                    'ｴﾗｰ発生確認
            If Mid(RecvDataLE, 5, 1) = "1" Then Throw New ApplicationException(mErrorComment)
        Catch ex As Exception
            mErrorComment = "ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信要求異常"
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString("D3") & "35,001,0", RecvDataLE)            'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ送信要求解除
        End Try

        'ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ受信
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0651,01," & CInt(mADDRESS).ToString("D4"), RecvDataLE)  '送信機器ｱﾄﾞﾚｽの設定
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0652,01,0001", RecvDataLE)        '受信機器数の設定
        Call cLnk1.F3PCLinkCommunication("01SWW" & mGBslot.ToString("D3") & ",0653,01,0000", RecvDataLE)        '受信機器GP-IBｱﾄﾞﾚｽの設定
        
        Try
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString("D3") & "36,001,1", RecvDataLE)
            StartTime = Now
            Do
                If (Now - StartTime).Seconds > 5 Then
                    mErrorComment = "ｲﾝﾀﾌｪｰｽﾒｯｾｰｼﾞ送信完了ﾘﾚｰ確認ﾀｲﾑｱｳﾄ"
                    Throw New ApplicationException(mErrorComment)
                Else
                    'System.Threading.Thread.Sleep(50)
                    My.Application.DoEvents()
                End If
                Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "04,001", RecvDataLE)
            Loop Until Mid(RecvDataLE, 5, 1) = "1"

            Call cLnk1.F3PCLinkCommunication("01BRDX" & mGBslot.ToString & "07,001", RecvDataLE)                    'ｴﾗｰ発生確認
            If Mid(RecvDataLE, 5, 1) = "1" Then Throw New ApplicationException(mErrorComment)
            Call cLnk1.F3PCLinkCommunication("01SWR" & mGBslot.ToString("D3") & ",0683,01", RecvDataLE)             '受信ﾃﾞｰﾀﾊﾞｲﾄ数
            Dim recByte As Integer = Convert.ToInt32(Mid(RecvDataLE, 5, 4), 16)
            'Dim recText As String = ""
            For i = 1 To Math.Ceiling(recByte / 2)
                Call cLnk1.F3PCLinkCommunication("01SWR" & mGBslot.ToString("D3") & ",06" & (83 + i).ToString & ",01", RecvDataLE)
                'recText &= Chr(Convert.ToInt32(Mid(RecvDataLE, 5, 2), 16)) & Chr(Convert.ToInt32(Mid(RecvDataLE, 7, 2), 16))
                sb.Append(Chr(Convert.ToInt32(Mid(RecvDataLE, 5, 2), 16))) : sb.Append(Chr(Convert.ToInt32(Mid(RecvDataLE, 7, 2), 16)))
            Next
            GpibDevMessageCom = sb.ToString
        Catch ex As Exception
            mErrorComment = "ﾃﾞﾊﾞｲｽﾒｯｾｰｼﾞ受信要求異常"
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            sb = Nothing
            Call cLnk1.F3PCLinkCommunication("01BWRY" & mGBslot.ToString & "36,001,0", RecvDataLE)
        End Try

        cLnk1 = Nothing

    End Function

    '---------------------------------------------
    '   プライベート関数
    '---------------------------------------------
    Private Sub Wait(ByVal Sec As Integer)
        System.Threading.Thread.Sleep((Sec * 1000))   'Sec秒のWait
    End Sub

    Private Sub Waitms(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)   'mSec秒のWait
    End Sub


End Class
