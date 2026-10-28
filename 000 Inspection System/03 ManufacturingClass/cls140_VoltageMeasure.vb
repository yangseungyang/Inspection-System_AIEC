


Public Class cls140_VoltageMeasure

    Dim cDMM As cls101_ComRS232C
    Dim cADO As cls010_MDB_Operation
    Dim cLnk1 As cls080_F3LinkToLE

    Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim rs232cDMMPort As String = "COM3"        'DMM7555のRS232Cﾎﾟｰﾄ
    Dim mErrorComment As String

    Dim mSerialPort As New System.IO.Ports.SerialPort
    Dim RecvDataLE As String

    Public Function VoltMeasure(ByVal ch As String) As Single

        'ﾘﾚｰ設定ON
        Call cLnk1.F3PCLinkCommunication("01" & "BWRY" & "50" & ch & ",001,1", RecvDataLE)
        If cLnk1.ERR_COMMENT <> "" Then
            mErrorComment = cLnk1.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        End If

        'DMM測定
        System.Threading.Thread.Sleep(700)         '700msecのWait
        VoltMeasure = DMM_7555_Voltage_Measure()

        'ﾘﾚｰ設定OFF
        Call cLnk1.F3PCLinkCommunication("01" & "BWRY" & "50" & ch & ",001,0", RecvDataLE)
        If cLnk1.ERR_COMMENT <> "" Then
            mErrorComment = cLnk1.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        End If

        System.Threading.Thread.Sleep(500)         '0.5秒のWait

    End Function

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
            mErrorComment = "機器構成一覧に設定が登録されていません。"
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

    Private Function DMM_7555_Voltage_Measure() As Single

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        DMM_7555_Voltage_Measure = Nothing

        Try
            'DMM設定
            cDMM = New cls101_ComRS232C
            cDMM.RS_COMM = mSerialPort
            cDMM.PORT = rs232cDMMPort
            If RS232CSettingRead(rs232cDMMPort, mSetting, mDelim) = False Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
            cDMM.SETTING = mSetting : cDMM.TERMINAL_CHAR = mDelim
            cDMM.PortSetup()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'Header del
            cDMM.SendCmd = "COMM:HEAD OFF" : Call cDMM.FuncSendCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'auto range off
            cDMM.SendCmd = "SENS:FUNC:VOLT:DC:AUTO OFF" : Call cDMM.FuncSendCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            '20V range
            cDMM.SendCmd = "SENS:FUNC:VOLT:DC:RANG 2E+1V" : Call cDMM.FuncSendCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'Measure
            cDMM.SendCmd = "MEAsure:VOLT:DC?" : Call cDMM.FuncSendCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'Receve
            Call cDMM.FuncRecvCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            DMM_7555_Voltage_Measure = CSng(cDMM.RECEIVE_DATA)

        Catch
            If mErrorComment = "" Then mErrorComment = "DMM7555測定異常"
            Throw New ApplicationException(mErrorComment)

        Finally
            Call cDMM.PortClose()
            cDMM = Nothing

        End Try

    End Function

    Public Function DMM_7555_Resistance_Measure() As Single

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        DMM_7555_Resistance_Measure = Nothing

        Try
            'DMM設定
            cDMM = New cls101_ComRS232C
            cDMM.RS_COMM = mSerialPort
            cDMM.PORT = rs232cDMMPort
            If RS232CSettingRead(rs232cDMMPort, mSetting, mDelim) = False Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
            cDMM.SETTING = mSetting : cDMM.TERMINAL_CHAR = mDelim
            cDMM.PortSetup()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'Header del
            cDMM.SendCmd = "COMM:HEAD OFF" : Call cDMM.FuncSendCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'auto range 
            cDMM.SendCmd = "SENS:FUNC:RES:AUTO ON" : Call cDMM.FuncSendCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'Measure
            cDMM.SendCmd = "MEAsure:RES?" : Call cDMM.FuncSendCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            'Receve
            Call cDMM.FuncRecvCmd()
            If cDMM.ERR_COMMENT <> "" Then
                mErrorComment = cDMM.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            DMM_7555_Resistance_Measure = CSng(cDMM.RECEIVE_DATA)

        Catch
            If mErrorComment = "" Then mErrorComment = "DMM7555測定異常"
            Throw New ApplicationException(mErrorComment)

        Finally
            Call cDMM.PortClose()
            cDMM = Nothing

        End Try


    End Function


    Public Sub New()

        'ﾘﾚｰ設定（GND）
        cLnk1 = New cls080_F3LinkToLE
        cLnk1.IPAdd = IP_U1
        Call cLnk1.F3PCLinkCommunication("01" & "BWRY" & "509" & ",001,1", RecvDataLE)   'GND側設定
        If cLnk1.ERR_COMMENT <> "" Then
            mErrorComment = cLnk1.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        End If

    End Sub

    Protected Overrides Sub Finalize()

        'ﾘﾚｰ設定（GND）OFF
        Call cLnk1.F3PCLinkCommunication("01" & "BWRY" & "509" & ",001,0", RecvDataLE)   'GND側設定
        cLnk1 = Nothing

        MyBase.Finalize()

    End Sub
End Class
