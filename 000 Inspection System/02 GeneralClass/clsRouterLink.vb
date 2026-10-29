Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : clsRouterLink
'Function   : 바코드 분배 프로그램(Barcode_Router) 연동 - 검사 종료 신호 회신 (A8)
'Language   : VB2019　Framework4.8
'Memo       : ----------------------------- Ver0.05 YSH ADD (클래스 전체 신규) -----------------------------
'             흐름 : Router → 가상 COM(EquipmentSetting.mdb > BarcodeReaderPort) 으로 바코드 수신
'                    → 기존 frm001_Main.SerialPort1_DataReceived 가 그대로 검사 → 끝나면 이 클래스로 "END + CR" 회신
'             - 사용 / 미사용 : EquipmentSetting.mdb > tblEquipmentSetting > RouterLink ("Yes" / "No")
'                               첫 글자가 Y 가 아니거나, 컬럼이 없거나, 읽기 오류면 미사용 (원본 프로그램과 같은 동작)
'                               처음 한 번만 읽고, 결과는 메시지 없이 로그로만 남긴다
'             - 공통 규약 : .claude\projects\Barcode_Router 공통\Barcode_Router 연동 공통 규약.md (Rev 0) 3항
'             - 미사용이면 아무것도 하지 않는다 → 라우터 없이 실제 바코드리더(COM1)로 써도 리더 쪽으로 END 를 보내지 않음
'             - 기존 cls101_ComRS232C.FuncSendCmd 는 WriteLine(LF 종단) 이라 Router(CR 종단)와 맞지 않으므로
'               SerialPort 에 직접 "END" + CR 을 쓴다
'             - END 를 보내기 전에 AllowSetForegroundWindow 로 Router 가 다시 맨 앞으로 올 수 있게 권한을 넘긴다
'             - 로그는 frm001_Main.SysLogWrite 와 같은 방식(cls013_TextFile_Operation)으로 시스템 로그에 직접 쓴다
'Issue      : 2026/10/27  Ver0.05    YSH       New Release
'             근거 : FAM3 증산 대응_CPU검사장치 ADD-ON.md Rev 2  6-4, 6-9, 7-3, 7-4 / CPU검사장치_Router대응_분석_v1.1.xlsx 6_대응계획 No 2
'             2026/10/28  Ver0.05    YSH       사용 / 미사용 모드 (RouterLink 컬럼) 추가, 로그를 클래스 안에서 처리
'             근거 : FAM3 증산 대응_CPU검사장치 ADD-ON.md Rev 7  11-3, 11-7 (P2)
'================================================================================================================================
Imports System.Runtime.InteropServices

Public Class clsRouterLink

    Private Const ROUTER_END As String = "END"          'Router 로 보내는 검사 종료 신호 (Barcode_Router clsRouterPort : CR 종단)
    Private Const ASFW_ANY As Integer = -1

    Private Shared mChecked As Boolean = False          'RouterLink 를 읽었으면 True
    Private Shared mEnabled As Boolean = False          'RouterLink = Yes 이면 True

    <DllImport("user32.dll")>
    Private Shared Function AllowSetForegroundWindow(ByVal dwProcessId As Integer) As Boolean
    End Function

    Public Shared Function IsEnabled(Optional ByVal logPath As String = "") As Boolean
        '==========================================================================================
        'Function : 라우터 사용 여부 (EquipmentSetting.mdb > tblEquipmentSetting > RouterLink)
        'Return   : 값의 첫 글자가 "Y" 이면 True. 컬럼 없음 / 읽기 오류 / 그 밖의 값이면 False (미사용)
        'Remarks  : 공통 규약 3-2, 3-4 - 처음 한 번만 읽고 기억한다. 메시지 없이 결과를 로그로만 남긴다
        '==========================================================================================
        If mChecked Then Return mEnabled

        Dim cADO As cls010_MDB_Operation = Nothing
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sValue As String = ""
        Dim msg As String
        Dim appPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = appPath & "010 EquipmentData\EquipmentSetting.mdb"
            cADO.TB_NAME = "tblEquipmentSetting"
            cADO.DB_SQL = "SELECT RouterLink FROM tblEquipmentSetting"
            Call cADO.DB_READ_JET(fName, sData)

            If cADO.ERR_COMMENT <> "" Then
                msg = "RouterLink : No (read error " & cADO.ERR_COMMENT & ")"
            Else
                If sData IsNot Nothing AndAlso sData.Length > 0 AndAlso sData(0) IsNot Nothing Then sValue = sData(0).Trim
                mEnabled = (sValue.ToUpper.StartsWith("Y"))
                msg = "RouterLink : " & If(mEnabled, "Yes", "No") & " (value '" & sValue & "')"
            End If

        Catch ex As Exception
            mEnabled = False
            msg = "RouterLink : No (column not found / read error " & ex.Message & ")"

        Finally
            cADO = Nothing

        End Try

        mChecked = True
        Call WriteLog(logPath, msg)
        Return mEnabled

    End Function

    Public Shared Sub SendEnd(ByVal port As System.IO.Ports.SerialPort, ByVal logPath As String)
        '==========================================================================================
        'Function : 검사 종료 신호 "END + CR" 을 Router 로 회신 (md 6-9, 11-3)
        'Data     : port = 바코드 수신에 쓰는 SerialPort (frm001_Main.SerialPort1)
        '           logPath = 시스템 로그 파일 (frm001_Main.mSystemLog)
        'Remarks  : RouterLink 미사용이면 아무것도 하지 않는다. 이 Sub 는 예외를 밖으로 내보내지 않는다
        '==========================================================================================
        If Not IsEnabled(logPath) Then Exit Sub

        Dim msg As String
        Try
            If port Is Nothing OrElse Not port.IsOpen Then
                msg = "Router END Send skip : port closed"
            Else
                Call AllowSetForegroundWindow(ASFW_ANY)
                port.Write(ROUTER_END & vbCr)
                msg = "Router END Send OK : " & port.PortName
            End If

        Catch ex As Exception
            msg = "Router END Send NG : " & ex.Message

        End Try

        Call WriteLog(logPath, msg)

    End Sub

    Private Shared Sub WriteLog(ByVal logPath As String, ByVal msg As String)
        '==========================================================================================
        'Function : 시스템 로그 기록 (frm001_Main.SysLogWrite 와 같은 방식)
        '==========================================================================================
        Dim cTxt As cls013_TextFile_Operation = Nothing

        Try
            If logPath Is Nothing OrElse logPath = "" Then Exit Sub
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = logPath
            Call cTxt.WRITE_TEXT_FILE(True, msg)

        Catch
            '로그 실패는 검사에 영향을 주지 않는다

        Finally
            cTxt = Nothing

        End Try

    End Sub

End Class
