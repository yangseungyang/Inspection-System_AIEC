'------------------------------------------------------------------------
'   clsNeptune
'   Copyright(c)  2008 Yokogawa Electric Corporation.
'
'   Comment:  <Neptune(UT55A/52A, UT35A32A) Class>
'   Modify 
'   > Ver 0.00 | 2008/10/13 | New Release                               | K.Minami
'   > Ver 0.86 | 2015/05/08 | MDL Maintenancce통신검사 대응             | Park.C.J
'   > Ver 
'
'------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System
Imports System.Math
Imports Microsoft.VisualBasic
Imports CommonCommunicationLibrary




Public Class clsNeptune
    Inherits clsCommMain
    '----------------------------------------
    '   Const Define
    '----------------------------------------

    Private Const EquipmentName As String = "UTX Series"                                '# 제품명
    Private Const RegisterFilename As String = "C:\Neptune ATE\Register_mapR0.csv"      '# Register명칭／Register No.변환 (csv)FILE Name
    Private Const RegisterFilename_xls As String = "C:\Neptune ATE\Register_mapR0.xls"  '# Register명칭／Register No.변환 (xls) FILE Name
    Private Const Sheetname As String = "Register"
    Private Const intTimeout As Integer = 100                                           '# 송수신 Timeout 시간[ms]
    '                                                                                   '# 항목명칭 검색 범위
    Private Const RowSearchMax As Integer = 100                                         '# 行 검색 Cell 최대치
    Private Const ColumnSearchMax As Integer = 100                                      '# 列 검색 Cell 최대치
    '                                                                                   '# Error Message
    Private Const ErrorMessageItem As String = " ：Register 항목명이 존재하지 않습니다."
    Private Const ErrorMessageCal As String = " ：문자수가 바르게 설정되지 않았습니다."


    '----------------------------------------
    '   Local Variable
    '----------------------------------------
    Private myID As Integer                                  '식별용ID(複数台 사용 시, 필요)

    '----------------------------------------
    '   Constructor ／ Distructor
    '----------------------------------------
    Public Sub New(ByVal id As Integer, ByVal filename As String)
        MyBase.New()
        myID = id
        TargetName = EquipmentName
        Me.Initialize(id, filename)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    '----------------------------------------
    '    Property
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
    '    Enum
    '----------------------------------------
    '필요에 따라 기술


    '----------------------------------------
    '   Method
    '----------------------------------------

    'ＯＰＥＮ
    Public Overrides Sub Open(Optional ByVal Param As String = Nothing)
        MyBase.Open()
        MyBase.Device.SetTimeout(intTimeout)                                'Neptune전용 TimeOut 설정（他기기에는 적용 안함）
        Dim RS As clsCommRs232c = CType(MyBase.Device, clsCommRs232c)       'dll FIie에는 「7bit EVEN」or「8bit non-parity」밖에 
        RS.SetParity(clsCommRs232c.COMM_RS_PARITY.Even)                     '설정이 없기 때문에 OPEN時,「8bit EVEN」으로 설정변경 필요有
    End Sub
    '#------[#Mim ADD : Ethernet用 Neptune OPEN ]------#
    Public Sub EtherOpen()
        MyBase.Open()
    End Sub
    '#-----------------------------------------------#

    '【필요에 따라..】초기화(본체 Reset)
    Public Sub Preset()

    End Sub


    'Neptune과의 Data 송수신
    Public Function TxRxUTX(ByRef strGetData As String, ByVal strHeadDt As String, Optional ByVal strRegName1 As String = "", _
        Optional ByVal strRegName2 As String = "", Optional ByVal strCmdSendDt As String = "", _
        Optional ByVal wtTime As Double = 0.005, Optional ByVal blEcho As Boolean = True, Optional ByVal MDLComm As Boolean = False) As Boolean 'ver0.86chg(MDLComm add)
        'DATA 송수신 예
        '  Neptune에 "0165[TESTMODE]000200040000162E"을 송신하는 경우.
        '      str1stData = "0165"
        '      str2ndData = "TESTMODE"
        '      str3rdData = "000200040000162E"
        '  을 기술함.
        '  Return Value 로써 Neptune으로 부터의 수신 DATA를 strGetData에 격납한다.
        '      strGetData = "0165****0002"
        '      단、"****" 은 [TESTMODE]에 대응하는 Register No.

        '----------------------------------------
        '   초기설정
        '----------------------------------------
        Dim strSendDataString As String
        Dim strSendDataBinary As String = ""
        Dim intDatalength As Integer
        Dim strRegister1 As String = ""
        Dim strRegister2 As String = ""
        Dim strBuff As String
        Dim intBuff As Integer
        Dim setlength As Integer = 1000
        Dim getDataCRC As String = ""
        Dim strcalCRC As String = ""
        Dim flg As Boolean
        '====Ver1.19ADD[S]====
        'If InStr(frmMain.TextBox_MSCode.Text, "UM33A-S") > 0 Then
        '    blEcho = False
        'End If
        '====Ver1.19ADD[E]====

        '#VER0.83ADD[S]+++++++++++++++++++++++++++++++++++++++++
        'If InStr(frmMain.TextBox_MSCode.Text, "/MDL") > 0 Then
        '    blEcho = False
        'End If
        'VER0.83ADD[E]+++++++++++++++++++++++++++++++++++++++++++

        '----------------------------------------
        '   송신DATA（문자열）작성
        '----------------------------------------
        '항목명 ⇒ Register No. 검색
        If strRegName1 <> "" Then
            Call ChangeRegisterName(strRegName1, strRegister1)

            '#------[#: Register Address변환 → Hex(strRegister) ]--#
            strRegister1 = Hex(Convert.ToInt32(strRegister1, 10))
            strRegister1 = strRegister1.PadLeft(4, "0"c)
            '#------------------------------------------------------#
        End If
        If strRegName2 <> "" Then
            Call ChangeRegisterName(strRegName2, strRegister2)

            '#------[#: Register Address변환 → Hex(strRegister) ]--#
            strRegister2 = Hex(Convert.ToInt32(strRegister2, 10))
            strRegister2 = strRegister2.PadLeft(4, "0"c)
            '#------------------------------------------------------#
        End If

        'CRC정보 제외, 송신 DATA 작성
        strBuff = strHeadDt & strRegister1 & strRegister2 & strCmdSendDt
        'Send Data（String）: CRC정보 결합.
        strSendDataString = strBuff & calCRC(strBuff)

        '----------------------------------------
        '   Data Send to.. Neptune.
        '----------------------------------------
        Me.SendBinary(strSendDataString)

        'If blEcho = True Then wtTime = 0.01 '# LL通信時待ち時間

        Call Wait(wtTime)                   '# [clsTestBase]: Public Defined..

        '----------------------------------------
        '   Data Receive From.. Neptune.
        '----------------------------------------
        Me.ReceiveBinary(strBuff, setlength, intBuff, flg)

        If MDLComm = True And strBuff = "" Then    'VER0.86ADD
            strGetData = "NoResponse"  'VER0.86ADD
            Return True           'VER0.86ADD
        End If                    'VER0.86ADD
        '----------------------------------------
        '   CRC비교
        '----------------------------------------
        '#: LL EchoBack処理 ※ LL通信可能であれば『blEcho = True』にOptional Default値変更要!! minami 
        If blEcho = True Then
            strBuff = Mid(strBuff, strSendDataString.Length + 1)
        End If

        intDatalength = Len(strBuff)
        strGetData = Mid(strBuff, 1, (intDatalength - 4))
        getDataCRC = Mid(strBuff, (intDatalength - 3), 4)

        If getDataCRC = calCRC(strGetData) Then
            Return True
        Else
            strGetData = "ERROR!"
            Return False
        End If

    End Function


    'Ethernet경유, Neptune과 Data 송수신
    Public Function SendReceiveE(ByRef strGetData As String, ByVal str1stData As String, _
        Optional ByVal str2ndData As String = "", Optional ByVal str3rdData As String = "", _
        Optional ByVal wtTime As Single = 0.1, Optional ByVal RecevieNone As Boolean = False) As Boolean
        'DATA 송수신 예
        '  Neptune에 "0165[TESTMODE]000200040000162E"을 송신하는 경우.
        '      str1stData = "0165"
        '      str2ndData = "TESTMODE"
        '      str3rdData = "000200040000162E"
        '  을 기술함.
        '  Return Value 로써 Neptune으로 부터의 수신 DATA를 strGetData에 격납한다.
        '      strGetData = "0165****0002"
        '      단、"****" 은 [TESTMODE]에 대응하는 Register No.

        '----------------------------------------
        '   초기설정
        '----------------------------------------
        Dim strSendDataString As String
        Dim strSendDataBinary As String = ""
        Dim strRegister As String = ""
        Dim strBuff As String = ""
        Dim intBuff As Integer
        Dim setlength As Integer = 1000
        Dim getDataCRC As String = ""
        Dim strcalCRC As String = ""
        Dim flg As Boolean

        SendReceiveE = False

        '----------------------------------------
        '   송신Data（String）작성
        '----------------------------------------
        '항목명 ⇒ Register No. 검색
        If str2ndData <> "" Then
            Call ChangeRegisterName(str2ndData, strRegister)
            strRegister = Hex(Convert.ToInt32(strRegister, 10))
            '왼쪽부터 빈공간에 "0"을 채움
            strRegister = strRegister.PadLeft(4, "0"c)
        End If

        '송신 Data（String）CRC정보 없음.
        strSendDataString = str1stData & strRegister & str3rdData

        '----------------------------------------
        '   Data Send to.. Neptune.
        '----------------------------------------
        Me.SendBinary(strSendDataString)
        Call Wait(wtTime)

        '----------------------------------------
        '   Data Receive From.. Neptune.
        '----------------------------------------
        '#[IP Address変更後,受信無し場合の処理]
        If RecevieNone = False Then
            Me.ReceiveBinary(strBuff, setlength, intBuff, flg)
        End If

        '----------------------------------------
        '   Return Value from Neptune
        '----------------------------------------
        strGetData = strBuff
        Return True

    End Function
    '-----------------------------------BSB ADD START--------------------------------------------
    'Ethernet경유, Neptune과 Data 송수신
    Public Function SendReceive2(ByRef strGetData As String, ByVal str1stData As String, _
        Optional ByVal str2ndData As String = "", Optional ByVal str3rdData As String = "", _
        Optional ByVal wtTime As Single = 0.1, Optional ByVal RecevieNone As Boolean = False) As Boolean
        'DATA 송수신 예
        '  Neptune에 "0165[TESTMODE]000200040000162E"을 송신하는 경우.
        '      str1stData = "0165"
        '      str2ndData = "TESTMODE"
        '      str3rdData = "000200040000162E"
        '  을 기술함.
        '  Return Value 로써 Neptune으로 부터의 수신 DATA를 strGetData에 격납한다.
        '      strGetData = "0165****0002"
        '      단、"****" 은 [TESTMODE]에 대응하는 Register No.

        '----------------------------------------
        '   초기설정
        '----------------------------------------
        Dim strSendDataString As String
        Dim strSendDataBinary As String = ""
        Dim strRegister As String = ""
        Dim strBuff As String = ""
        Dim intBuff As Integer
        Dim setlength As Integer = 1000
        Dim getDataCRC As String = ""
        Dim strcalCRC As String = ""
        Dim flg As Boolean

        SendReceive2 = False

        '----------------------------------------
        '   송신Data（String）작성
        '----------------------------------------
        '항목명 ⇒ Register No. 검색
        If str2ndData <> "" Then
            Call ChangeRegisterName(str2ndData, strRegister)
            strRegister = Hex(Convert.ToInt32(strRegister, 10))
            '왼쪽부터 빈공간에 "0"을 채움
            strRegister = strRegister.PadLeft(4, "0"c)
        End If

        '송신 Data（String）CRC정보 없음.
        strSendDataString = str1stData & strRegister & str3rdData

        '----------------------------------------
        '   Data Send to.. Neptune.
        '----------------------------------------
        Me.SendBinary(strSendDataString)
        Call Wait(wtTime)

        '----------------------------------------
        '   Data Receive From.. Neptune.
        '----------------------------------------
        '#[IP Address変更後,受信無し場合の処理]
        If RecevieNone = False Then
            Me.ReceiveBinary(strBuff, setlength, intBuff, flg)
        End If

        '----------------------------------------
        '   Return Value from Neptune
        '----------------------------------------
        strGetData = strBuff
        Return True

    End Function
    '-----------------------------------BSB ADD END--------------------------------------------
    'IEEE754(Float)浮動小数点形式 → 10進数形式변환
    Public Function IEEEtoDECIMAL(ByVal strDataIEEE As String) As Double
        'Parameter [strDataIEEE] : 浮動小数点 String(8文字)
        'Return Value [IEEEtoDECIMAL] : 10進数値(倍精度実数)

        '----------------------------------------
        '   초기설정
        '----------------------------------------
        Dim intPolarity As Integer            '부호部 Variable
        Dim dblExponent As Double        '指数部 Variable
        Dim dblSignificand As Double      '仮数部 Variable

        Dim strHex1 As String
        Dim strHex2 As String
        Dim strHex3 As String

        Dim intExpUpper As Integer
        Dim intExpLower As Integer
        Dim lngSignifUpper As Integer
        Dim lngSignifLower As Long

        '----------------------------------------
        '   형식변환
        '----------------------------------------
        'strDataIEEE 분해
        strHex1 = "&H" & Left(strDataIEEE, 2)         'strDataIEEE 왼쪽２文字
        strHex2 = "&H" & Mid(strDataIEEE, 3, 2)     'strDataIEEE ３번째 부터 2文字
        strHex3 = "&H0" & Mid(strDataIEEE, 5)       'strDataIEEE ５번째 부터 4文字

        '부호(1bit) 추출
        If (CInt(strHex1) And &H80) = 0 Then
            intPolarity = 1     '+ 부호
        Else
            intPolarity = -1    '- 부호
        End If

        '指数部(8bit) 추출
        intExpUpper = CInt(strHex1) And &H7F            '上位7bit
        intExpLower = CInt(strHex2) And &H80            '下位1bit
        dblExponent = intExpUpper * 2 ^ 1 + intExpLower / &H80 - 127    '指数部

        '仮数部(23bit) 추출
        lngSignifUpper = CInt(strHex2) And &H7F         '上位7bit
        lngSignifLower = CLng(strHex3)                    '下位16bit
        dblSignificand = lngSignifUpper * 2 ^ 16 + lngSignifLower       '仮数部

        '10進형식으로 결합
        IEEEtoDECIMAL = intPolarity * (2 ^ dblExponent) * (1 + dblSignificand * 2 ^ -23)

    End Function
    '-----------------------------------BSB ADD START--------------------------------------------
    Public Function IEEEtoDECIMAL_change(ByVal strDataIEEE As String) As Double
        'Parameter [strDataIEEE] : 浮動小数点 String(8文字)
        'Return Value [IEEEtoDECIMAL] : 10進数値(倍精度実数)

        '----------------------------------------
        '   초기설정
        '----------------------------------------
        Dim intPolarity As Integer            '부호部 Variable
        Dim dblExponent As Double        '指数部 Variable
        Dim dblSignificand As Double      '仮数部 Variable

        Dim strHex1 As String
        Dim strHex2 As String
        Dim strHex3 As String

        Dim intExpUpper As Integer
        Dim intExpLower As Integer
        Dim lngSignifUpper As Integer
        Dim lngSignifLower As Long

        '----------------------------------------
        '   형식변환
        '----------------------------------------
        'strDataIEEE 분해
        strHex1 = "&H" & Mid(strDataIEEE, 5, 2)         'strDataIEEE 왼쪽２文字
        strHex2 = "&H" & Mid(strDataIEEE, 7, 2)     'strDataIEEE ３번째 부터 2文字
        strHex3 = "&H0" & Left(strDataIEEE, 4)       'strDataIEEE ５번째 부터 4文字

        '부호(1bit) 추출
        If (CInt(strHex1) And &H80) = 0 Then
            intPolarity = 1     '+ 부호
        Else
            intPolarity = -1    '- 부호
        End If

        '指数部(8bit) 추출
        intExpUpper = CInt(strHex1) And &H7F            '上位7bit
        intExpLower = CInt(strHex2) And &H80            '下位1bit
        dblExponent = intExpUpper * 2 ^ 1 + intExpLower / &H80 - 127    '指数部

        '仮数部(23bit) 추출
        lngSignifUpper = CInt(strHex2) And &H7F         '上位7bit
        lngSignifLower = CLng(strHex3)                    '下位16bit
        dblSignificand = lngSignifUpper * 2 ^ 16 + lngSignifLower       '仮数部

        '10進형식으로 결합
        IEEEtoDECIMAL_change = intPolarity * (2 ^ dblExponent) * (1 + dblSignificand * 2 ^ -23)

    End Function
    '-----------------------------------BSB ADD END--------------------------------------------
    '10進数形式 → IEEE754(Float)浮動小数点形式변환
    Public Function DECIMALtoIEEE(ByVal dblDataDEC As Double) As String
        'Parameter [dblDataDEC] : 10進数値(倍精度実数)
        'Return Value [DECIMALtoIEEE] : 浮動小数点 String(8文字)

        '----------------------------------------
        '   초기설정
        '----------------------------------------
        Dim intPolarity As Integer                 '부호部 Variable
        Dim intExponent As Integer = &H7F    '指数部 Variable
        Dim dblSignificand As Double           '仮数部 Variable
        Dim intExpCount As Integer = 0         '指数部 Count 초기치
        Dim intExpLSB As Integer
        Dim strLower As String = "000000"    '３文字부터 ８文字分
        Dim intUpper As Integer                   '１文字부터２文字分（整数型）
        Dim strUpper As String = "00"           '１文字부터２文字分（文字型）

        '----------------------------------------
        '   형식변환
        '----------------------------------------
        '부호(1bit) 추출
        intPolarity = System.Math.Sign(dblDataDEC)      'Sign은 부호를 표시하는 값으로 Return
        If intPolarity < 0 Then
            intPolarity = 1
        Else
            intPolarity = 0
        End If

        '仮数部 &  指数部 작성
        dblSignificand = System.Math.Abs(dblDataDEC)
        If dblSignificand = 0 Then
            intExponent = 0
        Else
            '仮数部정규화
            'dblsignificand이 １以上２未満이면 빠져나옴.
            Do
                If 2 <= dblSignificand Then
                    dblSignificand = dblSignificand / 2
                    intExpCount = intExpCount + 1
                End If
                If dblSignificand < 1 Then
                    dblSignificand = dblSignificand * 2
                    intExpCount = intExpCount - 1
                End If
            Loop Until (1 <= dblSignificand) And (dblSignificand < 2)

            '仮数部 소수점을 오른쪽 23bit Shift
            dblSignificand = dblSignificand * (2 ^ 23)

            '仮数部가 16進数로 ７桁가 되는 경우、指数部의 桁 올림.
            If Len(Hex(dblSignificand)) = 7 Then intExpCount = intExpCount + 1

            '仮数部의 MSB(2^23에 해당하는 bit. 즉, 24bit번째를 １로 가정해、
            '그 직후에 소수점이 있는 것으로 해석하기 때문에、실제 仮数部는 仮数로 부터
            '１ 뺀 수치가 됨. 따라서、1～23bit 번째 까지가 유효숫자가 된다.
            dblSignificand = CLng(dblSignificand) And &H7FFFFF      '← *** 仮数部数値 ***

            intExpLSB = (intExponent + intExpCount) Mod 2           '指数部 LSB을 구함.
            intExpLSB = CInt(intExpLSB * (2 ^ 23))                      '指数部LSB의 桁으로 이동.
            '指数部LSB=0、仮数部上位bit=0의 경우、24bit(3Byte)만큼의 桁가 확보 불가능하게 됨.
            '대책으로서 좌측에 문자열 0을 ６桁 붙여 우측으로 부터６文字를 추출한다.
            strLower = Right(("000000" & Hex(intExpLSB + dblSignificand)), 6)

            '符号bit 와 指数部上位7bit 처리
            intExponent = (intExponent + intExpCount) \ 2            '指数部上位7bit 桁에 이동
            intUpper = CInt(intPolarity * 2 ^ 7 + intExponent)           '符号bit와 指数部上位7bit결합
            strUpper = Right(("00" & Hex(intUpper)), 2)
        End If

        'Return Value 작성
        If intExpCount > 127 Then
            DECIMALtoIEEE = "Over Flow!!"
        ElseIf intExpCount <= -127 Then
            DECIMALtoIEEE = "00000000"
        Else
            DECIMALtoIEEE = strUpper & strLower
        End If

    End Function

    '항목명（Register명칭） ⇒ Register No. 검색
    '대응표가 Excel File 인 경우
    Public Sub ChangeRegisterName_xls(ByVal strItemName As String, ByRef strRegister As String)
        'strItemName ： 項目名（Register명칭）
        'strRegister ： Register番号

        '----------------------
        '   초기설정
        '----------------------
        'Dim myExcelRW As New clsExcelRW()
        'Dim introw As Integer
        'Dim intcolumn As Integer
        'Dim strResult As String
        Dim SearchOKflug As Boolean = False

        '----------------------
        '   변환처리
        '----------------------
        'myExcelRW.OpenWorkBook(RegisterFilename_xls)    'Register명칭／Register No.변환표(Excel File )OPEN
        'myExcelRW.SelectSheet(Sheetname)                      'Sheet선택
        'Register No. 검색
        '    (1,1) - (RowSearchMax, ColumnSearchMax)  Cell로 부터 검색
        '    Cell 배열은 하기와 같을 것.
        '        奇数列이 Register Name、偶数列이 항목명
        '        항목명에 대응하는 Register Name은 반드시 좌측에 기술한다.
        ''        |Register名Ａ｜項目名Ａ|Register名Ｂ|項目名Ｂ|Register名Ｃ|項目名Ｃ|・・・・・
        'For intcolumn = 2 To ColumnSearchMax Step 2
        '    For introw = 1 To RowSearchMax
        '        strResult = myExcelRW.GetDataFromCell(introw, intcolumn)              '항목명 Read
        '        'System.Diagnostics.Debug.Print(introw & intcolumn)
        '        If strResult = strItemName Then
        '            strRegister = myExcelRW.GetDataFromCell(introw, intcolumn - 1)  'Register No. Read
        '            SearchOKflug = True
        '            Exit For
        '        End If
        '    Next introw
        '    If SearchOKflug = True Then Exit For
        'Next intcolumn
        'Call myExcelRW.CloseWorkBook()       'Excel File(Workbook) Close
        'Call myExcelRW.QuitExcel()                'Excel 종료

        'Error처리
        If SearchOKflug = False Then
            Throw New System.ApplicationException(strItemName & ErrorMessageItem)
        End If

    End Sub

    '항목명(Register 명칭) ⇒ Register No. 검색
    '대응표가 csv File 또는 Text File 인 경우.
    Public Sub ChangeRegisterName(ByVal strItemName As String, ByRef strRegister As String)
        'strItemName ： 항목명（Register명칭）
        'strRegister ： Register No.

        '----------------------
        '   초기설정
        '----------------------
        Dim strRegBuff As String = ""
        Dim strItemBuff As String = ""
        Dim SearchOKflug As Boolean = False
        Dim FileNumber As Integer = FreeFile()

        '----------------------
        '   변환처리
        '----------------------

        strItemName = "PNO_" + strItemName

        FileOpen(FileNumber, RegisterFilename, OpenMode.Input)

        'Register No. 검색
        '    Register No.의 순서는 하기와 같을 것.
        '        항목명은 偶数項으로、대응하는 Register명은 그 좌측의 奇数項에 기술한다.
        '        Register名Ａ, 項目名Ａ, Register名Ｂ, 項目名Ｂ, Register名Ｃ, 項目名Ｃ,・・・・・
        Do While Not EOF(FileNumber)
            Call Input(FileNumber, strRegBuff)
            Call Input(FileNumber, strItemBuff)
            If strItemName = strItemBuff Then
                strRegister = strRegBuff
                SearchOKflug = True
                Exit Do
            End If
        Loop
        Call FileClose(FileNumber)      'txt File or csv File Close

        'Error처리
        If SearchOKflug = False Then
            Throw New System.ApplicationException(strItemName & ErrorMessageItem)
        End If

    End Sub


    'Error Check(CRC-16) 계산
    'entry : strChrs as String
    'exit  : calCRC as String(4 characters)
    Public Function calCRC(ByVal strChrs As String) As String
        '---------------------------------
        '   초기설정
        '---------------------------------
        Dim intCRC As Integer = &HFFFF                  'CRC값.
        Dim intCRC1 As Integer                          '
        Dim intCRC2 As Integer                          '
        Dim strChr2 As String                           '
        Dim strCRC1 As String                           '
        Dim strCRC1l As String                          '
        Dim strCRC1r As String                          '
        Dim strCRC As String                            '
        Dim strCRCl As String                           '
        Dim strCRCr As String                           '
        Dim valLen As Long
        Dim i As Long
        'CRC Table
        Dim CRCTable() As Integer = { _
        &H0, &HC0C1, &HC181, &H140, &HC301, &H3C0, &H280, &HC241, &HC601, &H6C0, &H780, &HC741, &H500, &HC5C1, &HC481, &H440, _
        &HCC01, &HCC0, &HD80, &HCD41, &HF00, &HCFC1, &HCE81, &HE40, &HA00, &HCAC1, &HCB81, &HB40, &HC901, &H9C0, &H880, &HC841, _
        &HD801, &H18C0, &H1980, &HD941, &H1B00, &HDBC1, &HDA81, &H1A40, &H1E00, &HDEC1, &HDF81, &H1F40, &HDD01, &H1DC0, &H1C80, &HDC41, _
        &H1400, &HD4C1, &HD581, &H1540, &HD701, &H17C0, &H1680, &HD641, &HD201, &H12C0, &H1380, &HD341, &H1100, &HD1C1, &HD081, &H1040, _
        &HF001, &H30C0, &H3180, &HF141, &H3300, &HF3C1, &HF281, &H3240, &H3600, &HF6C1, &HF781, &H3740, &HF501, &H35C0, &H3480, &HF441, _
        &H3C00, &HFCC1, &HFD81, &H3D40, &HFF01, &H3FC0, &H3E80, &HFE41, &HFA01, &H3AC0, &H3B80, &HFB41, &H3900, &HF9C1, &HF881, &H3840, _
        &H2800, &HE8C1, &HE981, &H2940, &HEB01, &H2BC0, &H2A80, &HEA41, &HEE01, &H2EC0, &H2F80, &HEF41, &H2D00, &HEDC1, &HEC81, &H2C40, _
        &HE401, &H24C0, &H2580, &HE541, &H2700, &HE7C1, &HE681, &H2640, &H2200, &HE2C1, &HE381, &H2340, &HE101, &H21C0, &H2080, &HE041, _
        &HA001, &H60C0, &H6180, &HA141, &H6300, &HA3C1, &HA281, &H6240, &H6600, &HA6C1, &HA781, &H6740, &HA501, &H65C0, &H6480, &HA441, _
        &H6C00, &HACC1, &HAD81, &H6D40, &HAF01, &H6FC0, &H6E80, &HAE41, &HAA01, &H6AC0, &H6B80, &HAB41, &H6900, &HA9C1, &HA881, &H6840, _
        &H7800, &HB8C1, &HB981, &H7940, &HBB01, &H7BC0, &H7A80, &HBA41, &HBE01, &H7EC0, &H7F80, &HBF41, &H7D00, &HBDC1, &HBC81, &H7C40, _
        &HB401, &H74C0, &H7580, &HB541, &H7700, &HB7C1, &HB681, &H7640, &H7200, &HB2C1, &HB381, &H7340, &HB101, &H71C0, &H7080, &HB041, _
        &H5000, &H90C1, &H9181, &H5140, &H9301, &H53C0, &H5280, &H9241, &H9601, &H56C0, &H5780, &H9741, &H5500, &H95C1, &H9481, &H5440, _
        &H9C01, &H5CC0, &H5D80, &H9D41, &H5F00, &H9FC1, &H9E81, &H5E40, &H5A00, &H9AC1, &H9B81, &H5B40, &H9901, &H59C0, &H5880, &H9841, _
        &H8801, &H48C0, &H4980, &H8941, &H4B00, &H8BC1, &H8A81, &H4A40, &H4E00, &H8EC1, &H8F81, &H4F40, &H8D01, &H4DC0, &H4C80, &H8C41, _
        &H4400, &H84C1, &H8581, &H4540, &H8701, &H47C0, &H4680, &H8641, &H8201, &H42C0, &H4380, &H8341, &H4100, &H81C1, &H8081, &H4040 _
        }

        '---------------------------------
        '   문자수 확인
        '---------------------------------
        '문자수가 偶数 일 것.
        valLen = Len(strChrs)
        If valLen Mod 2 <> 0 Then
            calCRC = ""
            Throw New System.ApplicationException(strChrs & ErrorMessageCal)
            Exit Function
        End If

        For i = 1 To CLng((valLen / 2))
            strChr2 = "&H" & Mid(strChrs, CInt(i * 2 - 1), 2)   '1번째 문자 부터２문자씩 추출.(Comment만 추가)

            intCRC1 = intCRC Xor CInt(strChr2)                  'Integer Type은 2byte(부호付).(8bit를 비교. strChr2의 bit 반전)
            strCRC1 = Right("0000" & Hex(intCRC1), 4)
            strCRC1l = "&H" & Left(strCRC1, 2)                   '上位
            strCRC1r = "&H" & Right(strCRC1, 2)                 '下位
            intCRC2 = CRCTable(CInt(strCRC1r))                '下位1byte 부터 Table 참조
            intCRC = CInt(strCRC1l) Xor intCRC2                'Table의 값과 上位1byte의 XOR
        Next i
        strCRC = Right("0000" & Hex(intCRC), 4)
        strCRCl = Left(strCRC, 2)                                    '上位
        strCRCr = Right(strCRC, 2)                                  '下位
        calCRC = strCRCr & strCRCl                                 '上位・下位 바꿔치기

    End Function

    '----------------------------------------
    '    Private Function
    '----------------------------------------

    'Error Message Header 작성
    Protected Overridable Function GetErrorHeader() As String
        Return String.Format("%s_%d:", Me.TargetName, Me.myID)
    End Function

    'WaitTime(ms)
    Private Sub Wait(ByVal Sec As Double)
        System.Threading.Thread.Sleep(CInt(Sec) * 1000)
    End Sub

End Class
