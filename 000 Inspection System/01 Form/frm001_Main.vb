Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm001_Main
    'Program Ver ==========================================
    'Private Const Version As String = "0.00"
    'Private Const Version As String = "0.01" 'Ver 0.01 HSJ ADD
    'Private Const Version As String = "0.02" 'Ver 0.02 HSJ ADD
    'Private Const Version As String = "0.03" 'Ver 0.03 HSJ ADD
    'Private Const Version As String = "0.04" 'Ver 0.04 HSJ ADD     'Ver0.05 YSH DEL
    Private Const Version As String = "0.05" 'Ver0.05 YSH ADD : Barcode_Router 연동 (END 회신) / md Rev 2 6-9, 7-11, SystemSetting.mdb 0.05 레코드 추가

    Private EtherComNDll As New EtherComN.Dll

    'Class ================================================
    Dim cWin As cls001_Window_Control
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cFrm As cls090_DesignForm_Type01
    Dim cSrv As cls110_ServerConnect

    Dim cBcd As cls101_ComRS232C

    'Local (Standard form control) ========================
    Dim sBar() As PictureBox
    Dim sBtn() As Button
    Dim sLbl() As Label

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mMousePoint As Point
    Dim mErrorComment As String = ""
    Dim mBarCode As Boolean = False
    Dim mSetting As String = ""
    Dim mPort As String = ""
    Dim mDelim As String = ""

    'Local (Special form control) ========================
    Dim lblSysTitle() As Label
    Dim btnPanel() As Label
    Dim webBrow() As WebBrowser
    Dim lblNowInfoItem() As Label
    Dim lblNowInfo() As Label
    Dim lblOperator() As Label
    Dim cmbOperator() As ComboBox
    Dim lblEquipmentStatus() As Label
    Dim lblLastDataItem() As Label
    Dim lblLastData() As Label
    Dim lblCountDataItem() As Label
    Dim lblCountData() As Label
    Dim lblBarcode() As Label
    Public txtBarcode() As TextBox

    Dim timTimer() As Timer

    'Local Property =======================================
    Dim mSystemLog As String = ""
    Dim mEquipmentStatus As Integer = 0
    Dim mError As String
    Dim mPLASMABarcode As String

    'Property =====================================================================================
    Public ReadOnly Property SYSTEM_LOG() As String
        Get
            SYSTEM_LOG = mSystemLog
        End Get
    End Property

    Public ReadOnly Property EQUIPMENT_STATUS() As Integer
        Get
            EQUIPMENT_STATUS = mEquipmentStatus
        End Get
    End Property

    Public ReadOnly Property COM_BARCODE() As Boolean
        Get
            COM_BARCODE = mBarCode
        End Get
    End Property

    Public ReadOnly Property COM_BARCODE_SETTING() As String
        Get
            COM_BARCODE_SETTING = mSetting
        End Get
    End Property

    Public ReadOnly Property COM_BARCODE_PORT() As String
        Get
            COM_BARCODE_PORT = mPort
        End Get
    End Property

    Public ReadOnly Property COM_BARCODE_TERMINAL_CHAR() As String
        Get
            COM_BARCODE_TERMINAL_CHAR = mDelim
        End Get
    End Property

    Private Delegate Sub COM_Recieve()
            ''==========================================================================================
            ''Function : COM recieve event
            ''Remarks  :
            ''==========================================================================================
    Private Sub SerialPort1_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived

        AddHandler txtBarcode(0).TextChanged, AddressOf txtBarcode_TextChanged

        Try
            Me.Invoke(New COM_Recieve(AddressOf BeginningProcess), New Object() {})
        Catch

        Finally

        End Try

        cBcd.PortClose()

        Try
            txtBarcode(0).Invoke(New COM_Recieve(AddressOf DataWriteProcess), New Object() {})
        Catch

        Finally

        End Try

        cBcd.PortSetup()

        Call clsRouterLink.SendEnd(Me.SerialPort1, mSystemLog)     'Ver0.05 YSH ADD : Barcode_Router 로 검사 종료(END) 회신. EquipmentSetting.mdb RouterLink = Yes 일 때만 (md Rev 7 11-7 P3)

        RemoveHandler txtBarcode(0).TextChanged, AddressOf txtBarcode_TextChanged

    End Sub
    Private Sub ProgramVerCheck()
        '==========================================================================================
        'Function : Check program version between Date in VB.NET and Systemsetting.mdb
        'Remarks  :
        '==========================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""

        'Read of System setting ==================================================================
        Try
            dbName = mAppPath & "100 System Data\SystemSetting.mdb"
            tbName = "tblSystemSetting"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            kField = "SystemVersion"

            sSQL = "SELECT " & "Max(" & kField & ") as maxver" & " FROM " & tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)

            If Version <> sData(0) Then
                mErrorComment = "최신 Program이 아닙니다."
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            If mErrorComment = "" Then
                mErrorComment = "Systemsetting.mdb를 확인하세요."
            End If
            MessageBox.Show(mErrorComment & vbCrLf & "생산기술 호출바랍니다." & vbCrLf & "VB Version : " & Version & vbCrLf & "DB Version : " & sData(0), "Program Check NG")
            End
        Finally
            cADO = Nothing
        End Try
    End Sub

    Private Sub BeginningProcess()
        '==========================================================================================
        'Function : Beginning process of inspection
        'Remarks  :
        '==========================================================================================

        Try
            Call cBcd.FuncRecvCmd()
            mPLASMABarcode = cBcd.RECEIVE_DATA

        Catch oExcept As Exception
            mError = cBcd.ERR_COMMENT : If mError = "" Then oExcept.ToString()

        End Try

        If mError <> "" Then
            MessageBox.Show(mError, "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            'MessageBox.Show("Get Barcode Text OK !", "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If


    End Sub

    Private Sub DataWriteProcess()

        txtBarcode(0).Text = mPLASMABarcode

        'lblBarcode(0).Text = mPLASMABarcode

    End Sub

    'Form control *****************************************************************************************************
    Private Sub frm001_Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '==============================================================================================================
        'Function : Display start
        'Remarks  :
        '==============================================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""
        'Program Ver Check ========================================================================
        Call ProgramVerCheck()

        'Draw of standaed form ====================================================================
        'Title & Under bar ------------------------------------------
        Me.sBar = New System.Windows.Forms.PictureBox(3) {}
        For i As Integer = 0 To 3
            Me.sBar(i) = New System.Windows.Forms.PictureBox
            Me.Controls.Add(sBar(i))
            AddHandler sBar(i).MouseDown, AddressOf Window_MouseDown
            AddHandler sBar(i).MouseMove, AddressOf Window_MouseMove
        Next

        'Minimum & Close button -------------------------------------
        Me.sBtn = New System.Windows.Forms.Button(1) {}
        For i As Integer = 0 To 1
            Me.sBtn(i) = New System.Windows.Forms.Button
            Me.Controls.Add(sBtn(i))
        Next
        AddHandler sBtn(0).Click, AddressOf btnMinimum_Click
        AddHandler sBtn(1).Click, AddressOf btnClose_Click

        'Label ------------------------------------------------------
        Me.sLbl = New System.Windows.Forms.Label(0) {}
        For i As Integer = 0 To 0
            Me.sLbl(i) = New System.Windows.Forms.Label
            Me.Controls.Add(sLbl(i))
            AddHandler sLbl(i).MouseDown, AddressOf Window_MouseDown
            AddHandler sLbl(i).MouseMove, AddressOf Window_MouseMove
        Next
        Me.ResumeLayout()

        'Form -------------------------------------------------------
        With Me
            .Text = "frm001 Main Panel"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 550 : .Height = 374
            .Top = 0
            .Left = (Screen.PrimaryScreen.Bounds.Width) - (.Width)
        End With

        Try
            cFrm = New cls090_DesignForm_Type01

            cFrm.FORM_NAME = Me
            cFrm.PIC_FOLDER = mAppPath & "900 Picture\Picture\GrayForm\"
            cFrm.TITLE_BAR_L = sBar(0) : cFrm.TITLE_BAR_R = sBar(1)
            cFrm.UNDER_BAR_L = sBar(2) : cFrm.UNDER_BAR_R = sBar(3)
            cFrm.MIN_BUTTON = sBtn(0)
            cFrm.CLOSE_BUTTON = sBtn(1)
            cFrm.PANEL_TITLE = sLbl(0)

            Call cFrm.FORM_DRAW()

            Call cFrm.TEXT_DRAW(sLbl(0), Me.Text, "Times new roman", 8, 5, 7, Color.DimGray)

        Catch ex As Exception
            If mErrorComment <> "" Then mErrorComment = cFrm.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            MessageBox.Show(mErrorComment, Me.Text & " : Open Form", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            cFrm = Nothing
            If mErrorComment <> "" Then End

        End Try

        'Read of BarcodeReader Equipment No. ====================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "BarcodeReaderPort , BarcodeReaderSetting" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "バーコードリーダー：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try

        If rCount <> 0 Then

            If mErrorComment <> "" Then Exit Sub
            Dim BarSetting() As String = Split(sData(1), "-")

            cBcd = New cls101_ComRS232C
            cBcd.TERMINAL_CHAR = vbCr : mDelim = vbCr
            cBcd.RS_COMM = Me.SerialPort1 : cBcd.PORT = sData(0) : mPort = sData(0)
            cBcd.SETTING = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)
            mSetting = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)
            mBarCode = True

            cBcd.PortSetup()

        Else
            txtBarcode(0).Text = ""

        End If


        'Call Me.SysLogWrite(True, Me.Name & " : バーコードリーダーセットアップ完了")

        'Process Start ============================================================================
        Call Me.MainProcess()

    End Sub

    Private Sub frm001_Main_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        '==============================================================================================================
        'Function : After display
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

        DoEvents()

        Me.Opacity = 1.0

        Me.Activate()

    End Sub

    Private Sub Window_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        '==============================================================================================================
        'Function : Form move
        'Remarks  :
        '==============================================================================================================

        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then mMousePoint = New Point(e.X, e.Y)

    End Sub

    Private Sub Window_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        '==============================================================================================================
        'Function : Form move
        'Remarks  :
        '==============================================================================================================

        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then Me.Left = Me.Left + e.X - mMousePoint.X : Me.Top += e.Y - mMousePoint.Y

    End Sub

    Private Sub btnMinimum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Minimum window 
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()
        WindowState = FormWindowState.Minimized

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Close window 
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()
        Me.Close()

    End Sub

    Private Sub frm001_Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '==========================================================================================
        'Function : Plogram end
        'Remarks  :
        '==========================================================================================
        'Dim ret As DialogResult = MessageBox.Show("終了します。よろしいですか？", "Program End", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        Dim ret As DialogResult = MessageBox.Show("종료합니다. 하시겠습니까?", "Program End", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)

        If ret = MsgBoxResult.Cancel Then e.Cancel = True : Exit Sub

        Call Me.ProgramEnd()

        Dim ps As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("POWERPNT")
        For i As Integer = 0 To ps.Length - 1
            ps(i).Kill()
        Next

        End

    End Sub

    Private Sub ProgramEnd()
        '==============================================================================================================
        ' Function : Program end
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim cSys As cls102_SignalTower

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "●　システム終了処理 --------------------")

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "サーバー切断　・・・")

        Try
            cSrv = New cls110_ServerConnect

            Call cSrv.ServerDisconnect()

        Catch ex As Exception
            mErrorComment = cSrv.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cSrv = Nothing

        End Try

        If mErrorComment = "" Then
            Call Me.SysLogWrite(True, "サーバー切断：完了")
        Else
            Call Me.SysLogWrite(True, "サーバー切断：" & mErrorComment)
        End If

        cSys = New cls102_SignalTower
        Call cSys.SignalTowerReset()
        cSys = Nothing

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "■　システム終了 ==========================================================")
        Call Me.SysLogWrite(False, "")


    End Sub

    'Process (Start up) ***********************************************************************************************
    Private Sub MainProcess()
        '==============================================================================================================
        ' Function : Main process of this panel
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        'Application check ========================================================================
        If UBound(GetProcessesByName(GetCurrentProcess.ProcessName)) > 0 Then
            'MessageBox.Show("このプログラムは既に起動しています。", "Application start NG", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("이 프로그램은 이미 시작하고 있습니다", "Application start NG", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End If

        'Make of System Log =======================================================================
        Call Me.MakeSystemLog()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, "Make System Log", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End If
        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "●　メイン画面　起動 -------------------- (" & Me.Name & "起動)")
        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "二重起動チェック　完了")

        'Check of Equipment Condition =============================================================
        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "装置状況確認　開始　・・・")

        Dim frm As New frm100_EquipmentCondition
        Try
            frm.START_OPEN = True
            Call frm.ShowDialog()

        Finally
            mEquipmentStatus = frm.EQUIPMENT_STATUS
            frm = Nothing

        End Try

        Dim dfrm As New frm201_DispWait
        Try
            Call dfrm.Show()

            'Form Draw ================================================================================
            Call Me.FormDraw()
            If mErrorComment <> "" Then
                MessageBox.Show(mErrorComment, Me.Name & " Form Draw", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End
            End If

            'Form data set ============================================================================
            Call Me.FormDataSet()
            If mErrorComment <> "" Then
                MessageBox.Show(mErrorComment, Me.Name & " Data Set", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End
            End If

            'Ver 0.02 HSJ Add Start===
            Dim cSys As cls102_SignalTower
            cSys = New cls102_SignalTower
            cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
            cSys = Nothing
            'Ver 0.02 HSJ Add End===

            dfrm.Close()

        Finally
            dfrm = Nothing

        End Try

    End Sub

    Private Sub MakeSystemLog()
        '==============================================================================================================
        ' Function : System log
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim res As String() = {}

        'Make Folder ==============================================================================
        Dim SearchArea As String = mAppPath & "300 System Log"
        Dim SearchFolder As String = Format(DateTime.Now, "yyyy-MM")
        Dim SearchFile As String = Format(DateTime.Now, "yyyy-MM-dd")

        'Search date folder -----------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea
            cFop.FOLDER_NAME = SearchFolder

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        Dim sFolder As String = String.Empty
        For Each stFilePath As String In res
            sFolder &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then make folder -------------------------------------
        If sFolder = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.FOLDER_PATH = SearchArea & "\" & SearchFolder

                Call cFop.FOLDER_MAKE()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()

            Finally
                cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Search Log file ------------------------------------------- 
        mSystemLog = SearchArea & "\" & SearchFolder & "\" & Format(DateTime.Now, "yyyy-MM-dd") & ".txt"

        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\" & SearchFolder & "\"
            cFop.FILE_NAME = Format(DateTime.Now, "yyyy-MM-dd") & ".txt"
            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        Dim sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Log Write --------------------------------------------------
        If sFile = String.Empty Then
            Call Me.SysLogWrite(False, mSystemLog)
        End If

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "■　システム起動 ==========================================================")

    End Sub

    Private Sub FormDraw()
        '==============================================================================================================
        ' Function : Form draw
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        'System title =============================================================================
        Me.lblSysTitle = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblSysTitle(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblSysTitle(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblSysTitle(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 9, FontStyle.Bold)
                If i = 0 Then
                    .TextAlign = ContentAlignment.MiddleLeft
                    .Width = CInt(Me.Width * 0.66) : .Height = 20
                    .Top = sBar(0).Height + 2 : .Left = 2
                Else
                    .TextAlign = ContentAlignment.MiddleRight
                    .Width = CInt(Me.Width * 0.32) + 5 : .Height = lblSysTitle(0).Height
                    .Top = lblSysTitle(0).Top : .Left = lblSysTitle(0).Left + lblSysTitle(0).Width + 2
                End If
                .ForeColor = Color.White
                .BorderStyle = BorderStyle.None
                .Text = ""

                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\BlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
            End With
        Next

        'Panel button =============================================================================
        Dim PanelBtnCount As Integer = 7
        Me.btnPanel = New System.Windows.Forms.Label(PanelBtnCount) {}
        For i As Integer = 0 To PanelBtnCount
            Me.btnPanel(i) = New System.Windows.Forms.Label
            Me.Controls.Add(btnPanel(i))
            AddHandler btnPanel(i).Click, AddressOf PanelBtn_Click
            AddHandler btnPanel(i).MouseMove, AddressOf PanelBtn_MouseMove
            AddHandler btnPanel(i).MouseLeave, AddressOf PanelBtn_MouseLeave
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To PanelBtnCount
            With btnPanel(i)
                .Name = "btnPanel(" & CStr(i) & ")"
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleCenter
                .Height = CInt((Me.Height - sBar(0).Height - sBar(2).Height - lblSysTitle(0).Height - (2 * PanelBtnCount) - 6) / (PanelBtnCount + 1))
                .Width = CInt(.Height * 1.8)
                If i = 0 Then
                    .Top = lblSysTitle(i).Top + lblSysTitle(i).Height + 2
                Else
                    .Top = btnPanel(i - 1).Top + btnPanel(i - 1).Height + 2
                End If
                .Left = 2

                .ForeColor = Color.FromArgb(255, 255, 255)
                .Cursor = Cursors.Hand

                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\GrayForm\ButtonOff.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                'If i = 0 Then .Text = "装置環境"
                'If i = 1 Then .Text = "検査実績"
                'If i = 2 Then .Text = "始業点検" 'v0.01→名称変更 "自己診断"
                'If i = 3 Then .Text = "手動操作"
                'If i = 4 Then .Text = "保守"
                'If i = 5 Then .Text = "オーダー" & vbCr & "入力"
                'If i = 6 Then .Text = "システム" & vbCr & "情報"
                'If i = 7 Then .Text = "終了"
                If i = 0 Then .Text = "장치환경"
                If i = 1 Then .Text = "검사실적"
                If i = 2 Then .Text = "시업점검"
                If i = 3 Then .Text = "수동조작"
                'If i = 4 Then .Text = "보수"     'rev0.07 del
                If i = 4 Then .Text = "주명판" & vbCr & "재출력" 'rev0.07 add
                If i = 5 Then .Text = "ORDER" & vbCr & "입력"
                If i = 6 Then .Text = "SYSTEM" & vbCr & "정보"
                If i = 7 Then .Text = "종료"
            End With
        Next

        'WebBrowser ===============================================================================
        Me.webBrow = New System.Windows.Forms.WebBrowser(0) {}
        For i As Integer = 0 To 0
            Me.webBrow(i) = New System.Windows.Forms.WebBrowser
            Me.Controls.Add(webBrow(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With webBrow(i)
                .Width = 230 : .Height = CInt(.Width * 0.75)
                .Top = lblSysTitle(0).Top + lblSysTitle(0).Height + 2
                .Left = btnPanel(0).Left + btnPanel(0).Width + 4
                .BringToFront()
            End With
        Next

        'Information of Now =======================================================================
        Me.lblNowInfoItem = New System.Windows.Forms.Label(2) {}
        For i As Integer = 0 To 2
            Me.lblNowInfoItem(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblNowInfoItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 2
            With lblNowInfoItem(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleLeft
                .Width = 70 : .Height = 16
                .Left = webBrow(0).Left
                If i = 0 Then
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Width = webBrow(0).Width : .Height = 16
                    .Top = webBrow(0).Top + webBrow(0).Height + 10
                    .BorderStyle = BorderStyle.None
                    .ForeColor = Color.FromArgb(255, 255, 255)
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                Else
                    .Top = (lblNowInfoItem(i - 1).Top + lblNowInfoItem(i - 1).Height) + 2
                    .BorderStyle = BorderStyle.None
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.White
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                .BringToFront()
                'If i = 0 Then .Text = "現在の情報"
                If i = 0 Then .Text = "현재의 정보"
                If i = 1 Then .Text = "DATE/TIME"
                If i = 2 Then .Text = "TEMP/HUMI"

            End With
        Next

        Me.lblNowInfo = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblNowInfo(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblNowInfo(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblNowInfo(i)
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleLeft
                .Height = 16 : .Width = lblNowInfoItem(0).Width - lblNowInfoItem(1).Width - 2
                .Left = lblNowInfoItem(i + 1).Left + lblNowInfoItem(i + 1).Width + 2
                .Top = lblNowInfoItem(i + 1).Top
                .BorderStyle = BorderStyle.None
                .BackColor = Color.FromArgb(236, 233, 216)
                .ForeColor = Color.FromArgb(10, 10, 10)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\WhiteLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
            End With
        Next

        'Ope. & App. ==============================================================================
        Me.lblOperator = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblOperator(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblOperator(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblOperator(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = CInt(lblNowInfoItem(0).Width / 2 - 1) : .Height = 16
                .Top = lblNowInfo(1).Top + lblNowInfo(1).Height + 10
                .Left = webBrow(0).Left
                If i = 0 Then
                    .Left = webBrow(0).Left
                    '.Text = "作業者"
                    .Text = "작업자"
                Else
                    .Left = (lblOperator(i - 1).Left + lblOperator(i - 1).Width) + 2
                    '.Text = "責任者"
                    .Text = "승인자"
                End If
                .BackColor = Color.FromArgb(50, 50, 60)
                .ForeColor = Color.White
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
            End With
        Next

        Me.cmbOperator = New System.Windows.Forms.ComboBox(1) {}
        For i As Integer = 0 To 1
            Me.cmbOperator(i) = New System.Windows.Forms.ComboBox
            Me.Controls.Add(cmbOperator(i))
            AddHandler cmbOperator(i).KeyDown, AddressOf Operator_KeyDown
            AddHandler cmbOperator(i).KeyPress, AddressOf Operator_KeyPress
            AddHandler cmbOperator(i).SelectedIndexChanged, AddressOf Operator_SelectedIndexChanged
            AddHandler cmbOperator(i).TextChanged, AddressOf Operator_TextChanged
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With cmbOperator(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Name = "cmbOperator(" & CStr(i) & ")"
                .Width = lblOperator(i).Width
                .Top = lblOperator(i).Top + lblOperator(i).Height + 2
                .Cursor = Cursors.Hand
                .Left = lblOperator(i).Left
                .BackColor = Color.FromArgb(236, 233, 216)
                .ForeColor = Color.FromArgb(30, 30, 30)
                .BringToFront()
            End With
        Next

        'Equipment Status =========================================================================
        Me.lblEquipmentStatus = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblEquipmentStatus(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblEquipmentStatus(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblEquipmentStatus(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleCenter
                .Height = 16
                If i = 0 Then
                    .Left = webBrow(0).Left + webBrow(0).Width + 4
                    .Top = webBrow(0).Top
                    .Width = CInt((Me.Width - 4 - .Left) / 1)
                    '.Text = "装置の稼動状況"
                    .Text = "장치의 가동상황"
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.White
                    .BorderStyle = BorderStyle.None
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                Else
                    .Width = lblEquipmentStatus(0).Width
                    .Left = CInt(lblEquipmentStatus(0).Left + (lblEquipmentStatus(0).Width) / 2) - CInt(.Width / 2)
                    .Top = lblEquipmentStatus(i - 1).Top + lblEquipmentStatus(i - 1).Height + 2
                    .Text = ""
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.White
                    .BorderStyle = BorderStyle.None
                End If

                .BringToFront()
            End With
        Next

        'Last Data ================================================================================
        Me.lblLastDataItem = New System.Windows.Forms.Label(6) {}
        For i As Integer = 0 To 6
            Me.lblLastDataItem(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblLastDataItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 6
            With lblLastDataItem(i)
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleLeft
                .Height = 16
                .Width = 70
                If i = 0 Then
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Left = webBrow(0).Left + webBrow(0).Width + 4
                    .Width = Me.Width - 4 - .Left
                    .Top = 100
                    .BorderStyle = BorderStyle.None
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.FromArgb(255, 255, 255)
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                Else
                    .Left = webBrow(0).Left + webBrow(0).Width + 4
                    .Top = lblLastDataItem(i - 1).Top + lblLastDataItem(i - 1).Height + 2
                    .BorderStyle = BorderStyle.None
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.FromArgb(255, 255, 255)
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                .BringToFront()
                'If i = 0 Then .Text = "最新の検査データ"
                If i = 0 Then .Text = "최신의 검사DATA"
                If i = 1 Then .Text = "DATE TIME"
                If i = 2 Then .Text = "MS-CODE"
                If i = 3 Then .Text = "XJ NO."
                If i = 4 Then .Text = "O-I-C"
                If i = 5 Then .Text = "SERIAL"
                If i = 6 Then .Text = "Result"

            End With

        Next

        Me.lblLastData = New System.Windows.Forms.Label(5) {}
        For i As Integer = 0 To 5
            Me.lblLastData(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblLastData(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 5
            With lblLastData(i)
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleLeft
                .Height = 16 : .Width = lblLastDataItem(0).Width - lblLastDataItem(1).Width - 2
                .Left = lblLastDataItem(i + 1).Left + lblLastDataItem(i + 1).Width + 2
                .Top = lblLastDataItem(i + 1).Top
                .BorderStyle = BorderStyle.None
                .BackColor = Color.FromArgb(236, 233, 216)
                .ForeColor = Color.FromArgb(10, 10, 10)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\WhiteLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
            End With
        Next

        'Count Data ===============================================================================
        Me.lblCountDataItem = New System.Windows.Forms.Label(4) {}
        For i As Integer = 0 To 4
            Me.lblCountDataItem(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblCountDataItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 4
            With lblCountDataItem(i)
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleLeft
                .Height = 16
                .Width = 60
                If i = 0 Then
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Width = 150 : .Top = Me.Height - 120
                    .Left = Me.Width - .Width - 4
                    .BorderStyle = BorderStyle.None
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.FromArgb(255, 255, 255)
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                Else
                    .Left = lblCountDataItem(0).Left
                    .Top = lblCountDataItem(i - 1).Top + lblCountDataItem(i - 1).Height + 2
                    .BorderStyle = BorderStyle.None
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.FromArgb(255, 255, 255)
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                .BringToFront()
                'If i = 0 Then .Text = "本日の検査回数"
                'If i = 1 Then .Text = "始業点検"
                'If i = 2 Then .Text = "合格"
                'If i = 3 Then .Text = "不合格"
                'If i = 4 Then .Text = "手動操作"
                If i = 0 Then .Text = "오늘의검사횟수"
                If i = 1 Then .Text = "시업점검"
                If i = 2 Then .Text = "합격"
                If i = 3 Then .Text = "불합격"
                If i = 4 Then .Text = "수동조작"
            End With
        Next

        Me.lblCountData = New System.Windows.Forms.Label(3) {}
        For i As Integer = 0 To 3
            Me.lblCountData(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblCountData(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 3
            With lblCountData(i)
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleRight
                .Height = 16 : .Width = lblCountDataItem(0).Width - lblCountDataItem(1).Width - 2
                .Left = lblCountDataItem(i + 1).Left + lblCountDataItem(i + 1).Width + 2
                .Top = lblCountDataItem(i + 1).Top
                .BorderStyle = BorderStyle.None
                .BackColor = Color.FromArgb(236, 233, 216)
                .ForeColor = Color.FromArgb(10, 10, 10)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\WhiteLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
            End With
        Next

        'BarCodeTextBox ============================================================================
        Me.txtBarcode = New System.Windows.Forms.TextBox(0) {}
        Me.txtBarcode(0) = New System.Windows.Forms.TextBox
        Me.Controls.Add(txtBarcode(0))
        AddHandler txtBarcode(0).KeyPress, AddressOf txtBarcode_KeyPress
        Me.ResumeLayout()

        With txtBarcode(0)
            .Font = New Font("MS Gothic", 8, FontStyle.Regular)
            .AutoSize = False
            .Height = 16
            .Width = lblCountDataItem(0).Width - lblCountDataItem(1).Width - 2
            .Top = Me.Height - 136
            .Left = lblCountDataItem(1).Left + lblCountDataItem(1).Width + 2
            .BorderStyle = BorderStyle.None
            .ForeColor = Color.Black
            .BackColor = Color.White
            '.BackgroundImageLayout = ImageLayout.Stretch
            .BringToFront()
            .Text = ""
        End With

        Me.lblBarcode = New System.Windows.Forms.Label(0) {}
        Me.lblBarcode(0) = New System.Windows.Forms.Label
        Me.Controls.Add(lblBarcode(0))
        'AddHandler lblBarcode(0).TextChanged, AddressOf lblBarcode_TextChanged
        Me.ResumeLayout()

        With lblBarcode(0)
            .Font = New Font("MS Gothic", 8, FontStyle.Regular)
            .AutoSize = False
            .TextAlign = ContentAlignment.MiddleLeft
            .Height = 16
            .Width = 60
            .TextAlign = ContentAlignment.MiddleCenter
            .Top = Me.Height - 136
            .Left = lblCountDataItem(0).Left
            .BorderStyle = BorderStyle.None
            .BackColor = Color.FromArgb(50, 50, 60)
            .ForeColor = Color.FromArgb(255, 255, 255)
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch
            .BringToFront()
            .Text = "Barcode"
        End With

        'Reflesh timer ============================================================================
        Me.timTimer = New System.Windows.Forms.Timer(0) {}
        For i As Integer = 0 To 0
            Me.timTimer(i) = New System.Windows.Forms.Timer
            AddHandler timTimer(i).Tick, AddressOf Reflesh_Tick
        Next

        For i As Integer = 0 To 0
            With timTimer(i)
                .Interval = 500
                .Enabled = True
            End With
        Next

    End Sub

    Private Sub FormDataSet()
        '==============================================================================================================
        ' Function : FormDataSet
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "●　メイン画面　データ読込と表示")

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""

        'Read of System setting ==================================================================
        Try
            dbName = mAppPath & "100 System Data\SystemSetting.mdb"
            tbName = "tblSystemSetting"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            rField = "SystemName,SystemVersion"
            kField = "SystemVersion"

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & kField & " in "
            sSQL = sSQL & "(SELECT " & "Max(" & kField & ") as maxver" & " FROM " & tbName & ")"

            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "システム情報読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        lblSysTitle(0).Text = sData(0)
        lblSysTitle(1).Text = "System Version " & sData(1)

        Call Me.SysLogWrite(True, sData(0))
        Call Me.SysLogWrite(True, "Ver. " & sData(1))

        'Navigate of WebBrowser ===================================================================
        webBrow(0).Navigate(mAppPath & "900 Picture\ppt\SheetRead.ppt")
        'webBrow(0).Stop()

        'Read of Temp. & Humi =====================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblDefTempHumi"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "Temperature,Humidity" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "温・湿度読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        lblNowInfo(1).Text = sData(0) & " ℃     " & sData(1) & " %"

        Call Me.SysLogWrite(True, "温度：" & sData(0) & " / " & "湿度：" & sData(1))

        'Read of name list ========================================================================
        'Operator ---------------------------------------------------
        Call Me.SysLogWrite(False, "")
        Try
            dbName = mAppPath & "400 Common File\NameEntry.mdb"
            tbName = "tblMenuOperator"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "Name" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "作業者リスト読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        cmbOperator(0).Items.Clear()
        For i As Integer = 0 To rCount - 1
            cmbOperator(0).Items.Add(sData(i))
        Next

        Call Me.SysLogWrite(True, "作業者リスト読込：完了")

        'Approver ---------------------------------------------------
        Try
            dbName = mAppPath & "400 Common File\NameEntry.mdb"
            tbName = "tblMenuApprover"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "Name" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "責任者リスト読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        cmbOperator(1).Items.Clear()
        For i As Integer = 0 To rCount - 1
            cmbOperator(1).Items.Add(sData(i))
        Next

        Call Me.SysLogWrite(True, "責任者リスト読込：完了")

        'Read of select name ======================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblDefNameEntry"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "OperatorName,ApproverName" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "作業者選択読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        cmbOperator(0).Text = sData(0)
        cmbOperator(1).Text = sData(1)

        Call Me.SysLogWrite(True, "作業者：" & sData(0))
        Call Me.SysLogWrite(True, "責任者：" & sData(1))

        'Equipment　Status ========================================================================
        Select Case mEquipmentStatus
            Case 0
                'lblEquipmentStatus(1).Text = "異常" : lblEquipmentStatus(1).ForeColor = Color.DarkRed
                lblEquipmentStatus(1).Text = "이상" : lblEquipmentStatus(1).ForeColor = Color.DarkRed
                lblEquipmentStatus(1).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\RedLabel.bmp")
                lblEquipmentStatus(1).BackgroundImageLayout = ImageLayout.Stretch
                Call Me.SysLogWrite(True, "装置稼動状況：異常")
            Case 1
                'lblEquipmentStatus(1).Text = "注意" : lblEquipmentStatus(1).ForeColor = Color.Yellow
                lblEquipmentStatus(1).Text = "주의" : lblEquipmentStatus(1).ForeColor = Color.Yellow
                lblEquipmentStatus(1).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\YellowLabel.bmp")
                lblEquipmentStatus(1).BackgroundImageLayout = ImageLayout.Stretch
                Call Me.SysLogWrite(True, "装置稼動状況：注意")
            Case 2
                'lblEquipmentStatus(1).Text = "正常" : lblEquipmentStatus(1).ForeColor = Color.Lime
                lblEquipmentStatus(1).Text = "정상" : lblEquipmentStatus(1).ForeColor = Color.Lime
                lblEquipmentStatus(1).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GreenLabel.bmp")
                lblEquipmentStatus(1).BackgroundImageLayout = ImageLayout.Stretch
                Call Me.SysLogWrite(True, "装置稼動状況：正常")
        End Select

        'Read of last data ========================================================================
        Dim res As String() = {}

        'Year-month folder ------------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = mAppPath & "210 Result Data\"
            cFop.FOLDER_NAME = "*"

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        If res.Length = 0 Then
            'Year-month folder 0 ----------------------------------------
            For i As Integer = 0 To 5
                lblLastData(i).Text = ""
            Next
            Call Me.SysLogWrite(True, "最新データ読込　データフォルダ　0")

        Else
            'Last file --------------------------------------------------
            Try
                cFop = New cls012_File_Operation

                cFop.ROOT_FOLDER = res(res.Length - 1)
                cFop.FILE_NAME = (Format(DateTime.Now, "yyyy-MM")) & "*"

                Call cFop.FILE_SEARCH(res)

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()

            Finally
                cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

            If res.Length = 0 Then
                'Only year-month folder ---------------------------------
                For i As Integer = 0 To 5
                    lblLastData(i).Text = ""
                Next
                Call Me.SysLogWrite(True, "最新データ読込　データファイル　0")
            Else
                'Read of last file ------------------------------------
                Try
                    dbName = res(res.Length - 1)
                    tbName = "tblResult"
                    rField = "StartDateTime,MSCODE,XJNo,SerialNo,OrderNo,ItemNo,CompNo,Result"
                    kField = "StartDateTime"
                    cADO = New cls010_MDB_Operation

                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName

                    sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & kField & " in "
                    sSQL = sSQL & "(SELECT " & "Max(" & kField & ") as maxver" & " FROM " & tbName & ")"

                    cADO.DB_SQL = sSQL

                    Call cADO.DB_READ_JET(fName, sData)

                    fCount = cADO.DB_FIELD_COUNT
                    rCount = cADO.DB_RECORD_COUNT

                Catch ex As Exception
                    mErrorComment = cADO.ERR_COMMENT
                    If mErrorComment = "" Then mErrorComment = ex.ToString()
                    Call Me.SysLogWrite(True, "最新データ読込：" & mErrorComment)

                Finally
                    cADO = Nothing

                End Try
                If mErrorComment <> "" Then Exit Sub

                If rCount = 0 Then
                    For i As Integer = 0 To 5
                        lblLastData(i).Text = ""
                    Next
                    Call Me.SysLogWrite(True, "最新データ読込　データ　0")

                Else
                    lblLastData(0).Text = sData(0)
                    lblLastData(1).Text = sData(1)
                    lblLastData(2).Text = sData(2)
                    lblLastData(3).Text = sData(4) & "-" & sData(5) & "-" & sData(6)
                    lblLastData(4).Text = sData(3)
                    lblLastData(5).Text = sData(7)

                    Call Me.SysLogWrite(True, "最新データ読込　完了：" & sData(0))

                End If

            End If

        End If

        'Today's result count =====================================================================
        'Today folder -----------------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = mAppPath & "210 Result Data\"
            cFop.FOLDER_NAME = Format(DateTime.Now, "yyyy-MM")

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Today file --------------------------------------------------
        If res.Length = 0 Then
            For i As Integer = 0 To 3
                lblCountData(i).Text = "0"
            Next
            Call Me.SysLogWrite(True, "今月のフォルダ　無し")
        Else
            Try
                cFop = New cls012_File_Operation

                cFop.ROOT_FOLDER = mAppPath & "210 Result Data\" & Format(DateTime.Now, "yyyy-MM") & "\"
                cFop.FILE_NAME = (Format(DateTime.Now, "yyyy-MM-dd")) & ".mdb"

                Call cFop.FILE_SEARCH(res)

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()

            Finally
                cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

            If res.Length = 0 Then
                For i As Integer = 0 To 3
                    lblCountData(i).Text = "0"
                Next
                Call Me.SysLogWrite(True, "今日のファイル　無し")
            Else
                'Read of today file -----------------------------------------
                Try
                    dbName = mAppPath & "210 Result Data\" & Format(DateTime.Now, "yyyy-MM") & "\" & Format(DateTime.Now, "yyyy-MM-dd") & ".mdb"
                    tbName = "tblResult"
                    rField = "StartDateTime"

                    cADO = New cls010_MDB_Operation

                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName

                    'PASS & CHK
                    sSQL = "SELECT " & rField & " FROM " & tbName
                    sSQL &= " WHERE Result LIKE 'PASS' AND RunMode LIKE 'CHK'"
                    cADO.DB_SQL = sSQL

                    Call cADO.DB_READ_JET(fName, sData)
                    lblCountData(0).Text = CStr(cADO.DB_RECORD_COUNT)

                    'PASS & AUT
                    sSQL = "SELECT " & rField & " FROM " & tbName
                    sSQL &= " WHERE Result LIKE 'PASS' AND RunMode LIKE 'AUT'"
                    cADO.DB_SQL = sSQL

                    Call cADO.DB_READ_JET(fName, sData)
                    lblCountData(1).Text = CStr(cADO.DB_RECORD_COUNT)

                    'NG & AUT
                    sSQL = "SELECT " & rField & " FROM " & tbName
                    sSQL &= " WHERE Result LIKE 'NG' AND RunMode LIKE 'AUT'"
                    cADO.DB_SQL = sSQL

                    Call cADO.DB_READ_JET(fName, sData)
                    lblCountData(2).Text = CStr(cADO.DB_RECORD_COUNT)

                    'MAN
                    sSQL = "SELECT " & rField & " FROM " & tbName
                    sSQL &= " WHERE RunMode LIKE 'MAN'"
                    cADO.DB_SQL = sSQL

                    Call cADO.DB_READ_JET(fName, sData)
                    lblCountData(3).Text = CStr(cADO.DB_RECORD_COUNT)

                Catch ex As Exception
                    mErrorComment = cADO.ERR_COMMENT
                    If mErrorComment = "" Then mErrorComment = ex.ToString()
                    Call Me.SysLogWrite(True, "本日データ読込：" & mErrorComment)

                Finally
                    cADO = Nothing

                End Try
                If mErrorComment <> "" Then Exit Sub

                Call Me.SysLogWrite(True, "本日データ読込　完了")

            End If

        End If

        'Read of Equipment No. ====================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "EquipmentNumber" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "装置番号読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub
        Call Me.SysLogWrite(True, "装置番号：" & sData(0))

        lblSysTitle(0).Text &= "   No.:" & sData(0)

        Call Me.SysLogWrite(True, Me.Name & " : 表示データ　読込完了")

        txtBarcode(0).Text = ""

    End Sub

    'Process (Event) **************************************************************************************************
    Private Sub PanelBtn_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        '==============================================================================================================
        'Function : Pannel select
        'Remarks  :
        '==============================================================================================================

        With DirectCast(sender, Label)
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\GrayForm\ButtonON.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch
        End With

    End Sub

    Private Sub PanelBtn_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Pannel select
        'Remarks  :
        '==============================================================================================================

        With DirectCast(sender, Label)
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\GrayForm\ButtonOFF.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch
        End With

    End Sub

    Private Sub PanelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Pannel select
        'Remarks  :
        '==============================================================================================================

        Dim sRet As String = ""

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : " & DirectCast(sender, Label).Text & " ON")

        'End ======================================================================================
        If DirectCast(sender, Label).Name = "btnPanel(7)" Then
            Me.Close()
            Exit Sub
        End If

        'Panel open ===============================================================================
        Me.Hide()
        timTimer(0).Enabled = False

        Call cBcd.PortClose()

        Try
            'Equipment condition ----------------------------------------
            If DirectCast(sender, Label).Name = "btnPanel(0)" Then
                Dim frm As New frm100_EquipmentCondition
                Try
                    Call frm.ShowDialog()
                    mEquipmentStatus = frm.EQUIPMENT_STATUS
                Finally
                    frm = Nothing
                End Try
            End If

            'Inspection result ------------------------------------------
            If DirectCast(sender, Label).Name = "btnPanel(1)" Then
                '-----rev0.07 add-----
                Dim p2frm As New frm200_InputPass
                'Password 
                Try
                    Call p2frm.ShowDialog()
                    sRet = p2frm.DIALOG_RESULT
                Finally
                    p2frm = Nothing
                End Try

                If sRet = "PASS OK" Then
                    '---------------------
                    Dim frm As New frm110_Result
                    Try
                        Call frm.ShowDialog()
                    Finally
                        frm = Nothing
                    End Try
                End If
            End If     'rev0.07 add

            'Self check -------------------------------------------------
            If DirectCast(sender, Label).Name = "btnPanel(2)" Then
                Dim frm As New frm120_SelfCheck
                Try
                    Call frm.ShowDialog()
                Finally
                    frm = Nothing
                End Try
            End If

            'Manual operation -------------------------------------------
            If DirectCast(sender, Label).Name = "btnPanel(3)" Then
                'Password 
                Dim p1frm As New frm200_InputPass
                Try
                    Call p1frm.ShowDialog()
                    sRet = p1frm.DIALOG_RESULT
                Finally
                    p1frm = Nothing
                End Try

                If sRet = "PASS OK" Then
                    Dim frm As New frm130_ManualOperation
                    Try
                        Call frm.ShowDialog()
                    Finally
                        frm = Nothing
                    End Try
                End If
            End If

            ''Maintenance menu -------------------------------------------     'rev0.07 del
            'Reprint menu -------------------------------------------     'rev0.07 add
            If DirectCast(sender, Label).Name = "btnPanel(4)" Then
                Dim p2frm As New frm200_InputPass
                'Password 
                Try
                    Call p2frm.ShowDialog()
                    sRet = p2frm.DIALOG_RESULT
                Finally
                    p2frm = Nothing
                End Try

                If sRet = "PASS OK" Then
                    'Dim frm As New frm140_MaintenanceMenu     'rev0.07 del
                    Dim frm As New frm600_Reprint     'rev0.07 add
                    Try
                        Call frm.ShowDialog()
                    Finally
                        frm = Nothing
                    End Try
                End If
            End If

            'Order input ------------------------------------------------
            If DirectCast(sender, Label).Name = "btnPanel(5)" Then
                Dim frm As New frm010_OrderInput
                Try
                    Call frm.ShowDialog()
                Finally
                    frm = Nothing
                End Try
            End If

            'System information -----------------------------------------
            If DirectCast(sender, Label).Name = "btnPanel(6)" Then
                Dim frm As New frm150_SystemInformation
                Try
                    Call frm.ShowDialog()
                Finally
                    frm = Nothing
                End Try
            End If

        Finally
            Call ReloadWaitFormShow()

        End Try

        Call cBcd.PortOpen()

    End Sub

    Private Sub Operator_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        '==============================================================================================================
        ' Function : Cancel of Key input(operator) 
        ' Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

        If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Or _
           e.KeyCode = Keys.Right Or e.KeyCode = Keys.Left Or _
           e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Or _
           e.KeyCode = Keys.Prior Or e.KeyCode = Keys.Next Then

            e.Handled = True

        End If

    End Sub

    Private Sub Operator_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        '==============================================================================================================
        ' Function : Cancel of Key input(operator) 
        ' Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()
        e.KeyChar = Nothing

    End Sub

    Private Sub Operator_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Cancel of select(operator) 
        ' Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

    End Sub

    Private Sub Operator_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Select change of operator 
        ' Remarks  :
        '==============================================================================================================

        If cmbOperator(0).Focused = False And cmbOperator(1).Focused = False Then Exit Sub

        Call Me.SysLogWrite(False, "")
        Call Me.FocusControl()

        Select Case DirectCast(sender, ComboBox).Name
            Case "cmbOperator(0)"
                If cmbOperator(0).Text = "" Then Exit Sub
                Call Me.SysLogWrite(True, "作業者：" & cmbOperator(0).Text & "に変更")

            Case "cmbOperator(1)"
                If cmbOperator(1).Text = "" Then Exit Sub
                Call Me.SysLogWrite(True, "責任者：" & cmbOperator(1).Text & "に変更")

        End Select

        Call Me.SaveData()
        If mErrorComment <> "" Then
            'MessageBox.Show("人名保存失敗：" & mErrorComment, "Select name in Main Panel", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            MessageBox.Show("이름 보존 실패：" & mErrorComment, "Select name in Main Panel", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            mErrorComment = ""
        End If

    End Sub

    Private Sub txtBarcode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : 
        ' Remarks  :
        '==============================================================================================================

        If txtBarcode(0).TextLength = 6 Then

            Call PLASMABarCodeInspectionStart()

        ElseIf txtBarcode(0).TextLength = 8 Then

            Call OP_EXC()

        ElseIf txtBarcode(0).TextLength > 7 Then

            Call MODELBarCodeInspectionStart()

        End If

    End Sub

    Private Sub txtBarcode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar = Chr(Keys.Enter) Then '-------------エンターキーが入力されたら

            If mBarCode = True Then cBcd.PortClose()

            If txtBarcode(0).TextLength = 6 Then

                Call PLASMABarCodeInspectionStart()

            ElseIf txtBarcode(0).TextLength = 8 Then

                Call OP_EXC()

            ElseIf txtBarcode(0).TextLength > 7 Then

                Call MODELBarCodeInspectionStart()

            End If

            If mBarCode = True Then cBcd.PortSetup()

        End If

    End Sub

    Private Sub PLASMABarCodeInspectionStart()

        Me.Hide()

        Dim sRet As String = ""
        Dim frm As New frm020_InspectionStartEnd
        Try
            txtBarcode(0).Text = StrConv(txtBarcode(0).Text, VbStrConv.Uppercase)
            frm.PLASMA_BARCODE = txtBarcode(0).Text
            'frm.PLASMA_BARCODE = Mid(txtBarcode(0).Text, 1, Len(txtBarcode(0).Text) - Len(vbCr))
            frm.RUN_MODE = "AUT"
            frm.INSPECTION_MODEL = ""

            Call frm.InspectionStart()

            sRet = frm.DIALOG_RESULT
        Catch
            MsgBox("ERROR発生！")
        Finally
            frm = Nothing
            Call ReloadWaitFormShow()
        End Try

        txtBarcode(0).Text = ""
        txtBarcode(0).Focus()

        Me.Show()

    End Sub

    Private Sub OP_EXC()

        'Me.Hide()
        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""
        Dim OP_Name As String = ""
        Dim OP_ID As String = txtBarcode(0).Text

        Dim mPlasmaResult As String = ""
        Dim mErrorCode As String = ""

        'Read of Operator Name. ====================================================================
        Try

            sSQL = OP_ID & "-E"
            mPlasmaResult = EtherComNDll.ComHost("10.7.15.254", "10.7.15.140", 2031, "EMP-" & sSQL)

            Debug.Print(mPlasmaResult)

            Select Case Mid(mPlasmaResult, 1, 8)
                Case "ERROR:A0"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "生産情報がありません。このバ-コ-ドは無効です。"
                Case "ERROR:A1"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "生産情報がｷｬﾝｾルされているか。無効の着工札です。"
                Case "ERROR:A2" : Case "ERROR:A3" : Case "ERROR:A4"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "プログラム エラーです。"
                Case "ERROR:A5"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "ＯＲＡＣＬＥエラーです。"
                Case "ERROR:13"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "タイムアウトエラーです。"
                Case ""
                    mErrorCode = mPlasmaResult
                    mErrorComment = "データがありません。"
            End Select
            If mErrorCode <> "" Then Throw New ApplicationException(mErrorComment)

            OP_Name = mPlasmaResult

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

        'Write of Operator. setting ===========================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblDefNameEntry"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "OperatorName = '" & OP_Name & "' , "
            sSQL &= "Operatorid = '" & OP_ID & "'"

            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "作業者情報 更新失敗")
        Finally
            cADO = Nothing

        End Try

        Call Me.SysLogWrite(True, "■ 作業者交代")
        Call Me.SysLogWrite(True, "変更後 ID：" & OP_ID)
        Call Me.SysLogWrite(True, "変更後 Name：" & OP_Name)

        cmbOperator(0).Text = OP_Name
        txtBarcode(0).Text = ""
        txtBarcode(0).Focus()

        'Me.Show()

    End Sub


    Private Sub MODELBarCodeInspectionStart()

        Me.Hide()

        Dim sRet As String = ""
        Dim frm As New frm020_InspectionStartEnd
        Try
            txtBarcode(0).Text = StrConv(txtBarcode(0).Text, VbStrConv.Uppercase)
            frm.PLASMA_BARCODE = ""
            frm.RUN_MODE = "AUT"
            frm.INSPECTION_MODEL = txtBarcode(0).Text

            Call frm.InspectionStart()

            sRet = frm.DIALOG_RESULT
        Catch
            MsgBox("ERROR発生！")
        Finally
            frm = Nothing
            Call ReloadWaitFormShow()
        End Try

        txtBarcode(0).Text = ""
        txtBarcode(0).Focus()

        Me.Show()

    End Sub

    Private Sub ReloadWaitFormShow()

        Dim waitfrm As New frm201_DispWait
        Try
            Call waitfrm.Show()

            Me.FormDataSet() : Me.Opacity = 0 : Me.Show()
            'Me.TopMost = True
            DoEvents()
            Me.Opacity = 1.0
            timTimer(0).Enabled = True

            waitfrm.Close()

        Finally
            waitfrm = Nothing
            Me.txtBarcode(0).Focus()
        End Try

    End Sub

    Private Sub SaveData()
        '==============================================================================================================
        ' Function : Data save
        ' Remarks  :
        '==============================================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        'Write of Temp. & Humi. setting ===========================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblDefNameEntry"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "OperatorName = '" & cmbOperator(0).Text & "' , "
            sSQL &= "ApproverName = '" & cmbOperator(1).Text & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cADO = Nothing

        End Try

        Call Me.SysLogWrite(True, "データ保存：" & cmbOperator(0).Text & " / " & cmbOperator(1).Text)

    End Sub

    Private Sub Reflesh_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Reflesh timer
        ' Remarks  :
        '==============================================================================================================

        lblNowInfo(0).Text = Format(DateTime.Now, "yyyy.MM.dd  HH:mm:ss")

    End Sub

    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : PictureBox focus
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        'sBar(0).Focus()
        Me.txtBarcode(0).Focus()

    End Sub

    Private Sub SysLogWrite(ByVal tStmp As Boolean, ByVal msg As String)
        '==============================================================================================================
        ' Function : System Log write
        ' Data     : Time stapY/N , Message
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Try
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = mSystemLog

            Call cTxt.WRITE_TEXT_FILE(tStmp, msg)

        Finally
            cTxt = Nothing

        End Try

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
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

End Class