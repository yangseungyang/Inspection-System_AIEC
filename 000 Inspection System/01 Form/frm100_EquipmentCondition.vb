Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm001_Main
Imports VB = Microsoft.VisualBasic  'Ver0.37 GYS ADD


Public Class frm100_EquipmentCondition

    Public Shared Insfrm As frm001_Main

    'Class ================================================
    Dim cWin As cls001_Window_Control
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cSrv As cls110_ServerConnect
    Dim cFrm As cls090_DesignForm_Type01

    'Local (Standard form control) ========================
    Dim sBar() As PictureBox
    Dim sBtn() As Button
    Dim sLbl() As Label

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mMousePoint As Point
    Dim mErrorComment As String = ""
    Dim mWarningComment As String = ""
    Dim mNormalComment As String = ""
    Dim mSystemLog As String = ""
    Dim mServerDrive As String = ""

    'Local (Special form control) ========================
    Dim btnDialogResult() As Button
    Dim lblTitle() As Label
    Dim lblItem() As Label
    Dim lblResult() As Label
    Dim lblComment() As Label
    Dim lblSetItem() As Label
    Dim cmbTempHumi() As ComboBox
    Dim lblDateTime() As Label
    Dim btnDateTimeExe() As Button
    Dim btnEquipmentSetting() As Button
    Dim timTimer() As Timer

    'Local Property =======================================
    Dim mStartOpen As Boolean = False
    Dim mDialogResult As String = ""
    Dim mEquipmentStatus As Integer

    '-----Ver0.37 GYS ADD Start-----
    Dim Revdata As Double
    Dim a As String
    Dim myModbus1 As clsNeptune
    Dim Temp As Integer
    Dim Hum As Integer
    '-----Ver0.37 GYS ADD End-----

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    Public WriteOnly Property START_OPEN() As Boolean
        Set(ByVal Value As Boolean)
            mStartOpen = Value
        End Set
    End Property

    Public ReadOnly Property EQUIPMENT_STATUS() As Integer
        Get
            EQUIPMENT_STATUS = mEquipmentStatus
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm100_Condition_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '==============================================================================================================
        'Function : Display start
        'Remarks  :
        '==============================================================================================================

        'System Log Path ==========================================================================
        mSystemLog = frm001_Main.SYSTEM_LOG

        'Title & Under bar ========================================================================
        Me.sBar = New System.Windows.Forms.PictureBox(3) {}
        For i As Integer = 0 To 3
            Me.sBar(i) = New System.Windows.Forms.PictureBox
            Me.Controls.Add(sBar(i))
            AddHandler sBar(i).MouseDown, AddressOf Window_MouseDown
            AddHandler sBar(i).MouseMove, AddressOf Window_MouseMove
        Next

        'Minimum & Close button ===================================================================
        Me.sBtn = New System.Windows.Forms.Button(1) {}
        For i As Integer = 0 To 1
            Me.sBtn(i) = New System.Windows.Forms.Button
            Me.Controls.Add(sBtn(i))
        Next
        AddHandler sBtn(0).Click, AddressOf btnMinimum_Click
        AddHandler sBtn(1).Click, AddressOf btnClose_Click

        'Label ====================================================================================
        Me.sLbl = New System.Windows.Forms.Label(0) {}
        For i As Integer = 0 To 0
            Me.sLbl(i) = New System.Windows.Forms.Label
            Me.Controls.Add(sLbl(i))
            AddHandler sLbl(i).MouseDown, AddressOf Window_MouseDown
            AddHandler sLbl(i).MouseMove, AddressOf Window_MouseMove
        Next
        Me.ResumeLayout()

        'Standard Form ============================================================================
        With Me
            .Text = "Equipment condition"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 610 : .Height = 380
            .Top = (Screen.PrimaryScreen.Bounds.Height \ 2) - (.Height \ 2)
            .Left = (Screen.PrimaryScreen.Bounds.Width \ 2) - (.Width \ 2)
        End With

        Try
            cFrm = New cls090_DesignForm_Type01

            cFrm.FORM_NAME = Me
            cFrm.PIC_FOLDER = mAppPath & "900 Picture\Picture\WhiteForm\"
            cFrm.TITLE_BAR_L = sBar(0) : cFrm.TITLE_BAR_R = sBar(1)
            cFrm.UNDER_BAR_L = sBar(2) : cFrm.UNDER_BAR_R = sBar(3)
            cFrm.MIN_BUTTON = sBtn(0)
            cFrm.CLOSE_BUTTON = sBtn(1)
            cFrm.PANEL_TITLE = sLbl(0)

            Call cFrm.FORM_DRAW()

            Call cFrm.TEXT_DRAW(sLbl(0), Me.Text, "Times new roman", 8, 5, 7, Color.Gray)

        Catch ex As Exception
            mErrorComment = cFrm.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cFrm = Nothing

        End Try
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, "Condition Form Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        'Only this form ===========================================================================
        Call Me.MainProcess()

    End Sub

    Private Sub frm100_Condition_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        '==============================================================================================================
        'Function : After display
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

        DoEvents()

        Me.Opacity = 1.0

        Me.Activate()

    End Sub

    Private Sub Window_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) _
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
        mDialogResult = ""

        Me.Close()

    End Sub

    Private Sub frm100_Condition_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '==========================================================================================
        'Function : Close
        'Remarks  :
        '==========================================================================================

        Dim ret As DialogResult

        If mStartOpen = True And mDialogResult = "" Then
            If mErrorComment = "" Then
                'ret = MessageBox.Show("終了します。よろしいですか？", "Program End", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
                ret = MessageBox.Show("종료합니다. 진행합니까？", "Program End", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
                If ret = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                Else
                    Call Me.ProgramEnd()
                    End
                End If
            Else
                'MessageBox.Show("終了します。", "Program End", MessageBoxButtons.OK, MessageBoxIcon.Information)
                MessageBox.Show("종료합니다.", "Program End", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Call Me.ProgramEnd()
                End
            End If
        End If

        timTimer(0).Enabled = False

    End Sub

    Private Sub ProgramEnd()
        '==============================================================================================================
        ' Function : Program end
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

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

        Dim pNum As Integer = 0
        Dim pName As String = ""

        mEquipmentStatus = 2

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "●　装置状況確認 ------------------------ (" & Me.Name & "起動)")

        'Form Draw ================================================================================
        Call Me.FormDraw()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, "Condition Form Draw", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        'Get Server file ===========================================================================
        'pNum = 0 : pName = "サーバデータ取得"
        pNum = 0 : pName = "SERVER DATA 취득"

        Call Me.GetServerFile()

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        mErrorComment = "" : mWarningComment = "" : mNormalComment = ""

        'Date & Time of adjustment ================================================================
        'pNum = 1 : pName = "自動時刻調整"
        pNum = 1 : pName = "자동시각조정"

        Call Me.DateTimeAdjust()

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        mErrorComment = "" : mWarningComment = "" : mNormalComment = ""

        'Check of signal tower ====================================================================
        'pNum = 2 : pName = "サインタワーの確認"
        pNum = 2 : pName = "사인타워의 확인"

        Call Me.SignalTowerCheck()

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        mErrorComment = "" : mWarningComment = "" : mNormalComment = ""

        'Read of Temperature & Humidity ===========================================================
        'pNum = 3 : pName = "温・湿度の設定読込"
        pNum = 3 : pName = "온・습도의 설정읽기"

        'Call Me.ReadTempHumi()     'Ver 0.37 GYS DEL
        Call Me.ReadTempHumi2()     'Ver 0.37 GYS ADD

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        mErrorComment = "" : mWarningComment = "" : mNormalComment = ""

        Call Me.SysLogWrite(False, "")
        Select Case mEquipmentStatus
            Case 0 : Call Me.SysLogWrite(True, "装置状況:異常")
            Case 1 : Call Me.SysLogWrite(True, "装置状況:注意")
            Case 2 : Call Me.SysLogWrite(True, "装置状況:正常")
        End Select

    End Sub

    Private Sub FormDraw()
        '==============================================================================================================
        ' Function : Form draw
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        'DialogResult button ======================================================================
        Me.btnDialogResult = New System.Windows.Forms.Button(1) {}
        For i As Integer = 0 To 1
            Me.btnDialogResult(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnDialogResult(i))
            AddHandler btnDialogResult(i).Click, AddressOf DialogResult_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With btnDialogResult(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 12, FontStyle.Bold)
                .Width = 140 : .Height = 40
                .Top = Me.Height - (.Height + Me.sBar(2).Height + 20)
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\WhiteForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                If i = 0 Then
                    .Left = CInt((Me.Width / 2) - .Width - 20)
                    .Text = "OK"
                End If
                If i = 1 Then
                    .Left = CInt((Me.Width / 2) + 20)
                    .Text = "CANCEL"
                End If
            End With
        Next

        'Base point ===============================================================================
        Dim bTop As Integer = sBar(0).Top + sBar(0).Height + 20
        Dim bLeft As Integer = 50
        Dim bWidth As Integer = Me.Width - (bLeft * 2)

        'Title Label ==============================================================================
        Me.lblTitle = New System.Windows.Forms.Label(2) {}
        For i As Integer = 0 To 2
            Me.lblTitle(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblTitle(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 2
            With lblTitle(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Top = bTop : .Height = 16
                .TextAlign = ContentAlignment.MiddleCenter
                .BackColor = Color.FromArgb(0, 0, 0)
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\Graylabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()

                If i = 0 Then
                    .Width = CInt(bWidth * 0.25) : .Left = bLeft
                    .Text = "項目"
                End If
                If i = 1 Then
                    .Width = CInt(bWidth * 0.15) : .Left = lblTitle(i - 1).Left + lblTitle(i - 1).Width + 1
                    .Text = "結果"
                End If
                If i = 2 Then
                    .Width = CInt(bWidth * 0.6) : .Left = lblTitle(i - 1).Left + lblTitle(i - 1).Width + 1
                    .Text = "記事"
                End If

            End With
        Next

        Dim rec As Integer = 4
        'Item Label ==============================================================================
        Me.lblItem = New System.Windows.Forms.Label(rec) {}
        For i As Integer = 0 To rec
            Me.lblItem(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To rec
            With lblItem(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = lblTitle(0).Width : .Left = lblTitle(0).Left : .Height = 16
                .TextAlign = ContentAlignment.MiddleLeft
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()

                If i = 0 Then
                    .Top = bTop + lblTitle(0).Height + 2
                Else
                    .Top = lblItem(i - 1).Top + lblItem(i - 1).Height + 2
                End If
            End With
        Next

        'Result Label ==============================================================================
        Me.lblResult = New System.Windows.Forms.Label(rec) {}
        For i As Integer = 0 To rec
            Me.lblResult(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblResult(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To rec
            With lblResult(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = lblTitle(1).Width : .Left = lblTitle(1).Left : .Height = 16
                .TextAlign = ContentAlignment.MiddleCenter
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()

                If i = 0 Then
                    .Top = bTop + lblTitle(1).Height + 2
                Else
                    .Top = lblItem(i - 1).Top + lblItem(i - 1).Height + 2
                End If
            End With
        Next

        'Comment Label ==============================================================================
        Me.lblComment = New System.Windows.Forms.Label(rec) {}
        For i As Integer = 0 To rec
            Me.lblComment(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblComment(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To rec
            With lblComment(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = lblTitle(2).Width : .Left = lblTitle(2).Left : .Height = 16
                .TextAlign = ContentAlignment.MiddleLeft
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()

                If i = 0 Then
                    .Top = bTop + lblTitle(2).Height + 2
                Else
                    .Top = lblItem(i - 1).Top + lblItem(i - 1).Height + 2
                End If
            End With
        Next

        'Set Item Label ===========================================================================
        Me.lblSetItem = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblSetItem(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblSetItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblSetItem(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Top = lblItem(rec).Top + lblItem(0).Height + 20
                .Width = 162 : .Height = 16
                .TextAlign = ContentAlignment.MiddleCenter
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()

                If i = 0 Then
                    .Left = lblItem(0).Left
                Else
                    .Left = lblSetItem(i - 1).Left + lblSetItem(i - 1).Width + 25
                End If
                If i = 0 Then .Text = " 周囲温度(℃) / 湿度(%)"
                If i = 1 Then .Text = " 日付 / 時刻"
            End With
        Next

        'Temperature & Humidity ===================================================================
        Me.cmbTempHumi = New System.Windows.Forms.ComboBox(1) {}
        For i As Integer = 0 To 1
            Me.cmbTempHumi(i) = New System.Windows.Forms.ComboBox
            Me.Controls.Add(cmbTempHumi(i))
            AddHandler cmbTempHumi(i).KeyDown, AddressOf TempHumi_KeyDown
            AddHandler cmbTempHumi(i).KeyPress, AddressOf TempHumi_KeyPress
            AddHandler cmbTempHumi(i).SelectedIndexChanged, AddressOf TempHumi_SelectedIndexChanged
            AddHandler cmbTempHumi(i).TextChanged, AddressOf TempHumi_TextChanged
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With cmbTempHumi(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Name = "cmbTempHumi(" & CStr(i) & ")"
                .Width = CInt((lblSetItem(0).Width - 2) / 2)
                .Top = lblSetItem(0).Top + lblSetItem(0).Height + 2
                .BackColor = Color.FromArgb(0, 0, 150)
                .ForeColor = Color.FromArgb(255, 255, 255)
                .Cursor = Cursors.Hand
                .BringToFront()

                If i = 0 Then
                    .Left = lblSetItem(0).Left
                Else
                    .Left = cmbTempHumi(i - 1).Left + cmbTempHumi(i - 1).Width + 2
                End If
            End With
        Next

        'Date & Time ==============================================================================
        Me.lblDateTime = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblDateTime(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblDateTime(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblDateTime(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 9, FontStyle.Regular)
                .Width = CInt((lblSetItem(0).Width - 2) / 2) : .Height = 18
                .Top = lblSetItem(1).Top + lblSetItem(1).Height + 4
                .TextAlign = ContentAlignment.MiddleCenter
                .BackColor = Color.CornflowerBlue
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\BlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()

                If i = 0 Then
                    .Left = lblSetItem(1).Left
                Else
                    .Left = lblDateTime(i - 1).Left + lblDateTime(i - 1).Width + 2
                End If
                If i = 0 Then .Text = Format(DateTime.Now, "yyyy.MM.dd")
                If i = 1 Then .Text = Format(DateTime.Now, "HH.mm.ss")

            End With
        Next

        'Date & Time Application ==================================================================
        Me.btnDateTimeExe = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnDateTimeExe(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnDateTimeExe(i))
            AddHandler btnDateTimeExe(i).Click, AddressOf btnDateTimeExe_Click
        Next

        For i As Integer = 0 To 0
            With btnDateTimeExe(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Bold)
                .Width = lblSetItem(1).Width : .Height = 20
                .Left = lblSetItem(1).Left
                .Top = lblDateTime(0).Top + lblDateTime(0).Height + 4
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\WhiteForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                '.Text = "時刻設定"
                .Text = "시각설정"
            End With
        Next

        'Setting Panel ============================================================================
        Me.btnEquipmentSetting = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnEquipmentSetting(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnEquipmentSetting(i))
            AddHandler btnEquipmentSetting(i).Click, AddressOf btnEquipmentSetting_Click
        Next

        For i As Integer = 0 To 0
            With btnEquipmentSetting(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Bold)
                .Width = 100 : .Height = 40
                .Left = bLeft + bWidth - .Width
                .Top = lblSetItem(0).Top
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\WhiteForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                '.Text = "装置の" & vbCr & "条件設定"
                .Text = "장치의" & vbCr & "조건설정"
            End With
        Next

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

    Private Sub GetServerFile()
        '==============================================================================================================
        ' Function : Get Server file
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        'Server Connect ===========================================================================
        Try
            cSrv = New cls110_ServerConnect

            cSrv.FORM_NAME = Me
            Call cSrv.ServerConnect()

            mServerDrive = cSrv.SERVER_DRIVE

        Catch ex As Exception
            mWarningComment = cSrv.ERR_COMMENT
            If mWarningComment = "" Then mWarningComment = ex.ToString()

            mServerDrive = ""

        Finally
            cSrv = Nothing

        End Try
        'If mServerDrive = "" Then mWarningComment = "サーバーに接続できません。"
        If mServerDrive = "" Then mWarningComment = "SERVER에 접속이 불가능합니다."

        If mWarningComment = "" Then
            Call Me.SysLogWrite(True, "サーバー接続：" & mServerDrive)
        Else
            Call Me.SysLogWrite(True, "サーバー接続：" & mWarningComment)
            Exit Sub
        End If




        'Get of Common information ================================================================
        Call Me.SysLogWrite(False, "")

        Dim oFileName() As String = {}
        Dim mFoldName() As String = {}
        Dim mFileName() As String = {}

        Dim mFileCount As Integer = 1

        ReDim oFileName(mFileCount)
        ReDim mFoldName(mFileCount)
        ReDim mFileName(mFileCount)

        'Path -------------------------------------------------------
        oFileName(0) = Mid(mServerDrive, 1, 2) & "\001 Run system\Common File\aiec\Common Data\NameEntry.mdb "
        mFoldName(0) = mAppPath & "400 Common File\"
        mFileName(0) = "NameEntry.mdb"

        oFileName(1) = Mid(mServerDrive, 1, 2) & "\001 Run system\Common File\aiec\Common Data\MSCODE.mdb "
        ' A8 GET SERVE FILE
        oFileName(1) = Mid(mServerDrive, 1, 2) & "\001 Run system\Common File\aiec\Common Data\MSCODE.mdb "
        mFoldName(1) = mAppPath & "400 Common File\"
        mFileName(1) = "MSCODE.mdb"

        'Copy -------------------------------------------------------
        For i As Integer = 0 To mFileCount

            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = oFileName(i)
                cFop.FOLDER_PATH = mFoldName(i)
                cFop.FILE_NAME = mFileName(i)

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()
                Call Me.SysLogWrite(True, "共通ファイル取得：" & mErrorComment)

                mNormalComment = ""

            Finally
                cFop = Nothing

            End Try

            If mErrorComment = "" Then
                Call Me.SysLogWrite(True, "共通ファイル取得：" & mFileName(i))
                mNormalComment &= mFileName(i) & "/"
            Else
                Call Me.SysLogWrite(True, "共通ファイル取得：" & mFileName(i) & "　失敗")
                Exit Sub
            End If
        Next

        Call Me.SysLogWrite(True, "共通ファイル取得：完了")

        'Server Disconnect ========================================================================
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

    End Sub

    Private Sub DateTimeAdjust()
        '==============================================================================================================
        ' Function : Date & Time of adjustment
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        'Read of App. setting =====================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "DateTimeApp,DateTimeAppPath,DateTimeAppName" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "時刻調整関連の設定読出し：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Check of setting data ==================================================================== 
        If sData(0) = "" Or (sData(0) <> "" And sData(1) = "") Or (sData(0) <> "" And sData(2) = "") Then
            mWarningComment = "時刻調整の設定情報が異常です。確認して下さい。"
            mServerDrive = ""
            Call Me.SysLogWrite(True, "時刻調整設定：" & mWarningComment)
            Exit Sub
        End If

        If sData(0) = "No" Then
            'mWarningComment = "時刻調整設定：実施しない。"      'Rev 0.29 KYS DEL
            'Call Me.SysLogWrite(True, mWarningComment)          'Rev 0.29 KYS DEL

            mNormalComment = "時刻調整設定：実施しない。"      'Rev 0.29 KYS ADD
            Call Me.SysLogWrite(True, mNormalComment)          'Rev 0.29 KYS ADD
            Exit Sub
        End If

        'App. Start ===============================================================================
        Dim AppPath As String = mAppPath & sData(1)
        Dim AppTitle As String = sData(2)

        'Auto adjust start ------------------------------------------
        Try
            cWin = New cls001_Window_Control

            cWin.WINDOW_TITLE = AppTitle

            Call cWin.GET_WINDOW_HNDL()

            If cWin.WIN_HNDL = 0 Then
                Call Me.SysLogWrite(True, "時刻調整　起動　・・・")

                cWin.PROCESS_NAME = AppPath
                Call cWin.CREATE_PROCESS()

                'App start check --------------------------
                Dim i As Integer
                For i = 0 To 9
                    Call Me.WaitTimer(0.5)

                    Call cWin.GET_WINDOW_HNDL()
                    If cWin.WIN_HNDL <> 0 Then Exit For
                Next

                If i >= 10 Then
                    mWarningComment = "時刻調整のプログラムが起動できません。"
                    Call Me.SysLogWrite(True, mWarningComment)
                Else
                    mNormalComment = "時刻調整のプログラム　起動確認"
                    Call Me.SysLogWrite(True, mNormalComment)
                End If

                'App end check ----------------------------
                For i = 0 To 9
                    Call Me.WaitTimer(1)

                    Call cWin.GET_WINDOW_HNDL()
                    If cWin.WIN_HNDL = 0 Then Exit For
                Next

                If i >= 10 Then
                    'mWarningComment = "時刻調整が終了しません。動作を確認して下さい。"
                    mWarningComment = "시각조정이 종료하지 않습니다. 동작을 확인해주세요."
                    Call Me.SysLogWrite(True, "時刻調整：自動終了しないので強制終了")
                    Call cWin.WINDOW_QUIT()
                Else
                    'mNormalComment = "時刻調整：完了"
                    mNormalComment = "시각조정：완료"
                    Call Me.SysLogWrite(True, mNormalComment)
                End If

            Else
                mWarningComment = AppTitle & "は既に起動しています。"
                Call Me.SysLogWrite(True, "時刻調整：起動済みのため終了のみ実施")
                Call cWin.WINDOW_QUIT()
            End If

        Catch ex As Exception
            mWarningComment = "時刻調整のプログラムが起動できません。"
            Call Me.SysLogWrite(True, mWarningComment)

        Finally
            cWin = Nothing

        End Try

    End Sub

    Private Sub SignalTowerCheck()
        '==============================================================================================================
        ' Function : Signal tower information
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        Call Me.SysLogWrite(False, "")

        Call Me.SysLogWrite(True, "サインタワー確認　・・・")

        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "サインタワーの動作が確認できないため、未使用に設定。")
            Exit Sub
        End If

        'mNormalComment = "動作確認：完了"
        mNormalComment = "동작확인：완료"
        Call Me.SysLogWrite(True, "サインタワーの" & mNormalComment)

    End Sub
    '-----Ver0.37 GYS DEL START-----
    'Private Sub ReadTempHumi()
    '    '==============================================================================================================
    '    ' Function : Form draw by process result
    '    ' Data     : Nothing
    '    ' Return   : Nothing
    '    ' Remarks  :
    '    '==============================================================================================================

    '    Call Me.SysLogWrite(False, "")

    '    Dim dbName As String = ""
    '    Dim tbName As String = ""
    '    Dim fName() As String = {""}
    '    Dim sData() As String = {""}
    '    Dim sSQL As String = ""
    '    Dim fCount As Integer = 0
    '    Dim rCount As Integer = 0

    '    For i As Integer = 0 To 1
    '        cmbTempHumi(i).Items.Clear()
    '    Next
    '    For i As Integer = 160 To 300 Step 5
    '        cmbTempHumi(0).Items.Add(Format(i / 10, "##.0"))
    '    Next
    '    For i As Integer = 200 To 800 Step 5
    '        cmbTempHumi(1).Items.Add(Format(i / 10, "##.0"))
    '    Next


    '    'Read of select temp & humi ===============================================================
    '    Call Me.SysLogWrite(True, "温度・湿度設定読込　・・・")

    '    Try
    '        dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
    '        tbName = "tblDefTempHumi"
    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName

    '        sSQL = "SELECT " & "Temperature,Humidity" & " FROM " & tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString()
    '        Call Me.SysLogWrite(True, "温度・湿度読込：" & mErrorComment)

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    cmbTempHumi(0).Text = sData(0)
    '    cmbTempHumi(1).Text = sData(1)

    '    'mNormalComment = "温度・湿度設定読込：" & sData(0) & "℃　" & sData(1) & "%"
    '    mNormalComment = "온도・습도설정：" & sData(0) & "℃　" & sData(1) & "%"
    '    Call Me.SysLogWrite(True, mNormalComment)

    'End Sub
    '-----Ver0.37 GYS DEL END-----

    '-----Ver0.37 GYS ADD START-----
    Private Sub ReadTempHumi2()
        '==============================================================================================================
        ' Function : Form draw by process result
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        For i As Integer = 0 To 1
            cmbTempHumi(i).Items.Clear()
        Next
        For i As Integer = 160 To 300 Step 5
            cmbTempHumi(0).Items.Add(Format(i / 10, "##.0"))
        Next
        For i As Integer = 200 To 800 Step 5
            cmbTempHumi(1).Items.Add(Format(i / 10, "##.0"))
        Next


        'Read of select temp & humi ===============================================================
        Call Me.SysLogWrite(True, "温度・湿度設定読込　・・・")


        cmbTempHumi(0).Text = humid_temp("1916")
        cmbTempHumi(1).Text = humid_temp("1918")

        mNormalComment = "온도・습도설정：" & cmbTempHumi(0).Text & "℃　" & cmbTempHumi(1).Text & "%"

        Call Me.SysLogWrite(True, mNormalComment)

    End Sub
    Public Function humid_temp(ByVal rg_num As String) As String


        Dim SetupFileName As String = mAppPath & "010 EquipmentData\SetupNeptune.ini"
        myModbus1 = New clsNeptune(4, SetupFileName)

        Try
            myModbus1.EtherOpen()
            Call WaitTimer(0.5)

            '온도, 습도 취득
            myModbus1.SendReceive2(a, "000100000006" & "0103" & rg_num & "0002")

            Revdata = myModbus1.IEEEtoDECIMAL_change(VB.Right(a, 8))

            ' 온도,습도 정상범위 이외
            If rg_num = "1916" Then ' 온도
                If Revdata < 18 Or Revdata > 28 Then
                    mWarningComment = "온도 범위 벗어남!온도를 다시 확인해주세요!!"
                    Call Me.SysLogWrite(True, mWarningComment)
                    MsgBox("적정범위를 벗어났습니다!!  L/L를 불러주세요!! " & VB.vbCrLf & "현재 온도 : " & Revdata & ", Spec : 18 ~ 28", MsgBoxStyle.Critical)

                End If

            ElseIf rg_num = "1918" Then '습도

                If Revdata < 20 Or Revdata > 80 Then
                    mWarningComment = "습도 범위 벗어남!습도를 다시 확인해주세요!!"
                    Call Me.SysLogWrite(True, mWarningComment)
                    MsgBox("적정범위를 벗어났습니다!! L/L를 불러주세요!! " & VB.vbCrLf & "현재 습도 : " & Revdata & ", Spec : 20 ~ 80", MsgBoxStyle.Critical)

                End If
            End If
            myModbus1.Close()

        Catch ex As System.Exception
            mErrorComment = "온습도 데이터 취득에 실패했습니다. 온습도 수기입 해주세요"
            Call Me.SysLogWrite(True, mErrorComment)
            'MyBase.Close()
        End Try

        Return Format(Revdata, "00.0")

    End Function
    '-----Ver0.37 GYS ADD END-----


    Private Sub ResultDisp(ByVal n As Integer, ByVal s As String, ByVal e As String, ByVal w As String, ByVal c As String)
        '==============================================================================================================
        ' Function : Form draw by process result
        ' Data     : Process number , Process name , Error comment , Warning comment , Normal comment
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        lblItem(n).Text = " " & s

        'ERROR ====================================================================================
        If e <> "" Then
            mEquipmentStatus = 0

            lblItem(n).ForeColor = Color.Red
            lblResult(n).ForeColor = Color.Red
            lblComment(n).ForeColor = Color.Red
            lblResult(n).Text = "NG"
            lblComment(n).Text = " " & Mid(e, 1, 40)
        End If

        'WARNING ==================================================================================
        If w <> "" Then
            If mEquipmentStatus = 2 Then mEquipmentStatus = 1

            lblItem(n).ForeColor = Color.Yellow
            lblResult(n).ForeColor = Color.Yellow
            lblComment(n).ForeColor = Color.Yellow
            lblResult(n).Text = "WARNING"
            lblComment(n).Text = " " & Mid(w, 1, 40)
        End If

        'NORMAL ===================================================================================
        If e = "" And w = "" Then
            lblResult(n).Text = "OK"
            lblComment(n).Text = " " & Mid(c, 1, 40)
        End If

        '-----Ver0.37 GYS ADD Start-----
        If lblResult(0).Text = "OK" And lblResult(1).Text = "OK" And lblResult(2).Text = "OK" And lblResult(3).Text = "OK" Then

            mDialogResult = "OK"

            Call Me.SaveData()
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "データ保存失敗：" & mErrorComment)
                MessageBox.Show("Saving data NG : " & mErrorComment, "System Condition", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                mErrorComment = ""
                Exit Sub
            End If
            Call Me.SysLogWrite(True, "データ保存　完了")

            Me.Close()

        End If
        '-----Ver0.37 GYS ADD END-----

    End Sub

    'Process (Event) **************************************************************************************************
    Private Sub DialogResult_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : btnOK
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : " & DirectCast(sender, Button).Text & " ON")

        Call Me.FocusControl()

        Select Case DirectCast(sender, Button).Text

            Case "OK"
                mDialogResult = "OK"

                Call Me.SaveData()
                If mErrorComment <> "" Then
                    Call Me.SysLogWrite(True, "データ保存失敗：" & mErrorComment)
                    MessageBox.Show("Saving data NG : " & mErrorComment, "System Condition", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    mErrorComment = ""
                    Exit Sub
                End If
                Call Me.SysLogWrite(True, "データ保存　完了")

            Case "CANCEL"
                mDialogResult = ""

        End Select

        Me.Close()

    End Sub

    Private Sub SaveData()
        '==============================================================================================================
        ' Function : Data save
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")

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
            tbName = "tblDefTempHumi"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "Temperature = '" & cmbTempHumi(0).Text & "' , "
            sSQL &= "Humidity = '" & cmbTempHumi(1).Text & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

            '-----Ver0.37 GYS ADD-----
            Temp = CInt(cmbTempHumi(0).Text)
            Hum = CInt(cmbTempHumi(1).Text)

            If Temp > 18 And Temp < 28 And Hum > 20 And Hum < 80 And mEquipmentStatus = 1 Then
                mEquipmentStatus = 2
            End If

            '-----Ver0.37 GYS End-----

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cADO = Nothing

        End Try

        Call Me.SysLogWrite(True, "データ保存：" & cmbTempHumi(0).Text & "℃ " & cmbTempHumi(1).Text & "%")

    End Sub

    Private Sub TempHumi_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        '==============================================================================================================
        ' Function : Cancel of Key input(Temp. Humi.) 
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

    Private Sub TempHumi_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        '==============================================================================================================
        ' Function : Cancel of Key input(Temp. Humi.) 
        ' Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()
        e.KeyChar = Nothing

    End Sub

    Private Sub TempHumi_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Cancel of Key input(Temp. Humi.) 
        ' Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

    End Sub

    Private Sub TempHumi_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Select change of Temperature or Humidity 
        ' Remarks  :
        '==============================================================================================================

        If cmbTempHumi(0).Focused = False And cmbTempHumi(1).Focused = False Then Exit Sub

        Call Me.SysLogWrite(False, "")
        Call Me.FocusControl()

        Select Case DirectCast(sender, ComboBox).Name
            Case "cmbTempHumi(0)"
                If cmbTempHumi(0).Text = "" Then Exit Sub
                Call Me.SysLogWrite(True, "温度：" & cmbTempHumi(0).Text & "℃に変更")

            Case "cmbTempHumi(1)"
                If cmbTempHumi(1).Text = "" Then Exit Sub
                Call Me.SysLogWrite(True, "湿度：" & cmbTempHumi(1).Text & "%に変更")

        End Select

    End Sub

    Private Sub btnDateTimeExe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Windows calender open
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Text & " カレンダ表示　ON")

        Call Me.FocusControl()

        Try
            cWin = New cls001_Window_Control

            cWin.PROCESS_NAME = "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"
            cWin.WIN_X = 0
            cWin.WIN_Y = 0
            cWin.WIN_H = 0
            cWin.WIN_W = 0

            Call cWin.CREATE_PROCESS()

        Catch ex As Exception
            mErrorComment = cWin.ERR_COMMENT : If mErrorComment = "" Then mErrorComment = ex.ToString
            MessageBox.Show(mErrorComment, "Date Time Property", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            cWin = Nothing

        End Try

        Try
            cWin = New cls001_Window_Control

            cWin.WINDOW_TITLE = "日付と時刻のプロパティ"

            Do
                Call cWin.GET_WINDOW_HNDL()
                If cWin.WIN_HNDL <> 0 Then Exit Do
                Call WaitTimer(0.5)
            Loop

            Call cWin.GET_WINDOW_POSITION()

            cWin.WIN_Z = cWin.Z_TOPMOST

            Call cWin.SET_WINDOW_POSITION()

        Catch ex As Exception
            mErrorComment = cWin.ERR_COMMENT : If mErrorComment = "" Then mErrorComment = ex.ToString
            MessageBox.Show(mErrorComment, "Date Time Property", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            cWin = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

    End Sub

    Private Sub btnEquipmentSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Equipment setting panel start
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Text & " 装置設定画面の表示　ON")

        Dim sRet As String = ""

        Call Me.FocusControl()
        Me.Hide()

        'Password ---------------------------------------------------
        Dim pfrm As New frm200_InputPass
        Try
            Call pfrm.ShowDialog()
            sRet = pfrm.DIALOG_RESULT

        Finally
            pfrm = Nothing

        End Try
        If sRet <> "PASS OK" Then
            Call Me.SysLogWrite(False, "")
            Call Me.SysLogWrite(True, "●　" & Me.Name & " に戻る")
            Me.Show()
            Exit Sub
        End If

        'Open Setting panel -----------------------------------------
        Dim sfrm As New frm101_EquipmentSetting
        Try
            Call sfrm.ShowDialog()
            sRet = sfrm.DIALOG_RESULT

        Finally
            sfrm = Nothing

        End Try
        If sRet <> "OK" And sRet <> "" Then
            mErrorComment = sRet
            Me.Close()
        End If
        Call Me.SysLogWrite(False, "")

        Me.MainProcess()
        Me.Show()

    End Sub

    Private Sub Reflesh_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Reflesh timer
        ' Remarks  :
        '==============================================================================================================

        lblDateTime(0).Text = Format(DateTime.Now, "yyyy.MM.dd")
        lblDateTime(1).Text = Format(DateTime.Now, "HH:mm:ss")

    End Sub

    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : Button focus
        ' Remarks  :
        '==============================================================================================================

        If btnDialogResult(0).Visible = True Then
            btnDialogResult(0).Focus()
        Else
            btnDialogResult(1).Focus()
        End If

    End Sub

    Private Sub SysLogWrite(ByVal tStmp As Boolean, ByVal msg As String)
        '==============================================================================================================
        ' Function : System Log write
        ' Data     : Message
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