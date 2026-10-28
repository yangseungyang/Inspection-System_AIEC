Option Strict On : Option Explicit On

'Reference  :Microsoft HTML Oject Library

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm030_InspectionInformation

    'Class ================================================
    Dim cWin As cls001_Window_Control
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cDrv As cls020_Drive_Control
    Dim cFrm As cls090_DesignForm_Type01
    Dim cIns As cls2P0_Inspection
    Dim cSys As cls102_SignalTower
    Dim cBcd1 As cls101_ComRS232C

    'Local (Standard form control) ========================
    Dim sBar() As PictureBox
    Dim sBtn() As Button
    Dim sLbl() As Label

    '主銘板再印字用ボタン
    Public btnRePrintLabel As Button
    'リトライボタン
    Dim btnTestReTry As Button
    'テスト終了ボタン
    Dim btnTestEnd As Button
    '時間記録用
    Public time1 As Double
    Public time2 As Double
    Public time3 As Double
    'DLL
    Private EthercomNdll As New EtherComN.Dll

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mMousePoint As Point
    Dim mErrorComment As String = ""
    Dim mSystemLog As String = ""
    Dim mServerDrive As String = ""
    Dim mEquipmentStatus As Integer = 0

    'Process Status:0=No,1=RUN,2=STOP,3=END,4=NG STOP,5=NEXT,6=PREV,11=Retry
    Public mProcessStatus As Integer = 0

    Dim mInspectionResult As String = ""
    Dim mInspectionComment As String = ""
    Dim mInspectionTestItem As String = ""
    Dim mInspectionSpec As String = ""
    Dim mInspectionMeasure As String = ""


    'Local (Special form control) ========================
    Dim btnDialogResult() As Button
    Dim picBox() As PictureBox
    Dim dgvInspectionItem() As DataGridView
    Dim wBrowser() As WebBrowser
    Dim radManualSet() As RadioButton
    Dim pnlRepeat() As Panel
    Dim radRepeatSet() As RadioButton
    Dim nudRepeatFreq() As NumericUpDown
    Dim lblItem() As Label
    Dim lblData() As Label

    Dim mSetting As String = ""
    Dim mPort As String = ""
    Dim mDelim As String = ""

    'Local Property =======================================
    Dim mDialogResult As String = ""

    Public WriteOnly Property LogNavi() As Boolean
        Set(ByVal value As Boolean)
            Call Me.LogBrowserNavi()
        End Set
    End Property

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm030_InspectionInformation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '==============================================================================================================
        'Function : Display start
        'Remarks  :
        '==============================================================================================================

        'Equipment Information ====================================================================
        mSystemLog = frm001_Main.SYSTEM_LOG
        mEquipmentStatus = frm001_Main.EQUIPMENT_STATUS

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

        '主銘板再印字ボタン
        Me.btnRePrintLabel = New System.Windows.Forms.Button
        Me.Controls.Add(btnRePrintLabel)

        'リトライ検査ボタン
        Me.btnTestReTry = New System.Windows.Forms.Button
        Me.Controls.Add(btnTestReTry)

        '検査STOPボタン
        Me.btnTestEnd = New System.Windows.Forms.Button
        Me.Controls.Add(btnTestEnd)

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
            .Text = "frm030 Inspection Information"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            '.TopMost = True

            .Width = 1000 : .Height = 700
            .Top = (Screen.PrimaryScreen.Bounds.Height \ 2) - (.Height \ 2) - 10
            .Left = (Screen.PrimaryScreen.Bounds.Width \ 2) - (.Width \ 2)
        End With

        Try
            cFrm = New cls090_DesignForm_Type01

            cFrm.FORM_NAME = Me
            cFrm.PIC_FOLDER = mAppPath & "900 Picture\Picture\BlueForm\"
            cFrm.TITLE_BAR_L = sBar(0) : cFrm.TITLE_BAR_R = sBar(1)
            cFrm.UNDER_BAR_L = sBar(2) : cFrm.UNDER_BAR_R = sBar(3)
            cFrm.MIN_BUTTON = sBtn(0)
            cFrm.CLOSE_BUTTON = sBtn(1)
            cFrm.PANEL_TITLE = sLbl(0)

            Call cFrm.FORM_DRAW()

            Call cFrm.TEXT_DRAW(sLbl(0), Me.Text, "Times new roman", 8, 5, 7, Color.SteelBlue)

        Catch ex As Exception
            If mErrorComment <> "" Then mErrorComment = cFrm.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            MessageBox.Show(mErrorComment, Me.Text & " : Open Form", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            cFrm = Nothing
            If mErrorComment <> "" Then End

        End Try

        'Process Start ============================================================================
        Call Me.MainProcess()

    End Sub

    Private Sub frm030_InspectionInformation_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        '==============================================================================================================
        'Function : After display
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

        Call Me.LogBrowserNavi()

        Me.Opacity = 1.0

        Me.Activate()

        If lblData(5).Text = "MAN" Then
            dgvInspectionItem(0).Enabled = True
            Panel2.Enabled = True
            Exit Sub
        Else
            'dgvInspectionItem(0).Enabled = False
            btnStart.Enabled = False
            btnStop.Enabled = False
            btnPrev.Enabled = False
            btnNext.Enabled = False
            radRepeatSet(0).Enabled = False
            radRepeatSet(1).Enabled = False
            radManualSet(0).Enabled = False
            radManualSet(1).Enabled = False
            nudRepeatFreq(0).Enabled = False

            Call ModeControl()

        End If

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

        If mProcessStatus <> 0 Then Exit Sub

        mDialogResult = ""
        'If SerialPort2.IsOpen = True Then SerialPort2.Close()

        Me.Close()

    End Sub

    Private Sub frm030_InspectionInformation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '==========================================================================================
        'Function : Close
        'Remarks  :
        '==========================================================================================

        If mDialogResult = "" And mErrorComment <> "" Then mDialogResult = mErrorComment

    End Sub

    'Process **********************************************************************************************************
    Private Sub MainProcess()
        '==============================================================================================================
        ' Function : Main process of this panel
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "●　検査画面 ------------------------ (" & Me.Name & "起動)")

        'Form Draw ================================================================================
        Call Me.FormDraw()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, Me.Name & " Form Draw", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        'Starting Preprocess =================================================================
        'Folder & File ------------------------------------------
        Call MakeFolderFile()
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "フォルダ/ファイル準備：" & mErrorComment)
            Exit Sub
        Else
            Call Me.SysLogWrite(True, "フォルダ/ファイル準備：" & "完了")
        End If

        'Temporary data -----------------------------------------
        Call WriteTemporaryStartingData()
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "一時保存開始データ書込：" & mErrorComment)
            Exit Sub
        Else
            Call Me.SysLogWrite(True, "一時保存開始データ書込：" & "完了")
        End If

        'Temporary log ------------------------------------------
        Call WriteTemporaryStartingLog()
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "一時保存開始ログ書込：" & mErrorComment)
            Exit Sub
        Else
            Call Me.SysLogWrite(True, "一時保存開始ログ書込：" & "完了")
        End If

        'Form data set ============================================================================
        Call Me.FormDataSet()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, Me.Name & " Data Set", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

    End Sub

    Private Sub FormDraw()
        '==============================================================================================================
        ' Function : Form draw
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        'DialogResult button ======================================================================
        Me.btnDialogResult = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnDialogResult(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnDialogResult(i))
            AddHandler btnDialogResult(i).Click, AddressOf DialogResult_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With btnDialogResult(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Bold)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = 50 : .Height = 50
                .Top = Me.Height - (.Height + Me.sBar(2).Height + 20)
                '.ForeColor = Color.FromArgb(70, 70, 100)
                '.BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\BlueForm\BackColor.bmp")
                '.BackgroundImageLayout = ImageLayout.Stretch
                '.Cursor = Cursors.Hand
                '.BringToFront()
                .Left = CInt((Me.Width) - .Width - 20)
                .Text = "CLOSE"
                .Enabled = False
            End With
        Next

        'PictureBox ========================================================================
        Me.picBox = New System.Windows.Forms.PictureBox(3) {}
        For i As Integer = 0 To 3
            Me.picBox(i) = New System.Windows.Forms.PictureBox
            Me.Controls.Add(picBox(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 3
            With picBox(i)
                If i = 0 Then
                    .Top = sBar(0).Top + sBar(0).Height : .Left = 0 : .Width = Me.Width : .Height = 30
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\UPic.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                If i = 1 Then
                    .Top = picBox(0).Top + picBox(0).Height : .Left = 0 : .Width = 260 : .Height = Me.Height - sBar(0).Height - picBox(0).Height - sBar(2).Height
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\LPic.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                If i = 2 Then
                    .Top = picBox(0).Top + picBox(0).Height : .Left = picBox(1).Left + picBox(1).Width : .Width = 433 : .Height = picBox(1).Height
                    '.BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\CPic.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                If i = 3 Then
                    .Top = picBox(0).Top + picBox(0).Height : .Left = picBox(2).Left + picBox(2).Width : .Width = Me.Width - picBox(1).Width - picBox(2).Width : .Height = picBox(1).Height
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\RPic.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If

            End With
        Next

        'DataGridView ======================================================================
        Me.dgvInspectionItem = New System.Windows.Forms.DataGridView(0) {}
        For i As Integer = 0 To 0
            Me.dgvInspectionItem(i) = New System.Windows.Forms.DataGridView
            Me.Controls.Add(dgvInspectionItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With dgvInspectionItem(i)
                'Display ------------------------------------------------
                .Columns.Clear()
                .Top = picBox(0).Top + picBox(0).Height + 2 : .Height = Me.Height - .Top - sBar(2).Height
                .Left = 6 : .Width = picBox(1).Width - 9
                .Dock = DockStyle.None
                .BorderStyle = BorderStyle.FixedSingle

                .CellBorderStyle = DataGridViewCellBorderStyle.Single

                .RowCount = 2
                .RowHeadersVisible = False

                .ColumnHeadersVisible = True
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                .ColumnHeadersDefaultCellStyle.Font = New Font("MS Gothic", 8, FontStyle.Bold)

                .ScrollBars = ScrollBars.Vertical

                'Edit mode ----------------------------------------------
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = False
                .AllowUserToResizeColumns = False
                .AllowUserToResizeRows = False
                .ReadOnly = True        'False
                .Enabled = True

                'Selection ----------------------------------------------
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .MultiSelect = False

                'Color --------------------------------------------------
                .BackgroundColor = Color.White

                .RowsDefaultCellStyle.BackColor = Color.White
                .RowsDefaultCellStyle.ForeColor = Color.RoyalBlue

                .RowsDefaultCellStyle.SelectionBackColor = Color.White
                .RowsDefaultCellStyle.SelectionForeColor = Color.RoyalBlue

                .GridColor = Color.Gray

                'Column -------------------------------------------------
                .ColumnCount = 0

                Dim column As New DataGridViewCheckBoxColumn
                .Columns.Add(column)
                .RowCount = 1
                .ColumnCount = 2
                .Columns(0).Name = ""
                .Columns(1).Name = "Inspection Item"

                .Columns(0).Width = 30
                .Columns(1).Width = .Width - .Columns(0).Width

                .BringToFront()
            End With
        Next

        'wBrowserser ===============================================================================
        Me.wBrowser = New System.Windows.Forms.WebBrowser(0) {}
        For i As Integer = 0 To 0
            Me.wBrowser(i) = New System.Windows.Forms.WebBrowser
            Me.Controls.Add(wBrowser(i))
            AddHandler wBrowser(i).DocumentCompleted, AddressOf wBrowser_DocumentCompleted
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With wBrowser(i)
                .Width = picBox(2).Width - 1 : .Height = dgvInspectionItem(0).Height - 115
                .Top = dgvInspectionItem(0).Top + 1
                .Left = picBox(2).Left
                .BringToFront()
            End With
        Next

        'Timer ===============================================================================
        With Panel1
            .Width = wBrowser(0).Width : .Height = 46
            .Top = wBrowser(0).Top + wBrowser(0).Height + 2
            .Left = wBrowser(0).Left
            .BringToFront()
        End With

        With tmWaite
            .Font = New Font("MS Gothic", 12, FontStyle.Regular)
            .SettingValue = 0.0
            .CountUpDown = ToY_StopWatch.ToY_StopWatch.cUpDown.Down
            .ForeColor = Color.DarkSlateGray
            .BringToFront()
        End With

        With tmCount
            .Font = New Font("MS Gothic", 12, FontStyle.Regular)
            .SettingValue = 0.0
            .CountUpDown = ToY_StopWatch.ToY_StopWatch.cUpDown.Up
            .ForeColor = Color.DarkSlateGray
            .BringToFront()
        End With

        With tmInspectionTime
            .Font = New Font("MS Gothic", 12, FontStyle.Regular)
            .SettingValue = 0.0
            .CountUpDown = ToY_StopWatch.ToY_StopWatch.cUpDown.Up
            .ForeColor = Color.DarkSlateGray
            .BringToFront()
        End With

        'Manual operation =========================================================================
        With Panel2
            .Width = wBrowser(0).Width : .Height = 62
            .Top = Panel1.Top + Panel1.Height + 2
            .Left = Panel1.Left
            .BringToFront()
        End With

        With btnStart
            .Cursor = Cursors.Hand
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\RUN.bmp")
            .BackgroundImageLayout = ImageLayout.Center
        End With

        With btnEnd
            .Cursor = Cursors.Hand
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\END.bmp")
            .BackgroundImageLayout = ImageLayout.Center
        End With

        With btnStop
            .Cursor = Cursors.Hand
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\STOP.bmp")
            .BackgroundImageLayout = ImageLayout.Center
        End With

        With btnPrev
            .Cursor = Cursors.Hand
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\PREV.bmp")
            .BackgroundImageLayout = ImageLayout.Center
        End With

        With btnNext
            .Cursor = Cursors.Hand
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NEXT.bmp")
            .BackgroundImageLayout = ImageLayout.Center
        End With

        'Manual Operation =========================================================================
        Me.radManualSet = New System.Windows.Forms.RadioButton(1) {}
        For i As Integer = 0 To 1
            Me.radManualSet(i) = New System.Windows.Forms.RadioButton
            Panel2.Controls.Add(radManualSet(i))
            AddHandler radManualSet(i).Click, AddressOf radManualSet_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With radManualSet(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = 80 : .Height = 16
                .Left = btnNext.Left + btnNext.Width + 2
                .TextAlign = ContentAlignment.MiddleLeft
                If i = 0 Then
                    .Top = 2
                Else
                    .Top = radManualSet(i - 1).Top + radManualSet(i - 1).Height + 1
                End If
                .BackColor = Color.FromArgb(50, 50, 60)
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
                .Cursor = Cursors.Hand
                If i = 0 Then
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                    .ForeColor = Color.White
                    .Text = "繰返　無"
                    .Checked = True
                End If
                If i = 1 Then
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                    .ForeColor = Color.White
                    .Text = "繰返　有"
                    .Checked = False
                End If
            End With
        Next

        'Manual Operation Panel ===================================================================
        Me.pnlRepeat = New System.Windows.Forms.Panel(0) {}
        For i As Integer = 0 To 0
            Me.pnlRepeat(i) = New System.Windows.Forms.Panel
            Panel2.Controls.Add(pnlRepeat(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With pnlRepeat(i)
                .Width = 161 : .Height = 16
                .Left = radManualSet(0).Left + radManualSet(0).Width + 2
                .Top = 18
                .BackColor = Color.FromArgb(50, 50, 70)
                .BringToFront()
            End With
        Next

        'Manual Operation Label ===================================================================
        Me.radRepeatSet = New System.Windows.Forms.RadioButton(1) {}
        For i As Integer = 0 To 1
            Me.radRepeatSet(i) = New System.Windows.Forms.RadioButton
            pnlRepeat(0).Controls.Add(radRepeatSet(i))
            AddHandler radRepeatSet(i).Click, AddressOf radRepeatSet_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With radRepeatSet(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = 80 : .Height = 16
                .Top = 0
                .TextAlign = ContentAlignment.MiddleCenter
                If i = 0 Then
                    .Left = 0
                Else
                    .Left = radRepeatSet(i - 1).Left + radRepeatSet(i - 1).Width + 1
                End If
                .BackColor = Color.FromArgb(50, 50, 60)
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
                .Cursor = Cursors.Hand
                If i = 0 Then
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\BlueLabel.bmp")
                    .ForeColor = Color.White
                    .Text = "リピート"
                    .Checked = True
                End If
                If i = 1 Then
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\BlueLabel.bmp")
                    .ForeColor = Color.White
                    .Text = "回数指定"
                    .Checked = False
                End If
            End With
        Next

        'Manual Operation Rep. Freq. ==============================================================
        Me.nudRepeatFreq = New System.Windows.Forms.NumericUpDown(0) {}
        For i As Integer = 0 To 0
            Me.nudRepeatFreq(i) = New System.Windows.Forms.NumericUpDown
            Panel2.Controls.Add(nudRepeatFreq(i))
            AddHandler nudRepeatFreq(i).Click, AddressOf nudRepeatFreq_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With nudRepeatFreq(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Minimum = 1 : .Maximum = 10000
                .Width = 80
                .Left = pnlRepeat(0).Left + pnlRepeat(0).Width \ 2
                .Top = pnlRepeat(0).Top + pnlRepeat(0).Height + 1
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BackColor = Color.FromArgb(50, 50, 70)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
            End With
        Next

        'Inspection Data ================================================================================
        Me.lblItem = New System.Windows.Forms.Label(7) {}
        For i As Integer = 0 To 7
            Me.lblItem(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 7
            With lblItem(i)
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleLeft
                .Height = 16
                .Width = 70
                .Left = picBox(3).Left + 10
                If i = 0 Then
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Width = picBox(3).Width - 20
                    .Top = picBox(3).Top + 10
                    .BorderStyle = BorderStyle.None
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.FromArgb(255, 255, 255)
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                Else
                    .Top = lblItem(i - 1).Top + lblItem(i - 1).Height + 2
                    .BorderStyle = BorderStyle.None
                    .BackColor = Color.FromArgb(50, 50, 60)
                    .ForeColor = Color.FromArgb(255, 255, 255)
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                .BringToFront()
                'If i = 0 Then .Text = "基本情報"
                'If i = 1 Then .Text = "検査型式"
                'If i = 2 Then .Text = "オーダ型式"
                'If i = 3 Then .Text = "特注型式"
                'If i = 4 Then .Text = "O-I-C"
                'If i = 5 Then .Text = "SERIAL"
                'If i = 6 Then .Text = "動作モード"
                'If i = 7 Then .Text = "開始日時"
                If i = 0 Then .Text = "기본정보"
                If i = 1 Then .Text = "검사 형식"
                If i = 2 Then .Text = "ORDER 형식"
                If i = 3 Then .Text = "특주 형식"
                If i = 4 Then .Text = "O-I-C"
                If i = 5 Then .Text = "SERIAL"
                If i = 6 Then .Text = "동작MODE"
                If i = 7 Then .Text = "개시 일시"
            End With

        Next

        Me.lblData = New System.Windows.Forms.Label(6) {}
        For i As Integer = 0 To 6
            Me.lblData(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblData(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 6
            With lblData(i)
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleLeft
                .Height = 16 : .Width = lblItem(0).Width - lblItem(1).Width - 2
                .Left = lblItem(1).Left + lblItem(1).Width + 2
                .Top = lblItem(i + 1).Top
                .BorderStyle = BorderStyle.None
                .ForeColor = Color.FromArgb(10, 10, 10)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\WhiteLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
            End With
        Next

        'リトライ検査ボタン配置
        With Me.btnTestReTry
            .Visible = True
            .AutoSize = False
            .Font = New Font("MS Mincho", 8, FontStyle.Bold)
            .TextAlign = ContentAlignment.MiddleCenter
            .Width = 50 : .Height = 50
            .Top = Me.Height - (.Height + Me.sBar(2).Height + 20)
            .FlatStyle = FlatStyle.Standard
            .Text = "ReTRY"
            .Left = picBox(3).Left + 10
            .Cursor = Cursors.Hand
            .Enabled = False
        End With
        AddHandler btnTestReTry.Click, AddressOf btnTestReTry_Click

        '検査ENDボタン配置
        With Me.btnTestEnd
            .Visible = True
            .Font = New Font("MS Mincho", 8, FontStyle.Bold)
            .Width = 50 : .Height = 50
            .Top = Me.Height - (.Height + Me.sBar(2).Height + 20)
            .FlatStyle = FlatStyle.Standard
            .Text = "END"
            .Left = picBox(3).Left + 10 + 50 + 10
            .Cursor = Cursors.Hand
            .Enabled = False
        End With
        AddHandler btnTestEnd.Click, AddressOf btnEnd_Click

        '主銘板再印字ボタン配置
        With Me.btnRePrintLabel
            .Visible = True
            .Width = 50 : .Height = 50
            .Top = btnDialogResult(0).Top - 50 - 5
            .FlatStyle = FlatStyle.Standard
            .Text = "主銘板再印字"
            .Left = btnDialogResult(0).Left
            .Cursor = Cursors.Hand
            .Enabled = False
        End With

    End Sub

    Private Sub FormDataSet()
        '==============================================================================================================
        ' Function : Form data setting
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
        Dim rField As String = ""
        Dim kField As String = ""

        'Inspection Model =========================================================================
        Try
            dbName = mAppPath & "210 Result Data\Temporary.mdb"
            tbName = "tblResult"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            rField = "InspectionModel"
            kField = ""

            sSQL = "SELECT " & rField & " FROM " & tbName

            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "検査型式読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub
        Call Me.SysLogWrite(True, "検査型式読込：" & sData(0))

        'Inspection Item ==========================================================================
        Try
            dbName = mAppPath & "500 Model Information\ModelInformation.mdb"
            tbName = "tblModelInformation"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            rField = "InspectionFileName"
            kField = "Model"

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & kField & " LIKE '" & sData(0) & "'"

            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "検査項目ファイル名読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub
        If rCount = 0 Then mErrorComment = "該当型式が登録されていません。" : Exit Sub
        Call Me.SysLogWrite(True, "検査項目ファイル名読込：" & sData(0))

        'Inspection File ==========================================================================
        Try
            dbName = mAppPath & "510 Inspection Item\" & sData(0)
            tbName = "tblInspectionItem"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            rField = "InspectionItemName"
            kField = "InspectionNo"

            sSQL = "SELECT " & rField & " FROM " & tbName & " ORDER BY " & kField

            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "検査項目読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub
        If rCount = 0 Then mErrorComment = "検査項目が登録されていません。" : Exit Sub
        Call Me.SysLogWrite(True, "検査項目読込：" & "完了")

        With dgvInspectionItem(0)
            .RowCount = rCount
            For i As Integer = 1 To rCount
                .Rows(i - 1).Selected = True
                .Rows.Item(i - 1).Cells.Item(0).Value = 1
                .Rows.Item(i - 1).Cells.Item(1).Value = sData(i - 1)
            Next
        End With

        'Navigate of wBrowserser ===================================================================
        Call Me.LogBrowserNavi()

        Call Me.SysLogWrite(True, Me.Name & " : 表示データ　読込完了")

    End Sub

    Private Sub MakeFolderFile()
        '==============================================================================================================
        ' Function : Preparation of folder & file 
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim res As String() = {}

        'Make Folder & Date file (ResultData) ***************************************************************
        'Make Folder ==============================================================================
        Dim SearchArea As String = mAppPath & "210 Result Data"
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
            If mErrorComment = "" Then mErrorComment = ex.ToString

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
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try

            If mErrorComment <> "" Then Exit Sub
        End If

        'Search date file ------------------------------------------- 日付ﾌｧｲﾙ
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\" & SearchFolder & "\"
            cFop.FILE_NAME = SearchFile & ".mdb"

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
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then copy file ---------------------------------------
        If sFile = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = SearchArea & "\" & "ResultMaster.mdb"
                cFop.FOLDER_PATH = SearchArea & "\" & SearchFolder & "\"
                cFop.FILE_NAME = SearchFile & ".mdb"

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub


        End If

        'Search month date file ------------------------------------------- 月別ﾌｧｲﾙ
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\MonthData\"             '& SearchFolder & "\"
            cFop.FILE_NAME = SearchFolder & ".mdb"

            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then copy file ---------------------------------------
        If sFile = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = SearchArea & "\" & "ResultMaster.mdb"
                cFop.FOLDER_PATH = SearchArea & "\MonthData\"         '& SearchFolder & "\"
                cFop.FILE_NAME = SearchFolder & ".mdb"

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Temporary data file ======================================================================
        'Old temporary file is deleted ------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.FILE_PATH = SearchArea & "\" & "Temporary.mdb"

            Call cFop.FILE_DELETE()

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Copy of ResultMaster to New temporary file -----------------
        Try
            cFop = New cls012_File_Operation

            cFop.ORIGINAL_PATH = SearchArea & "\" & "ResultMaster.mdb"
            cFop.FOLDER_PATH = SearchArea & "\"
            cFop.FILE_NAME = "Temporary.mdb"

            Call cFop.FILE_COPY()

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        '------------------------------------------- Rev 0.32 BSB ADD START ----------------------------------------

        'Make NG Date file (ResultData) ***************************************************************
        Dim SearchAreaNg As String = "\\10.36.14.215\Fam3\Err_Data"
        Dim SearchFileNg As String = Format(DateTime.Now, "yyyy-MM") & "_NG"

        'Search NG date file ------------------------------------------- 日付ﾌｧｲﾙ
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchAreaNg & "\"
            cFop.FILE_NAME = SearchFileNg & ".mdb"

            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then copy file ---------------------------------------
        If sFile = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = SearchAreaNg & "\" & "NG_Master.mdb"
                cFop.FOLDER_PATH = SearchAreaNg & "\"
                cFop.FILE_NAME = SearchFileNg & ".mdb"

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Make NG Date file (TestSearch) ***************************************************************
        'Search date file ------------------------------------------- 日付ﾌｧｲﾙ
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchAreaNg & "\"
            cFop.FILE_NAME = "TestSearch.mdb"

            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then copy file ---------------------------------------
        If sFile = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = SearchAreaNg & "\" & "TestSearch_Master.mdb"
                cFop.FOLDER_PATH = SearchAreaNg & "\"
                cFop.FILE_NAME = "TestSearch.mdb"

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        '------------------------------------------- Rev 0.32 BSB ADD END ----------------------------------------
        'Make Folder & log file (ResultLog) *****************************************************************
        'Make Folder ==============================================================================
        SearchArea = mAppPath & "220 Result Log"

        'Search date folder -----------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea
            cFop.FOLDER_NAME = SearchFolder

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        sFolder = String.Empty
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
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Temporary log file =======================================================================
        'Old temporary file is deleted ------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.FILE_PATH = SearchArea & "\" & "Temporary.txt"

            Call cFop.FILE_DELETE()

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cFop = Nothing

        End Try

    End Sub

    Private Sub WriteTemporaryStartingData()
        '==============================================================================================================
        ' Function : Writing of starting data
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim i As Integer
        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        Dim ProductHardREV As String = ""
        Dim ProductSoftREV As String = ""

        Dim ReleaseNum As String = ""
        Dim StyleNum As String = ""

        lblData(6).Text = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")

        'Get of start information =================================================================
        dbName = mAppPath & "200 Start Information\StartInformation.mdb"
        tbName = "tblInspectionStart"

        Try
            sSQL = "SELECT MSCODE,XJNo,SerialNo,OrderNo,ItemNo,CompNo,QICRequest,RunMode,InspectionModel,Linkage FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Dim StartData() As String = DirectCast(sData.Clone(), String())

        lblData(0).Text = sData(8)
        lblData(1).Text = sData(0)
        lblData(2).Text = sData(1)
        lblData(3).Text = sData(3) & "-" & sData(4) & "-" & sData(5)
        lblData(4).Text = sData(2)
        lblData(5).Text = sData(7)

        'Get of model information =================================================================
        dbName = mAppPath & "500 Model Information\ModelInformation.mdb"
        tbName = "tblModelInformation"

        Try
            sSQL = "SELECT InspectionFilename,StyleNum,ReleaseNum FROM " & tbName & " WHERE Model = '" & StartData(8) & "'"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        If rCount = 0 Then
            mErrorComment = "型式が定義されていません。: " & StartData(8) & " ???"
            mDialogResult = "型式が定義されていません。: " & StartData(8) & " ???"
            Exit Sub
        End If

        Dim ProductData() As String = DirectCast(sData.Clone(), String())

        'Get of REV information =================================================================
        dbName = mAppPath & "400 Common File\MSCODE.mdb"
        tbName = "tblMSCODE"

        Dim Model() As String = Split(StartData(8), "-")

        Try
            If UBound(Model) = 2 Then
                sSQL = "SELECT Model,StyleNum,ReleaseNum FROM " & tbName & " WHERE Model like '" & Model(0) & "-" & Model(1) & "%'"
            Else
                sSQL = "SELECT Model,StyleNum,ReleaseNum FROM " & tbName & " WHERE Model LIKE '" & StartData(8) & "%'"
            End If

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Select Case rCount
            Case 0
                StyleNum = ""
                ReleaseNum = ""
            Case 1
                StyleNum = sData(1)  ' StyleNum
                ReleaseNum = sData(2)  ' ReleaseNum
            Case Is > 1             '候補が複数あった場合
                For i = 0 To rCount - 1
                    If StartData(8) = sData(i * 3) Then          '全く同じ形名かどうか比較
                        StyleNum = sData(i * 3 + 1)  ' StyleNum
                        ReleaseNum = sData(i * 3 + 2)  ' ReleaseNum
                        Exit Select
                    End If
                    If InStr(sData(i * 3), "/") <> 0 Then        'SUFFIXを照合して存在すればそれを使用
                        If InStr(StartData(8), Split(sData(i * 3), "/")(1)) <> 0 Then
                            StyleNum = sData(i * 3 + 1)  ' StyleNum
                            ReleaseNum = sData(i * 3 + 2)  ' ReleaseNum
                            Exit Select
                        End If
                    End If
                Next
                If lblData(5).Text = "AUT" And rCount > 0 Then
                    mErrorComment = "該当する形名が見つかりませんでした。"
                    'Throw New ApplicationException(mErrorComment)
                End If
        End Select

        'MSCODE.mdbのStyleNum,ReleaseNumとModelInformation.mdbのStyleNum,ReleaseNumが合っているかを照合
        If lblData(5).Text = "AUT" And rCount > 0 Then
            If StyleNum <> ProductData(1) Or ReleaseNum <> ProductData(2) Then
                MsgBox("StyleNum,ReleaseNumを確認してください。" & vbCrLf & "MSCODE.mdb->StyleNum:" & StyleNum & " ReleaseNum:" & ReleaseNum & vbCrLf _
                       & "ModelInforrmation.mdb->StyleNum:" & ProductData(1) & " ReleaseNum:" & ProductData(2), MsgBoxStyle.Critical)
                mErrorComment = "StyleNum,ReleaseNum相違"
                Throw New ApplicationException(mErrorComment)
            End If
        End If

        'Get of system information ================================================================
        dbName = mAppPath & "100 System Data\SystemSetting.mdb"
        tbName = "tblSystemSetting"

        Try
            sSQL = "SELECT SystemName,SystemVersion,SystemVerDate FROM " & tbName & " WHERE SystemVersion in "
            sSQL = sSQL & "(SELECT " & "Max(SystemVersion) as maxver" & " FROM " & tbName & ")"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Dim SystemData() As String = DirectCast(sData.Clone(), String())

        'Get of equipment information =============================================================
        dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        tbName = "tblEquipmentSetting"

        Try
            sSQL = "SELECT EquipmentNumber FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Dim EquipmentData() As String = DirectCast(sData.Clone(), String())

        'Operator,Approver --------------------------------------- 
        tbName = "tblDefNameEntry"

        Try
            sSQL = "SELECT OperatorName,ApproverName FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Dim NameData() As String = DirectCast(sData.Clone(), String())

        'Temperature,Humidity --------------------------------------- 
        tbName = "tblDefTempHumi"

        Try
            sSQL = "SELECT Temperature,Humidity FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Dim TempHumiData() As String = DirectCast(sData.Clone(), String())

        'Write of temporary file ==================================================================
        'Get of field name ------------------------------------------
        dbName = mAppPath & "210 Result Data\ResultMaster.mdb"
        tbName = "tblResult"

        Try
            sSQL = "SELECT * FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Write of data ----------------------------------------------
        sSQL = "INSERT INTO " & tbName & " ("
        sSQL = sSQL & fName(0) & " , "
        sSQL = sSQL & fName(1) & " , "
        For i = 4 To 20
            sSQL = sSQL & fName(i) & " , "
        Next
        sSQL = sSQL & fName(24) & " , "
        sSQL = sSQL & fName(25) & " , "
        sSQL = sSQL & fName(31)
        sSQL = sSQL & ") VALUES ("

        sSQL = sSQL & "'" & lblData(6).Text & "' , "
        sSQL = sSQL & "'" & StartData(7) & "' , "
        For i = 0 To 6
            sSQL = sSQL & "'" & StartData(i) & "' , "
        Next
        For i = 0 To 1
            sSQL = sSQL & "'" & NameData(i) & "' , "
        Next
        For i = 0 To 1
            sSQL = sSQL & "'" & TempHumiData(i) & "' , "
        Next
        For i = 0 To 2
            sSQL = sSQL & "'" & SystemData(i) & "' , "
        Next
        sSQL = sSQL & "'" & EquipmentData(0) & "' , "
        sSQL = sSQL & "'" & StartData(8) & "' , "
        sSQL = sSQL & "'" & ProductData(0) & "' , "
        sSQL = sSQL & "'" & StyleNum & "' , "
        sSQL = sSQL & "'" & ReleaseNum & "' , "
        sSQL = sSQL & "'" & StartData(9) & "'"
        sSQL = sSQL & ");"

        dbName = mAppPath & "210 Result Data\Temporary.mdb"
        tbName = "tblResult"

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
    'Private Sub WriteTemporaryStartingData()
    '    '==============================================================================================================
    '    ' Function : Writing of starting data
    '    ' Data     : Nothing
    '    ' Return   : Nothing
    '    ' Remarks  :
    '    '==============================================================================================================
    '    Dim i As Integer
    '    Dim dbName As String = ""
    '    Dim tbName As String = ""
    '    Dim fName() As String = {""}
    '    Dim sData() As String = {""}
    '    Dim sSQL As String = ""
    '    Dim fCount As Integer = 0
    '    Dim rCount As Integer = 0

    '    Dim ProductHardREV As String = ""
    '    Dim ProductSoftREV As String = ""

    '    lblData(6).Text = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")

    '    'Get of start information =================================================================
    '    dbName = mAppPath & "200 Start Information\StartInformation.mdb"
    '    tbName = "tblInspectionStart"

    '    Try
    '        sSQL = "SELECT MSCODE,XJNo,SerialNo,OrderNo,ItemNo,CompNo,QICRequest,RunMode,InspectionModel,Linkage FROM " & tbName

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    Dim StartData() As String = DirectCast(sData.Clone(), String())

    '    lblData(0).Text = sData(8)
    '    lblData(1).Text = sData(0)
    '    lblData(2).Text = sData(1)
    '    lblData(3).Text = sData(3) & "-" & sData(4) & "-" & sData(5)
    '    lblData(4).Text = sData(2)
    '    lblData(5).Text = sData(7)

    '    'Get of model information =================================================================
    '    dbName = mAppPath & "500 Model Information\ModelInformation.mdb"
    '    tbName = "tblModelInformation"

    '    Try
    '        sSQL = "SELECT InspectionFilename,SoftRev,HardRev FROM " & tbName & " WHERE Model = '" & StartData(8) & "'"
    '        'sSQL = "SELECT InspectionFilename FROM " & tbName & " WHERE Model = '" & StartData(8) & "'"     'REV管理をMSCODE.mdbで一括管理

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    If rCount = 0 Then
    '        mErrorComment = "型式が定義されていません。: " & StartData(8) & " ???"
    '        mDialogResult = "型式が定義されていません。: " & StartData(8) & " ???"
    '        Exit Sub
    '    End If

    '    Dim ProductData() As String = DirectCast(sData.Clone(), String())

    '    'Get of REV information =================================================================
    '    dbName = mAppPath & "400 Common File\MSCODE.mdb"
    '    tbName = "tblMSCODE"

    '    Dim Model() As String = Split(StartData(8), "-")

    '    Try
    '        If UBound(Model) = 2 Then
    '            sSQL = "SELECT Model,RevRom FROM " & tbName & " WHERE Model like '" & Model(0) & "-" & Model(1) & "%'"     'REV管理をMSCODE.mdbで一括管理
    '        Else
    '            sSQL = "SELECT Model,RevRom FROM " & tbName & " WHERE Model LIKE '" & StartData(8) & "%'"     'REV管理をMSCODE.mdbで一括管理
    '        End If

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    Select Case rCount
    '        Case 0
    '            ProductHardREV = ""
    '            ProductSoftREV = ""
    '        Case 1
    '            ProductHardREV = Split(sData(1), ":")(1)
    '            ProductSoftREV = Split(sData(1), ":")(0)
    '        Case Is > 1             '候補が複数あった場合
    '            For i = 0 To rCount - 1
    '                If StartData(8) = sData(i * 2) Then          '全く同じ形名かどうか比較
    '                    ProductHardREV = Split(sData(i * 2 + 1), ":")(1)
    '                    ProductSoftREV = Split(sData(i * 2 + 1), ":")(0)
    '                    Exit Select
    '                End If
    '                If InStr(sData(i * 2), "/") <> 0 Then        'SUFFIXを照合して存在すればそれを使用
    '                    If InStr(StartData(8), Split(sData(i * 2), "/")(1)) <> 0 Then
    '                        ProductHardREV = Split(sData(i * 2 + 1), ":")(1)
    '                        ProductSoftREV = Split(sData(i * 2 + 1), ":")(0)
    '                        Exit Select
    '                    End If
    '                End If
    '            Next
    '            If lblData(5).Text = "AUT" And rCount > 0 Then
    '                mErrorComment = "該当する形名が見つかりませんでした。"
    '                'Throw New ApplicationException(mErrorComment)
    '            End If
    '    End Select

    '    'MSCODE.mdbのレブとModelInformation.mdbのレブが合っているかを照合
    '    If lblData(5).Text = "AUT" And rCount > 0 Then
    '        If ProductSoftREV <> ProductData(1) Or ProductHardREV <> ProductData(2) Then
    '            MsgBox("REVを確認してください。" & vbCrLf & "MSCODE.mdb->Soft:" & ProductSoftREV & " Hard:" & ProductHardREV & vbCrLf _
    '                   & "ModelInforrmation.mdb->Soft:" & ProductData(1) & " Hard:" & ProductData(2), MsgBoxStyle.Critical)
    '            mErrorComment = "REV相違"
    '            'Throw New ApplicationException(mErrorComment)
    '        End If
    '    End If

    '    'Get of system information ================================================================
    '    dbName = mAppPath & "100 System Data\SystemSetting.mdb"
    '    tbName = "tblSystemSetting"

    '    Try
    '        sSQL = "SELECT SystemName,SystemVersion,SystemVerDate FROM " & tbName & " WHERE SystemVersion in "
    '        sSQL = sSQL & "(SELECT " & "Max(SystemVersion) as maxver" & " FROM " & tbName & ")"

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    Dim SystemData() As String = DirectCast(sData.Clone(), String())


    '    'Get of equipment information =============================================================
    '    dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
    '    tbName = "tblEquipmentSetting"

    '    Try
    '        sSQL = "SELECT EquipmentNumber FROM " & tbName

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    Dim EquipmentData() As String = DirectCast(sData.Clone(), String())

    '    'Operator,Approver --------------------------------------- 
    '    tbName = "tblDefNameEntry"

    '    Try
    '        sSQL = "SELECT OperatorName,ApproverName FROM " & tbName

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    Dim NameData() As String = DirectCast(sData.Clone(), String())

    '    'Temperature,Humidity --------------------------------------- 
    '    tbName = "tblDefTempHumi"

    '    Try
    '        sSQL = "SELECT Temperature,Humidity FROM " & tbName

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    Dim TempHumiData() As String = DirectCast(sData.Clone(), String())

    '    'Write of temporary file ==================================================================
    '    'Get of field name ------------------------------------------
    '    dbName = mAppPath & "210 Result Data\ResultMaster.mdb"
    '    tbName = "tblResult"

    '    Try
    '        sSQL = "SELECT * FROM " & tbName

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString

    '    Finally
    '        cADO = Nothing

    '    End Try
    '    If mErrorComment <> "" Then Exit Sub

    '    'Write of data ----------------------------------------------
    '    sSQL = "INSERT INTO " & tbName & " ("
    '    sSQL = sSQL & fName(0) & " , "
    '    sSQL = sSQL & fName(1) & " , "
    '    For i = 4 To 20
    '        sSQL = sSQL & fName(i) & " , "
    '    Next
    '    sSQL = sSQL & fName(24) & " , "
    '    sSQL = sSQL & fName(25) & " , "
    '    sSQL = sSQL & fName(31)
    '    sSQL = sSQL & ") VALUES ("

    '    sSQL = sSQL & "'" & lblData(6).Text & "' , "
    '    sSQL = sSQL & "'" & StartData(7) & "' , "
    '    For i = 0 To 6
    '        sSQL = sSQL & "'" & StartData(i) & "' , "
    '    Next
    '    For i = 0 To 1
    '        sSQL = sSQL & "'" & NameData(i) & "' , "
    '    Next
    '    For i = 0 To 1
    '        sSQL = sSQL & "'" & TempHumiData(i) & "' , "
    '    Next
    '    For i = 0 To 2
    '        sSQL = sSQL & "'" & SystemData(i) & "' , "
    '    Next
    '    sSQL = sSQL & "'" & EquipmentData(0) & "' , "
    '    sSQL = sSQL & "'" & StartData(8) & "' , "
    '    sSQL = sSQL & "'" & ProductData(0) & "' , "
    '    sSQL = sSQL & "'" & ProductHardREV & "' , "
    '    sSQL = sSQL & "'" & ProductSoftREV & "' , "
    '    sSQL = sSQL & "'" & StartData(9) & "'"
    '    sSQL = sSQL & ");"

    '    dbName = mAppPath & "210 Result Data\Temporary.mdb"
    '    tbName = "tblResult"

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

    Private Sub WriteTemporaryStartingLog()
        '==============================================================================================================
        ' Function : Writing of starting log
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

        'Get of temporary data ====================================================================
        dbName = mAppPath & "210 Result Data\Temporary.mdb"
        tbName = "tblResult"

        Try
            sSQL = "SELECT * FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Write of temporary log ===================================================================
        Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"
        Try
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = lFile
            cTxt.WRITE_TEXT_FILE(False, sData(0) & ".txt")
            cTxt.WRITE_TEXT_FILE(False, "")

            cTxt.WRITE_TEXT_FILE(False, "■ System information")
            cTxt.WRITE_TEXT_LINE(False, "*", 50)
            cTxt.WRITE_TEXT_FILE(False, "System Name : " & sData(15))
            cTxt.WRITE_TEXT_FILE(False, "System Date : " & sData(17) & "  " & "System Rev. : " & sData(16))
            cTxt.WRITE_TEXT_FILE(False, "System No.  : " & sData(18))
            cTxt.WRITE_TEXT_FILE(False, "")

            cTxt.WRITE_TEXT_FILE(False, "■ Inspection condition")
            cTxt.WRITE_TEXT_LINE(False, "*", 50)
            cTxt.WRITE_TEXT_FILE(False, "Start Date & Time : " & sData(0))
            cTxt.WRITE_TEXT_FILE(False, "Temperature       : " & sData(13) & " degC")
            cTxt.WRITE_TEXT_FILE(False, "Humidity          : " & sData(14) & " %")
            cTxt.WRITE_TEXT_FILE(False, "Operator          : " & sData(11))
            cTxt.WRITE_TEXT_FILE(False, "Approver          : " & sData(12))
            cTxt.WRITE_TEXT_FILE(False, "Run Mode          : " & sData(1))
            cTxt.WRITE_TEXT_FILE(False, "Inspection Model  : " & sData(19))
            cTxt.WRITE_TEXT_FILE(False, "Inspection File   : " & sData(20))
            cTxt.WRITE_TEXT_FILE(False, "")

            cTxt.WRITE_TEXT_FILE(False, "■ Order Information")
            cTxt.WRITE_TEXT_LINE(False, "*", 50)
            cTxt.WRITE_TEXT_FILE(False, "MS-CODE     : " & sData(4))
            cTxt.WRITE_TEXT_FILE(False, "XJ-No       : " & sData(5))
            cTxt.WRITE_TEXT_FILE(False, "Serial No   : " & sData(6))
            cTxt.WRITE_TEXT_FILE(False, "O-I-C       : " & sData(7) & "-" & sData(8) & "-" & sData(9))
            cTxt.WRITE_TEXT_FILE(False, "QIC Request : " & sData(10))
            cTxt.WRITE_TEXT_FILE(False, "")

            cTxt.WRITE_TEXT_FILE(False, "■ Inspection")
            cTxt.WRITE_TEXT_LINE(False, "*", 50)

        Catch ex As Exception
            mErrorComment = cTxt.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cTxt = Nothing

        End Try

    End Sub

    Private Sub ModeControl()
        '==============================================================================================================
        ' Function : RUN Mode control
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim MAN_RepCount As Integer = 0

        mProcessStatus = 1

        Do
            Call InspectionTimerStart()

            MAN_RepCount += 1

            'Inspection ===========================================================================
            Call Me.InspectionControl()

            'Ending Preprocess ====================================================================
            'Temporary data -----------------------------------------
            Call WriteTemporaryEndingData()
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "一時保存終了データ書込：" & mErrorComment)
                Exit Sub
            Else
                Call Me.SysLogWrite(True, "一時保存終了データ書込：" & "完了")
            End If

            'Temporary log ------------------------------------------
            Call WriteTemporaryEndingLog()
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "一時保存終了ログ書込：" & mErrorComment)
                Exit Sub
            Else
                Call Me.SysLogWrite(True, "一時保存終了ログ書込：" & "完了")
            End If

            'Save data for Server -----------------------------------
            Call SaveServer()
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "サーバ保存：" & mErrorComment)
                Exit Sub
            Else
                Call Me.SysLogWrite(True, "サーバ保存：" & "完了")
            End If

            'AUT or CHK then end ------------------------------------
            If lblData(5).Text <> "MAN" Then
                Call Me.SysLogWrite(True, "検査動作：" & "MANではないので動作終了")
                Exit Do

            Else
                'MAN - Repetition ---------------------------------------
                If radManualSet(1).Checked = False Then
                    Call Me.SysLogWrite(True, "検査動作：" & "MAN-繰返し指定無しなので動作終了")
                    Exit Do
                End If

                'MAN - Repetition - Count up ----------------------------
                If (radRepeatSet(0).Checked = False) And (CInt(nudRepeatFreq(0).Text) <= MAN_RepCount) Then
                    Call Me.SysLogWrite(True, "検査動作：" & "繰返し指定回数を実施したので動作終了")
                    Exit Do
                End If

            End If

            'Starting Preprocess =================================================================
            'Folder & File ------------------------------------------
            Call MakeFolderFile()
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "フォルダ/ファイル準備：" & mErrorComment)
                Exit Sub
            Else
                Call Me.SysLogWrite(True, "フォルダ/ファイル準備：" & "完了")
            End If

            'Temporary data -----------------------------------------
            Call WriteTemporaryStartingData()
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "一時保存開始データ書込：" & mErrorComment)
                Exit Sub
            Else
                Call Me.SysLogWrite(True, "一時保存開始データ書込：" & "完了")
            End If

            'Temporary log ------------------------------------------
            Call WriteTemporaryStartingLog()
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "一時保存開始ログ書込：" & mErrorComment)
                Exit Sub
            Else
                Call Me.SysLogWrite(True, "一時保存開始ログ書込：" & "完了")
            End If

            Call InspectionTimerStop()
        Loop

        Call InspectionTimerStop()

    End Sub

    Private Sub InspectionControl()
        '==============================================================================================================
        ' Function : Control of Inspection
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim pCount As Integer = 0
        Dim cValue As String = ""
        Dim iName As String = ""

        mInspectionResult = ""
        mInspectionComment = ""
        mErrorComment = ""

        cSys = New cls102_SignalTower
        cSys.SignalTower("GREEN", "OFF", "OFF", "OFF")
        cSys = Nothing

        Do
            cValue = dgvInspectionItem(0).Rows.Item(pCount).Cells(0).Value.ToString

            If cValue = "1" Or cValue = "True" Then
                iName = dgvInspectionItem(0).Rows.Item(pCount).Cells(1).Value.ToString
                Call Me.SetSelColor(pCount, "R")

                Call Me.Call_Inspection(iName)

                'NG ------------------------
                If mInspectionComment <> "" Then
                    mInspectionResult = "NG"
                    Call Me.SetSelColor(pCount, "N")

                    If lblData(5).Text = "MAN" Then
                        tmInspectionTime.CountStop()
                        mProcessStatus = 4
                        Do
                            Select Case mProcessStatus
                                Case 1
                                    pCount -= 1
                                    tmInspectionTime.CountStop()
                                    Exit Do
                                Case 3  'END
                                    mProcessStatus = 0
                                    tmInspectionTime.CountStop()
                                    cSys = New cls102_SignalTower
                                    cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
                                    cSys = Nothing

                                    Call PowerOFF()
                                    Exit Sub
                                Case 5
                                    mProcessStatus = 4
                                    pCount += 1
                                    If dgvInspectionItem(0).RowCount - 1 < pCount Then pCount = dgvInspectionItem(0).RowCount - 1
                                    Call Me.SetSelColor(pCount, "R")
                                Case 6
                                    mProcessStatus = 4
                                    pCount -= 1
                                    If pCount < 0 Then pCount = 0
                                    Call Me.SetSelColor(pCount, "R")
                            End Select

                            DoEvents()
                        Loop

                    ElseIf lblData(5).Text = "AUT" Or lblData(5).Text = "CHK" Then
                        tmInspectionTime.CountStop()
                        mProcessStatus = 4
                        btnTestReTry.Enabled = True
                        btnTestEnd.Enabled = True
                        Do
                            Select Case mProcessStatus
                                Case 3  'END
                                    'btnTestReTry.Enabled = False
                                    'btnTestEnd.Enabled = False
                                    btnTestReTry.Enabled = True
                                    btnTestEnd.Enabled = True
                                    'ver 0.02 hsj add start===
                                    cSys = New cls102_SignalTower
                                    cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
                                    cSys = Nothing
                                    'ver 0.02 hsj add end===
                                    Call PowerOFF()
                                    Exit Sub
                                Case 11 'Retry
                                    cSys = New cls102_SignalTower
                                    cSys.SignalTower("GREEN", "OFF", "OFF", "OFF")
                                    cSys = Nothing

                                    'btnTestReTry.Enabled = False
                                    'btnTestEnd.Enabled = False
                                    btnTestReTry.Enabled = True
                                    btnTestEnd.Enabled = True
                                    mProcessStatus = 1  '運転再開
                                    pCount -= 1
                                    tmInspectionTime.CountStop()
                                    Exit Do

                            End Select

                            DoEvents()
                        Loop

                    Else
                        Exit Sub
                    End If
                End If

                'End ----------------------
                If mProcessStatus = 3 Then
                    mProcessStatus = 0
                    mInspectionResult = "STOP"
                    mInspectionComment = "強制終了"
                    Call Me.SetSelColor(pCount - 1, "E")
                    'ver 0.02 hsj add start===
                    cSys = New cls102_SignalTower
                    cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
                    cSys = Nothing
                    'ver 0.02 hsj add end===
                    Call PowerOFF()
                    Exit Sub
                End If

                'STOP ---------------------
                If mProcessStatus = 2 Then
                    tmInspectionTime.CountStop()
                    Do
                        Select Case mProcessStatus
                            Case 1      'RUN時
                                pCount -= 1
                                tmInspectionTime.CountStop()
                                Exit Do

                            Case 3      'END時
                                mProcessStatus = 0
                                tmInspectionTime.CountStop()
                                mInspectionResult = "STOP"
                                mInspectionComment = "強制終了"
                                Call Me.SetSelColor(pCount - 1, "E")
                                Call PowerOFF()
                                Exit Sub

                            Case 5      'NEXT時
                                mProcessStatus = 2
                                pCount += 1
                                If dgvInspectionItem(0).RowCount - 1 < pCount Then pCount = dgvInspectionItem(0).RowCount - 1
                                Call Me.SetSelColor(pCount, "R")

                            Case 6      'PREV時
                                mProcessStatus = 2
                                pCount -= 1
                                If pCount < 0 Then pCount = 0
                                Call Me.SetSelColor(pCount, "R")

                        End Select

                        DoEvents()
                    Loop


                End If
            End If

            pCount += 1

            If dgvInspectionItem(0).RowCount <= pCount Then Exit Do
            DoEvents()
        Loop

        Call Me.SetSelColor(pCount - 1, "E")

        'Signal tower All OFF -----------------
        cSys = New cls102_SignalTower
        cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
        cSys = Nothing

        mInspectionResult = "PASS"
        mInspectionComment = ""

        mProcessStatus = 0

    End Sub

    Private Sub SetSelColor(ByVal RW As Integer, ByVal sts As String)
        '==========================================================================================
        ' Function : Grid Color
        ' Data     : RowNo,Status
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim i As Integer = 0
        Dim j As Integer = 0

        If RW < 0 Then RW = 0

        With dgvInspectionItem(0)

            .Rows(RW).Selected = False

            For i = 0 To .RowCount - 1
                For j = 0 To 1
                    .Rows.Item(i).Cells.Item(j).Style.BackColor = Color.White
                    .Rows.Item(i).Cells.Item(j).Style.ForeColor = Color.RoyalBlue
                Next
            Next

            If sts = "R" Then
                For i = 0 To 1
                    .Rows.Item(RW).Cells.Item(i).Style.BackColor = Color.DarkCyan
                    .Rows.Item(RW).Cells.Item(i).Style.ForeColor = Color.White
                Next
                .Refresh()
            End If
            If sts = "N" Then
                For i = 0 To 1
                    .Rows.Item(RW).Cells.Item(i).Style.BackColor = Color.DarkRed
                    .Rows.Item(RW).Cells.Item(i).Style.ForeColor = Color.White
                Next
            End If
            If sts = "E" Then
                For i = 0 To .RowCount - 1
                    For j = 0 To 1
                        .Rows.Item(i).Cells.Item(j).Style.BackColor = Color.White
                        .Rows.Item(i).Cells.Item(j).Style.ForeColor = Color.RoyalBlue
                    Next
                Next
            End If

            .Refresh()
            .FirstDisplayedScrollingRowIndex = RW

        End With

        DoEvents()

    End Sub

    Private Sub Call_Inspection(ByVal iName As String)
        '==============================================================================================================
        ' Function : Call Inspection
        ' Data     : Inspection Name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        'Dim kannsuu As cls200_Inspection = CType("cIns." & Mid(iName, 1, 3) & "(" & iName & ")", cls200_Inspection)
        mInspectionResult = ""
        mInspectionComment = ""
        Dim dfrm As frm500_MessagePopUp

        Try
            cIns = New cls2P0_Inspection
            cIns.InspectionMODEL = Me.lblData(0).Text
            cIns.InspectionNumber = Mid(iName, 1, 3)
            cIns.InspectionName = iName

            Call cIns.TestExecute()

        Catch ex As Exception
            mInspectionComment = cIns.ERR_COMMENT
            If mInspectionComment = "" Then mInspectionComment = ex.ToString()

            '----- Rev0.32 BSB ADD START -----
            mInspectionTestItem = iName
            mInspectionSpec = cIns.TestSpec
            If mInspectionSpec = "" Then mInspectionSpec = "-----"
            mInspectionMeasure = cIns.TestMeasure
            If mInspectionMeasure = "" Then mInspectionMeasure = "-----"
            '----- Rev0.32 BSB ADD END -----

            Call Me.SysLogWrite(True, "検査実施：" & iName & " " & mInspectionComment)

            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\NGWin.ppt"
            'dfrm.LED_COLOR = "RED" : dfrm.BUZZ1 = "ON" : dfrm.BUZZ2 = "OFF" : dfrm.FLASH = "ON" : dfrm.BUZZ_TIME = 5 'Ver0.02 hsj del
            dfrm.LED_COLOR = "RED" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "OFF" : dfrm.FLASH = "OFF" : dfrm.BUZZ_TIME = 5 'Ver0.02 hsj add
            dfrm.Button1.Visible = False
            dfrm.Button2.Text = "OK"
            dfrm.Button2.Left = CInt(dfrm.Width / 2) - CInt(dfrm.Button2.Width / 2)
            dfrm.Button2.Focus()

            Call dfrm.frmShow()

            Do
                If dfrm.DIALOG_RESULT <> "" Then
                    Exit Do
                End If
                DoEvents()
            Loop

            dfrm.Close()
            dfrm = Nothing
            'ver 0.02 hsj add start===
            cSys = New cls102_SignalTower
            Call cSys.SignalTowerReset()
            cSys = Nothing
            'ver 0.02 hsj add end===

        Finally
            cIns = Nothing

        End Try

    End Sub

    Private Sub PowerOFF()
        Try

            cIns = New cls2P0_Inspection
            cIns.InspectionMODEL = Me.lblData(0).Text
            Call cIns.PowerOFF()

        Catch ex As Exception
        Finally
            cIns = Nothing
        End Try

    End Sub

    Public Sub WriteTemporaryEndingData()
        '==============================================================================================================
        ' Function : Temporary file to Date file 
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
        Dim rField As String = ""
        Dim kField As String = ""

        'Time cal =================================================================================
        Dim TIME1 As DateTime = CDate(lblData(6).Text)
        Dim TIME2 As DateTime = CDate(Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss"))
        Dim objTIME1 As TimeSpan = TIME2.Subtract(TIME1)
        Dim difTime As Integer = CInt(objTIME1.TotalSeconds)
        Dim digCountSec As Integer = 60
        Dim digCountMin As Integer = 3600
        Dim strH As String = Format(CInt(Int(difTime / digCountMin)), "0#")
        Dim strM As String = Format(CInt(Int(difTime / digCountSec) - (Int(difTime / digCountMin) * digCountSec)), "0#")
        Dim strS As String = Format(CInt(Int((difTime - Int((difTime / digCountSec)) * digCountSec))), "0#")
        Dim AllTime As String = strH & ":" & strM & ":" & strS
        Dim sH As String = Format(CInt(Int(tmInspectionTime.CountValue / digCountMin)), "0#")
        Dim sM As String = Format(CInt(Int(tmInspectionTime.CountValue / digCountSec) - (Int(tmInspectionTime.CountValue / digCountMin) * digCountSec)), "0#")
        Dim sS As String = Format(CInt(Int((tmInspectionTime.CountValue - Int((tmInspectionTime.CountValue / digCountSec)) * digCountSec))), "0#")
        Dim CountTime As String = sH & ":" & sM & ":" & sS

        'Write of End Data ========================================================================
        Try
            If InStr(mInspectionComment, "'") <> 0 Then
                mInspectionComment = Replace(mInspectionComment, "'", "-")
            End If

            dbName = mAppPath & "210 Result Data\Temporary.mdb"
            tbName = "tblResult"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "Result = '" & mInspectionResult & "' , "
            sSQL &= "ResultComment = '" & mInspectionComment & "' , "
            sSQL &= "EndDateTime = '" & TIME2 & "' , "
            sSQL &= "InspectionRunTime = '" & CountTime & "' , "
            sSQL &= "InspectionAllTime = '" & AllTime & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cADO = Nothing

        End Try

        'Read of Temp All data ====================================================================
        Try
            sSQL = "SELECT * FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Write of Date file ======================================================================= 
        Dim ymName As String = Mid(lblData(6).Text, 1, 7)
        Dim ymdName As String = Mid(lblData(6).Text, 1, 10) & ".mdb"

        dbName = mAppPath & "210 Result Data\" & ymName & "\" & ymdName
        tbName = "tblResult"

        Try
            sSQL = "INSERT INTO " & tbName & " ("

            For i = 0 To fCount - 1
                sSQL &= fName(i) & " , "
            Next
            sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 3)

            sSQL &= ") VALUES ("

            For i = 0 To fCount - 1
                sSQL &= "'" & sData(i) & "' , "
            Next
            sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 3)

            sSQL &= ")"

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
        If mErrorComment <> "" Then Exit Sub

        'Write of month Date file ======================================================================= 
        dbName = mAppPath & "210 Result Data\MonthData\" & ymName & ".mdb"
        tbName = "tblResult"

        Try
            sSQL = "INSERT INTO " & tbName & " ("

            For i = 0 To fCount - 1
                sSQL &= fName(i) & " , "
            Next
            sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 3)

            sSQL &= ") VALUES ("

            For i = 0 To fCount - 1
                sSQL &= "'" & sData(i) & "' , "
            Next
            sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 3)

            sSQL &= ")"

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
        If mErrorComment <> "" Then Exit Sub

        '--------------------------------------- Rev 0.32 BSB ADD START ---------------------------------
        'Write of Error Date file ======================================================================= 
        If mInspectionResult <> "PASS" Then
            dbName = "\\10.36.14.211\aiec\Err_Data\" & ymName & "_NG.mdb"
            tbName = "tblResult"
            Dim iName As String

            Try

                cIns = New cls2P0_Inspection
                iName = mInspectionTestItem
                '既存mdbと同じ方法でDataをCopy
                sSQL = "INSERT INTO " & tbName & " ("

                For i = 0 To fCount - 1
                    sSQL &= fName(i) & " , "
                Next
                sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 3)

                sSQL &= ") VALUES ("

                For i = 0 To fCount - 1
                    sSQL &= "'" & sData(i) & "' , "
                Next
                sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 3)

                sSQL &= ")"

                cADO = New cls010_MDB_Operation

                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName
                cADO.DB_SQL = sSQL

                Call cADO.DB_EXEC_JET()

                'Error_mdb用の検査Spec、測定値、検査名を追加
                sSQL = "UPDATE " & tbName & " SET "
                sSQL &= "TestSpec = '" & mInspectionSpec & "' , "
                sSQL &= "TestMeasure = '" & mInspectionMeasure & "' , "
                sSQL &= "TestName = '" & iName & "'"
                sSQL &= " WHERE StartDateTime = '" & sData(0) & "'"
                cADO.DB_SQL = sSQL

                Call cADO.DB_EXEC_JET()

                'NG品の“StartDateTime、MSCODE、SerialNo”をmdbに保存、処理票印字SystemのSearch用
                dbName = "\\10.36.14.211\aiec\Err_Data\TestSearch.mdb"
                tbName = "tblDate"

                sSQL = "INSERT INTO " & tbName & " (StartDateTime , MSCODE , SerialNo) VALUES ('" & sData(0) & "' , '" & sData(4) & "' , '" & sData(6) & "')"
                cADO.DB_SQL = sSQL

                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName
                Call cADO.DB_EXEC_JET()

            Catch ex As Exception
                mErrorComment = cADO.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                cADO = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        '--------------------------------------- Rev 0.32 BSB ADD END ---------------------------------

        '--------------------------------------- Ver 0.40 GYS ADD Start -------------------------------------------
        Dim mPlasmaResult As String = ""
        Dim Send_Table As String = ""
        Dim Send_clonm As String = ""
        Dim Send_Data As String = ""

        'Write of Error Date to AWS ===================================================================== 

        If mInspectionResult <> "PASS" Then

            Send_Table = "YMFK_FAM3_NG_LOG_TB"

            ' 1:RECNO        2      3        4          5       6      7       8           9            10        11     12    13
            Send_clonm = "MS_CODE,SERIAL,STARTDATE,STARTTIME,RUNMODE,RESULT,TESTNAME,RESULTCOMMENT,TESTMEASURE,TESTSPEC,XJNO,ORDERNO"
            Send_clonm = Send_clonm & ",ITEMNO,COMPNO,QICREQUEST,OPERATOR,APPROVER,TEMPERATURE,HUMIDITY,SYSTEMNAME,SYSTEMVERSION,SYSTEMVERSIONDATE"
            '                             14     15      16         17       18         19        20         21         22              23

            '                                24           25              26          27      28          29                 30        
            Send_clonm = Send_clonm & ",EQUIPMENTNO,INSPECTIONMODEL,INSPECTIONFILE,ENDDATE,ENDTIME,INSPECTIONRUNTIME,INSPECTIONALLTIME"
            Send_clonm = Send_clonm & ",PRODUCTHARDREV,PRODUCTSOFTREV,MACADDRESS,MACADDRESS2,IEEE1394ADDRESS,USBADDRESS,LINESERIAL,LINKAGE"
            '                                  31           32            33         34            35            36         37       38


            '                    2                  3                     4                                  5                       6                  7
            Send_Data = "'" & sData(4) & "','" & sData(6) & "','" & Mid(sData(0), 1, 10) & "','" & Mid(sData(0), 12, 8) & "','" & sData(1) & "','" & sData(2)
            Send_Data = Send_Data & "','" & mInspectionTestItem & "','" & sData(3) & "','" & mInspectionMeasure & "','" & mInspectionSpec & "','" & sData(5)
            '                                        8                      9                        10                         11                   12     
            '                                  13                 14                 15                 16                  17                  18     
            Send_Data = Send_Data & "','" & sData(7) & "','" & sData(8) & "','" & sData(9) & "','" & sData(10) & "','" & sData(11) & "','" & sData(12)
            Send_Data = Send_Data & "','" & sData(13) & "','" & sData(14) & "','" & sData(15) & "','" & sData(16) & "','" & sData(17) & "','" & sData(18)
            '                                  19                  20                  21                  22                  23                  24

            '                                  25                 26                        27                                28                      29     
            Send_Data = Send_Data & "','" & sData(19) & "','" & sData(20) & "','" & Mid(sData(21), 1, 10) & "','" & Mid(sData(21), 12, 8) & "','" & sData(22)
            Send_Data = Send_Data & "','" & sData(23) & "','" & sData(24) & "','" & sData(25) & "','" & sData(26) & "','" & sData(27) & "','" & sData(28)
            '                                  30                  31                  32                  33                  34                  35

            '                                  36                  37                  38      
            Send_Data = Send_Data & "','" & sData(29) & "','" & sData(30) & "','" & sData(31) & "'"

            sSQL = Send_Table & "|" & Send_clonm & "|" & Send_Data

            'mPlasmaResult = EthercomNdll.ComHost("10.7.15.71", "10.7.15.71", 2005, "QDBW-" & sSQL)     'Ver 0.43 GYS Del
            ' mPlasmaResult = EthercomNdll.ComHost("10.7.15.25", "10.7.15.140", 2031, "QDBW-" & sSQL)     'Ver 0.43 GYS ADD
            mPlasmaResult = EthercomNdll.ComHost("10.7.15.254", "10.7.15.140", 2031, "QDBW-" & sSQL)     'Ver 0.04 MKS ADD
            Debug.Print(mPlasmaResult)

        End If

        'Write of Test Date To AWS ====================================================================== 
        If mInspectionResult = "PASS" And sData(1) = "AUT" Then
            Dim RDate_S As String = Mid(sData(0), 1, 10)
            Dim RTime_S As String = Mid(sData(0), 12, 8)
            Dim RDate_E As String = Mid(sData(21), 1, 10)
            Dim RTime_E As String = Mid(sData(21), 12, 8)
            Dim WorkerID As String = ""
            Dim BYORD_Start_No As String = sData(4)
            Dim Soft_Name As String = sData(15)
            Dim Soft_Ver As String = sData(16)
            Dim sData2() As String = {""}

            'Get of OperaterID data ============================
            Try
                cADO = New cls010_MDB_Operation

                cADO.DB_PATH = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
                tbName = "tblDefNameEntry"
                cADO.TB_NAME = tbName
                rField = "Operatorid"

                sSQL = "SELECT " & rField & " FROM " & tbName
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(fName, sData2)

                If cADO.DB_DATA_COUNT = 0 Then
                    mErrorComment = "作業者情報取得ＮＧ"
                    Throw New ApplicationException(mErrorComment)
                End If

                WorkerID = sData2(0)

            Catch ex As Exception
                mErrorComment = cADO.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString
                Call Me.SysLogWrite(True, "データ取得失敗：" & mErrorComment)
            End Try


            'Get of temporary data ============================
            dbName = mAppPath & "210 Result Data\Traceability_Temp.mdb"
            tbName = "tblResult"

            sSQL = "SELECT * FROM " & tbName

            Try
                cADO = New cls010_MDB_Operation

                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(fName, sData)

                fCount = cADO.DB_FIELD_COUNT
                rCount = cADO.DB_RECORD_COUNT

                If rCount = 0 Or fCount = 0 Then
                    mErrorComment = "Traceability Data Nothing"
                    Throw New ApplicationException(mErrorComment)
                End If
            Catch ex As Exception
                mErrorComment = cADO.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString
            Finally
                cADO = Nothing
            End Try

            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "データ保存失敗：" & mErrorComment)
                Exit Sub
            End If


            'Write of Test Date To AWS ======================== 
            Send_Table = "YMFK_FAM3_PROD_LOG_TB"
            If InStr(sData(7), "F3SC") = 0 Then BYORD_Start_No = ""

            Try
                ' 1:RECNO      2    3       4          5            6           7         8       9          10               11        12       13
                Send_clonm = "BAR,SERIAL,PROD_NO,SAP_LINKAGE_NO,SAP_COMP_NO,CRP_GR_NO,START_NO,MS_CODE,FINISH_SCHDULE_D,INST_FINISH_D,XJ_NO,PROCESS_CODE"
                Send_clonm = Send_clonm & ",RDATE_S,RTIME_S,RDATE_E,RTIME_E,WORKERID,INDEX_NO,BYORD_START_NO,PC_NAME,SOFT,VER"
                '                             14      15      16      17       18       19        20           21     22   23

                '                    2                  3                  4                  5                  6                  7                  8
                Send_Data = "'" & sData(0) & "','" & sData(1) & "','" & sData(2) & "','" & sData(3) & "','" & sData(4) & "','" & sData(5) & "','" & sData(6)
                Send_Data = Send_Data & "','" & sData(7) & "','" & sData(8) & "','" & sData(9) & "','" & sData(10) & "','" & sData(11) & "','" & RDate_S
                '                                   9                10                  11                 12                   13                14
                '                                  15                16                17                18                 19                    20     
                Send_Data = Send_Data & "','" & RTime_S & "','" & RDate_E & "','" & RTime_E & "','" & WorkerID & "','" & sData(15) & "','" & BYORD_Start_No
                Send_Data = Send_Data & "','" & sData(17) & "','" & Soft_Name & "','" & Soft_Ver & "'"
                '                                  21                  22                  23

                sSQL = Send_Table & "|" & Send_clonm & "|" & Send_Data

                'mPlasmaResult = EthercomNdll.ComHost("10.7.15.71", "10.7.15.71", 2005, "QDBW-" & sSQL)     'Ver 0.43 GYS Del
                ' mPlasmaResult = EthercomNdll.ComHost("10.7.15.25", "10.7.15.140", 2031, "QDBW-" & sSQL)     'Ver 0.43 GYS ADD
                mPlasmaResult = EthercomNdll.ComHost("10.7.15.254", "10.7.15.140", 2031, "QDBW-" & sSQL)     'Ver 0.04 MKS ADD
                Debug.Print(mPlasmaResult)
                Me.SysLogWrite(True, mPlasmaResult)

                Select Case Mid(mPlasmaResult, 1, 8)
                    Case "ERROR:A0"
                        mErrorComment = "生産情報がありません。このバーコードは無効です。"
                    Case "ERROR:A1"
                        mErrorComment = "生産情報がｷｬﾝｾﾙされています。無効の着工札です。"
                    Case "ERROR:A2" : Case "ERROR:A3" : Case "ERROR:A4"
                        mErrorComment = "プログラムエラーです。"
                    Case "ERROR:A5"
                        mErrorComment = "ＯＲＡＣＬＥエラーです。"
                    Case "ERROR:13"
                        mErrorComment = "タイムアウトエラーです。"
                    Case ""
                        mErrorComment = "データがありません。"
                    Case Else

                End Select

                If mErrorComment <> "" Then Throw New ApplicationException(mErrorComment)

            Catch ex As Exception
                Me.SysLogWrite(True, "AWS Data Up Load NG _ Ethercom Err")
            Finally

            End Try

        End If
        '---------------------------------------- Ver 0.40 GYS ADD End --------------------------------------------

    End Sub

    Public Sub WriteTemporaryEndingLog()
        '==============================================================================================================
        ' Function : Temporary file to Log file  
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
        Dim rField As String = ""
        Dim kField As String = ""

        'Read of Temp data ========================================================================
        dbName = mAppPath & "210 Result Data\Temporary.mdb"
        tbName = "tblResult"

        Try
            sSQL = "SELECT Result,ResultComment,InspectionRunTime,StartDateTime FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Write of temporary log ===================================================================
        Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"
        Try
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = lFile

            cTxt.WRITE_TEXT_FILE(False, "")
            cTxt.WRITE_TEXT_FILE(False, "■ Inspection End")
            cTxt.WRITE_TEXT_LINE(False, "*", 50)

            cTxt.WRITE_TEXT_FILE(False, "Result           : " & sData(0))
            cTxt.WRITE_TEXT_FILE(False, "Memo             : " & sData(1))
            cTxt.WRITE_TEXT_FILE(False, "Inspection Time  : " & sData(2))

        Catch ex As Exception
            mErrorComment = cTxt.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cTxt = Nothing

        End Try

        Call LogBrowserNavi()

        Try
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = lFile

            If sData(0) = "PASS" Then
                cTxt.WRITE_TEXT_FILE(False, "")
                cTxt.WRITE_TEXT_FILE(False, "PASS：検査合格=========")
                'cTxt.WRITE_TEXT_FILE(False, "PAS    S	   PASS  PASS")
                'cTxt.WRITE_TEXT_FILE(False, "P  A  S  S   P	    A")
                'cTxt.WRITE_TEXT_FILE(False, "PAS  SPASSP  ASSP	ASSP")
                'cTxt.WRITE_TEXT_FILE(False, "P    A    S     S	   P")
                'cTxt.WRITE_TEXT_FILE(False, "P   A      S SPAS	SPAS")
                cTxt.WRITE_TEXT_FILE(False, "=======================")

                'Ver0.02 hsj del start===
                'cSys = New cls102_SignalTower
                'Call cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
                'cSys = Nothing
                'Ver0.02 hsj del end===

            End If

            If sData(0) = "NG" Then
                cTxt.WRITE_TEXT_FILE(False, "")
                cTxt.WRITE_TEXT_FILE(False, "× 検査不合格 ===================================")
                cTxt.WRITE_TEXT_FILE(False, "　①リーダを呼んで下さい")

                'Ver0.02 hsj del start===
                'cSys = New cls102_SignalTower
                'Call cSys.SignalTower("RED", "OFF", "OFF", "OFF")
                'cSys = Nothing
                'Ver0.02 hsj del end===

            End If

            Call LogBrowserNavi()

        Catch ex As Exception
            mErrorComment = cTxt.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cTxt = Nothing

        End Try

        'Temp to DateTime.txt =====================================================================
        Dim ymName As String = Mid(lblData(6).Text, 1, 7)
        Dim ymdName As String = Format(CDate(lblData(6).Text), "yyyyMMdd HHmmss") & ".txt"
        Try
            cFop = New cls012_File_Operation

            cFop.ORIGINAL_PATH = lFile
            cFop.FOLDER_PATH = mAppPath & "220 Result Log\" & ymName & "\"
            cFop.FILE_NAME = ymdName

            Call cFop.FILE_COPY()

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'CLOSEボタン有効化
        btnDialogResult(0).Enabled = True
        btnDialogResult(0).Focus()

        'Call BarcodeSetup()

    End Sub

    'Private Delegate Sub COM_Recieve()
    ' ''==========================================================================================
    ' ''Function : COM recieve event
    ' ''Remarks  :
    ' ''==========================================================================================
    'Private Sub SerialPort2_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort2.DataReceived

    '    'AddHandler TextBox1.TextChanged, AddressOf TextBox1_TextChanged

    '    Try
    '        Me.Invoke(New COM_Recieve(AddressOf BeginningProcess), New Object() {})
    '    Catch

    '    Finally

    '    End Try

    '    cBcd1.PortClose()
    '    cBcd1 = Nothing

    'End Sub

    'Private Sub BeginningProcess()
    '    '==========================================================================================
    '    'Function : Beginning process of inspection
    '    'Remarks  :
    '    '==========================================================================================
    '    Dim mBarcodeReadData As String = ""

    '    Try
    '        Call cBcd1.FuncRecvCmd()
    '        mBarcodeReadData = cBcd1.RECEIVE_DATA

    '    Catch oExcept As Exception
    '        mErrorComment = cBcd1.ERR_COMMENT : If mErrorComment = "" Then oExcept.ToString()

    '    End Try

    '    If mErrorComment <> "" Then
    '        MessageBox.Show(mErrorComment, "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Else
    '        'MessageBox.Show("Get Barcode Text OK !", "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    End If

    '    frm001_Main.txtBarcode(0).Text = mBarcodeReadData

    'End Sub

    'Private Sub BarcodeSetup()

    '    Dim dbName As String
    '    Dim tbName As String
    '    Dim fName() As String = {""}
    '    Dim sData() As String = {""}
    '    Dim sSQL As String = ""
    '    Dim fCount As Integer = 0
    '    Dim rCount As Integer = 0
    '    Dim rField As String = ""
    '    Dim kField As String = ""

    '    'Read of BarcodeReader Equipment No. ====================================================================
    '    Try
    '        dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
    '        tbName = "tblEquipmentSetting"

    '        cADO = New cls010_MDB_Operation

    '        cADO.DB_PATH = dbName
    '        cADO.TB_NAME = tbName

    '        sSQL = "SELECT " & "BarcodeReaderPort , BarcodeReaderSetting" & " FROM " & tbName
    '        cADO.DB_SQL = sSQL

    '        Call cADO.DB_READ_JET(fName, sData)

    '        fCount = cADO.DB_FIELD_COUNT
    '        rCount = cADO.DB_RECORD_COUNT

    '    Catch ex As Exception
    '        mErrorComment = cADO.ERR_COMMENT
    '        If mErrorComment = "" Then mErrorComment = ex.ToString()

    '    Finally
    '        cADO = Nothing

    '    End Try

    '    If rCount <> 0 Then

    '        If mErrorComment <> "" Then Exit Sub
    '        Dim BarSetting() As String = Split(sData(1), "-")

    '        cBcd1 = New cls101_ComRS232C
    '        cBcd1.TERMINAL_CHAR = vbCr : mDelim = vbCr
    '        cBcd1.RS_COMM = Me.SerialPort2 : cBcd1.PORT = sData(0) : mPort = sData(0)
    '        cBcd1.SETTING = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)
    '        mSetting = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)

    '        cBcd1.PortSetup()

    '    End If

    'End Sub

    Public Sub SaveServer()
        '==============================================================================================================
        ' Function : Data is saved in the server 
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================


    End Sub

    'Process (Event) **************************************************************************************************
    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        '==============================================================================================================
        'Function : Manual Start
        'Remarks  :
        '==============================================================================================================

        If mProcessStatus <> 0 Then mProcessStatus = 1 : Exit Sub

        'Starting Preprocess =================================================================
        'Folder & File ------------------------------------------
        Call MakeFolderFile()
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "フォルダ/ファイル準備：" & mErrorComment)
            Exit Sub
        Else
            Call Me.SysLogWrite(True, "フォルダ/ファイル準備：" & "完了")
        End If

        'Temporary data -----------------------------------------
        Call WriteTemporaryStartingData()
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "一時保存開始データ書込：" & mErrorComment)
            Exit Sub
        Else
            Call Me.SysLogWrite(True, "一時保存開始データ書込：" & "完了")
        End If

        'Temporary log ------------------------------------------
        Call WriteTemporaryStartingLog()
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "一時保存開始ログ書込：" & mErrorComment)
            Exit Sub
        Else
            Call Me.SysLogWrite(True, "一時保存開始ログ書込：" & "完了")
        End If

        Call Me.ModeControl()

    End Sub

    Private Sub btnEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnd.Click
        '==============================================================================================================
        'Function : Inspection end
        'Remarks  :
        '==============================================================================================================

        If mProcessStatus <> 0 Then mProcessStatus = 3
        btnDialogResult(0).Enabled = True

    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        '==============================================================================================================
        'Function : Inspection stop
        'Remarks  :
        '==============================================================================================================

        If mProcessStatus = 1 Then mProcessStatus = 2

    End Sub

    Private Sub btnPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrev.Click
        '==============================================================================================================
        'Function : Prev. process
        'Remarks  :
        '==============================================================================================================

        If mProcessStatus = 2 Or mProcessStatus = 4 Then mProcessStatus = 6

    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        '==============================================================================================================
        'Function : Next process
        'Remarks  :
        '==============================================================================================================

        If mProcessStatus = 2 Or mProcessStatus = 4 Then mProcessStatus = 5

    End Sub

    Private Sub radManualSet_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Select of repeats
        'Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, DirectCast(sender, RadioButton).Text & " ON")

        Call Me.FocusControl()

    End Sub

    Private Sub radRepeatSet_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Select of repeats
        'Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, DirectCast(sender, RadioButton).Text & " ON")

        Call Me.FocusControl()

    End Sub

    Private Sub nudRepeatFreq_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Setting of Repeat Freq.
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

    End Sub

    Private Sub btnTestReTry_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        mProcessStatus = 11
        
    End Sub

    Private Sub DialogResult_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Dialog result
        'Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : " & DirectCast(sender, Button).Text & " ON")

        Call Me.FocusControl()

        Select Case DirectCast(sender, Button).Text

            Case "CLOSE"
                mDialogResult = ""
                Me.Close()

        End Select

    End Sub

    Private Sub wBrowser_DocumentCompleted(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Log read
        'Remarks  :
        '==============================================================================================================

        wBrowser(0).Visible = False
        wBrowser(0).Document.Encoding = "Shift_JIS"
        'wBrowser(0).Refresh()
        'DoEvents()

        Dim ScrollWindow As HtmlWindow
        ScrollWindow = wBrowser(0).Document.Window

        'ScrollWindow.ScrollTo(Screen.PrimaryScreen.WorkingArea.Width, Screen.PrimaryScreen.WorkingArea.Height)
        'ScrollWindow.ScrollTo(Screen.PrimaryScreen.WorkingArea.Width, 1000000)
        'ScrollWindow.ScrollTo(Screen.PrimaryScreen.WorkingArea.Width, wBrowser(0).Document.Body.ScrollRectangle.Height)
        ScrollWindow.ScrollTo(0, wBrowser(0).Document.Body.ScrollRectangle.Height)
        wBrowser(0).Visible = True

        DoEvents()

    End Sub

    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : Button focus
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        sBar(0).Focus()

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
            DoEvents()
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

    Public Sub LogBrowserNavi()
        '==========================================================================================
        ' Function : Browser navigate
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Me.wBrowser(0).Navigate(mAppPath & "220 Result Log\Temporary.txt")

        DoEvents()

    End Sub

    Public Sub WaitStart(ByVal sVal As Double)
        '==========================================================================================
        ' Function : Wait Start
        ' Data     : Setting time
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        tmWaite.SettingValue = sVal
        tmWaite.CountStart()
        tmWaite.ForeColor = Color.White

        Do
            If tmWaite.CountStatus = ToY_StopWatch.ToY_StopWatch.cStatus.CountUp Then Exit Do
            DoEvents()
        Loop

        tmWaite.ForeColor = Color.DarkSlateGray

    End Sub

    Public Sub WaitStop()
        '==========================================================================================
        ' Function : Wait Stop
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        tmWaite.CountStopReset()
        tmWaite.ForeColor = Color.DarkSlateGray

    End Sub

    Public Sub CountStart(ByVal sVal As Double)
        '==========================================================================================
        ' Function : Count Start
        ' Data     : Setting time
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        tmCount.SettingValue = sVal
        tmCount.CountStart()
        tmCount.ForeColor = Color.White

    End Sub

    Public Sub CountStop()
        '==========================================================================================
        ' Function : Count Stop
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        tmCount.CountStopReset()
        tmCount.ForeColor = Color.DarkSlateGray

    End Sub

    Public Sub InspectionTimerStart()
        '==========================================================================================
        ' Function : Inspection timer start
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        tmInspectionTime.CountStopReset()

        tmInspectionTime.CountStart()
        tmInspectionTime.ForeColor = Color.White

    End Sub

    Public Sub InspectionTimerStop()
        '==========================================================================================
        ' Function : Inspection timer Stop
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        tmInspectionTime.CountStop()
        tmInspectionTime.ForeColor = Color.DarkSlateGray

    End Sub

End Class