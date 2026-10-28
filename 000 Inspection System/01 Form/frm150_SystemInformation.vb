Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm150_SystemInformation

    'Class ================================================
    Dim cWin As cls001_Window_Control
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cFrm As cls090_DesignForm_Type01

    'Local (Standard form control) ========================
    Dim sBar() As PictureBox
    Dim sBtn() As Button
    Dim sLbl() As Label

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mMousePoint As Point
    Dim mErrorComment As String = ""
    Dim mSystemLog As String = ""
    Dim mEquipmentStatus As Integer = 0

    'Local (Special form control) ========================
    Dim btnDialogResult() As Button
    Dim lblTitle() As Label
    Dim dgvGrid() As DataGridView

    'Local Property =======================================
    Dim mDialogResult As String = ""

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm150_SystemInformation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            .Text = "System information"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 750 : .Height = 310
            .Top = (Screen.PrimaryScreen.Bounds.Height \ 2) - (.Height \ 2)
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

    Private Sub frm150_SystemInformation_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
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

        mDialogResult = ""

        Me.Close()

    End Sub

    Private Sub frm150_SystemInformation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        Call Me.SysLogWrite(True, "●　システム情報 ------------------------- (" & Me.Name & "起動)")

        'Form Draw ================================================================================
        Call Me.FormDraw()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, Me.Name & " Form Draw", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
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
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = 140 : .Height = 40
                .Top = Me.Height - (.Height + Me.sBar(2).Height + 20)
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\BlueForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                If i = 0 Then
                    .Left = CInt((Me.Width / 2) - (.Width) - 20)
                    .Text = "OK"
                    .Visible = False
                End If
                If i = 1 Then
                    .Left = CInt((Me.Width / 2) - (.Width / 2))
                    .Text = "CLOSE"
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
                    .Width = CInt(bWidth * 0.15) : .Left = bLeft
                    .Text = "バージョン"
                End If
                If i = 1 Then
                    .Width = CInt(bWidth * 0.15) : .Left = lblTitle(i - 1).Left + lblTitle(i - 1).Width + 1
                    .Text = "日付"
                End If
                If i = 2 Then
                    .Width = CInt(bWidth * 0.7) : .Left = lblTitle(i - 1).Left + lblTitle(i - 1).Width + 1
                    .Text = "記事"
                End If

            End With
        Next

        'Grid =====================================================================================
        Me.dgvGrid = New System.Windows.Forms.DataGridView(0) {}
        For i As Integer = 0 To 0
            Me.dgvGrid(i) = New System.Windows.Forms.DataGridView
            Me.Controls.Add(dgvGrid(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With dgvGrid(i)
                'Display ------------------------------------------------
                .Columns.Clear()
                .Width = bWidth + 3 : .Height = 122
                .Top = lblTitle(0).Top + lblTitle(1).Height + 1
                .Left = bLeft
                .Dock = DockStyle.None
                .BorderStyle = BorderStyle.FixedSingle

                .CellBorderStyle = DataGridViewCellBorderStyle.Single

                .RowHeadersVisible = False
                .RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

                .ColumnHeadersVisible = False
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                .ColumnHeadersDefaultCellStyle.Font = New Font("MS Gothic", 10, FontStyle.Bold)

                .ScrollBars = ScrollBars.Vertical

                'Edit mode ----------------------------------------------
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = False
                .AllowUserToResizeColumns = False
                .AllowUserToResizeRows = False
                .ReadOnly = True
                .Enabled = True

                'Selection ----------------------------------------------
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .MultiSelect = False

                'Color --------------------------------------------------
                .BackgroundColor = Color.FromArgb(50, 50, 70)

                .RowsDefaultCellStyle.BackColor = Color.FromArgb(50, 50, 70)
                .RowsDefaultCellStyle.ForeColor = Color.White

                .RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(50, 50, 70)
                .RowsDefaultCellStyle.SelectionForeColor = Color.White

                .GridColor = Color.FromArgb(200, 200, 255)

                'Column -------------------------------------------------
                .ColumnCount = 3

                .Columns(0).Name = "バージョン"
                .Columns(1).Name = "日付"
                .Columns(2).Name = "記事"

                .Columns(0).Width = lblTitle(0).Width
                .Columns(1).Width = lblTitle(1).Width
                .Columns(2).Width = lblTitle(2).Width

                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

                'Row ----------------------------------------------------
                .RowCount = 7
                For j As Integer = 0 To .RowCount - 1
                    .Rows(j).Height = 17
                Next

            End With
        Next

    End Sub

    Private Sub FormDataSet()
        '==============================================================================================================
        ' Function : Form data setting
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

        'Read of System setting ===================================================================
        Try
            dbName = mAppPath & "100 System Data\SystemSetting.mdb"
            tbName = "tblSystemSetting"
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "SystemVersion,SystemVerDate,SystemVerComment" & " FROM " & tbName
            sSQL &= " ORDER BY SystemVersion DESC"
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
        If rCount >= 8 Then
            dgvGrid(0).RowCount = rCount
        End If

        For i As Integer = 0 To rCount - 1
            With dgvGrid(0)
                .Rows.Item(i).Cells.Item(0).Value = sData(i * 3)
                .Rows.Item(i).Cells.Item(1).Value = sData(i * 3 + 1)
                .Rows.Item(i).Cells.Item(2).Value = sData(i * 3 + 2)
            End With
        Next

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : 表示データ　読込完了")

    End Sub

    'Process (Event) **************************************************************************************************
    Private Sub DialogResult_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Dialog result
        'Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : " & DirectCast(sender, Button).Text & " ON")

        Call Me.FocusControl()

        Me.Close()

    End Sub

    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : Button focus
        ' Data     : Nothing
        ' Return   : Nothing
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