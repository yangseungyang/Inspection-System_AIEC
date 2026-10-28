Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm110_Result

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
    Dim pnlFrame() As Panel
    Dim dgvGrid() As DataGridView
    Dim wBrowser() As WebBrowser
    Dim cmbItem() As ComboBox
    Dim cmbCondition() As ComboBox
    Dim cmbAndOr() As ComboBox
    Dim txtCondition() As TextBox
    Dim btnSearch() As Button
    Dim lblSearchDate() As Label

    Dim btnLabelPrint() As Button

    'Local Property =======================================
    Dim mDialogResult As String = ""

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm110_Result_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            .Text = "Inspection Result"
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

    Private Sub frm110_Result_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
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

    Private Sub frm110_Result_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        Call Me.SysLogWrite(True, "●　検査実績 ------------------------ (" & Me.Name & "起動)")

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
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\BlueForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                .Left = CInt((Me.Width) - .Width - 20)
                .Text = "CLOSE"
            End With
        Next

        Dim bTop As Integer = sBar(0).Top + sBar(0).Height + 10
        Dim bLeft As Integer = 20
        Dim bWidth As Integer = Me.Width - (bLeft * 2)

        'Panel ====================================================================================
        Me.pnlFrame = New System.Windows.Forms.Panel(2) {}
        For i As Integer = 0 To 2
            Me.pnlFrame(i) = New System.Windows.Forms.Panel
            Me.Controls.Add(pnlFrame(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 2
            With pnlFrame(i)
                If i = 0 Then .Top = bTop : .Left = bLeft : .Height = 200 : .Width = bWidth
                If i = 1 Then .Top = pnlFrame(0).Top + pnlFrame(0).Height + 2 : .Left = bLeft : .Height = (sBar(2).Top - 10) - .Top : .Width = 450
                If i = 2 Then .Top = pnlFrame(1).Top : .Left = pnlFrame(1).Left + pnlFrame(1).Width + 2 : .Height = (sBar(2).Top - 200) - .Top : .Width = (bLeft + bWidth) - .Left

                .BackColor = Color.Black
                If i = 2 Then
                    .BackColor = Color.SteelBlue
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\BlueForm\BackColor.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If

                .BringToFront()
            End With
        Next

        'Grid =====================================================================================
        Me.dgvGrid = New System.Windows.Forms.DataGridView(0) {}
        For i As Integer = 0 To 0
            Me.dgvGrid(i) = New System.Windows.Forms.DataGridView
            pnlFrame(0).Controls.Add(dgvGrid(i))
            AddHandler dgvGrid(i).SelectionChanged, AddressOf dgvGrid_SelectionChanged
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With dgvGrid(i)

                'Display ------------------------------------------------
                .Columns.Clear()
                .Width = pnlFrame(0).Width - 4 : .Height = pnlFrame(0).Height - 4
                .Top = 2
                .Left = 2
                .Dock = DockStyle.None
                .BorderStyle = BorderStyle.FixedSingle

                .CellBorderStyle = DataGridViewCellBorderStyle.Single

                .RowHeadersVisible = False
                .RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

                .ColumnHeadersVisible = True
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                .ColumnHeadersDefaultCellStyle.Font = New Font("MS Gothic", 8, FontStyle.Bold)

                .ScrollBars = ScrollBars.Both

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

                .RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(30, 70, 30)
                .RowsDefaultCellStyle.SelectionForeColor = Color.White

                .GridColor = Color.FromArgb(200, 200, 255)

            End With
        Next

        'WebBrowser ===============================================================================
        Me.wBrowser = New System.Windows.Forms.WebBrowser(0) {}
        For i As Integer = 0 To 0
            Me.wBrowser(i) = New System.Windows.Forms.WebBrowser
            pnlFrame(1).Controls.Add(wBrowser(i))
            AddHandler wBrowser(i).DocumentCompleted, AddressOf wBrowser_DocumentCompleted
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With wBrowser(i)
                .Width = pnlFrame(1).Width - 4 : .Height = pnlFrame(1).Height - 4
                .Top = 2
                .Left = 2
                .BringToFront()
            End With
        Next

        'cmbItem ==================================================================================
        Me.cmbItem = New System.Windows.Forms.ComboBox(3) {}
        For i As Integer = 0 To 3
            Me.cmbItem(i) = New System.Windows.Forms.ComboBox
            pnlFrame(2).Controls.Add(cmbItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 3
            With cmbItem(i)
                .Width = 120 : .Left = 20 : .Top = (.Height + .Height + 10) * i + (i + 30)
                .BringToFront()
            End With
        Next

        'cmbCondition =============================================================================
        Me.cmbCondition = New System.Windows.Forms.ComboBox(3) {}
        For i As Integer = 0 To 3
            Me.cmbCondition(i) = New System.Windows.Forms.ComboBox
            pnlFrame(2).Controls.Add(cmbCondition(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 3
            With cmbCondition(i)
                .Width = 80 : .Left = cmbItem(i).Left + cmbItem(i).Width + 5 : .Top = cmbItem(i).Top
                .BringToFront()
            End With
        Next

        'txtCondition =============================================================================
        Me.txtCondition = New System.Windows.Forms.TextBox(3) {}
        For i As Integer = 0 To 3
            Me.txtCondition(i) = New System.Windows.Forms.TextBox
            pnlFrame(2).Controls.Add(txtCondition(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 3
            With txtCondition(i)
                .Width = 120 : .Left = cmbCondition(i).Left + cmbCondition(i).Width + 5 : .Top = cmbItem(i).Top
                .BringToFront()
            End With
        Next

        'cmbAndOr =================================================================================
        Me.cmbAndOr = New System.Windows.Forms.ComboBox(2) {}
        For i As Integer = 0 To 2
            Me.cmbAndOr(i) = New System.Windows.Forms.ComboBox
            pnlFrame(2).Controls.Add(cmbAndOr(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 2
            With cmbAndOr(i)
                .Width = 80 : .Left = cmbCondition(i).Left : .Top = cmbCondition(i).Top + cmbCondition(i).Height + 5
                .BringToFront()
            End With
        Next

        'Search button ============================================================================
        Me.btnSearch = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnSearch(i) = New System.Windows.Forms.Button
            pnlFrame(2).Controls.Add(btnSearch(i))
            AddHandler btnSearch(i).Click, AddressOf Search_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With btnSearch(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Bold)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = 60 : .Height = 50
                .Top = pnlFrame(2).Height - (.Height + 20)
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\BlueForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                .Left = CInt((pnlFrame(2).Width) - .Width - 20)
                .Text = "Search"
            End With
        Next

        Me.lblSearchDate = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblSearchDate(i) = New System.Windows.Forms.Label
            pnlFrame(2).Controls.Add(lblSearchDate(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblSearchDate(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Top = 10
                .Width = 120 : .Height = 16
                .TextAlign = ContentAlignment.MiddleCenter
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()

                If i = 0 Then
                    .Left = cmbItem(0).Left
                Else
                    .Left = txtCondition(0).Left + txtCondition(0).Width - .Width
                End If
            End With
        Next

        '主銘板印字ボタン
        Me.btnLabelPrint = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnLabelPrint(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnLabelPrint(i))
            AddHandler btnLabelPrint(i).Click, AddressOf LabelPrint_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With btnLabelPrint(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Bold)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = 60 : .Height = 50
                .Top = Me.Height - (.Height + Me.sBar(2).Height + 20)
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\BlueForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                .Left = CInt((pnlFrame(2).Width) - .Width - 20) + .Width
                '.Left = 930
                .Text = "Label" & vbCrLf & "Print"
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

        Call Me.SysLogWrite(False, "")

        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""

        Dim oFile As String = ""
        Dim nFile As String = ""

        'OldDate - NewDate ======================================================================== 
        Call DirDateSearch(oFile, nFile)
        lblSearchDate(0).Text = "最古：" & oFile
        lblSearchDate(1).Text = "最新：" & nFile

        'Last date search =========================================================================
        cmbItem(0).Text = "StartDateTime"
        cmbCondition(0).Text = "="
        txtCondition(0).Text = nFile & "%"
        Call SearchProcess()
        If mErrorComment <> "" Then
            MessageBox.Show("画面が表示できません。", "Data Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        cmbItem(0).Text = ""
        cmbCondition(0).Text = ""
        txtCondition(0).Text = ""

        'Data read  ===============================================================================
        dbName = mAppPath & "\210 Result Data\SearchResult.mdb"
        tbName = "tblResult"

        Try
            rField = "*" : kField = "StartDateTime"

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE "
            sSQL = sSQL & kField & " LIKE '%' " & "ORDER BY " & kField & " DESC;"

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
            MessageBox.Show("該当データがありません。", "Data Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        'Display =================================================================================
        Call GridSetup(fCount, rCount, fName, sData)

        Call LogDisp(dgvGrid(0).Rows.Item(0).Cells(0).Value.ToString)

        For i = 0 To 3
            cmbItem(i).Text = ""
            cmbItem(i).Items.Add("")

            For j = 0 To fCount - 1
                cmbItem(i).Items.Add(fName(j))
            Next
        Next

        For i = 0 To 3
            cmbCondition(i).Text = "="
            cmbCondition(i).Items.Add("=")
            cmbCondition(i).Items.Add("<>")
            cmbCondition(i).Items.Add("<=")
            cmbCondition(i).Items.Add(">=")
            cmbCondition(i).Items.Add("<")
            cmbCondition(i).Items.Add(">")
        Next

        For i = 0 To 2
            cmbAndOr(i).Text = "AND"
            cmbAndOr(i).Items.Add("AND")
            cmbAndOr(i).Items.Add("OR")
        Next

        Call Me.SysLogWrite(True, Me.Name & " : 表示データ　読込完了")

    End Sub

    Private Sub DirDateSearch(ByRef OldDate As String, ByRef NewDate As String)
        '==========================================================================================
        ' Function : File date serch. most old and most new in ResultDB
        ' Data     : OldDate , NewDate(Return)
        ' Return   : Nothing 
        ' Remarks  :
        '==========================================================================================

        Dim NewDir As String = ""
        Dim NewFile As String = ""
        Dim OldDir As String = ""
        Dim OldFile As String = ""
        Dim SearchDir As String = ""
        Dim SearchFile As String = ""

        'search folder ---------------------------------------
        OldDir = "" : NewDir = ""
        SearchDir = Dir(mAppPath & "\210 Result Data\", FileAttribute.Directory)

        Do
            If InStr(SearchDir, "-") <> 0 Then

                If OldDir = "" Then OldDir = SearchDir
                If NewDir = "" Then NewDir = SearchDir

                If OldDir > SearchDir Then OldDir = SearchDir
                If NewDir < SearchDir Then NewDir = SearchDir

            End If

            SearchDir = Dir()
            If SearchDir = "" Then Exit Do

        Loop

        'Old File --------------------------------------------
        OldFile = ""
        SearchFile = Dir(mAppPath & "\210 Result Data\" & OldDir & "\*")
        Do
            If InStr(SearchFile, "-") <> 0 Then

                If OldFile = "" Then OldFile = SearchFile
                If OldFile > SearchFile Then OldFile = SearchFile

            End If

            SearchFile = Dir()
            If SearchFile = "" Then Exit Do

        Loop
        OldDate = Mid(OldFile, 1, 4) & "-" & Mid(OldFile, 6, 2) & "-" & Mid(OldFile, 9, 2)

        'New File --------------------------------------------
        NewFile = ""
        SearchFile = Dir(mAppPath & "\210 Result Data\" & NewDir & "\*")
        Do
            If InStr(SearchFile, "-") <> 0 Then

                If NewFile = "" Then NewFile = SearchFile
                If NewFile < SearchFile Then NewFile = SearchFile

            End If

            SearchFile = Dir()
            If SearchFile = "" Then Exit Do

        Loop
        NewDate = Mid(NewFile, 1, 4) & "-" & Mid(NewFile, 6, 2) & "-" & Mid(NewFile, 9, 2)

    End Sub

    Private Sub SerachresultDisp()

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""

        'Data read  ===============================================================================
        dbName = mAppPath & "\210 Result Data\SearchResult.mdb"
        tbName = "tblResult"

        Try
            rField = "*" : kField = "StartDateTime"

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE "
            sSQL = sSQL & kField & " LIKE '%' " & "ORDER BY " & kField & " DESC;"

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
            MessageBox.Show("該当データがありません。", "Data Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        'Display =================================================================================
        Call GridSetup(fCount, rCount, fName, sData)

        Call LogDisp(dgvGrid(0).Rows.Item(0).Cells(0).Value.ToString)

        Call Me.FocusControl()

    End Sub

    Private Sub SearchProcess()
        '==========================================================================================
        ' Function : Serach process
        ' Data     : Nothing
        ' Return   : true , false
        ' Remarks  :
        '==========================================================================================

        Dim i As Integer = 0

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""

        Dim SearchItem(3) As String
        Dim SearchCond(3) As String
        Dim SearchWord(3) As String
        Dim SearchFileName As String = ""
        Dim MasterName As String = ""
        Dim sDir As String = ""
        Dim sDirs As String = ""
        Dim sDirName() As String = {""}
        Dim sDirCount As Integer = 0

        'Input condition ==========================================================================
        For i = 0 To 3
            SearchItem(i) = cmbItem(i).Text
            SearchCond(i) = cmbCondition(i).Text
            If SearchCond(i) = "=" Then SearchCond(i) = "LIKE"
            SearchWord(i) = txtCondition(i).Text
        Next

        'Search ===================================================================================

        'old SearchResult Kill --------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.FILE_PATH = mAppPath & "\210 Result Data\SearchResult.mdb"

            Call cFop.FILE_DELETE()

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Copy of ResultMaster to New SearchResult file --------------
        Try
            cFop = New cls012_File_Operation

            cFop.ORIGINAL_PATH = mAppPath & "\210 Result Data\ResultMaster.mdb"
            cFop.FOLDER_PATH = mAppPath & "\210 Result Data" & "\"
            cFop.FILE_NAME = "SearchResult.mdb"

            Call cFop.FILE_COPY()

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        dbName = mAppPath & "\210 Result Data\SearchResult.mdb"
        tbName = "tblResult"

        'Table clear ------------------------------------------------
        Try
            sSQL = "DELETE * FROM " & "tblResult"

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

        'Search SQL -------------------------------------------------
        tbName = "tblResult"
        kField = "StartDateTime"

        sSQL = "SELECT " & "*" & " FROM " & tbName & " "
        If SearchWord(0) = "" And SearchWord(1) = "" And SearchWord(2) = "" And SearchWord(3) = "" Then
            sSQL = sSQL & "WHERE " & kField & " LIKE '*' "
        Else
            sSQL = sSQL & "WHERE "

            If SearchWord(0) <> "" Then
                sSQL = sSQL & SearchItem(0) & " " & SearchCond(0) & " '" & SearchWord(0) & "' "
            End If
            If SearchWord(1) <> "" Then
                sSQL = sSQL & cmbAndOr(0).Text & " " & SearchItem(1) & " " & SearchCond(1) & " '" & SearchWord(1) & "' "
            End If
            If SearchWord(2) <> "" Then
                sSQL = sSQL & cmbAndOr(1).Text & " " & SearchItem(2) & " " & SearchCond(2) & " '" & SearchWord(2) & "' "
            End If
            If SearchWord(3) <> "" Then
                sSQL = sSQL & cmbAndOr(2).Text & " " & SearchItem(3) & " " & SearchCond(3) & " '" & SearchWord(3) & "' "
            End If
        End If
        sSQL = sSQL & "ORDER BY " & kField & " DESC;"

        'Dir name and Count -----------------------------------------
        sDir = Dir(mAppPath & "\210 Result Data\MonthData\", FileAttribute.Directory)
        Do
            If InStr(sDir, "-") <> 0 Then
                sDirs = sDirs & sDir & ","
            End If

            sDir = Dir()

            If sDir = "" Then
                sDirs = Mid(sDirs, 1, Len(sDirs) - 1)
                Exit Do
            End If
        Loop
        sDirName = Split(sDirs, ",")
        sDirCount = UBound(sDirName)

        For i = 0 To sDirCount
            Call SearchResult(mAppPath & "\210 Result Data\MonthData\" & sDirName(i), sSQL)
            If mErrorComment <> "" Then
                MessageBox.Show("検索エラー発生", "Data Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Next

    End Sub

    Private Sub SearchResult(ByRef sDir As String, ByRef seSQL As String)
        '==========================================================================================
        ' Function : Serach data
        ' Data     : Search Dir path , tableName
        ' Return   : true , false
        ' Remarks  :
        '==========================================================================================

        Dim j, i, k As Integer
        Dim dbName As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim fCount As Integer
        Dim rCount As Integer
        'Dim SearchFile As String

        'SearchFile = Dir(sDir & "\*")

        'Do
        'If SearchFile = "" Then Exit Do

        dbName = sDir           '& ".mdb"          ' "\" & SearchFile

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = "tblResult"
            cADO.DB_SQL = seSQL
            Debug.Print(dbName)

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

        If rCount <> 0 Then
            'Result write to SearchFile ------
            dbName = mAppPath & "\210 Result Data\SearchResult.mdb"

            k = 0
            For j = 0 To rCount - 1

                sSQL = "INSERT INTO " & "tblResult" & " ("
                For i = 0 To fCount - 1
                    If i = fCount - 1 Then
                        sSQL = sSQL & fName(i)
                    Else
                        sSQL = sSQL & fName(i) & " , "
                    End If
                Next

                sSQL = sSQL & ") VALUES ("

                For i = 0 To fCount - 1
                    If i = fCount - 1 Then
                        sSQL = sSQL & "'" & sData(k) & "'"
                    Else
                        sSQL = sSQL & "'" & sData(k) & "' , "
                    End If
                    k = k + 1
                Next
                sSQL = sSQL & ");"

                'Table clear ------------------------------------------------
                Try
                    cADO = New cls010_MDB_Operation

                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = "tblResult"
                    cADO.DB_SQL = sSQL

                    Call cADO.DB_EXEC_JET()

                Catch ex As Exception
                    mErrorComment = cADO.ERR_COMMENT
                    If mErrorComment = "" Then mErrorComment = ex.ToString

                Finally
                    cADO = Nothing

                End Try

            Next

        End If

        '        SearchFile = Dir()

        'Loop

    End Sub

    Private Sub GridSetup(ByRef fCount As Integer, ByRef rCount As Integer, ByRef fName() As String, ByRef sData() As String)
        '==========================================================================================
        ' Function : MSFlexGrid Setup
        ' Data     : fcount,rCount,fName(),sData()
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = 0

        With dgvGrid(0)
            .Columns.Clear()

            .RowCount = rCount
            .ColumnCount = fCount

            'FieldName ---------------------
            For i = 0 To fCount - 1
                .Columns(i).Name = fName(i)
            Next

            'Data =======================================================
            For i = 0 To rCount - 1
                For j = 0 To fCount - 1
                    .Rows.Item(i).Cells.Item(j).Value = sData(k)
                    k += 1
                Next
            Next

            '.Columns(0).Width = 120
            '.Columns(3).Width = 360

            .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        End With

    End Sub

    Private Sub LogDisp(ByRef StartDateTime As String)
        '==========================================================================================
        ' Function : Serach Log and display
        ' Data     : date time
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim FindDate As String = ""
        Dim FindTime As String = ""
        Dim DirPath As String = ""
        Dim FileName As String = ""

        FindDate = Format(CDate(StartDateTime), "yyyyMMdd")
        FindTime = Format(CDate(StartDateTime), "HHmmss")

        DirPath = mAppPath & "220 Result Log\" & Mid(FindDate, 1, 4) & "-" & Mid(FindDate, 5, 2)
        FileName = Dir(DirPath & "\*" & FindDate & " " & FindTime & ".txt")

        If FileName = "" Then
            Me.wBrowser(0).Navigate(mAppPath & "220 Result Log\Dummy.txt")
        Else
            Me.wBrowser(0).Navigate(DirPath & "\" & FileName)
        End If

        DoEvents()

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

        Select Case DirectCast(sender, Button).Text

            Case "OK"
                mDialogResult = DirectCast(sender, Button).Text
                Me.Close()

            Case "CLOSE"
                mDialogResult = ""
                Me.Close()

        End Select

    End Sub

    Private Sub Search_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Data search
        'Remarks  :
        '==============================================================================================================
        Dim waitfrm As New Form

        With waitfrm
            Try
                .FormBorderStyle = Windows.Forms.FormBorderStyle.None
                .Text = "Please Wait..."
                .Width = 187 : .Height = 100
                .Top = (Screen.PrimaryScreen.Bounds.Height \ 2) - (.Height \ 2)
                .Left = (Screen.PrimaryScreen.Bounds.Width \ 2) - (.Width \ 2)
                .StartPosition = FormStartPosition.CenterScreen
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\Wait.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch

                Call .Show()
                DoEvents()

                Call SearchProcess()
                If mErrorComment <> "" Then
                    MessageBox.Show("データが更新できません。", "Data Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Me.Close()
                End If

                Call SerachresultDisp()

                .Close()

            Finally
                waitfrm = Nothing

            End Try

        End With
        
        'Call SearchProcess()
        'If mErrorComment <> "" Then
        '    MessageBox.Show("データが更新できません。", "Data Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Me.Close()
        'End If

        'Call SerachresultDisp()

    End Sub

    Private Sub LabelPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim cPrn As cls2P0_Inspection
        Dim dbName As String
        Dim rDTkey As String

        Dim rowNo As Integer = dgvGrid(0).CurrentRow.Index
        Dim MODEL As String

        dbName = mAppPath & "\210 Result Data\SearchResult.mdb"
        rDTkey = dgvGrid(0).Rows.Item(rowNo).Cells(0).Value.ToString()
        MODEL = dgvGrid(0).Rows.Item(rowNo).Cells(4).Value.ToString()

        Try
            cPrn = New cls2P0_Inspection
            cPrn.InspectionMODEL = MODEL
            Call cPrn.LabelPrint(dbName, rDTkey)
        Catch ex As Exception
        Finally
            cPrn = Nothing
        End Try

    End Sub

    Private Sub dgvGrid_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==========================================================================================
        'Function : Select change
        'Remarks  :
        '==========================================================================================

        Try
            Dim rowNo As Integer = dgvGrid(0).CurrentRow.Index

            If dgvGrid(0).Rows.Item(rowNo).Cells(1).Value.ToString = "AUT" And _
            dgvGrid(0).Rows.Item(rowNo).Cells(2).Value.ToString = "PASS" Then
                btnLabelPrint(0).Enabled = True
            Else
                btnLabelPrint(0).Enabled = False
            End If

            Call LogDisp(dgvGrid(0).Rows.Item(rowNo).Cells(0).Value.ToString)

        Catch ex As Exception

        End Try

        Call Me.FocusControl()

    End Sub

    Private Sub wBrowser_DocumentCompleted(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Log read
        'Remarks  :
        '==============================================================================================================

        wBrowser(0).Document.Encoding = "Shift_JIS"
        wBrowser(0).Refresh()

    End Sub

    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : Button focus
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        dgvGrid(0).Focus()

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