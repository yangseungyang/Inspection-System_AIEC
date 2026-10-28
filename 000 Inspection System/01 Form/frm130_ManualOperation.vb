Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm130_ManualOperation

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
    Dim cmbInsModel() As ComboBox
    Dim lblTitle() As Label
    Dim btnDetails() As Button

    'Local Property =======================================
    Dim mDialogResult As String = ""

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm130_ManualOperation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            .Text = "Manual Operation"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 350 : .Height = 300
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

    Private Sub frm130_ManualOperation_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
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

        Me.sBar(0).Focus()
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

    Private Sub frm130_ManualOperation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        Call Me.SysLogWrite(True, "●　手動操作 ------------------------ (" & Me.Name & "起動)")

        'Form Draw ================================================================================
        'Process 0
        Call Me.FormDraw()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, Me.Name & " Form Draw", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        'Form data set ============================================================================
        'Process 1
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
                    .Left = CInt((Me.Width / 2) - .Width - 20)
                    .Text = "START"
                End If
                If i = 1 Then
                    .Left = CInt((Me.Width / 2) + 20)
                    .Text = "CLOSE"
                End If
            End With
        Next

        'Title Label ==============================================================================
        Me.lblTitle = New System.Windows.Forms.Label(0) {}
        For i As Integer = 0 To 0
            Me.lblTitle(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblTitle(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With lblTitle(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = 250 : .Height = 16
                .Left = CInt(Me.Width / 2 - .Width / 2)
                .TextAlign = ContentAlignment.MiddleCenter
                .BackColor = Color.FromArgb(50, 50, 60)
                .ForeColor = Color.White
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
                If i = 0 Then
                    .Top = sBar(0).Top + sBar(0).Height + 20
                    .Text = "検査型式"
                End If
            End With
        Next

        'Model ComboBox ===========================================================================
        Me.cmbInsModel = New System.Windows.Forms.ComboBox(0) {}
        For i As Integer = 0 To 0
            Me.cmbInsModel(i) = New System.Windows.Forms.ComboBox
            Me.Controls.Add(cmbInsModel(i))
            AddHandler cmbInsModel(i).SelectedIndexChanged, AddressOf cmbInsModel_SelectedIndexChanged
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With cmbInsModel(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Width = lblTitle(0).Width
                .Left = CInt(Me.Width / 2 - .Width / 2)
                .BackColor = Color.White
                .ForeColor = Color.Black
                .BringToFront()
                .Cursor = Cursors.Hand
                If i = 0 Then .Top = lblTitle(i).Top + lblTitle(i).Height + 2
            End With
        Next

        'Details button ===========================================================================
        Me.btnDetails = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnDetails(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnDetails(i))
            AddHandler btnDetails(i).Click, AddressOf btnDetails_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With btnDetails(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 12, FontStyle.Bold)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = 100 : .Height = 40
                .Top = btnDialogResult(0).Top - .Height - 10
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\BlueForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                If i = 0 Then
                    .Left = CInt((Me.Width / 2) - (.Width / 2))
                    .Text = "Details"
                End If
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

        'Read of MS-CODE file =====================================================================
        Try
            dbName = mAppPath & "500 Model Information\ModelInformation.mdb"
            tbName = "tblModelInformation"
            rField = "Model" : kField = "Product"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            sSQL &= " ORDER BY " & rField
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "検査用型式読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        cmbInsModel(0).Items.Clear()
        For i As Integer = 0 To rCount - 1
            cmbInsModel(0).Items.Add(sData(i))
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

        Select Case DirectCast(sender, Button).Text

            Case "START"
                mDialogResult = DirectCast(sender, Button).Text

                Me.Hide()

                Dim sRet As String = ""
                Dim frm As New frm020_InspectionStartEnd
                Try
                    frm.PLASMA_ORDER = "," & "," & "," & "," & "," & "," & ","
                    frm.RUN_MODE = "MAN"
                    frm.INSPECTION_MODEL = cmbInsModel(0).Text

                    Call frm.InspectionStart()

                    sRet = frm.DIALOG_RESULT

                Finally
                    frm = Nothing

                End Try
                'Me.TopMost = True
                Me.Opacity = 0 : Me.Show()
                DoEvents()
                Me.Opacity = 1.0

            Case "CLOSE"
                mDialogResult = ""
                Me.Close()

        End Select

    End Sub

    Private Sub cmbInsModel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Select of Inspection model
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

    End Sub

    Private Sub btnDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Display of Details Form
        'Remarks  :
        '==============================================================================================================

        Me.Hide()

        Dim frm As New frm131_ManualDetails
        Try
            Call frm.ShowDialog()

        Finally
            frm = Nothing

        End Try

        Me.Show()
        Call Me.FocusControl()

    End Sub

    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : Button focus
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        btnDialogResult(1).Focus()

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