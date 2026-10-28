Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm101_EquipmentSetting

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

    'Local (Special form control) ========================
    Dim btnDialogResult() As Button
    Dim lblTitle() As Label
    Dim cmbInput() As ComboBox
    Dim txtInput() As TextBox

    'Local Property =======================================
    Dim mDialogResult As String = ""

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm101_EquipmentSetting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '==============================================================================================================
        'Function : Display start
        'Remarks  :
        '==============================================================================================================

        'Equipment Information ====================================================================
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
            .Text = "Equipment Setting"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 500 : .Height = 450
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

    Private Sub frm101_EquipmentSetting_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        '==============================================================================================================
        'Function : After display
        'Remarks  :
        '==============================================================================================================

        Call Me.FocusControl()

        DoEvents()

        Me.Opacity = 1.0

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

    Private Sub frm101_EquipmentSetting_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        Call Me.SysLogWrite(True, "●　装置関連の設定画面 --------------------- (" & Me.Name & "起動)")

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
                    .Left = CInt((Me.Width / 2) - .Width - 20)
                    .Text = "OK"
                End If
                If i = 1 Then
                    .Left = CInt((Me.Width / 2) + 20)
                    .Text = "CANCEL"
                End If
            End With
        Next

        'Title Label ==============================================================================
        Me.lblTitle = New System.Windows.Forms.Label(10) {}
        For i As Integer = 0 To 10
            Me.lblTitle(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblTitle(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 10
            With lblTitle(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = 170 : .Height = 21
                .Left = 20
                .TextAlign = ContentAlignment.MiddleLeft
                If i = 0 Then
                    .Top = sBar(0).Top + sBar(0).Height + 20
                Else
                    .Top = lblTitle(i - 1).Top + lblTitle(i - 1).Height + 5
                End If
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BackColor = Color.FromArgb(50, 50, 60)
                .ForeColor = Color.White
                .BorderStyle = BorderStyle.None
                .BringToFront()
                If i = 0 Then .Text = "自動時刻合わせ"
                If i = 1 Then .Text = "時刻合わせソフトのパス"
                If i = 2 Then .Text = "時刻合わせソフトのタイトル"
                If i = 3 Then .Text = "サインタワーの使用"
                If i = 4 Then .Text = "バーコードリーダのポート"
                If i = 5 Then .Text = "バーコードリーダの設定"
                If i = 6 Then .Text = "サーバのドライブ"
                If i = 7 Then .Text = "サーバのパス"
                If i = 8 Then .Text = "装置ナンバー"
                If i = 9 Then .Text = "QICプリンタ"
                If i = 10 Then .Text = "銘板プリンタ"
            End With
        Next

        'ComboBox =================================================================================
        Me.cmbInput = New System.Windows.Forms.ComboBox(3) {}
        For i As Integer = 0 To 3
            Me.cmbInput(i) = New System.Windows.Forms.ComboBox
            Me.Controls.Add(cmbInput(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 3
            With cmbInput(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Width = 280
                .Left = lblTitle(0).Left + lblTitle(0).Width + 10
                .BackColor = Color.FromArgb(255, 255, 255)
                .ForeColor = Color.Black
                .BringToFront()
                If i = 0 Then .Top = lblTitle(0).Top
                If i = 1 Then .Top = lblTitle(3).Top
                If i = 2 Then .Top = lblTitle(4).Top
                If i = 3 Then .Top = lblTitle(6).Top
            End With
        Next

        'TextBox ==================================================================================
        Me.txtInput = New System.Windows.Forms.TextBox(6) {}
        For i As Integer = 0 To 6
            Me.txtInput(i) = New System.Windows.Forms.TextBox
            Me.Controls.Add(txtInput(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 6
            With txtInput(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Width = 280
                .Left = lblTitle(0).Left + lblTitle(0).Width + 10
                .BackColor = Color.FromArgb(255, 255, 255)
                .ForeColor = Color.Black
                .BringToFront()
                If i = 0 Then .Top = lblTitle(1).Top
                If i = 1 Then .Top = lblTitle(2).Top
                If i = 2 Then .Top = lblTitle(5).Top
                If i = 3 Then .Top = lblTitle(7).Top
                If i = 4 Then .Top = lblTitle(8).Top
                If i = 5 Then .Top = lblTitle(9).Top
                If i = 6 Then .Top = lblTitle(10).Top
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

        'ComboBox menu ============================================================================
        For i As Integer = 0 To 3
            cmbInput(i).Items.Clear()
        Next
        For i As Integer = 0 To 1
            cmbInput(i).Items.Add("Yes")
            cmbInput(i).Items.Add("No")
        Next
        cmbInput(2).Items.Add("COM1")
        cmbInput(2).Items.Add("COM2")
        cmbInput(2).Items.Add("COM3")
        cmbInput(2).Items.Add("COM4")
        cmbInput(2).Items.Add("COM5")
        cmbInput(2).Items.Add("COM6")
        cmbInput(2).Items.Add("USB/KeyBoard")

        For i = 65 To 90
            cmbInput(3).Items.Add(Chr(i))
        Next

        'Read of equipment setting ================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            rField = "DateTimeApp,DateTimeAppPath,DateTimeAppName,SignalTowerControl,BarcodeReaderPort,BarcodeReaderSetting,BackupDriveSign,BackupDrivePath,EquipmentNumber,QICPrinter,LabelPrinter"

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "装置設定：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        cmbInput(0).Text = sData(0)
        txtInput(0).Text = sData(1)
        txtInput(1).Text = sData(2)
        cmbInput(1).Text = sData(3)
        cmbInput(2).Text = sData(4)
        txtInput(2).Text = sData(5)
        cmbInput(3).Text = sData(6)
        txtInput(3).Text = sData(7)
        txtInput(4).Text = sData(8)
        txtInput(5).Text = sData(9)
        txtInput(6).Text = sData(10)

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : 表示データ　読込完了")

    End Sub

    'Process (Event) **************************************************************************************************
    Private Sub DialogResult_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Select of order or barcode
        'Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : " & DirectCast(sender, Button).Text & " ON")

        Call Me.FocusControl()

        Select Case DirectCast(sender, Button).Text

            Case "OK"
                Call Me.SaveData()
                If mErrorComment <> "" Then
                    MessageBox.Show(mErrorComment, "Equipment setting", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    mErrorComment = ""
                    Exit Sub
                End If
                mDialogResult = DirectCast(sender, Button).Text
                Me.Close()

            Case "CANCEL"
                mDialogResult = ""
                Me.Close()

        End Select

    End Sub

    Private Sub SaveData()
        '==============================================================================================================
        ' Function : Data save InputOrder.mdb & StartInformation.mdb
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

        Dim wData() As String = {""}

        'Read of EquipmentSetting.mdb =============================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            rField = "*"

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "装置設定：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Check of change data ===================================================================== 
        ReDim wData(fCount - 1)
        wData(0) = cmbInput(0).Text
        wData(1) = txtInput(0).Text
        wData(2) = txtInput(1).Text
        wData(3) = cmbInput(1).Text
        wData(4) = cmbInput(2).Text
        wData(5) = txtInput(2).Text
        wData(6) = cmbInput(3).Text
        wData(7) = txtInput(3).Text
        wData(8) = txtInput(4).Text
        wData(9) = txtInput(5).Text
        wData(10) = txtInput(6).Text

        Dim dCount As Integer = 0
        For i As Integer = 0 To fCount - 1
            If sData(i) <> wData(i) Then
                dCount += 1
                Call Me.SysLogWrite(True, "装置設定データ:[" & lblTitle(i).Text & "] " & sData(i) & " を " & wData(i) & " に変更")
            End If
        Next
        If dCount = 0 Then Call Me.SysLogWrite(True, "装置設定データ:変更無し")

        'Write of InputOrder.mdb ==================================================================
        Try
            dbName = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
            tbName = "tblEquipmentSetting"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "UPDATE " & tbName & " SET "
            For i As Integer = 0 To fCount - 1
                sSQL &= fName(i) & " = '" & wData(i) & "' , "
            Next
            sSQL = Mid(sSQL, 1, Len(sSQL) - 3)
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Call Me.SysLogWrite(True, "装置設定データ保存 完了")

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