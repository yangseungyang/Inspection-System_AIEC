Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm010_OrderInput

    'Class ================================================
    Dim cADO As cls010_MDB_Operation
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
    Dim mServerDrive As String = ""
    Dim mEquipmentStatus As Integer = 0

    'Local (Special form control) ========================
    Dim lblTitle() As Label
    Dim lblItem() As Label
    Dim btnDialogResult() As Button
    Dim cmbInput() As ComboBox

    'Local Property =======================================
    Dim mDialogResult As String = ""

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm010_OrderInput_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            .Text = "frm010 Order Input"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 400 : .Height = 450
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

    Private Sub frm010_OrderInput_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
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

    Private Sub frm010_OrderInput_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        Call Me.SysLogWrite(True, "●　オーダー入力 ------------------------ (" & Me.Name & "起動)")

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
        Me.lblTitle = New System.Windows.Forms.Label(1) {}
        For i As Integer = 0 To 1
            Me.lblTitle(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblTitle(i))
            AddHandler lblTitle(i).Click, AddressOf Title_Click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 1
            With lblTitle(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .Width = 280 : .Height = 16
                .Left = CInt(Me.Width / 2 - .Width / 2)
                .TextAlign = ContentAlignment.MiddleCenter
                If i = 0 Then
                    .Top = sBar(0).Top + sBar(0).Height + 20
                Else
                    .Top = lblTitle(0).Top + lblTitle(0).Height + 200
                End If
                .BackColor = Color.FromArgb(50, 50, 60)
                .ForeColor = Color.White
                .BorderStyle = BorderStyle.None
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
                .Cursor = Cursors.Hand
                If i = 0 Then
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\BlueLabel.bmp")
                    .Text = "オーダー入力"
                End If
                If i = 1 Then
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
                    .Text = "バーコード入力"
                End If
            End With
        Next

        'Item Label ===============================================================================
        Me.lblItem = New System.Windows.Forms.Label(8) {}
        For i As Integer = 0 To 8
            Me.lblItem(i) = New System.Windows.Forms.Label
            Me.Controls.Add(lblItem(i))
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 7
            With lblItem(i)
                .AutoSize = False
                .Font = New Font("MS Gothic", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleLeft
                .Width = 80 : .Height = 21
                .Left = lblTitle(0).Left
                If i = 0 Then
                    .Top = lblTitle(0).Top + lblTitle(0).Height + 2
                Else
                    .Top = lblItem(i - 1).Top + lblItem(i - 1).Height + 2
                End If
                .BackColor = Color.FromArgb(50, 50, 60)
                .ForeColor = Color.White
                .BorderStyle = BorderStyle.None
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .BringToFront()
                If i = 0 Then .Text = "MS-CODE"
                If i = 1 Then .Text = "XJ-No."
                If i = 2 Then .Text = "Serial No."
                If i = 3 Then .Text = "Order No."
                If i = 4 Then .Text = "Item No."
                If i = 5 Then .Text = "Comp No."
                If i = 6 Then .Text = "QIC"
                If i = 7 Then .Text = "Linkage"
            End With
        Next

        With lblItem(8)
            .AutoSize = False
            .Font = New Font("MS Gothic", 8, FontStyle.Regular)
            .TextAlign = ContentAlignment.MiddleLeft
            .Width = 80 : .Height = 21
            .Top = lblTitle(1).Top + lblTitle(1).Height + 2
            .Left = lblItem(0).Left
            .BackColor = Color.FromArgb(50, 50, 60)
            .ForeColor = Color.White
            .BorderStyle = BorderStyle.None
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\WhiteLabel.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch
            .BringToFront()
            .Text = "バーコード"
        End With

        'Item Combo ===============================================================================
        Me.cmbInput = New System.Windows.Forms.ComboBox(8) {}
        For i As Integer = 0 To 8
            Me.cmbInput(i) = New System.Windows.Forms.ComboBox
            Me.Controls.Add(cmbInput(i))
            AddHandler cmbInput(i).KeyPress, AddressOf cmbInput_KeyPress
            AddHandler cmbInput(i).SelectedIndexChanged, AddressOf cmbInput_IndexChanged
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 8
            With cmbInput(i)
                .Name = "cmbInput(" & CStr(i) & ")"
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Width = lblTitle(0).Width - lblItem(i).Width - 2 : .Height = 16
                .Left = lblItem(0).Left + lblItem(i).Width + 2
                .Top = lblItem(i).Top
                .BackColor = Color.FromArgb(236, 233, 216)
                .ForeColor = Color.FromArgb(30, 30, 30)
                .BringToFront()
                If i = 8 Then cmbInput(i).Enabled = False
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

        For i As Integer = 0 To 7
            cmbInput(i).Items.Clear()
        Next
        'Read of last file ========================================================================
        Try
            dbName = mAppPath & "200 Start Information\InputOrder.mdb"
            tbName = "tblInputOrder"
            rField = "MSCODE,XJNo,SerialNo,OrderNo,ItemNo,CompNo,Barcode"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName

            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "前回入力オーダー読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        cmbInput(0).Text = sData(0)
        For i As Integer = 1 To 5
            cmbInput(i).Items.Add(sData(i))
        Next

        cmbInput(6).Items.Add("Yes")
        cmbInput(6).Items.Add("No")

        cmbInput(7).Items.Add(sData(6))

        'Read of MS-CODE file =====================================================================
        Try
            dbName = mAppPath & "500 Model Information\ModelInformation.mdb"
            tbName = "tblModelInformation"
            rField = "Model" : kField = "Product"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            sSQL &= " WHERE " & kField & " LIKE Yes"
            sSQL &= " ORDER BY " & rField
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "前回入力オーダー読込：" & mErrorComment)

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        cmbInput(0).Items.Clear()
        For i As Integer = 0 To rCount - 1
            cmbInput(0).Items.Add(sData(i))
        Next

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : 表示データ　読込完了")

    End Sub

    'Process (Event) **************************************************************************************************
    Private Sub Title_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Select of order or barcode
        'Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, Me.Name & " : " & DirectCast(sender, Label).Text & " ON")

        Dim ctext As String = DirectCast(sender, Label).Text
        If ctext = "オーダー入力" Then
            lblTitle(1).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
            lblTitle(0).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\BlueLabel.bmp")
            For i As Integer = 0 To 7
                lblItem(i).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
                cmbInput(i).Enabled = True
            Next
            lblItem(8).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\WhiteLabel.bmp")
            cmbInput(8).Enabled = False
        Else
            lblTitle(0).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
            lblTitle(1).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\BlueLabel.bmp")
            For i As Integer = 0 To 7
                lblItem(i).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\WhiteLabel.bmp")
                cmbInput(i).Enabled = False
            Next
            lblItem(8).BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\NavyBlueLabel.bmp")
            cmbInput(8).Enabled = True
        End If

        Call Me.FocusControl()

    End Sub

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
                mDialogResult = DirectCast(sender, Button).Text

                Call Me.InputCheck()
                If mErrorComment <> "" Then
                    MessageBox.Show(mErrorComment, "Order input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    mErrorComment = ""
                    Exit Sub
                End If

                Call Me.SaveData()
                If mErrorComment <> "" Then
                    MessageBox.Show(mErrorComment, "Order input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    mErrorComment = ""
                    Exit Sub
                End If

                Me.Hide()

                Dim sRet As String = ""
                Dim frm As New frm020_InspectionStartEnd
                Try
                    If cmbInput(8).Enabled = False Then
                        frm.PLASMA_ORDER = cmbInput(0).Text & "," & cmbInput(1).Text & "," & cmbInput(2).Text & "," & _
                                           cmbInput(3).Text & "," & cmbInput(4).Text & "," & cmbInput(5).Text & "," & cmbInput(6).Text & "," & _
                                           cmbInput(7).Text
                        frm.PLASMA_BARCODE = ""
                    Else
                        frm.PLASMA_ORDER = ""
                        frm.PLASMA_BARCODE = cmbInput(8).Text
                    End If

                    frm.RUN_MODE = "AUT"
                    frm.INSPECTION_MODEL = ""

                    Call frm.InspectionStart()

                    sRet = frm.DIALOG_RESULT

                Finally
                    frm = Nothing

                End Try
                Me.Close()

            Case "CANCEL"
                mDialogResult = ""
                Me.Close()

        End Select

    End Sub

    Private Sub InputCheck()
        '==============================================================================================================
        ' Function : Data save InputOrder.mdb & StartInformation.mdb
        ' Remarks  :
        '==============================================================================================================

        If cmbInput(0).Enabled = True Then

            If cmbInput(0).Text = "" And cmbInput(1).Text = "" Then
                mErrorComment = "Input MS-CODE or XJ No."
                cmbInput(0).Focus()
                Exit Sub
            End If

            If Len(cmbInput(2).Text) <> 9 Then
                mErrorComment = "Serial No. is 9 digit"
                cmbInput(2).Focus()
                Exit Sub
            End If

            If Len(cmbInput(3).Text) <> 10 Then
                mErrorComment = "Order No. is 10 digit"
                cmbInput(3).Focus()
                Exit Sub
            End If

            If Len(cmbInput(4).Text) <> 6 Then
                mErrorComment = "Item No. is 6 digit"
                cmbInput(0).Focus()
                Exit Sub
            End If

            If Len(cmbInput(5).Text) <> 4 Then
                mErrorComment = "Comp No. is 4 digit"
                cmbInput(0).Focus()
                Exit Sub
            End If

            If cmbInput(6).Text <> "Yes" And cmbInput(6).Text <> "No" Then
                mErrorComment = "QIC Select Yes or No"
                cmbInput(0).Focus()
                Exit Sub
            End If
        Else
            If Len(cmbInput(8).Text) <> 6 Then
                mErrorComment = "Barcode No. is 6 digit"
                cmbInput(0).Focus()
                Exit Sub
            End If
        End If

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

        'Write of InputOrder.mdb ==================================================================
        Try
            dbName = mAppPath & "200 Start Information\InputOrder.mdb"
            tbName = "tblInputOrder"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            If cmbInput(0).Enabled = True Then
                sSQL = "UPDATE " & tbName & " SET "
                sSQL &= "MSCODE = '" & cmbInput(0).Text & "' , "
                sSQL &= "XJNo = '" & cmbInput(1).Text & "' , "
                sSQL &= "SerialNo = '" & cmbInput(2).Text & "' , "
                sSQL &= "OrderNo = '" & cmbInput(3).Text & "' , "
                sSQL &= "ItemNo = '" & cmbInput(4).Text & "' , "
                sSQL &= "CompNo = '" & cmbInput(5).Text & "'"
            Else
                sSQL = "UPDATE " & tbName & " SET "
                sSQL &= "Barcode = '" & cmbInput(8).Text & "'"
            End If

            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        Call Me.SysLogWrite(True, Me.Name & " データ保存 完了")

    End Sub

    Private Sub cmbInput_IndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : Select change of Temperature or Humidity 
        ' Remarks  :
        '==============================================================================================================

        Select Case DirectCast(sender, ComboBox).Name
            Case "cmbInput(0)" : cmbInput(1).Focus()
            Case "cmbInput(1)" : cmbInput(2).Focus()
            Case "cmbInput(2)" : cmbInput(3).Focus()
            Case "cmbInput(3)" : cmbInput(4).Focus()
            Case "cmbInput(4)" : cmbInput(5).Focus()
            Case "cmbInput(5)" : cmbInput(6).Focus()
            Case "cmbInput(6)" : cmbInput(0).Focus()
            Case "cmbInput(7)" : cmbInput(7).Focus()
            Case "cmbInput(8)" : cmbInput(8).Focus()
        End Select

    End Sub

    Private Sub cmbInput_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        '==============================================================================================================
        ' Function : Input of order information
        ' Remarks  :
        '==============================================================================================================

        Dim kASC As Integer = Asc(e.KeyChar)

        If kASC = 13 Then
            Select Case DirectCast(sender, ComboBox).Name
                Case "cmbInput(0)" : cmbInput(1).Focus()
                Case "cmbInput(1)" : cmbInput(2).Focus()
                Case "cmbInput(2)" : cmbInput(3).Focus()
                Case "cmbInput(3)" : cmbInput(4).Focus()
                Case "cmbInput(4)" : cmbInput(5).Focus()
                Case "cmbInput(5)" : cmbInput(6).Focus()
                Case "cmbInput(6)" : cmbInput(0).Focus()
                Case "cmbInput(7)" : cmbInput(7).Focus()
                Case "cmbInput(8)" : cmbInput(8).Focus()
            End Select

        Else
            Dim sChr As String = e.KeyChar
            Dim uChr As String = sChr.ToUpper()

            e.KeyChar = CChar(uChr)
        End If

    End Sub

    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : Button focus
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        If cmbInput(0).Enabled = True Then
            cmbInput(0).Focus()
        Else
            cmbInput(8).Focus()
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