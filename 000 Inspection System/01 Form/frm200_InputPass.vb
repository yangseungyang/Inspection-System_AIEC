Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm200_InputPass

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
    Dim btnDialogResult() As Button
    Dim txtInput() As TextBox
    Dim btnKeyText() As Button
    Dim btnKeyNumber() As Button
    Dim btnKeyDel() As Button
    Dim btnKeyCaps() As Button

    'Local Property =======================================
    Dim mDialogResult As String = ""

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm200_InputPass_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '==============================================================================================================
        'Function : Display start
        'Remarks  :
        '==============================================================================================================

        'Equipment Information ====================================================================
        mSystemLog = frm001_Main.SYSTEM_LOG
        'mServerDrive = frm001_Main.SERVER_DRIVE
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
            .Text = "Input Password"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 450 : .Height = 320
            .Top = (Screen.PrimaryScreen.Bounds.Height \ 2) - (.Height \ 2)
            .Left = (Screen.PrimaryScreen.Bounds.Width \ 2) - (.Width \ 2)
        End With

        Try
            cFrm = New cls090_DesignForm_Type01

            cFrm.FORM_NAME = Me
            cFrm.PIC_FOLDER = mAppPath & "900 Picture\Picture\GreenForm\"
            cFrm.TITLE_BAR_L = sBar(0) : cFrm.TITLE_BAR_R = sBar(1)
            cFrm.UNDER_BAR_L = sBar(2) : cFrm.UNDER_BAR_R = sBar(3)
            cFrm.MIN_BUTTON = sBtn(0)
            cFrm.CLOSE_BUTTON = sBtn(1)
            cFrm.PANEL_TITLE = sLbl(0)

            Call cFrm.FORM_DRAW()

            Call cFrm.TEXT_DRAW(sLbl(0), Me.Text, "Times new roman", 8, 5, 7, Color.ForestGreen)

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

    Private Sub frm200_InputPass_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
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

    Private Sub frm200_InputPass_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '==========================================================================================
        'Function : Close
        'Remarks  :
        '==========================================================================================


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
        Call Me.SysLogWrite(True, "●　パスワード入力 ----------------------- (" & Me.Name & "起動)")

        'Form Draw ================================================================================
        Call Me.FormDraw()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, Me.Name & " Form Draw", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
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
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\GreenForm\BackColor.bmp")
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
                .Width = 280 : .Height = 16
                .Left = CInt(Me.Width / 2 - .Width / 2)
                .TextAlign = ContentAlignment.MiddleCenter
                If i = 0 Then
                    .Top = sBar(0).Top + sBar(0).Height + 20
                    .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\GrayForm\ButtonON.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End If
                .BackColor = Color.FromArgb(50, 50, 60)
                .ForeColor = Color.White
                .BorderStyle = BorderStyle.None
                .BringToFront()
                .Cursor = Cursors.Hand
                If i = 0 Then .Text = "パスワードを入力してOKを押して下さい。"
            End With
        Next

        'Input TextBox ============================================================================
        Me.txtInput = New System.Windows.Forms.TextBox(0) {}
        For i As Integer = 0 To 0
            Me.txtInput(i) = New System.Windows.Forms.TextBox
            Me.Controls.Add(txtInput(i))
            AddHandler txtInput(i).KeyPress, AddressOf txtInput_KeyPress
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With txtInput(i)
                .Name = "txtInput(" & CStr(i) & ")"
                .AutoSize = False
                .Font = New Font("MS Gothic", 10, FontStyle.Regular)
                .Width = lblTitle(0).Width
                .Left = lblTitle(0).Left
                .Top = lblTitle(i).Top + lblTitle(i).Height + 2
                .BackColor = Color.FromArgb(236, 233, 216)
                .ForeColor = Color.FromArgb(30, 30, 30)
                .PasswordChar = CChar("*")
                .BringToFront()
            End With
        Next

        'Key button(Text) =========================================================================
        Me.btnKeyText = New System.Windows.Forms.Button(25) {}
        For i As Integer = 0 To 25
            Me.btnKeyText(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnKeyText(i))
            AddHandler btnKeyText(i).Click, AddressOf btnKey_click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 25
            With btnKeyText(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = 30 : .Height = 25
                .ForeColor = Color.FromArgb(70, 70, 100)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\WhiteForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                .Text = Chr(i + 65)
                If i = 0 Then
                    .Top = txtInput(0).Top + txtInput(0).Height + 10
                    .Left = 40
                End If
                If i >= 1 And i <= 8 Then
                    .Top = txtInput(0).Top + txtInput(0).Height + 10
                    .Left = btnKeyText(i - 1).Left + btnKeyText(i - 1).Width + 1
                End If
                If i >= 9 And i <= 17 Then
                    .Top = btnKeyText(0).Top + btnKeyText(0).Height + 1
                    .Left = btnKeyText(i - 9).Left
                End If
                If i >= 18 And i <= 25 Then
                    .Top = btnKeyText(9).Top + btnKeyText(9).Height + 1
                    .Left = btnKeyText(i - 9).Left
                End If
            End With
        Next

        'Key button(Number) =======================================================================
        Me.btnKeyNumber = New System.Windows.Forms.Button(9) {}
        For i As Integer = 0 To 9
            Me.btnKeyNumber(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnKeyNumber(i))
            AddHandler btnKeyNumber(i).Click, AddressOf btnKey_click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 9
            With btnKeyNumber(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = 30 : .Height = 25
                .ForeColor = Color.FromArgb(255, 255, 255)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\GrayForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                .Text = CStr(i)
                If i = 0 Then
                    .Top = btnKeyText(18).Top
                    .Left = btnKeyText(17).Left
                End If
                If i = 1 Then
                    .Top = txtInput(0).Top + txtInput(0).Height + 10
                    .Left = btnKeyText(8).Left + btnKeyText(8).Width + 2
                End If
                If i >= 2 And i <= 3 Then
                    .Top = btnKeyNumber(1).Top
                    .Left = btnKeyNumber(i - 1).Left + btnKeyNumber(i - 1).Width + 1
                End If
                If i >= 4 And i <= 6 Then
                    .Top = btnKeyNumber(1).Top + btnKeyNumber(1).Height + 1
                    .Left = btnKeyNumber(i - 3).Left
                End If
                If i >= 7 And i <= 9 Then
                    .Top = btnKeyNumber(4).Top + btnKeyNumber(4).Height + 1
                    .Left = btnKeyNumber(i - 3).Left
                End If
            End With
        Next

        'Key button(Del) ==========================================================================
        Me.btnKeyDel = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnKeyDel(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnKeyDel(i))
            AddHandler btnKeyDel(i).Click, AddressOf btnKey_click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With btnKeyDel(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = btnKeyText(0).Width * 2 + 1 : .Height = 25
                .ForeColor = Color.FromArgb(0, 0, 0)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\WhiteForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                .Text = "Del"
                .Top = btnKeyNumber(0).Top + btnKeyNumber(0).Height + 1
                .Left = btnKeyText(0).Left
            End With
        Next

        'Key button(Caps) =========================================================================
        Me.btnKeyCaps = New System.Windows.Forms.Button(0) {}
        For i As Integer = 0 To 0
            Me.btnKeyCaps(i) = New System.Windows.Forms.Button
            Me.Controls.Add(btnKeyCaps(i))
            AddHandler btnKeyCaps(i).Click, AddressOf btnKey_click
        Next
        Me.ResumeLayout()

        For i As Integer = 0 To 0
            With btnKeyCaps(i)
                .AutoSize = False
                .Font = New Font("MS Mincho", 8, FontStyle.Regular)
                .TextAlign = ContentAlignment.MiddleCenter
                .Width = btnKeyDel(0).Width : .Height = 25
                .ForeColor = Color.FromArgb(0, 0, 0)
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\WhiteForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                .Text = "Caps"
                .Top = btnKeyDel(0).Top
                .Left = btnKeyDel(0).Left + btnKeyDel(0).Width + 1
            End With
        Next

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
                Call Me.InputCheck()
                If mErrorComment <> "" Then
                    MessageBox.Show(mErrorComment, "Input Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    mErrorComment = ""
                    Call Me.SysLogWrite(True, "パスワード入力　NG：" & txtInput(0).Text)
                    txtInput(0).Text = ""
                    Exit Sub
                End If
                mDialogResult = "PASS OK"
                Call Me.SysLogWrite(True, "パスワード入力　OK")

                Me.Close()

            Case "CANCEL"
                mDialogResult = ""
                Me.Close()

        End Select

    End Sub

    Private Sub InputCheck()
        '==============================================================================================================
        ' Function : Check of Input text
        ' Remarks  :
        '==============================================================================================================

        If txtInput(0).Text <> "aiec" Then
            mErrorComment = "パスワードが違います。再度入力して下さい。"
        End If
        'If txtInput(0).Text <> "fam3" Then
        '    mErrorComment = "パスワードが違います。再度入力して下さい。"
        'End If

    End Sub

    Private Sub btnKey_click(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : Key push
        'Remarks  :
        '==============================================================================================================

        Select Case DirectCast(sender, Button).Text
            Case "Del"
                txtInput(0).Text = ""
            Case "Caps"
                For i As Integer = 0 To 25
                    btnKeyText(i).Text = Chr(i + 97)
                Next
                btnKeyCaps(0).Text = "Lock"
            Case "Lock"
                For i As Integer = 0 To 25
                    btnKeyText(i).Text = Chr(i + 65)
                Next
                btnKeyCaps(0).Text = "Caps"
            Case Else
                txtInput(0).Text &= DirectCast(sender, Button).Text
        End Select
        txtInput(0).Focus()

    End Sub

    Private Sub txtInput_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        '==============================================================================================================
        'Function : Enter
        'Remarks  :
        '==============================================================================================================

        Dim kASC As Integer = Asc(e.KeyChar)

        If kASC = 13 Then
            Call Me.InputCheck()
            If mErrorComment <> "" Then
                MessageBox.Show(mErrorComment, "Input Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
                mErrorComment = ""
                Call Me.SysLogWrite(True, "パスワード入力　NG：" & txtInput(0).Text)
                txtInput(0).Text = ""
                Exit Sub
            End If
            mDialogResult = "PASS OK"
            Call Me.SysLogWrite(True, "パスワード入力　OK")

            Me.Close()

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

        txtInput(0).Focus()

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