Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm020_InspectionStartEnd

    'Class ================================================
    Dim cWin As cls001_Window_Control
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cDrv As cls020_Drive_Control
    Dim cFrm As cls090_DesignForm_Type01
    Dim cPsm As cls100_PLASMA_Order

    Dim cBcd1 As cls101_ComRS232C

    Public Shared Insfrm As frm030_InspectionInformation

    'Local (Standard form control) ========================
    Dim sBar() As PictureBox
    Dim sBtn() As Button
    Dim sLbl() As Label

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mMousePoint As Point
    Dim mSystemLog As String = ""
    Dim mServerDrive As String = ""
    Dim mEquipmentStatus As Integer = 0
    Dim mErrorComment As String = ""
    Dim mWarningComment As String = ""
    Dim mNormalComment As String = ""
    Dim mOIC As String = ""
    Dim mError As String = ""

    'Local (Special form control) ========================
    Dim btnDialogResult() As Button
    Dim lblTitle() As Label
    Dim lblItem() As Label
    Dim lblResult() As Label
    Dim lblComment() As Label
    Dim txtModel As TextBox
    Dim lblModel As Label

    Const Process_Code As String = "FAM3_011"       'Ver 0.40 GYS ADD

    'Local Property =======================================
    Dim mDialogResult As String = ""
    Dim mPlasmaOrder As String = ""
    Dim mPlasmaBarCode As String = ""
    Dim mRunMode As String = ""
    Dim mInspectionModel As String = ""

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property

    Public Property PLASMA_ORDER() As String
        Get
            PLASMA_ORDER = mPlasmaOrder
        End Get
        Set(ByVal value As String)
            mPlasmaOrder = value
        End Set
    End Property

    Public WriteOnly Property PLASMA_BARCODE() As String
        Set(ByVal value As String)
            mPlasmaBarCode = value
        End Set
    End Property

    Public WriteOnly Property RUN_MODE() As String
        Set(ByVal value As String)
            mRunMode = value
        End Set
    End Property

    Public WriteOnly Property INSPECTION_MODEL() As String
        Set(ByVal value As String)
            mInspectionModel = value
        End Set
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm020_InspectionStartEnd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '==============================================================================================================
        'Function : Display start
        'Remarks  :
        '==============================================================================================================

    End Sub

    Private Delegate Sub COM_Recieve()
    ''==========================================================================================
    ''Function : COM recieve event
    ''Remarks  :
    ''==========================================================================================
    Private Sub SerialPort1_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived

        Try
            Me.txtModel.Invoke(New COM_Recieve(AddressOf BeginningProcess), New Object() {})
        Catch

        Finally

        End Try

        cBcd1.PortClose()
        cBcd1 = Nothing

    End Sub

    Private Sub BeginningProcess()
        '==========================================================================================
        'Function : Beginning process of inspection
        'Remarks  :
        '==========================================================================================

        Try
            Call cBcd1.FuncBarRecvCmd()
            txtModel.Text = cBcd1.RECEIVE_DATA

        Catch oExcept As Exception
            mError = cBcd1.ERR_COMMENT : If mError = "" Then oExcept.ToString()

        End Try

        If mError <> "" Then
            MessageBox.Show(mError, "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            'MessageBox.Show("Get Barcode Text OK !", "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Public Sub InspectionStart()
        '==============================================================================================================
        ' Function : Inspection start
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
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
        '-----rev0.07 del-----
        'Me.sBtn = New System.Windows.Forms.Button(1) {}
        'For i As Integer = 0 To 1
        '    Me.sBtn(i) = New System.Windows.Forms.Button
        '    Me.Controls.Add(sBtn(i))
        'Next
        'AddHandler sBtn(0).Click, AddressOf btnMinimum_Click
        'AddHandler sBtn(1).Click, AddressOf btnClose_Click
        '---------------------
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
            .Text = "frm020 Inspection Start Control"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 750 : .Height = 350
            .StartPosition = FormStartPosition.CenterScreen
        End With

        Try
            cFrm = New cls090_DesignForm_Type01

            cFrm.FORM_NAME = Me
            cFrm.PIC_FOLDER = mAppPath & "900 Picture\Picture\RedForm\"
            cFrm.TITLE_BAR_L = sBar(0) : cFrm.TITLE_BAR_R = sBar(1)
            cFrm.UNDER_BAR_L = sBar(2) : cFrm.UNDER_BAR_R = sBar(3)
            '-----rev0.07 del-----
            'cFrm.MIN_BUTTON = sBtn(0)
            'cFrm.CLOSE_BUTTON = sBtn(1)
            '---------------------
            cFrm.PANEL_TITLE = sLbl(0)

            Call cFrm.FORM_DRAW()

            Call cFrm.TEXT_DRAW(sLbl(0), Me.Text, "Times new roman", 8, 5, 7, Color.DarkRed)

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

    Private Sub frm020_InspectionStartEnd_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
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

    Private Sub frm020_InspectionStartEnd_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

        Dim pNum As Integer = 0
        Dim pName As String = ""

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "●　検査動作制御 ------------------------ (" & Me.Name & "起動)")

        'Form Draw ================================================================================
        Call Me.FormDraw()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, "Inspection Start Form Draw", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        DoEvents()
        Me.Show()

        If mRunMode = "CHK" Then

        End If

        'PLASMA Communication =====================================================================
        pNum = 0 : pName = "PLASMA通信"

        If mRunMode = "AUT" And mPlasmaBarCode <> "" Then
            Call Me.PlasmaCommunication()
        ElseIf mRunMode = "AUT" And mPlasmaBarCode = "" Then
            mNormalComment = "オーダー手動入力"
        Else
            mNormalComment = "「AUT」でない、またはバーコードが無いため、未実施"
            'If mPlasmaOrder = "" Then
            '    mPlasmaOrder = mInspectionModel & ", , , , , , "
            'End If
            mPlasmaOrder = mInspectionModel & ",,,,,,,"              '",,,,,,"
        End If

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        If mErrorComment <> "" Or mWarningComment <> "" Then
            Call Me.subErrorDisp()
            If mDialogResult = "検査中止" Then Exit Sub
            mDialogResult = "" : mErrorComment = "" : mWarningComment = "" : mNormalComment = ""
        End If

        DoEvents()
        Me.Show()

        '形名照合 ==============================================================
        pNum = 1 : pName = "MSCODE照合"

        If mRunMode = "AUT" Then
            Call subSetModule(mPlasmaOrder)
        Else
            mNormalComment = "AUTでない為、未実施"
        End If

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        If mErrorComment <> "" Or mWarningComment <> "" Then
            Call Me.subErrorDisp()
            If mDialogResult = "検査中止" Then Exit Sub
            mDialogResult = "" : mErrorComment = "" : mWarningComment = "" : mNormalComment = ""
        End If

        DoEvents()
        Me.Show()

        '機種名照合 ==============================================================
        pNum = 2 : pName = "検査装置対応機種照合"

        Call ModuleNameSerch(mPlasmaOrder)

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        If mErrorComment <> "" Or mWarningComment <> "" Then
            Call Me.subErrorDisp()
            If mDialogResult = "検査中止" Then Exit Sub
            mDialogResult = "" : mErrorComment = "" : mWarningComment = "" : mNormalComment = ""
        End If

        DoEvents()
        Me.Show()

        'Write of Start Informmation ==============================================================
        pNum = 3 : pName = "開始データ保存"

        Call Me.WriteStartInformation()

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        If mErrorComment <> "" Or mWarningComment <> "" Then
            Call Me.subErrorDisp()
            If mDialogResult = "検査中止" Then Exit Sub
            mDialogResult = "" : mErrorComment = "" : mWarningComment = "" : mNormalComment = ""
        End If

        DoEvents()
        Me.Show()

        'Order Check ==============================================================================
        pNum = 4 : pName = "オーダー重複確認"

        If mRunMode = "AUT" Then
            Call Me.OrderCheck()    '要る？
        Else
            mNormalComment = "「AUT」でないため、未実施"
        End If

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        If mErrorComment <> "" Or mWarningComment <> "" Then
            Call Me.subErrorDisp()
            If mDialogResult = "検査中止" Then Exit Sub
            mDialogResult = "" : mErrorComment = "" : mWarningComment = "" : mNormalComment = ""
        End If

        'MODEL Check ==============================================================================
        pNum = 5 : pName = "検査用型式判定"

        Call Me.InspectionModel()

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        If mErrorComment <> "" Or mWarningComment <> "" Then
            Call Me.subErrorDisp()
            If mDialogResult = "検査中止" Then Exit Sub
            mDialogResult = "" : mErrorComment = "" : mWarningComment = "" : mNormalComment = ""
        End If

        'v0.06　検査中止フラグ確認を追加　==============================================================================
        pNum = 6 : pName = "検査中止フラグの確認"

        If mRunMode = "AUT" Then
            Call Me.InspectionStopFlagCheck()
        Else
            mNormalComment = "「AUT」でないため、未実施"
        End If

        Call Me.ResultDisp(pNum, pName, mErrorComment, mWarningComment, mNormalComment)
        If mErrorComment <> "" Or mWarningComment <> "" Then
            Call Me.subErrorDisp()
            If mDialogResult = "検査中止" Then Exit Sub
            mDialogResult = "" : mErrorComment = "" : mWarningComment = "" : mNormalComment = ""
        End If



        'v0.01始業点検してあるかどうかﾁｪｯｸを追加


        Me.Visible = False

        'Inspection Rutine ========================================================================
        'Open Setting panel -----------------------------------------
        Call Me.SysLogWrite(True, "検査画面呼出し・・・")
        Dim sRet As String = ""

        Insfrm = New frm030_InspectionInformation
        Try

            Call Insfrm.ShowDialog()

            sRet = Insfrm.DIALOG_RESULT

        Catch


        Finally
            Insfrm = Nothing

        End Try


        'Display of operation panel(Normal end) ===================================================
        Me.Close()

    End Sub

    Private Sub subErrorDisp()
        '==============================================================================================================
        ' Function : Error disp and wait button
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

        If mErrorComment <> "" Then
            btnDialogResult(0).Visible = False
            btnDialogResult(1).Left = CInt((Me.Width / 2) - (btnDialogResult(1).Width / 2))
        End If

        If mWarningComment <> "" Then
            btnDialogResult(0).Visible = True
            btnDialogResult(1).Left = CInt((Me.Width / 2) + 20)

            If InStr(mWarningComment, "セット形名") <> 0 Then

                Me.txtModel.Visible = True
                Me.lblModel.Visible = True

                'Read of BarcodeReader Equipment No. ====================================================================
                If frm001_Main.COM_BARCODE = True Then

                    cBcd1 = New cls101_ComRS232C
                    cBcd1.TERMINAL_CHAR = frm001_Main.COM_BARCODE_TERMINAL_CHAR
                    cBcd1.RS_COMM = Me.SerialPort1 : cBcd1.PORT = frm001_Main.COM_BARCODE_PORT
                    cBcd1.SETTING = frm001_Main.COM_BARCODE_SETTING

                    cBcd1.PortSetup()

                End If

                Me.txtModel.Text = ""
                Me.txtModel.Focus()

            End If

        End If

        Me.Show()

        DoEvents()

        If Me.txtModel.Visible = True Then
            Me.txtModel.Focus()
        ElseIf btnDialogResult(0).Visible = True Then
            Me.btnDialogResult(0).Focus()
        ElseIf btnDialogResult(1).Visible = True Then
            Me.btnDialogResult(1).Focus()
        End If
        
        'Wait Button ON ---------------------------------------------
        Do
            If mDialogResult <> "" Then Exit Do
            DoEvents()
        Loop

        Me.Hide()

        If InStr(mWarningComment, "セット形名") <> 0 Then

            If InStr(txtModel.Text, vbCr) <> 0 Then txtModel.Text = Replace(txtModel.Text, vbCr, "")
            txtModel.Text = Trim(txtModel.Text)

            Dim zModel As String = Split(mPlasmaOrder, ",")(0)
            If Split(mPlasmaOrder, ",")(1) <> "" Then zModel = Split(mPlasmaOrder, ",")(1)

            If InStr(zModel, "/") <> 0 Then
                txtModel.Text = txtModel.Text & Mid(zModel, InStr(zModel, "/"), zModel.Length - InStr(zModel, "/") + 1)
            End If

            If InStr(mPlasmaOrder, "T911") <> 0 Then
                mPlasmaOrder = Replace(mPlasmaOrder, "Fxxxxxxxx", txtModel.Text)
            Else
                mPlasmaOrder = Replace(mPlasmaOrder, zModel, txtModel.Text)
            End If


            Me.txtModel.Visible = False
                Me.lblModel.Visible = False
                If frm001_Main.COM_BARCODE = True Then

                    'cBcd.PortClose()
                    cBcd1 = Nothing

                End If

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
                .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\RedForm\BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
                .Cursor = Cursors.Hand
                .BringToFront()
                If i = 0 Then
                    .Left = CInt((Me.Width / 2) - .Width - 20)
                    .Text = "検査実施"
                    .Visible = False
                End If
                If i = 1 Then
                    .Left = CInt((Me.Width / 2) - (.Width / 2))
                    .Text = "検査中止"
                    .Visible = True
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

        Dim rec As Integer = 7
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

        '機種名入力テキストボックス========================================================================
        Me.txtModel = New System.Windows.Forms.TextBox
        Me.Controls.Add(txtModel)
        AddHandler txtModel.TextChanged, AddressOf txtModel_TextChanged
        AddHandler txtModel.KeyPress, AddressOf txtModel_KeyPress
        Me.ResumeLayout()

        With txtModel

            .Font = New Font("MS Gothic", 12, FontStyle.Regular)
            .AutoSize = False
            .Height = 16
            .Width = Me.btnDialogResult(1).Width
            .Top = Me.btnDialogResult(0).Top - .Height * 2
            .Left = lblTitle(1).Left + lblTitle(1).Width + 1
            .BorderStyle = BorderStyle.None
            .ForeColor = Color.Black
            .BackColor = Color.White
            '.BackgroundImageLayout = ImageLayout.Stretch
            .BringToFront()
            .Text = ""
            .Visible = False

        End With

        '機種名 Label ==============================================================================
        Me.lblModel = New System.Windows.Forms.Label
        Me.lblModel = New System.Windows.Forms.Label
        Me.Controls.Add(lblModel)
        Me.ResumeLayout()

        With lblModel
            .AutoSize = False
            .Font = New Font("MS Gothic", 8, FontStyle.Regular)
            .Height = 16
            .Top = Me.btnDialogResult(0).Top - txtModel.Height * 2
            .TextAlign = ContentAlignment.MiddleCenter
            .BackColor = Color.FromArgb(0, 0, 0)
            .ForeColor = Color.FromArgb(255, 255, 255)
            .BorderStyle = BorderStyle.None
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\Graylabel.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch
            .BringToFront()

            .Width = CInt(bWidth * 0.15) : .Left = lblTitle(0).Left + lblTitle(0).Width + 1
            .Text = "機種名"
            .Visible = False

        End With


    End Sub

    Private Sub txtModel_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtModel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar = Chr(Keys.Enter) Then '-------------エンターキーが入力されたら

            DoEvents()

            Me.btnDialogResult(0).Focus()

        End If

    End Sub

    Private Sub PlasmaCommunication()
        '==============================================================================================================
        ' Function : PlasmaCommunication
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim dbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim tbName As String = ""
        Dim mErrorComment As String = ""
        Dim rField As String = ""
        Dim sSQL As String = ""

        Call Me.SysLogWrite(False, "")

        'Get of Plasma data =======================================================================
        Try
            cPsm = New cls100_PLASMA_Order
            cPsm.BARCODE_TEXT = mPlasmaBarCode

            Call cPsm.GET_ORDER_INFORMAION()

            mPlasmaOrder = cPsm.MSCODE & "," & cPsm.XJNO & "," & cPsm.SERIAL & "," & _
                    cPsm.ORDER & "," & cPsm.ITEM & "," & cPsm.COMP & "," & cPsm.QICRQ & "," & cPsm.LINKAGE

            mNormalComment = mPlasmaOrder

            'Traceability Data Save ==================================================

            dbName = mAppPath & "210 Result Data\Traceability_Temp.mdb"
            tbName = "tblResult"

            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "BAR = '" & mPlasmaBarCode & "' , "
            sSQL &= "SERIAL = '" & cPsm.SERIAL & "' , "
            sSQL &= "PROD_NO = '" & cPsm.PRDNO & "' , "
            sSQL &= "SAP_LINKAGE_NO = '" & cPsm.LINKAGE & "' , "
            sSQL &= "SAP_COMP_NO = '" & cPsm.COMP & "' , "
            sSQL &= "CRP_GR_NO = '" & cPsm.CORPG & "' , "
            sSQL &= "START_NO = '" & cPsm.START & "' , "
            sSQL &= "MS_CODE = '" & cPsm.MSCODE & "' , "
            sSQL &= "FINISH_SCHDULE_D = '" & cPsm.FINSC & "' , "
            sSQL &= "INST_FINISH_D = '" & cPsm.INSFN & "' , "
            sSQL &= "XJ_NO = '" & cPsm.XJNO & "' , "
            sSQL &= "PROCESS_CODE = '" & Process_Code & "' , "
            sSQL &= "RDATE_S = '" & "" & "' , "
            sSQL &= "RTIME_S = '" & "" & "' , "
            sSQL &= "WORKERID = '" & "" & "' , "
            sSQL &= "INDEX_NO = '" & cPsm.Index & "' , "
            sSQL &= "BYORD_START_NO = '' , "
            sSQL &= "PC_Name = '" & System.Net.Dns.GetHostName & "' , "
            sSQL &= "SOFT = '" & "" & "' , "
            sSQL &= "VER = '" & "" & "'"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cPsm.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cPsm = Nothing
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "PLASMA通信：" & mErrorComment)
        Else
            Call Me.SysLogWrite(True, "PLASMA通信：" & mNormalComment)
        End If

    End Sub

    Private Sub WriteStartInformation()
        '==============================================================================================================
        ' Function : DB write of Start data.
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

        Dim nRec As Integer = 1
        Dim sRec() As String = {"", "", ""}

        'Write of starting infomation =============================================================
        dbName = mAppPath & "200 Start Information\StartInformation.mdb"
        tbName = "tblInspectionStart"

        'Get of field name ------------------------------------------
        Do
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
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "開始条件保存：" & mErrorComment)
                Exit Sub
            End If

            'Delete of data -----------------------------------------
            If rCount <> 0 Then

                Try
                    sSQL = "DELETE FROM " & tbName & " WHERE " & fName(0) & " = '" & sData(0) & "'"

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

            Else
                Exit Do

            End If
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "開始条件保存：" & mErrorComment)
                Exit Sub
            End If

        Loop

        'Write of order information ---------------------------------
        Dim sInformation As String = mPlasmaOrder & "," & mRunMode & "," & mInspectionModel
        Dim sInfo() As String = Split(sInformation, ",")

        mOIC = sInfo(3) & "-" & sInfo(4) & "-" & sInfo(5)

        Try

            sSQL = "INSERT INTO " & tbName & " ("

            For i = 0 To UBound(sInfo)
                sSQL &= fName(i) & " , "
            Next
            sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 3)

            sSQL &= ") VALUES ("

            For i = 0 To UBound(sInfo)
                sSQL &= "'" & sInfo(i) & "' , "
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
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "開始条件保存：" & mErrorComment)
        Else
            mNormalComment = sInformation
            Call Me.SysLogWrite(True, "開始条件保存：" & mNormalComment)
        End If

    End Sub

    Private Sub OrderCheck()
        '==============================================================================================================
        ' Function : Checking for duplication of order information 
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

        'Search of result order ===================================================================
        dbName = mAppPath & "200 Start Information\OrderResult.mdb"
        tbName = "tblOrderResult"

        'Check of record count --------------------------------------
        Try
            'sSQL = "SELECT * FROM " & tbName
            sSQL = "SELECT count(*) FROM " & tbName

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            'fCount = cADO.DB_FIELD_COUNT
            'rCount = cADO.DB_RECORD_COUNT
            rCount = CInt(sData(0))

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "オーダチェック：" & mErrorComment)
            Exit Sub
        End If

        Dim oTotalResultRecord As Integer = rCount

        'Search of result order -------------------------------------
        Try
            sSQL = "SELECT OIC FROM " & tbName & " WHERE OIC = '" & mOIC & "'"

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
            mErrorComment = cADO.ERR_COMMENT : cADO = Nothing

        End Try
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "オーダチェック：" & mErrorComment)
            Exit Sub
        End If

        'Write of Order result ====================================================================
        If rCount = 0 Then
            Call Me.SysLogWrite(True, "オーダチェック：オーダ重複無し")

            'OrderResultDB is MAX 10000 records ---------------------
            If oTotalResultRecord >= 10000 Then
                Try
                    sSQL = "DELETE FROM " & tbName & " WHERE " & "Startdate" & " in "
                    sSQL = sSQL & "(SELECT " & "MIN(" & "Startdate" & ") AS MinVeri FROM " & tbName & ")"

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

            End If
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "オーダチェック：" & mErrorComment)
                Exit Sub
            End If

            'Write of new data --------------------------------------
            Try
                sSQL = "INSERT INTO " & tbName & " (StartDate , OIC) VALUES ('"
                sSQL &= Format(DateTime.Now, "yyyy.MM.dd HH:mm:ss") & "' , '" & mOIC & "')"

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
            If mErrorComment <> "" Then
                Call Me.SysLogWrite(True, "オーダチェック：" & mErrorComment)
            Else
                mNormalComment = "オーダチェック：" & "新規登録完了"
                Call Me.SysLogWrite(True, mNormalComment)
            End If
            Exit Sub

        End If

        'Order result <> 0 ------------------------------------------
        mWarningComment = "過去に実施したオーダです。検査を実施しますか？"
        Call Me.SysLogWrite(True, mWarningComment)

    End Sub

    Private Sub InspectionModel()
        '==============================================================================================================
        ' Function : Decision of inspection model
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

        'Get of starting infomation ===============================================================
        dbName = mAppPath & "200 Start Information\StartInformation.mdb"
        tbName = "tblInspectionStart"

        'Get of Order info ------------------------------------------
        Try
            sSQL = "SELECT MSCODE,XJNo,RunMode,InspectionModel FROM " & tbName

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
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "検査用型式：" & mErrorComment)
            Exit Sub
        End If

        'Run mode check ===========================================================================
        If sData(2) = "MAN" Or sData(2) = "CHK" Then
            If sData(3) = "" Then
                mErrorComment = "型式が無いので、検査動作が開始できません。"
                Call Me.SysLogWrite(True, "検査用型式：" & mErrorComment)
                Exit Sub
            Else
                mNormalComment = "検査用型式：" & sData(3)
                Call Me.SysLogWrite(True, mNormalComment)
                Exit Sub
            End If
        End If

        'Order model to Inspection model ==========================================================
        Dim InspectionModel As String = ""

        InspectionModel = sData(0)

        If InspectionModel = "" Then
            mErrorComment = "型式が無いので、検査動作が開始できません。"
            Call Me.SysLogWrite(True, "検査用型式：" & mErrorComment)
            Exit Sub
        End If

        'Check Option specification ===============================================================
        Call subOptionCode(InspectionModel)
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "検査用型式：" & mErrorComment)
            Exit Sub
        End If

        'Write of Inspection model ================================================================
        Try
            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "InspectionModel = '" & InspectionModel & "'"
            sSQL &= " WHERE " & fName(0) & " = '" & sData(0) & "'"

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
        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "検査型式：" & mErrorComment)
        Else
            mNormalComment = "検査用型式：" & InspectionModel
            Call Me.SysLogWrite(True, mNormalComment)
        End If

        mInspectionModel = InspectionModel

    End Sub

    Private Sub InspectionStopFlagCheck()
        '==============================================================================================================
        ' Function : 検査中止ﾌﾗｸﾞ確認
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
        Dim rField, kField As String
        
        dbName = mAppPath & "400 Common File\MSCODE.mdb"

        cADO = New cls010_MDB_Operation

        tbName = "tblMSCODE"
        rField = "検査中止フラグ , 検査中止メッセージ" : kField = "Model"

        sSQL = "SELECT " & rField & " FROM " & tbName
        sSQL = sSQL & " WHERE " & kField & " = '" & mInspectionModel & "'"

        cADO.DB_PATH = dbName
        cADO.TB_NAME = tbName
        cADO.DB_SQL = sSQL

        Call cADO.DB_READ_JET(fName, sData)

        If sData(0) = "True" Then
            mErrorComment = "検査中止ﾌﾗｸﾞがONです。ﾒｯｾｰｼﾞ：" & sData(1)
            Exit Sub
        End If
        
        mNormalComment = "検査中止ﾌﾗｸﾞなしOK。"
        cADO = Nothing

    End Sub

    Private Sub subSetModule(ByRef sModel As String)
        '==============================================================================================================
        ' Function : Decision of inspection model
        ' Data     : Inspection model(Return)
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
        Dim rField, kField As String
        Dim model() As String = Split(sModel, ",")
        Dim Suffix As String
        Dim SerModel As String

        SerModel = model(0)

        dbName = mAppPath & "400 Common File\MSCODE.mdb"

        cADO = New cls010_MDB_Operation

        tbName = "tblMSCODE"
        rField = "SetValue , Model" : kField = "Model"

        If InStr(1, SerModel, "/CT") <> 0 Then
            SerModel = Replace(SerModel, "/CT", "")
        End If
        If InStr(1, SerModel, "/CE") <> 0 Then
            SerModel = Replace(SerModel, "/CE", "")
        End If

Reserch1:
        'sSQL = "SELECT " & rField & " FROM " & tbName
        'sSQL = sSQL & " WHERE " & kField & " LIKE '" & SerModel & "'"  '& "ORDER BY " & kField        '& " DESC"

        'cADO.DB_PATH = dbName
        'cADO.TB_NAME = tbName
        'cADO.DB_SQL = sSQL

        'Call cADO.DB_READ_JET(fName, sData)
        'If cADO.DB_DATA_COUNT = 0 Then

        '    If InStr(SerModel, "/") <> 0 Then

        '        Suffix = Mid(SerModel, InStr(1, SerModel, "/") + 1, Len(SerModel) - InStr(1, SerModel, "/") + 1)
        '        If InStr(Suffix, "/") <> 0 Then
        '            SerModel = Mid(SerModel, 1, InStr(1, SerModel, "/") - 1) & "/" & Mid(Suffix, 1, InStr(1, Suffix, "/") - 1)
        '        Else
        '            SerModel = Mid(SerModel, 1, InStr(1, SerModel, "/") - 1)
        '        End If
        '        GoTo Reserch1

        '    Else

        '        mErrorComment = "MSCODE.mdbに形名登録がありません。"
        '        cADO = Nothing
        '        Exit Sub

        '    End If

        'ElseIf sData(0) = "True" Then
        '    If Mid(SerModel, 1, 4) = "T911" Then
        '        mWarningComment = "セット形名です。Serial Noを入力してください。"
        '    Else
        '        mWarningComment = "セット形名です。機種名を入力してください。"
        '    End If

        '    cADO = Nothing
        '    Exit Sub

        'End If

        mNormalComment = "MSCODE.mdbとの形名照合OK"
        cADO = Nothing

    End Sub

    Private Sub ModuleNameSerch(ByRef sModel As String)
        '==============================================================================================================
        ' Function : Decision of inspection model
        ' Data     : Inspection model(Return)
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
        Dim rField, kField As String
        Dim model() As String = Split(sModel, ",")
        Dim Suffix As String
        Dim SerModel As String

        SerModel = model(0)

        dbName = mAppPath & "500 Model Information\ModelInformation.mdb"

        cADO = New cls010_MDB_Operation

        tbName = "tblModelInformation"
        rField = "Product" : kField = "Model"

        If InStr(1, SerModel, "/CT") <> 0 Then
            SerModel = Replace(SerModel, "/CT", "")
        End If
        If InStr(1, SerModel, "/CE") <> 0 Then
            SerModel = Replace(SerModel, "/CE", "")
        End If

Reserch1:
        sSQL = "SELECT " & rField & " FROM " & tbName
        sSQL = sSQL & " WHERE " & kField & " LIKE '" & SerModel & "'"  '& "ORDER BY " & kField        '& " DESC"

        cADO.DB_PATH = dbName
        cADO.TB_NAME = tbName
        cADO.DB_SQL = sSQL

        Call cADO.DB_READ_JET(fName, sData)
        If cADO.DB_DATA_COUNT = 0 Then

            If InStr(SerModel, "/") <> 0 Then

                Suffix = Mid(SerModel, InStr(1, SerModel, "/") + 1, Len(SerModel) - InStr(1, SerModel, "/") + 1)
                If InStr(Suffix, "/") <> 0 Then
                    SerModel = Mid(SerModel, 1, InStr(1, SerModel, "/") - 1) & "/" & Mid(Suffix, 1, InStr(1, Suffix, "/") - 1)
                Else
                    SerModel = Mid(SerModel, 1, InStr(1, SerModel, "/") - 1)
                End If
                GoTo Reserch1

            Else

                mErrorComment = "検査装置対応機種登録がありません。"
                cADO = Nothing
                Exit Sub

            End If

        Else

            mNormalComment = "検査装置機種登録確認"
            cADO = Nothing

        End If


    End Sub


    Private Sub subOptionCode(ByRef InspectionModel As String)
        '==============================================================================================================
        ' Function : Decision of inspection model
        ' Data     : Inspection model(Return)
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
        Dim rField, kField As String
        Dim OrderModel() As String = Split(InspectionModel, ",")
        Dim ReadModel() As String = {""}
        'Dim Suffix As String
        Dim SerModel() As String
        Dim fitMAX() As Integer

        Dim SyouGouModel() As String

        SerModel = Split(OrderModel(0), "/")

        dbName = mAppPath & "500 Model Information\ModelInformation.mdb"

        cADO = New cls010_MDB_Operation

        tbName = "tblModelInformation"
        rField = "Product,Model" : kField = "Model"

        'Reserch1:
        sSQL = "SELECT " & rField & " FROM " & tbName
        sSQL = sSQL & " WHERE " & kField & " LIKE '" & SerModel(0) & "%'"  '& "ORDER BY " & kField        '& " DESC"

        cADO.DB_PATH = dbName
        cADO.TB_NAME = tbName
        cADO.DB_SQL = sSQL

        Call cADO.DB_READ_JET(fName, sData)
        rCount = cADO.DB_RECORD_COUNT
        cADO = Nothing

        Select Case rCount
            Case 0
                mErrorComment = "検査型式確認異常"
            Case 1
                InspectionModel = sData(1)
            Case Is > 1
                ReDim fitMAX(rCount - 1)
                For i = 0 To rCount - 1
                    If sData(i * 2 + 1) = OrderModel(0) Then        'ドンピシャのときそれを使用
                        InspectionModel = sData(i * 2 + 1)
                        Exit Select
                    End If

                    ReadModel = Split(sData(i * 2 + 1), "/")
                    For j = 0 To ReadModel.Length - 1
                        If InStr(OrderModel(0), ReadModel(j)) <> 0 Then
                            fitMAX(i) = fitMAX(i) + 1
                        End If
                    Next j
                Next i

                Dim max As Integer = 0
                For Each int As Integer In fitMAX
                    max = Math.Max(int, max)
                Next

                '最終確認
                If max > 1 Then
                    SyouGouModel = Split(sData(Array.IndexOf(fitMAX, max) * 2 + 1), "/")

                    For i = 0 To UBound(SyouGouModel)
                        If InStr(OrderModel(0), SyouGouModel(i)) = 0 Then
                            mErrorComment = "検査型式照合異常。オーダー型式：" & OrderModel(0) & "　検査型式：" & sData(Array.IndexOf(fitMAX, max) * 2 + 1)
                            Exit Sub
                        End If
                    Next

                    InspectionModel = sData(Array.IndexOf(fitMAX, max) * 2 + 1)

                ElseIf max = 1 Then
                    InspectionModel = SerModel(0)
                End If

        End Select

        'If cADO.DB_DATA_COUNT = 0 Then

        '    If InStr(SerModel, "/") <> 0 Then

        '        Suffix = Mid(SerModel, InStr(1, SerModel, "/") + 1, Len(SerModel) - InStr(1, SerModel, "/") + 1)
        '        If InStr(Suffix, "/") <> 0 Then
        '            SerModel = Mid(SerModel, 1, InStr(1, SerModel, "/") - 1) & "/" & Mid(Suffix, 1, InStr(1, Suffix, "/") - 1)
        '        Else
        '            SerModel = Mid(SerModel, 1, InStr(1, SerModel, "/") - 1)
        '        End If
        '        GoTo Reserch1

        '    Else

        '        mErrorComment = "検査型式確認異常"
        '        cADO = Nothing
        '        Exit Sub

        '    End If

        'Else

        '    mNormalComment = "検査型式登録確認"
        '    cADO = Nothing

        'End If

        'InspectionModel = SerModel

    End Sub

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
            lblComment(n).Text = " " & Mid(e, 1, 50)
        End If

        'WARNING ==================================================================================
        If w <> "" Then
            If mEquipmentStatus = 2 Then mEquipmentStatus = 1

            lblItem(n).ForeColor = Color.Yellow
            lblResult(n).ForeColor = Color.Yellow
            lblComment(n).ForeColor = Color.Yellow
            lblResult(n).Text = "WARNING"
            lblComment(n).Text = " " & Mid(w, 1, 50)
        End If

        'NORMAL ===================================================================================
        If e = "" And w = "" Then
            lblResult(n).Text = "OK"
            lblComment(n).Text = " " & Mid(c, 1, 50)
        End If

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

            Case "検査実施"
                mDialogResult = DirectCast(sender, Button).Text

            Case "検査中止"
                mDialogResult = DirectCast(sender, Button).Text

        End Select

        Me.Hide()

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
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

End Class