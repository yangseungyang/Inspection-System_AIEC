'rev0.07 add
Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm600_Reprint
    'Class ================================================
    Dim cWin As cls001_Window_Control
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cFrm As cls090_DesignForm_Type01

    Dim cBcd As cls101_ComRS232C
    Dim cPsm As cls100_PLASMA_Order

    'Local (Standard form control) ========================
    Dim sBar() As PictureBox
    Dim sBtn() As Button
    Dim sLbl() As Label
    Dim mBarCode As Boolean = False
    Dim mSetting As String = ""
    Dim mPort As String = ""
    Dim mDelim As String = ""

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mMousePoint As Point
    Dim mErrorComment As String = ""
    Dim mSystemLog As String = ""
    Dim mNormalComment As String = ""

    'Local (Special form control) ========================
    Dim btnDialogResult() As Button
    Dim lblBarcode() As Label
    Public txtBarcode() As TextBox

    'Local Property =======================================
    Dim mDialogResult As String = ""
    Dim mPlasmaBarCode As String = ""
    Dim mPlasmaOrder As String = ""
    Dim mPlasmaSerial As String = ""
    Dim mError As String

    'Property =====================================================================================
    Public ReadOnly Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
    End Property
    Public WriteOnly Property PLASMA_BARCODE() As String
        Set(ByVal value As String)
            mPlasmaBarCode = value
        End Set
    End Property

    'Form control *****************************************************************************************************
    Private Sub frm600_Reprint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            .Text = "주명판 재출력"
            .Visible = False
            .Opacity = 0
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None

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
            cBcd.RS_COMM = Me.SerialPort2 : cBcd.PORT = sData(0) : mPort = sData(0)
            cBcd.SETTING = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)
            mSetting = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)
            mBarCode = True

            cBcd.PortSetup()

        Else
            txtBarcode(0).Text = ""

        End If


        'Process Start ============================================================================
        Call Me.MainProcess()

    End Sub

    Private Sub frm600_Reprint_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
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
        cBcd.PortClose()
        Me.Close()

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
        Call Me.SysLogWrite(True, "●　主銘板Reprint ------------------------- (" & Me.Name & "起動)")

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
                    .Visible = False
                End If
                If i = 1 Then
                    .Left = CInt((Me.Width / 2) - (.Width / 2))
                    .Text = "CLOSE"
                End If
            End With
        Next

        'BarCodeTextBox ============================================================================
        Me.txtBarcode = New System.Windows.Forms.TextBox(0) {}
        Me.txtBarcode(0) = New System.Windows.Forms.TextBox
        Me.Controls.Add(txtBarcode(0))
        AddHandler txtBarcode(0).KeyPress, AddressOf txtBarcode_KeyPress
        Me.ResumeLayout()

        With txtBarcode(0)
            .Font = New Font("MS Gothic", 14, FontStyle.Regular)
            .AutoSize = False
            .Height = 30
            .Width = 150
            .Top = Me.Height - 300
            .Left = CInt((Me.Width / 2) - (.Width / 2))
            .BorderStyle = BorderStyle.None
            .ForeColor = Color.Black
            .BackColor = Color.White
            .BringToFront()
            .Text = ""
        End With

        Me.lblBarcode = New System.Windows.Forms.Label(1) {}
        Me.lblBarcode(0) = New System.Windows.Forms.Label
        Me.lblBarcode(1) = New System.Windows.Forms.Label
        Me.Controls.Add(lblBarcode(0))
        Me.Controls.Add(lblBarcode(1))

        Me.ResumeLayout()

        With lblBarcode(0)
            .Font = New Font("MS Gothic", 14, FontStyle.Regular)
            .AutoSize = False
            .TextAlign = ContentAlignment.MiddleLeft
            .Height = 20
            .Width = 150
            .TextAlign = ContentAlignment.MiddleCenter
            .Top = Me.Height - 320
            .Left = CInt((Me.Width / 2) - (.Width / 2))
            .BorderStyle = BorderStyle.None
            .BackColor = Color.FromArgb(50, 50, 60)
            .ForeColor = Color.FromArgb(255, 255, 255)
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\GrayLabel.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch
            .BringToFront()
            .Text = "Barcode"
        End With

        With lblBarcode(1)
            .Font = New Font("MS Gothic", 25, FontStyle.Regular)
            .AutoSize = False
            .TextAlign = ContentAlignment.MiddleLeft
            .Height = 100
            .Width = 300
            .TextAlign = ContentAlignment.MiddleCenter
            .Top = Me.Height - 240
            .Left = CInt((Me.Width / 2) - (.Width / 2))
            .BorderStyle = BorderStyle.None
            .BackColor = Color.FromArgb(50, 50, 60)
            .ForeColor = Color.FromArgb(255, 255, 255)
            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\CPic.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch
            .BringToFront()
            .Text = "TAG의 INDEX NO를 Hit해주세요."
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

            Case "OK"
                mDialogResult = DirectCast(sender, Button).Text
                Me.Close()

            Case "CLOSE"
                mDialogResult = ""
                cBcd.PortClose()
                Me.Close()

        End Select

    End Sub
    'Sub process ******************************************************************************************************
    Private Sub FocusControl()
        '==============================================================================================================
        ' Function : Button focus
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

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
    Private Sub txtBarcode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        ' Function : 
        ' Remarks  :
        '==============================================================================================================

        Try
            If txtBarcode(0).TextLength = 6 Then
                txtBarcode(0).Text = StrConv(txtBarcode(0).Text, VbStrConv.Uppercase)
                Me.PLASMA_BARCODE = txtBarcode(0).Text

                Call Me.PlasmaCommunication()
                Call reprint()
            End If
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = ex.ToString
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            MsgBox(mErrorComment)
        Finally
            txtBarcode(0).Text = ""
        End Try

    End Sub
    Private Sub txtBarcode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar = Chr(Keys.Enter) Then '-------------エンターキーが入力されたら
            If mBarCode = True Then cBcd.PortClose()
            Try
                If txtBarcode(0).TextLength = 6 Then
                    txtBarcode(0).Text = StrConv(txtBarcode(0).Text, VbStrConv.Uppercase)
                    Me.PLASMA_BARCODE = txtBarcode(0).Text

                    Call Me.PlasmaCommunication()
                    Call reprint()
                End If
            Catch ex As Exception
                If mErrorComment = "" Then mErrorComment = ex.ToString
                mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                MsgBox(mErrorComment)
            Finally
                If mBarCode = True Then cBcd.PortSetup()
                txtBarcode(0).Text = ""
            End Try
        End If

    End Sub
    Private Sub PlasmaCommunication()
        '==============================================================================================================
        ' Function : PlasmaCommunication
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call Me.SysLogWrite(False, "")
        mErrorComment = ""

        'Get of Plasma data =======================================================================
        Try
            cPsm = New cls100_PLASMA_Order
            cPsm.BARCODE_TEXT = mPlasmaBarCode


            Call cPsm.GET_ORDER_INFORMAION()

            mPlasmaOrder = cPsm.MSCODE & "," & cPsm.XJNO & "," & cPsm.SERIAL & "," & _
                    cPsm.ORDER & "," & cPsm.ITEM & "," & cPsm.COMP & "," & cPsm.QICRQ & "," & cPsm.LINKAGE

            mPlasmaSerial = cPsm.SERIAL

            mNormalComment = mPlasmaOrder

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Call Me.SysLogWrite(True, "PLASMA通信Error：" & mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cPsm = Nothing
            If mErrorComment = "" Then
                Call Me.SysLogWrite(True, "PLASMA通信：" & mNormalComment)
            End If
        End Try
    End Sub
    Private Sub reprint()

        Dim cPrn As cls2P0_Inspection
        Dim dbName As String
        Dim rDTkey As String
        Dim tbName As String = ""
        Dim sData() As String = {""}
        Dim fName() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim MODEL As String

        Call SearchProcess()


        dbName = mAppPath & "\210 Result Data\ReprintSearch.mdb"

        'Read of StartDateTime, MSCODE =====================================================================
        Try
            tbName = "tblResult"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "StartDateTime,MSCODE" & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            fCount = cADO.DB_FIELD_COUNT
            rCount = cADO.DB_RECORD_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "ReprintSearch.mdb：" & mErrorComment)

        Finally
            cADO = Nothing
        End Try

        If sData(0) = "" Then
            MsgBox("검사Data가 없습니다.")
            Exit Sub
        Else

            rDTkey = sData(0)
            MODEL = sData(1)

            Try
                cPrn = New cls2P0_Inspection
                cPrn.InspectionMODEL = MODEL
                Call cPrn.LabelPrint(dbName, rDTkey)
            Catch ex As Exception
            Finally
                cPrn = Nothing
            End Try
        End If

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

        Dim SearchFileName As String = ""
        Dim MasterName As String = ""
        Dim sDir As String = ""
        Dim sDirs As String = ""
        Dim sDirName() As String = {""}
        Dim sDirCount As Integer = 0
        Dim Searchdata As String = ""
        Dim SearchDir As String = ""
        Dim SearchDirs As String = ""
        Dim SearchDirName() As String = {""}
        Dim SearchDirCount As Integer = 0

        'Search ===================================================================================

        'old SearchResult Kill --------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.FILE_PATH = mAppPath & "\210 Result Data\ReprintSearch.mdb"

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
            cFop.FILE_NAME = "ReprintSearch.mdb"

            Call cFop.FILE_COPY()

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        dbName = mAppPath & "\210 Result Data\ReprintSearch.mdb"
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


        'Search result month -----------------------------------------
        SearchDir = Dir(mAppPath & "\210 Result Data\", FileAttribute.Directory)
        Do
            If InStr(SearchDir, "-") <> 0 Then
                SearchDirs = SearchDirs & SearchDir & ","
            End If

            SearchDir = Dir()

            If SearchDir = "" Then
                SearchDirs = Mid(SearchDirs, 1, Len(SearchDirs) - 1)
                Exit Do
            End If
        Loop
        sDirName = Split(SearchDirs, ",")
        sDirCount = UBound(sDirName)


        If sDirCount = 0 Then
            Searchdata = sDirName(sDirCount)
        Else
            Searchdata = sDirName(sDirCount - 1)
        End If

        Searchdata = Searchdata & "-01"

        'Search SQL -------------------------------------------------
        tbName = "tblResult"
        kField = "StartDateTime"

        sSQL = "SELECT * FROM " & tbName & " WHERE StartDateTime >= '" & Searchdata
        sSQL = sSQL & "' AND SerialNo LIKE '" & mPlasmaSerial & "' AND Result LIKE 'PASS' "
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
                mErrorComment = "검사Data(MonthData)를 확인해주세요."
                MsgBox(mErrorComment)
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

        dbName = sDir

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
            dbName = mAppPath & "\210 Result Data\ReprintSearch.mdb"

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

    End Sub
    Private Delegate Sub COM_Recieve()
    ''==========================================================================================
    ''Function : COM recieve event
    ''Remarks  :
    ''==========================================================================================
    Private Sub SerialPort2_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort2.DataReceived

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

        RemoveHandler txtBarcode(0).TextChanged, AddressOf txtBarcode_TextChanged

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
        End If

    End Sub

    Private Sub DataWriteProcess()
        txtBarcode(0).Text = mPlasmaBarCode
    End Sub
End Class