Option Strict On : Option Explicit On








'Doevents
Imports System.Windows.Forms.Application



Public Class frm500_MessagePopUp

    Dim cADO As cls010_MDB_Operation
    Dim cBcd1 As cls101_ComRS232C

    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mSetting As String = ""
    Dim mPort As String = ""
    Dim mDelim As String = ""

    Dim mErrorComment As String

    Dim mDispFileName As String = ""
    Dim mDispTitleName As String = ""
    Dim mDispMessage As String = ""
    Dim mDialogResult As String
    Dim mLEDColor As String = ""
    Dim mBuzz1 As String = ""
    Dim mBuzz2 As String = ""
    Dim mFlash As String = ""
    Dim mBuzzTime As Integer = 0

    Dim mBarcodeCtrl As Boolean = False
    Dim mBarcodeReadData As String = ""

    Dim mediaPlayer As New WMPLib.WindowsMediaPlayer()

    Dim cSys As cls102_SignalTower

    Public WriteOnly Property LED_COLOR() As String
        Set(ByVal Value As String)
            mLEDColor = Value
        End Set
    End Property

    Public WriteOnly Property BUZZ1() As String
        Set(ByVal Value As String)
            mBuzz1 = Value
        End Set
    End Property

    Public WriteOnly Property BUZZ2() As String
        Set(ByVal Value As String)
            mBuzz2 = Value
        End Set
    End Property

    Public WriteOnly Property FLASH() As String
        Set(ByVal Value As String)
            mFlash = Value
        End Set
    End Property

    Public WriteOnly Property BUZZ_TIME() As Integer
        Set(ByVal Value As Integer)
            mBuzzTime = Value
        End Set
    End Property


    Public WriteOnly Property DISP_FILENAME() As String
        Set(ByVal Value As String)
            mDispFileName = Value
        End Set
    End Property

    Public WriteOnly Property BARCODE_CTRL() As Boolean
        Set(ByVal Value As Boolean)
            mBarcodeCtrl = Value
        End Set
    End Property

    Public WriteOnly Property DISP_TITLENAME() As String
        Set(ByVal Value As String)
            mDispTitleName = Value
        End Set
    End Property

    Public Property DISP_MESSAGE() As String
        Set(ByVal Value As String)
            mDispMessage = Value
        End Set
        Get
            DISP_MESSAGE = mDispMessage
        End Get
    End Property

    Public Property DIALOG_RESULT() As String
        Get
            DIALOG_RESULT = mDialogResult
        End Get
        Set(ByVal Value As String)
            mDialogResult = Value
        End Set
    End Property

    Public Sub frmShow()

        ''連続再生
        'mediaPlayer.settings.setMode("loop", True)
        ''オーディオファイルを指定する（自動的に再生される）
        'mediaPlayer.URL = mAppPath & "900 Picture\marche_tsc_pi.mid"
        ''再生する
        'mediaPlayer.controls.play()

        If mBarcodeCtrl = True Then
            Call BarcodeSetup()
        End If

        If mDispTitleName <> "" Then Me.Text = mDispTitleName

        If mLEDColor = "" Then mLEDColor = "GREEN"
        If mBuzz1 = "" Then mBuzz1 = "OFF"
        If mBuzz2 = "" Then mBuzz2 = "OFF"
        If mFlash = "" Then mFlash = "OFF"
        If mBuzzTime = 0 Then
            mBuzzTime = 300 * 1000      '5分間鳴りっぱなしにする
        Else
            mBuzzTime = mBuzzTime * 1000
        End If

        Timer1.Interval = mBuzzTime     'ブザー時間の設定

        cSys = New cls102_SignalTower
        'Call cSys.SignalTowerReset()
        Call cSys.SignalTower(mLEDColor, mBuzz1, mBuzz2, mFlash)
        cSys = Nothing

        Select Case mDispFileName
            Case ""
                Me.WebBrowser1.Visible = False
                Me.TextBox1.Visible = True

                Me.TextBox1.Text = mDispMessage

            Case Else
                Me.WebBrowser1.Visible = True
                Me.TextBox1.Visible = False

                Me.WebBrowser1.Navigate(mDispFileName)

        End Select

        tmCount.CountStart()

        Me.Show()
        DoEvents()
        Me.Button1.Focus()

    End Sub

    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        mDialogResult = Button1.Text
        Button1.BackColor = Color.Yellow
        Button1.Enabled = False
        DoEvents()

        cSys = New cls102_SignalTower
        Call cSys.SignalTower("GREEN", "OFF", "OFF", "OFF")
        cSys = Nothing

    End Sub

    Public Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        mDialogResult = Button2.Text
        Button2.BackColor = Color.Red
        Button2.Enabled = False
        DoEvents()

        cSys = New cls102_SignalTower
        Call cSys.SignalTower("RED", "OFF", "OFF", "OFF")
        cSys = Nothing

    End Sub

    Private Delegate Sub COM_Recieve()
    ''==========================================================================================
    ''Function : COM recieve event
    ''Remarks  :
    ''==========================================================================================
    Private Sub SerialPort1_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived

        'AddHandler TextBox1.TextChanged, AddressOf TextBox1_TextChanged

        Try
            Me.Invoke(New COM_Recieve(AddressOf BeginningProcess), New Object() {})
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
            Call cBcd1.FuncRecvCmd()
            mBarcodeReadData = cBcd1.RECEIVE_DATA

        Catch oExcept As Exception
            mErrorComment = cBcd1.ERR_COMMENT : If mErrorComment = "" Then oExcept.ToString()

        End Try

        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            'MessageBox.Show("Get Barcode Text OK !", "Get Barcode Text", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        If mBarcodeReadData = "OK" Then
            mDialogResult = "OK"
        Else
            mDialogResult = "NG"
        End If
        
    End Sub

    Private Sub BarcodeSetup()

        Dim dbName As String
        Dim tbName As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""

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

        Finally
            cADO = Nothing

        End Try

        If rCount <> 0 Then

            If mErrorComment <> "" Then Exit Sub
            Dim BarSetting() As String = Split(sData(1), "-")

            cBcd1 = New cls101_ComRS232C
            cBcd1.TERMINAL_CHAR = vbCr : mDelim = vbCr
            cBcd1.RS_COMM = Me.SerialPort1 : cBcd1.PORT = sData(0) : mPort = sData(0)
            cBcd1.SETTING = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)
            mSetting = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)

            cBcd1.PortSetup()

        End If

        Me.Button1.Visible = False
        Me.Button2.Visible = False

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        cSys = New cls102_SignalTower
        Call cSys.SignalTower("", "OFF", "OFF", "")
        cSys = Nothing

        Timer1.Enabled = False

    End Sub

    Private Sub frm500_MessagePopUp_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        ''再生しているオーディオを停止する
        'mediaPlayer.controls.stop()

    End Sub


End Class