Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls020_Network_Drive
'Function   : Connection & disconnect of Network drive
'Language   : VB2008　Framework3.5
'Reference  : Microsoft Scripting Runtime
'Memo       :
'
'Issue      : 2008/09/30  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

Imports System.Runtime.InteropServices

Public Class cls020_Drive_Control

    'Network connect parameter ============================
    Structure mNetResource
        Dim dwScope As Integer
        Dim dwType As Integer
        Dim dwDisplayName As Integer
        Dim dwUsage As Integer
        Dim lpLocalName As String
        Dim lpRemoteName As String
        Dim lpComment As String
        Dim lpProvider As String
    End Structure

    'Network drive connection =============================
    Private Declare Function WNetAddConnection2 Lib "mpr" Alias "WNetAddConnection2A" (ByRef lpNetResource As mNetResource, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer
    Private Declare Function WNetAddConnection3 Lib "mpr" Alias "WNetAddConnection3A" (ByVal hWnd As Integer, ByRef lpNetResource As mNetResource, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer
    Private Const RESURCETYPE_ANY As Integer = &H0
    Private Const RESURCETYPE_DISK As Integer = &H1
    Private Const RESURCETYPE_PRINT As Integer = &H2

    'Network drive disconnection ==========================
    Private Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnectionA" (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Boolean) As Integer

    'Get Drive Sign =======================================
    Private Declare Function GetLogicalDrives Lib "kernel32" () As Integer

    'Get Drive type =======================================
    Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal RootPathName As String) As Integer

    'Get Drive name(VolumeName) ===========================
    Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal RootPathName As String, ByVal VolumeName As String, ByVal VolumeNameSize As Integer, ByRef SerialNumber As Integer, ByRef MaxFileLen As Integer, ByRef FileSystemFlags As Integer, ByVal FileSystemName As String, ByVal FileSystemNameSize As Integer) As Integer

    'ドライブコントロール(EJECT用)
    Public Declare Function DeviceIoControl Lib "kernel32" _
      (ByVal hDevice As Long, _
       ByVal dwIoControlCode As Long, ByVal lpInBuffer As IntPtr, _
       ByVal nInBufferSize As Long, ByVal lpOutBuffer As IntPtr, _
       ByVal nOutBufferSize As Long, ByVal lpBytesReturned As Long, ByVal lpOverlapped As IntPtr) As Long

    Public Const IOCTL_STORAGE_LOAD_MEDIA As Long = &H2D480C
    Public Const IOCTL_STORAGE_EJECT_MEDIA As Long = &H2D4808
    Public Const IOCTL_STORAGE_MEDIA_REMOVAL As Long = &H2D4804

    Public Declare Function CreateFile Lib "kernel32" _
   Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, ByVal lpSecurityAttributes As IntPtr, _
   ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long

    Public Const INVALID_HANDLE_VALUE As Long = -1&
    Public Const GENERIC_READ As Long = &H80000000
    Public Const FILE_SHARE_READ As Long = &H1
    Public Const FILE_SHARE_WRITE As Long = &H2
    Public Const OPEN_EXISTING As Long = 3

    Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    'Local Property =======================================
    Dim mFormName As System.Windows.Forms.Form
    Dim mNetPath As String = ""
    Dim mDrive As String = ""
    Dim mUser As String = ""
    Dim mPass As String = ""
    Dim mErrorComment As String = ""
    Dim mErrorCode As String = ""

    'Property =====================================================================================
    'Form name ------------------------
    Public WriteOnly Property FORM_NAME() As System.Windows.Forms.Form
        Set(ByVal Value As System.Windows.Forms.Form)
            mFormName = Value
        End Set
    End Property

    'NetPath --------------------------
    Public WriteOnly Property NET_PATH() As String
        Set(ByVal Value As String)
            mNetPath = Value
        End Set
    End Property

    'Drive ----------------------------
    Public WriteOnly Property DRIVE() As String
        Set(ByVal Value As String)
            mDrive = Value
        End Set
    End Property

    'User -----------------------------
    Public WriteOnly Property USER() As String
        Set(ByVal Value As String)
            mUser = Value
        End Set
    End Property

    'Pass -----------------------------
    Public WriteOnly Property PASS() As String
        Set(ByVal Value As String)
            mPass = Value
        End Set
    End Property

    'Error Message --------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public ReadOnly Property ERR_CODE() As String
        Get
            ERR_CODE = mErrorCode
        End Get
    End Property


    'Process ======================================================================================
    Public Sub DRIVE_INFPRMATION(ByRef stDrives As String())
        '==========================================================================================
        ' Function : Use drive in my PC
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim fs As New Scripting.FileSystemObject

        Dim i As Integer = 0
        Dim drvType As String = ""
        Dim volName As String = ""
        Dim sDrvs() As String = {}

        mErrorComment = ""

        Try
            sDrvs = System.IO.Directory.GetLogicalDrives()

            For Each stDrive In sDrvs

                'Drive type ===============================
                Select Case GetDriveType(stDrive)
                    Case 0 : drvType = "Unknown"
                    Case 1 : drvType = "Nothing"
                    Case 2 : drvType = "RemovableDisk"
                    Case 3 : drvType = "HardDisk"
                    Case 4 : drvType = "RemoteDisk"
                    Case 5 : drvType = "CD-ROM"
                    Case 6 : drvType = "RAMDisk"
                End Select
                sDrvs(i) = sDrvs(i) & "-" & drvType

                'Volume name ==============================
                volName = fs.GetDrive(fs.GetDriveName(Left(stDrive, 1) & ":")).ShareName
                If volName <> "" Then sDrvs(i) = sDrvs(i) & "-" & volName
                i = i + 1

            Next

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        Finally
            stDrives = DirectCast(sDrvs.Clone(), String())
            fs = Nothing

        End Try


    End Sub

    Public Sub NETWORK_DRIVE_CONNECT()
        '==========================================================================================
        ' Function : Network drive connect
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim ret As Integer
        Dim strUser As String = mUser & ControlChars.NullChar
        Dim strPass As String = mPass & ControlChars.NullChar
        Dim mHwnd As Integer = mFormName.Handle.ToInt32
        Dim uNetResource As New mNetResource

        mErrorComment = ""

        Try
            'Setting ---------------------------
            uNetResource.dwType = RESURCETYPE_DISK
            'Drive Sign
            uNetResource.lpLocalName = UCase(mDrive) & ":" & ControlChars.NullChar
            'Net Path
            uNetResource.lpRemoteName = mNetPath & ControlChars.NullChar
            uNetResource.lpProvider = ControlChars.NullChar

            'Exec -------------------------------
            ret = WNetAddConnection2(uNetResource, strPass, strUser, 0)

            If ret <> 0 Then
                Call DriveConnectCode(ret)
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub NETWORK_DRIVE_DISCONNECT()
        '==========================================================================================
        ' Function : Network drive disconnect
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim strDrive As String = mDrive & ":" & vbNullChar
        Dim ret As Integer

        mErrorComment = ""

        Try
            'Exec -------------------------------
            ret = WNetCancelConnection2(strDrive, 0, True)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

        If ret <> 0 Then
            Call DriveConnectCode(ret)
            Throw New ApplicationException(mErrorComment)
        End If

    End Sub

    Private Sub DriveConnectCode(ByVal ret As Integer)
        '==========================================================================================
        ' Function : Error code of network drive connect 
        ' Data     : ret
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Select Case ret
            Case 5 : mErrorComment = "ERROR:" & CStr(ret) & ":ACCESS_DENIED"
            Case 85 : mErrorComment = "ERROR:" & CStr(ret) & ":ALREADY_ASSIGNED"
            Case 86 : mErrorComment = "ERROR:" & CStr(ret) & ":INVALID_PASSWORD"
            Case 53 : mErrorComment = "ERROR:" & CStr(ret) & ":BAD_NETPATH"
            Case 66 : mErrorComment = "ERROR:" & CStr(ret) & ":BAD_DEV_TYPE"
            Case 67 : mErrorComment = "ERROR:" & CStr(ret) & ":BAD_NET_NAME"
            Case 170 : mErrorComment = "ERROR:" & CStr(ret) & ":BUSY"
            Case 487 : mErrorComment = "ERROR:" & CStr(ret) & ":INVALID_ADDRESS"
            Case 1200 : mErrorComment = "ERROR:" & CStr(ret) & ":BAD_DEVICE"
            Case 1202 : mErrorComment = "ERROR:" & CStr(ret) & ":DEVICE_ALREADY_REMEMBERED"
            Case 1203 : mErrorComment = "ERROR:" & CStr(ret) & ":NO_NET_OR_BAD_PATH"
            Case 1204 : mErrorComment = "ERROR:" & CStr(ret) & ":BAD_PROVIDER"
            Case 1205 : mErrorComment = "ERROR:" & CStr(ret) & ":CANNOT_OPEN_PROFILE"
            Case 1206 : mErrorComment = "ERROR:" & CStr(ret) & ":BAD_PROFILE"
            Case 1208 : mErrorComment = "ERROR:" & CStr(ret) & ":EXTENDED_ERROR"
            Case 1219 : mErrorComment = "ERROR:" & CStr(ret) & ":CONNECTED_ERROR"
            Case 1222 : mErrorComment = "ERROR:" & CStr(ret) & ":NO_NETWORK"
            Case 1223 : mErrorComment = "ERROR:" & CStr(ret) & ":CANCELLED"
            Case 2250 : mErrorComment = "ERROR:" & CStr(ret) & ":NOT_CONNECTED"
            Case 2404 : mErrorComment = "ERROR:" & CStr(ret) & ":DEVICE_IN_USE"
        End Select

    End Sub

    Private Sub EjectDrive(ByVal ctrlCode As Long)

        Dim hDevice As Long
        Dim bytesReturned As Long
        Dim hRtn As Long

        'ﾊﾝﾄﾞﾙ取得
        hDevice = CreateFile("\\.\" & mDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                              IntPtr.Zero, OPEN_EXISTING, 0&, 0&)

        'ｲｼﾞｪｸﾄ
        If hDevice <> INVALID_HANDLE_VALUE Then
            hRtn = DeviceIoControl(hDevice, ctrlCode, IntPtr.Zero, 0&, IntPtr.Zero, 0&, bytesReturned, IntPtr.Zero)
            If hRtn = 0 Then
                MsgBox("ドライブ取り外し　失敗")
                Throw New ApplicationException(mErrorComment)
            End If
        End If

        'ﾊﾝﾄﾞﾙ開放
        Call CloseHandle(hDevice)

    End Sub

End Class
