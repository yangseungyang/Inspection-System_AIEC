Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls001_Window_Control
'Function   : Control of window position etc.
'Language   : VB2008Å@Framework3.5
'Reference  : 
'Memo       :
'
'Issue      : 2008/12/26  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

Public Class cls001_Window_Control

    'Process control ======================================
    Private Structure PROCESS_INFORMATION
        Dim hProcess As Integer
        Dim hThread As Integer
        Dim dwProcessId As Integer
        Dim dwThreadId As Integer
    End Structure

    Private Structure STARTUPINFO
        Dim cb As Integer
        Dim lpReserved As String
        Dim lpDesktop As String
        Dim lpTitle As String
        Dim dwX As Integer
        Dim dwY As Integer
        Dim dwXSize As Integer
        Dim dwYSize As Integer
        Dim dwXCountChars As Integer
        Dim dwYCountChars As Integer
        Dim dwFillAttribute As Integer
        Dim dwFlags As Integer
        Dim wShowWindow As Integer
        Dim cbReserved2 As Integer
        Dim lpReserved2 As Integer
        Dim hStdInput As Integer
        Dim hStdOutput As Integer
        Dim hStdError As Integer
    End Structure

    'dwFlags ----------------------------------------------
    Private Const STARTF_USEPOSITION As Integer = &H4
    Private Const STARTF_USESIZE As Integer = &H2
    Private Const STARTF_USECOUNTCHARS As Integer = &H8
    Private Const STARTF_FILLATTRIBUTE As Integer = &H10
    Private Const STARTF_USESHOWWINDOW As Integer = &H1
    Private Const STARTF_USESTDHANDLES As Integer = &H100
    Private Const STARTF_FORCEONFEEDBACK As Integer = &H40
    Private Const STARTF_FORCEOFFFEEDBACK As Integer = &H80

    Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Integer, _
    ByVal lpThreadAttributes As Integer, _
    ByVal bInheritHandles As Integer, _
    ByVal dwCreationFlags As Integer, _
    ByVal lpEnvironment As Integer, _
    ByVal lpCurrentDriectory As String, _
    ByRef lpStartupInfo As STARTUPINFO, _
    ByRef lpProcessInformation As PROCESS_INFORMATION _
    ) As Integer

    'dwCreationFlags --------------------------------------
    Private Const NORMAL_PRIORITY_CLASS As Integer = &H20
    Private Const CREATE_DEFAULT_ERROR_MODE As Integer = &H400000
    Private Const CREATE_NEW_CONSOLE As Integer = &H10
    Private Const CREATE_NEW_PROCESS_GROUP As Integer = &H200
    Private Const CRAETE_SUSPENED As Integer = &H4
    Private Const CRAETE_UNICODE_ENVIRONMENT As Integer = &H400
    Private Const DEBUG_PROCESS As Integer = &H1
    Private Const DEBUG_ONLY_THIS_PROCESS As Integer = &H2
    Private Const DETACHED_PROCESS As Integer = &H8

    Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwAccess As Integer, _
    ByVal fInherit As Integer, _
    ByVal hObject As Integer _
    ) As Integer

    Private Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Integer, _
    ByVal uExitCode As Integer _
    ) As Integer

    Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Integer _
    ) As Integer

    'Get Window Position ==================================
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer

    Structure RECT
        Dim left As Integer
        Dim top As Integer
        Dim right As Integer
        Dim bottom As Integer
    End Structure

    'Set Window Position Size =============================
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    'Z-Position
    Private Const HWND_BOTTOM As Integer = 1
    Private Const HWND_TOP As Integer = 0
    Private Const HWND_TOPMOST As Integer = -1
    Private Const HWND_NOTOPMOST As Integer = -2
    'Size and Position
    Private Const SWP_SHOWWINDOW As Integer = &H40
    Private Const SWP_NOMOVE As Integer = &H1
    Private Const SWP_NOSIZE As Integer = &H2

    'Find Window ==========================================
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private lHndl As Object
    Private lRtn As Integer
    Private Const lpClass As String = vbNullString
    Private SendProjNm As String

    'Message to Window ====================================
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Const WM_QUIT As Integer = &H12

    'Local ================================================
    Dim mProcessInfo As PROCESS_INFORMATION
    Dim mStartInfo As STARTUPINFO
    Dim sNull As String = vbNullString

    'Local Property =======================================
    Dim mProcessName As String = ""
    Dim mWindowTitle As String = ""
    Dim mHwnd As Integer = 0
    Dim mWinZ As Integer = 0
    Dim mWinX As Integer = 0
    Dim mWinY As Integer = 0
    Dim mWinW As Integer = 0
    Dim mWinH As Integer = 0
    Dim mHProcess As Integer = 0
    Dim mhThread As Integer = 0
    Dim mErrorComment As String = ""
    Dim mErrorCode As String = ""

    'Property =====================================================================================
    'Process name ---------------------
    Public WriteOnly Property PROCESS_NAME() As String
        Set(ByVal Value As String)
            mProcessName = Value
        End Set
    End Property

    'Window Title ---------------------
    Public WriteOnly Property WINDOW_TITLE() As String
        Set(ByVal Value As String)
            mWindowTitle = Value
        End Set
    End Property

    'Window handle --------------------
    Public ReadOnly Property WIN_HNDL() As Integer
        Get
            WIN_HNDL = mHwnd
        End Get
    End Property

    'Setting Position Z ---------------
    Public WriteOnly Property WIN_Z() As Integer
        Set(ByVal Value As Integer)
            mWinZ = Value
        End Set
    End Property

    'Setting Position X ---------------
    Public Property WIN_X() As Integer
        Get
            WIN_X = mWinX
        End Get
        Set(ByVal Value As Integer)
            mWinX = Value
        End Set
    End Property

    'Setting Position Y ---------------
    Public Property WIN_Y() As Integer
        Get
            WIN_Y = mWinY
        End Get
        Set(ByVal Value As Integer)
            mWinY = Value
        End Set
    End Property

    'Setting Position W ---------------
    Public Property WIN_W() As Integer
        Get
            WIN_W = mWinW
        End Get
        Set(ByVal Value As Integer)
            mWinW = Value
        End Set
    End Property

    'Setting Position H ---------------
    Public Property WIN_H() As Integer
        Get
            WIN_H = mWinH
        End Get
        Set(ByVal Value As Integer)
            mWinH = Value
        End Set
    End Property

    'H_PROCESS ------------------------
    Public Property H_PROCESS() As Integer
        Get
            H_PROCESS = mHProcess
        End Get
        Set(ByVal Value As Integer)
            mHProcess = Value
        End Set
    End Property

    'H_THREAD -------------------------
    Public Property H_THREAD() As Integer
        Get
            H_THREAD = mhThread
        End Get
        Set(ByVal Value As Integer)
            mhThread = Value
        End Set
    End Property

    'Z Position -----------------------
    Public ReadOnly Property Z_TOPMOST() As Integer
        Get
            Z_TOPMOST = HWND_TOPMOST
        End Get
    End Property

    'Z Position -----------------------
    Public ReadOnly Property Z_BOTTOM() As Integer
        Get
            Z_BOTTOM = HWND_BOTTOM
        End Get
    End Property

    'Z Position -----------------------
    Public ReadOnly Property Z_TOP() As Integer
        Get
            Z_TOP = HWND_TOP
        End Get
    End Property

    'Z Position -----------------------
    Public ReadOnly Property Z_NOTOPMOST() As Integer
        Get
            Z_NOTOPMOST = HWND_NOTOPMOST
        End Get
    End Property

    'Error message --------------------
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
    Public Sub CREATE_PROCESS()
        '==========================================================================================
        ' Function : Cleate process
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim ret As Integer

        mErrorComment = ""

        mStartInfo.cb = System.Runtime.InteropServices.Marshal.SizeOf(mStartInfo)
        mStartInfo.dwFlags = STARTF_USEPOSITION Or STARTF_USESIZE
        mStartInfo.dwX = mWinX
        mStartInfo.dwY = mWinY
        mStartInfo.dwXSize = mWinW
        mStartInfo.dwYSize = mWinH

        Try
            ret = CreateProcess(sNull, mProcessName, 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, sNull, mStartInfo, mProcessInfo)

            mHProcess = mProcessInfo.hProcess
            mhThread = mProcessInfo.hThread

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub END_PROCESS()
        '==========================================================================================
        ' Function : End process
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim ret As Integer

        mErrorComment = ""

        Try
            mProcessInfo.hProcess = mHProcess
            mProcessInfo.hThread = mhThread

            ret = TerminateProcess(mProcessInfo.hProcess, 0)
            ret = CloseHandle(mProcessInfo.hThread)
            ret = CloseHandle(mProcessInfo.hProcess)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub GET_WINDOW_HNDL()
        '==========================================================================================
        ' Function : Get of window Handle
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim mW_Title As String = mWindowTitle

        mErrorComment = ""
        mHwnd = 0

        If mWindowTitle = "" Then
            mErrorComment = "The title is not set"
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End If

        Try
            mHwnd = FindWindow(vbNullString, mW_Title)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub GET_WINDOW_POSITION()
        '==========================================================================================
        ' Function : Get of window position & size
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim mW_Title As String = mWindowTitle
        Dim mGetPos As Integer
        Dim iprect As RECT

        mErrorComment = ""

        If mWindowTitle = "" Then
            mErrorComment = "The title is not set"
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End If

        Try
            'Find Window ================================================
            mHwnd = FindWindow(vbNullString, mW_Title)

            'Set Position ===============================================
            If mHwnd <> 0 Then
                mGetPos = GetWindowRect(mHwnd, iprect)

            Else
                mErrorComment = "Window not found"
                Debug.Print(mErrorComment)

                Throw New ApplicationException(mErrorComment)

            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

        mWinX = iprect.left
        mWinY = iprect.top
        mWinW = iprect.right - iprect.left
        mWinH = iprect.bottom - iprect.top

    End Sub

    Public Sub SET_WINDOW_POSITION()
        '==========================================================================================
        ' Function : Set window position & size
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim mW_Title As String = mWindowTitle
        Dim mSetPos As Integer

        mErrorComment = ""

        If mWindowTitle = "" Then
            mErrorComment = "The title is not set"
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End If

        Try
            'Find Window ================================================
            mHwnd = FindWindow(vbNullString, mW_Title)

            'Set Position ===============================================
            If mHwnd <> 0 Then
                mSetPos = SetWindowPos(mHwnd, mWinZ, mWinX, mWinY, mWinW, mWinH, SWP_SHOWWINDOW Or 0 Or 0)
            Else
                mErrorComment = "Window not found"
                Debug.Print(mErrorComment)

                Throw New ApplicationException(mErrorComment)

            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub WINDOW_QUIT()
        '==========================================================================================
        ' Function : Window Quit
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim mW_Title As String = mWindowTitle
        Dim mQuit As Integer

        mErrorComment = ""

        If mWindowTitle = "" Then
            mErrorComment = "The title is not set"
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End If

        Try
            'Find Window ================================================
            mHwnd = FindWindow(vbNullString, mW_Title)

            'Send Message ===============================================
            If mHwnd <> 0 Then
                mQuit = PostMessage(mHwnd, WM_QUIT, 0, 0)

            Else
                mErrorComment = "Window not found"
                Debug.Print(mErrorComment)

                Throw New ApplicationException(mErrorComment)

            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

End Class