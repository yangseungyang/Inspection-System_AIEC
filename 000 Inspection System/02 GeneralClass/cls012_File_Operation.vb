Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls012_File_Operation
'Function   : Folder & File operation
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'
'Issue      : 2008/12/26  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

Public Class cls012_File_Operation

    'Local ================================================
    Dim mSearchFolder As String
    Dim mSearchFile As String

    'Local Property =======================================
    Dim mDispLabel As Label
    Dim mRootFolderPath As String = ""
    Dim mFolderName As String = ""
    Dim mFolderPath As String = ""
    Dim mOriginalPath As String = ""
    Dim mFileName As String = ""
    Dim mFilePath As String = ""
    Dim mErrorComment As String = ""
    Dim mErrorCode As String = ""

    'Property =====================================================================================
    'Display Label --------------------------------------------------
    Public WriteOnly Property DISPLAY_LABEL() As Label
        Set(ByVal Value As Label)
            mDispLabel = Value
        End Set
    End Property

    'Root Folder ----------------------------------------------------
    Public WriteOnly Property ROOT_FOLDER() As String
        Set(ByVal Value As String)
            mRootFolderPath = Value
        End Set
    End Property

    'Folder Name ----------------------------------------------------
    Public WriteOnly Property FOLDER_NAME() As String
        Set(ByVal Value As String)
            mFolderName = Value
        End Set
    End Property

    'Folder Path ----------------------------------------------------
    Public WriteOnly Property FOLDER_PATH() As String
        Set(ByVal Value As String)
            mFolderPath = Value
        End Set
    End Property

    'Original Path --------------------------------------------------
    Public WriteOnly Property ORIGINAL_PATH() As String
        Set(ByVal Value As String)
            mOriginalPath = Value
        End Set
    End Property

    'File Name ------------------------------------------------------
    Public WriteOnly Property FILE_NAME() As String
        Set(ByVal Value As String)
            mFileName = Value
        End Set
    End Property

    'File Path ------------------------------------------------------
    Public WriteOnly Property FILE_PATH() As String
        Set(ByVal Value As String)
            mFilePath = Value
        End Set
    End Property

    'Error message --------------------------------------------------
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
    Public Sub FOLDER_SEARCH(ByRef Result() As String)
        '==========================================================================================
        ' Function : Folder Search
        ' Data     : Result(Return)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Dim sData() As String = {}

        Call sFolderSearch(mRootFolderPath, mRootFolderPath, mFolderName, sData)

        Result = DirectCast(sData.Clone(), String())

        If mErrorComment <> "" Then Throw New ApplicationException(mErrorComment)

    End Sub

    Private Sub sFolderSearch(ByVal RootFolderName As String, ByVal SearchFolderName As String, ByVal FolderName As String, ByRef Result() As String)
        '==========================================================================================
        ' Function : Folder Search (sub routine)
        ' Data     : Root folder , Search folder , Target folder , Result(Return)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim i As Integer

        Dim sDir As String() = {}
        Dim nDir As String

        Try
            'Display of search folder ===================================
            If Not mDispLabel Is Nothing Then
                mDispLabel.Text = SearchFolderName
                System.Windows.Forms.Application.DoEvents()
            End If
            Debug.Print(SearchFolderName)

            'Search in SearchFolderName =================================
            sDir = System.IO.Directory.GetDirectories(SearchFolderName, FolderName)
            For i = 0 To sDir.Length - 1
                mSearchFolder &= sDir(i) & ","
            Next

            'Next folder (recursion) ====================================
            For Each nDir In System.IO.Directory.GetDirectories(SearchFolderName)
                Call sFolderSearch(RootFolderName, nDir, FolderName, Result)
            Next

            'Result =====================================================
            If mSearchFolder <> "" And RootFolderName = SearchFolderName Then
                mSearchFolder = Left(mSearchFolder, Len(mSearchFolder) - 1)
                Result = Split(mSearchFolder, ",")
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

        End Try

    End Sub

    Public Sub FOLDER_MAKE()
        '==========================================================================================
        ' Function : Folder making
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try
            System.IO.Directory.CreateDirectory(mFolderPath)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub FOLDER_COPY()
        '==========================================================================================
        ' Function : Folder copy 
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Dim sData() As String = {}

        Call sFolder_Copy(mOriginalPath, mFolderPath)

        If mErrorComment <> "" Then Throw New ApplicationException(mErrorComment)

    End Sub

    Private Sub sFolder_Copy(ByVal OrginalPath As String, ByVal CopyPath As String)
        '==========================================================================================
        ' Function : Folder copy (sub routine)
        ' Data     : Origin folder path , Copy folder path
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim eCount As Integer = 0
        Dim eCode As Integer = 0

        Do
            Try

                mErrorComment = ""

                'It makes it when there is no directory =====================
                If Not System.IO.Directory.Exists(CopyPath) Then
                    System.IO.Directory.CreateDirectory(CopyPath)
                    System.IO.File.SetAttributes(CopyPath, System.IO.File.GetAttributes(CopyPath))
                End If

                '"\" is applied to the end of the directory =================
                If CopyPath.Chars((CopyPath.Length - 1)) <> System.IO.Path.DirectorySeparatorChar Then
                    CopyPath = CopyPath & System.IO.Path.DirectorySeparatorChar
                End If

                'File Copy ==================================================
                Dim fs As String() = System.IO.Directory.GetFiles(OrginalPath)
                Dim f As String
                For Each f In fs
                    System.IO.File.Copy(f, CopyPath + System.IO.Path.GetFileName(f), True)
                Next

                'Copy of next directory (recursion) =========================
                Dim dirs As String() = System.IO.Directory.GetDirectories(OrginalPath)
                Dim dir As String
                For Each dir In dirs
                    sFolder_Copy(dir, CopyPath & System.IO.Path.GetFileName(dir))
                Next

                mErrorComment = ""

            Catch ex As Exception
                'NG =======================================
                eCode = Err.Number

                If eCount >= 4 Or Err.Number <> 57 Then
                    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                    mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
                    Debug.Print(mErrorComment)

                    'NG End ---------------------
                    Throw New ApplicationException(mErrorComment)

                End If

                'File Open -> Retry =======================
                If Err.Number = 57 Then
                    Call WaitTimer(0.5)
                    eCount += 1
                    mErrorComment = "ERROR"
                End If

            End Try

            If mErrorComment = "" Then Exit Do

        Loop

    End Sub

    Public Sub FOLDER_DELETE()
        '==========================================================================================
        ' Function : Folder delete
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim eCount As Integer = 0
        Dim eCode As Integer = 0

        Do

            Try
                'Folder delete ============================
                If IO.Directory.Exists(mFolderName) = True Then
                    System.IO.Directory.Delete(mFolderPath, True)
                Else
                End If
                mErrorComment = ""

            Catch ex As Exception
                'NG =======================================
                eCode = Err.Number

                If eCount >= 4 Or Err.Number <> 57 Then
                    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                    mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
                    Debug.Print(mErrorComment)

                    'NG End ---------------------
                    Throw New ApplicationException(mErrorComment)

                End If

                'File Open -> Retry =======================
                If Err.Number = 57 Then
                    Call WaitTimer(0.5)
                    eCount += 1
                    mErrorComment = "ERROR"
                End If

            End Try

            If mErrorComment = "" Then Exit Do

        Loop

    End Sub

    Public Sub FOLDER_MOVE()
        '==========================================================================================
        ' Function : Folder move
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim eCount As Integer = 0
        Dim eCode As Integer = 0

        Do

            Try
                'Folder Move ============================
                System.IO.Directory.Move(mOriginalPath, mFolderPath)
                mErrorComment = ""

            Catch ex As Exception
                'NG =======================================
                eCode = Err.Number

                If eCount >= 4 Or Err.Number <> 57 Then
                    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                    mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
                    Debug.Print(mErrorComment)

                    'NG End ---------------------
                    Throw New ApplicationException(mErrorComment)

                End If

                'File Open -> Retry =======================
                If Err.Number = 57 Then
                    Call WaitTimer(0.5)
                    eCount += 1
                    mErrorComment = "ERROR"
                End If

            End Try

            If mErrorComment = "" Then Exit Do

        Loop

    End Sub

    Public Sub FILE_SEARCH(ByRef Result() As String)
        '==========================================================================================
        ' Function : File Search
        ' Data     : result(Return)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Dim sData() As String = {}

        Call sFileSearch(mRootFolderPath, mRootFolderPath, mFileName, sData)

        Result = DirectCast(sData.Clone(), String())

        If mErrorComment <> "" Then Throw New ApplicationException(mErrorComment)

    End Sub

    Private Sub sFileSearch(ByVal RootFolderName As String, ByVal SearchFolderName As String, ByVal FileName As String, ByRef Result() As String)
        '==========================================================================================
        ' Function : File Search (sub routine)
        ' Data     : Root folder , Search folder , Target file , result(Return)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim i As Integer
        Dim sFile As String() = {}

        Dim sDir As String() = {}
        Dim nDir As String

        Try
            'Display of search folder ===================================
            If Not mDispLabel Is Nothing Then
                mDispLabel.Text = SearchFolderName
                System.Windows.Forms.Application.DoEvents()
            End If
            Debug.Print(SearchFolderName)

            'Search in SearchFolderName =================================
            sFile = System.IO.Directory.GetFiles(SearchFolderName, FileName)
            For i = 0 To sFile.Length - 1
                mSearchFile &= sFile(i) & ","
            Next

            'Next folder (recursion) ====================================
            For Each nDir In System.IO.Directory.GetDirectories(SearchFolderName)
                Call sFileSearch(RootFolderName, nDir, FileName, Result)
            Next

            'Result =====================================================
            If mSearchFile <> "" And RootFolderName = SearchFolderName Then
                mSearchFile = Left(mSearchFile, Len(mSearchFile) - 1)
                Result = Split(mSearchFile, ",")
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

        End Try

    End Sub

    Public Sub FILE_COPY()
        '==========================================================================================
        ' Function : File copy
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim eCount As Integer = 0
        Dim eCode As Integer = 0

        Do

            Try
                'Make folder ======================================
                System.IO.Directory.CreateDirectory(mFolderPath)
                mErrorComment = ""

                'File copy ======================================== 
                System.IO.File.Copy(mOriginalPath, mFolderPath & mFileName, True)
                mErrorComment = ""

            Catch ex As Exception
                'NG =======================================
                eCode = Err.Number

                If eCount >= 4 Or Err.Number <> 55 Then
                    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                    mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
                    Debug.Print(mErrorComment)

                    'NG End ---------------------
                    Throw New ApplicationException(mErrorComment)

                End If

                'File Open -> Retry =======================
                If Err.Number = 55 Then
                    Call WaitTimer(0.5)
                    eCount += 1
                    mErrorComment = "ERROR"
                End If

            End Try

            If mErrorComment = "" Then Exit Do

        Loop

    End Sub

    Public Sub FILE_DELETE()
        '==========================================================================================
        ' Function : File delete
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim eCount As Integer = 0
        Dim eCode As Integer = 0

        Do

            Try
                'File delete ==============================
                System.IO.File.Delete(mFilePath)
                mErrorComment = ""

            Catch ex As Exception
                'NG =======================================
                eCode = Err.Number

                If eCount >= 4 Or Err.Number <> 57 Then
                    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                    mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
                    Debug.Print(mErrorComment)

                    'NG End ---------------------
                    Throw New ApplicationException(mErrorComment)

                End If

                'File Open -> Retry =======================
                If Err.Number = 57 Then
                    Call WaitTimer(0.5)
                    eCount += 1
                    mErrorComment = "ERROR"
                End If

            End Try

            If mErrorComment = "" Then Exit Do

        Loop

    End Sub

    Public Sub FILE_MOVE()
        '==========================================================================================
        ' Function : File move
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim eCount As Integer = 0
        Dim eCode As Integer = 0

        Do

            Try
                'File Move ================================
                System.IO.File.Move(mOriginalPath, mFilePath)
                mErrorComment = ""

            Catch ex As Exception
                'NG =======================================
                eCode = Err.Number

                If eCount >= 4 Or Err.Number <> 57 Then
                    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                    mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
                    Debug.Print(mErrorComment)

                    'NG End ---------------------
                    Throw New ApplicationException(mErrorComment)

                End If

                'File Open -> Retry =======================
                If Err.Number = 57 Then
                    Call WaitTimer(0.5)
                    eCount += 1
                    mErrorComment = "ERROR"
                End If

            End Try

            If mErrorComment = "" Then Exit Do

        Loop

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
