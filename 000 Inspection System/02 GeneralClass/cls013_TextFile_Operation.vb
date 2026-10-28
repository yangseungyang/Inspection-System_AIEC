Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls013_TextFile_Operation
'Function   : Text file operation
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'
'Issue      : 2008/09/30  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

Public Class cls013_TextFile_Operation

    'Text File ============================================
    Dim w As System.IO.StreamWriter
    Dim r As System.IO.StreamReader

    'Local ================================================
    Dim sJisEnc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

    'Local Property =======================================
    Dim mFilePath As String = ""
    Dim mErrorComment As String = ""
    Dim mErrorCode As String = ""

    'Property =====================================================================================
    'db path --------------------------------------------------------
    Public WriteOnly Property FILE_PATH() As String
        Set(ByVal Value As String)
            mFilePath = Value
        End Set
    End Property

    'Error message --------------------------------------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_Comment = mErrorComment
        End Get
    End Property

    Public ReadOnly Property ERR_CODE() As String
        Get
            ERR_CODE = mErrorCode
        End Get
    End Property


    'Process ======================================================================================
    Public Sub READ_TEXT_FILE(ByRef msg As String)
        '==========================================================================================
        ' Function : Read of message : One parameter
        ' Data     : Message(Return)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""
        msg = ""

        Try
            r = New System.IO.StreamReader(mFilePath, sJisEnc)
            msg = r.ReadToEnd

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        Finally
            r.Close()

        End Try

    End Sub

    Public Sub READ_TEXT_ARRAY(ByRef msg() As String)
        '==========================================================================================
        ' Function : Read of message : One line
        ' Data     : Message(Return)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim line As String = ""
        Dim al As New ArrayList

        mErrorComment = ""
        ReDim msg(0)

        Try
            r = New System.IO.StreamReader(mFilePath, sJisEnc)

            line = r.ReadLine()
            Do Until line Is Nothing
                al.Add(line)
                line = r.ReadLine()
            Loop

            ReDim msg(al.Count - 1)
            For i As Integer = 0 To al.Count - 1
                msg(i) = al.Item(i).ToString
            Next i

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        Finally
            r.Close()

        End Try

    End Sub

    Public Sub WRITE_TEXT_FILE(ByVal tStmp As Boolean, ByVal msg As String)
        '==========================================================================================
        ' Function : Write of message
        ' Data     : Time stamp Y/N , Message
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Dim tNow As String = ""
        If tStmp = True Then
            tNow = Format(DateTime.Now, "HH:mm:ss") & " "
        Else
            tNow = ""
        End If

        Try
            w = New System.IO.StreamWriter(mFilePath, True, sJisEnc)
            w.WriteLine(tNow & msg)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        Finally
            w.Close()

        End Try

    End Sub

    Public Sub WRITE_TEXT_LINE(ByVal tStmp As Boolean, ByVal sin As String, ByVal cnt As Integer)
        '==========================================================================================
        ' Function : Write of line
        ' Data     : Time stamp Y/N , Message
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Dim lSin As String = ""
        For i As Integer = 0 To cnt - 1
            lSin &= sin
        Next

        Try
            w = New System.IO.StreamWriter(mFilePath, True, sJisEnc)
            w.WriteLine(lSin)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        Finally
            w.Close()

        End Try

    End Sub

End Class
