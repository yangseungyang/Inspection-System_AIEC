Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls090_DesignForm_Type01
'Function   : Design Form Type 01
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'
'Issue      : 2008/09/30  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

Public Class cls090_DesignForm_Type01

    'Local Property =======================================
    Dim mFormName As Form
    Dim mPicFolder As String = ""
    Dim mTitleBar_L As Control
    Dim mTitleBar_R As Control
    Dim mUnderBar_L As Control
    Dim mUnderBar_R As Control
    Dim mMinButton As Button
    Dim mCloseButton As Button
    Dim mPanelTitle As Label
    Dim mErrorComment As String = ""
    Dim mErrorCode As String = ""

    'Property =====================================================================================
    'Form Name ------------------------
    Public WriteOnly Property FORM_NAME() As Form
        Set(ByVal Value As Form)
            mFormName = Value
        End Set
    End Property

    'Pic Folder -----------------------
    Public WriteOnly Property PIC_FOLDER() As String
        Set(ByVal Value As String)
            mPicFolder = Value
        End Set
    End Property

    'Form Control ---------------------
    Public WriteOnly Property TITLE_BAR_L() As PictureBox
        Set(ByVal Value As PictureBox)
            mTitleBar_L = Value
        End Set
    End Property

    Public WriteOnly Property TITLE_BAR_R() As PictureBox
        Set(ByVal Value As PictureBox)
            mTitleBar_R = Value
        End Set
    End Property

    Public WriteOnly Property UNDER_BAR_L() As PictureBox
        Set(ByVal Value As PictureBox)
            mUnderBar_L = Value
        End Set
    End Property

    Public WriteOnly Property UNDER_BAR_R() As PictureBox
        Set(ByVal Value As PictureBox)
            mUnderBar_R = Value
        End Set
    End Property

    Public WriteOnly Property MIN_BUTTON() As Button
        Set(ByVal Value As Button)
            mMinButton = Value
        End Set
    End Property

    Public WriteOnly Property CLOSE_BUTTON() As Button
        Set(ByVal Value As Button)
            mCloseButton = Value
        End Set
    End Property

    Public WriteOnly Property PANEL_TITLE() As Label
        Set(ByVal Value As Label)
            mPanelTitle = Value
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
    Public Sub FORM_DRAW()
        '==============================================================================================================
        ' Function : Form Draw
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        mErrorComment = ""

        Try
            'Form =================================================================================
            With mFormName
                .BackgroundImage = System.Drawing.Image.FromFile(mPicFolder & "BackColor.bmp")
                .BackgroundImageLayout = ImageLayout.Stretch
            End With

            With mTitleBar_R
                .Width = 288 : .Height = 35 : .Top = 0 : .Left = mFormName.Width - .Width
                .SendToBack()

                .BackgroundImage = System.Drawing.Image.FromFile(mPicFolder & "Title_R.bmp")
                .BackgroundImageLayout = ImageLayout.None
            End With

            'Title bar ============================================================================
            With mTitleBar_L
                .Width = mFormName.Width - mTitleBar_R.Width : .Height = 35 : .Top = 0 : .Left = 0
                .SendToBack()

                .BackgroundImage = System.Drawing.Image.FromFile(mPicFolder & "Title_L.bmp")
                .BackgroundImageLayout = ImageLayout.None
            End With

            With mUnderBar_R
                .Width = 288 : .Height = 30 : .Top = mFormName.Height - .Height : .Left = mFormName.Width - .Width
                .SendToBack()

                .BackgroundImage = System.Drawing.Image.FromFile(mPicFolder & "Under_R.bmp")
                .BackgroundImageLayout = ImageLayout.None
            End With

            With mUnderBar_L
                .Width = mFormName.Width - mUnderBar_R.Width : .Height = 30 : .Top = mFormName.Height - .Height : .Left = 0
                .SendToBack()

                .BackgroundImage = System.Drawing.Image.FromFile(mPicFolder & "Under_L.bmp")
                .BackgroundImageLayout = ImageLayout.None
            End With

            'Button ===============================================================================
            If Not mMinButton Is Nothing Then
                With mMinButton
                    .Top = 4 : .Left = mFormName.Width - 58 : .Width = 19 : .Height = 19
                    .FlatStyle = FlatStyle.Standard
                    .Text = ""
                    .Cursor = Cursors.Hand

                    .BackgroundImage = System.Drawing.Image.FromFile(mPicFolder & "Minimum.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            End If

            If Not mCloseButton Is Nothing Then
                With mCloseButton
                    .Top = 4 : .Left = mFormName.Width - 22 : .Width = 19 : .Height = 19
                    .FlatStyle = FlatStyle.Standard
                    .Text = ""
                    .Cursor = Cursors.Hand

                    .BackgroundImage = System.Drawing.Image.FromFile(mPicFolder & "Close.bmp")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            End If

            'Label ================================================================================
            If Not mPanelTitle Is Nothing Then
                With mPanelTitle
                    .AutoSize = True
                    .Image = System.Drawing.Image.FromFile(mPicFolder & "Title.bmp")
                    .BackgroundImageLayout = ImageLayout.None
                End With
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub TEXT_DRAW(ByVal tControl As Control, ByVal tText As String, ByVal tFont As String, ByVal tSize As Integer, ByVal x As Integer, ByVal y As Integer, ByVal tColor As Color)
        '==============================================================================================================
        'Function : Text write
        'Argument : Control , Text , font , size , x , y , color
        'Return   : Nothing
        'Remarks  :
        '==============================================================================================================

        mErrorComment = ""

        Try
            With tControl
                .Text = tText
                .Font = New System.Drawing.Font(tFont, tSize, System.Drawing.FontStyle.Regular)
                .Top = y : .Left = x
                .ForeColor = tColor
            End With

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

End Class
