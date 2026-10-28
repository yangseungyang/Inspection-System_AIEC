Option Strict On : Option Explicit On


'Doevents
Imports System.Windows.Forms.Application

Public Class cls130_EI_GAZO

    Dim mIPAdd As String = ""
    Dim mMasterFileName As String = ""
    Dim mResult As String = ""
    Dim mResultData As String = ""

    Dim cLnk As cls080_F3LinkToLE

    'IPAddress -----------------------------
    Public Property IPAdd() As String
        Set(ByVal Value As String)
            mIPAdd = Value
        End Set
        Get
            IPAdd = mIPAdd
        End Get
    End Property

    'ﾏｽﾀﾌｧｲﾙ設定 ====================================
    Public WriteOnly Property MasterFile() As String
        Set(ByVal Value As String)
            mMasterFileName = Value
        End Set
    End Property

    'Result Text ==============================
    Public ReadOnly Property RESULT() As String
        Get
            RESULT = mResult
        End Get
    End Property

    'Result Text ==============================
    Public ReadOnly Property ResultData() As String
        Get
            ResultData = mResultData
        End Get
    End Property

    Public Function MasterFileRead() As Boolean
        'ﾏｽﾀﾌｧｲﾙの読み込み実行
        MasterFileRead = False
        If mMasterFileName = "" Then Exit Function

        Dim RecvCom As String = ""

        cLnk = New cls080_F3LinkToLE
        Try
            cLnk.IPAdd = mIPAdd : cLnk.Port = 8000
            Call cLnk.SocketCommunication("SetMasterImage,C:\Master\" & mMasterFileName, RecvCom)
            If Strings.Left(RecvCom, 1) <> "0" Then
                Exit Function
            End If
            If Strings.Left(RecvCom, 1) = "0" Then MasterFileRead = True
        Catch ex As Exception

        Finally
            cLnk = Nothing
        End Try

    End Function

    Public Sub GAZO_TEST()

        'Dim RecvCom As String = ""
        Dim strArry() As String = {""}
        Dim strArry1() As String = {""}
        Dim rcvEIData() As String = {""}

        mResult = "" : mResultData = ""

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = mIPAdd : cLnk.Port = 8000

        Try
            'Call cLnk.SocketCommunication("InspectSingleSync", mResultData) 
            Call cLnk.SocketCommunication("T2", mResultData)
            strArry = Split(mResultData, " ")
            ReDim rcvEIData(UBound(strArry))        '画像処理検査結果の格納
            For j = 0 To UBound(strArry)
                strArry1 = Split(strArry(j), ",")
                rcvEIData(j) = strArry1(1)
                If j = UBound(strArry) Then
                    mResult &= strArry1(1)
                Else
                    mResult &= strArry1(1) & ","
                End If
            Next j
        Catch ex As Exception

        Finally
            cLnk = Nothing
        End Try

    End Sub

End Class
