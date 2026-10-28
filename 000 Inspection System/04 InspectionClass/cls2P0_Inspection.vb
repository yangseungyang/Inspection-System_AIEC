Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls2D0_Inspection
'Function   : Adjustment and Inspection ATE X
'Language   : VB2019　Framework4.8
'Reference  : 
'Memo       :
'           :
'Issue      : 2021/07/30  Rev.0    DAEIL.JANG      New Release
'================================================================================================================================

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm020_InspectionStartEnd

Public Class cls2P0_Inspection

    'Class ================================================

    Dim cInsA8HCA As cls2X0_A8HCA_inspection

    Dim cLnk As cls080_F3LinkToLE
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl
    Dim cMainPoewr As cls2X0_A8HCA_inspection

    'Local ================================================
    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String
    Dim mTestSpec As String
    Dim mTestMeasure As String

    Dim iName As String

    'Property =====================================================================================
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

    Public ReadOnly Property TestSpec() As String
        Get
            TestSpec = mTestSpec
        End Get
    End Property

    Public ReadOnly Property TestMeasure() As String
        Get
            TestMeasure = mTestMeasure
        End Get
    End Property

    Public WriteOnly Property InspectionMODEL() As String
        Set(ByVal Value As String)
            mInspectionMODEL = Value
        End Set
    End Property

    Public WriteOnly Property InspectionNumber() As String
        Set(ByVal Value As String)
            mInspectionNo = Value
        End Set
    End Property

    Public WriteOnly Property InspectionName() As String
        Set(ByVal Value As String)
            iName = Value
        End Set
    End Property

    Public Sub TestExecute()

        If InStr(mInspectionMODEL, "A8HCA") <> 0 Then
            'A8HCA検査
            Try
                cInsA8HCA = New cls2X0_A8HCA_inspection
                Call cInsA8HCA.TestExecute(iName)
            Catch ex As Exception
                If cInsA8HCA.ERR_COMMENT <> "" Then
                    If InStr(cInsA8HCA.ERR_COMMENT, ",") <> 0 Then
                        mErrorComment = Replace(cInsA8HCA.ERR_COMMENT, ",", ".")
                    Else
                        mErrorComment = cInsA8HCA.ERR_COMMENT
                        mTestSpec = cInsA8HCA.TestSpec
                        mTestMeasure = cInsA8HCA.TestMeasure
                    End If
                End If
                Throw New ApplicationException(mErrorComment)
            Finally
                cInsA8HCA = Nothing
            End Try

        Else
            mErrorComment = "形名のプログラムが存在しません。確認してください。"
            MsgBox(mErrorComment)
            Throw New ApplicationException(mErrorComment)

        End If

    End Sub

    Public Sub PowerOFF() '*******異常があれば電源OFF、FA-M3で電源制御できるもの*******

        Try
            cMainPoewr = New cls2X0_A8HCA_inspection
            Call cMainPoewr.I91(iName)

            'cBUS = New cls122_FAM3_IOBUS_Ctrl
            'Call cBUS.BUSMONITOR_TB04_END()

            'cInsF3CU04 = New cls2X0_F3CU04_inspection
            'Call cInsF3CU04.rBox_Clear()

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Throw New ApplicationException(mErrorComment)
        Finally
            'cBUS = Nothing
            'cInsA8HCA = Nothing
            cMainPoewr = Nothing

        End Try

        'N号機のみ　IOﾊﾞｽ5Vを元に戻す
        'Call DC_SUPPLY_POWER_ON()

    End Sub

    Public Sub LabelPrint(ByVal dbname As String, ByVal dtkey As String)

        'v0.02で見直し
        Select Case Strings.Left(mInspectionMODEL, 4)

            'Case "F3CU"
            '    Try
            '        cInsF3CU04 = New cls2X0_F3CU04_inspection
            '        Call cInsF3CU04.LabelPrint(dbname, dtkey)
            '    Catch ex As Exception
            '    Finally
            '        cInsF3CU04 = Nothing
            '    End Try
            'Case "F3CX"
            '    Try
            '        cInsF3CU04 = New cls2X0_F3CU04_inspection
            '        Call cInsF3CU04.LabelPrint(dbname, dtkey)
            '    Catch ex As Exception
            '    Finally
            '        cInsF3CU04 = Nothing
            '    End Try

            Case Else

        End Select

    End Sub

    Public Sub BUS_MONITOR_POWER_OFF()

        Try
            'cInsF3CU04 = New cls2X0_F3CU04_inspection
            'Call cInsF3CU04.rBox_Clear()

            'cBUS = New cls122_FAM3_IOBUS_Ctrl
            'Call cBUS.BUS_MONITOR_POWER_OFF()
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Throw New ApplicationException(mErrorComment)
        Finally
            'cBUS = Nothing
            'cInsF3CU04 = Nothing
        End Try

    End Sub

    Public Sub BUS_MONITOR_POWER_ON()

        Try
            'cBUS = New cls122_FAM3_IOBUS_Ctrl
            'Call cBUS.BUS_MONITOR_POWER_ON()

            cInsA8HCA = New cls2X0_A8HCA_inspection
            'Call cInsA8HCA.rBox_Clear()
            Call cInsA8HCA.TestExecute("I95:測定器Preset")

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Throw New ApplicationException(mErrorComment)
        Finally
            cInsA8HCA = Nothing
            'cBUS = Nothing
        End Try

    End Sub

End Class
