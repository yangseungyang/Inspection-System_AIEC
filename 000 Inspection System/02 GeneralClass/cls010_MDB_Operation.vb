Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls010_MDB_Operation
'Function   : Jet-mdb Read Write
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'
'Issue      : 2008/09/30  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

Imports System.Data.OleDb

Public Class cls010_MDB_Operation

    'OLE DB ===============================================
    Dim oCnn As OleDbConnection
    Dim oCmd As OleDbCommand
    Dim oAdp As OleDbDataAdapter
    Dim dts As DataSet

    'Local Property =======================================
    Dim mdbPath As String = ""
    Dim mTblName As String = ""
    Dim mSQL As String = ""
    Dim mdbField As String = ""
    Dim mdbData As String = ""
    Dim mDataCount As Integer = 0
    Dim mFIeldCount As Integer = 0
    Dim mRecordCount As Integer = 0
    Dim mErrorComment As String = ""
    Dim mErrorCode As String = ""

    'Property =====================================================================================
    'db path --------------------------
    Public WriteOnly Property DB_PATH() As String
        Set(ByVal Value As String)
            mdbPath = Value
        End Set
    End Property

    'table Name -----------------------
    Public WriteOnly Property TB_NAME() As String
        Set(ByVal Value As String)
            mTblName = Value
        End Set
    End Property

    'SQL ------------------------------
    Public WriteOnly Property DB_SQL() As String
        Set(ByVal Value As String)
            mSQL = Value
        End Set
    End Property

    'DB field name --------------------
    Public ReadOnly Property DB_FIELD_NAME() As String
        Get
            DB_FIELD_NAME = mdbField
        End Get
    End Property

    'DB Data --------------------------
    Public ReadOnly Property DB_DATA() As String
        Get
            DB_DATA = mdbData
        End Get
    End Property

    'DB Data count --------------------
    Public ReadOnly Property DB_DATA_COUNT() As Integer
        Get
            DB_DATA_COUNT = mDataCount
        End Get
    End Property

    'DB field count -------------------
    Public ReadOnly Property DB_FIELD_COUNT() As Integer
        Get
            DB_FIELD_COUNT = mFIeldCount
        End Get
    End Property

    'DB Record count ------------------
    Public ReadOnly Property DB_RECORD_COUNT() As Integer
        Get
            DB_RECORD_COUNT = mRecordCount
        End Get
    End Property

    'Error message --------------------
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
    Public Sub DB_READ_JET(ByRef Field_Name() As String, ByRef DB_Data() As String)
        '==========================================================================================
        ' Function : Open Dataset - Get Field name & Data
        ' Data     : Nothing
        ' Return   : Field name(), data()
        ' Remarks  :
        '==========================================================================================

        Dim dtTable As DataTable
        Dim dtCol As DataColumn
        Dim oRow As DataRow

        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Dim i As Integer = 0
        Dim sb As New System.Text.StringBuilder

        mErrorComment = ""

        ReDim Field_Name(0)
        ReDim DB_Data(0)

        Try ' =======================================================

            'DB Cnn =================================================
            oCnn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & mdbPath & ";")

            'Set SQL -----------------------
            oCmd = oCnn.CreateCommand
            oCmd.CommandText = mSQL
            Debug.Print(mSQL)

            'Cnn Adapter -------------------
            oAdp = New OleDbDataAdapter(oCmd)

            'Data Set ----------------------
            dts = New DataSet
            oAdp.Fill(dts, mTblName)

            'Get Field name =========================================
            dtTable = dts.Tables(mTblName)
            For Each dtCol In dtTable.Columns
                'mdbField &= dtCol.ColumnName & ","
                sb.Append(dtCol.ColumnName) : sb.Append(",")
            Next

            'Make return field name --------
            'If mdbField <> "" Then
            If sb.ToString <> "" Then
                'mdbField = Left(mdbField, Len(mdbField) - 1)
                mdbField = Left(sb.ToString, Len(sb.ToString) - 1)
                fName = Split(mdbField, ",", -1, CompareMethod.Binary)
                mFIeldCount = fName.Length
            Else
                ReDim fName(0) : fName(0) = ""
            End If
            Field_Name = DirectCast(fName.Clone(), String())

            sb.Length = 0

            'Get Data ===============================================
            mRecordCount = 0
            For Each oRow In dtTable.Rows
                i = 0
                For Each dtCol In dtTable.Columns
                    If Not IsDBNull(oRow(i)) Then
                        'mdbData &= CStr(oRow(i)) & ",,"
                        sb.Append(CStr(oRow(i))) : sb.Append(",,")
                    Else
                        'mdbData &= "" & ",,"
                        sb.Append("") : sb.Append(",,")
                    End If
                    i += 1
                Next
                mRecordCount += 1
            Next

            'Make return data --------------
            'If mdbData <> "" Then
            If sb.ToString <> "" Then
                'mdbData = Left(mdbData, Len(mdbData) - 2)
                mdbData = Left(sb.ToString, Len(sb.ToString) - 2)
                sData = Split(mdbData, ",,", -1, CompareMethod.Binary)
                mDataCount = sData.Length
            Else
                ReDim sData(0) : sData(0) = ""
            End If
            DB_Data = DirectCast(sData.Clone(), String())

        Catch ex As Exception ' ================================

            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            mRecordCount = 0 : mFIeldCount = 0 : mdbData = "" : mDataCount = 0
            fName = Split(mdbField, ",", -1, CompareMethod.Binary)
            sData = Split(mdbData, ",,", -1, CompareMethod.Binary)

            Field_Name = DirectCast(fName.Clone(), String())
            DB_Data = DirectCast(sData.Clone(), String())

            Throw New ApplicationException(mErrorComment)


        Finally ' ===================================================

            Debug.Print("Field:" & mdbField)
            Debug.Print("Rec:" & mRecordCount & " Field:" & mFIeldCount & " Data:" & mdbData & " DataNum:" & mDataCount)

            dts = Nothing
            oAdp = Nothing

            oCnn.Close()
            oCnn = Nothing

            sb.Length = 0
            sb = Nothing

        End Try

    End Sub

    Public Sub DB_EXEC_JET()
        '==========================================================================================
        ' Function : SQL execute
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' =======================================================

            oCnn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & mdbPath & ";")
            oCnn.Open()

            ''Set SQL -----------------------
            oCmd = oCnn.CreateCommand
            oCmd.CommandText = mSQL
            Debug.Print(mSQL)

            oCmd.ExecuteNonQuery()


        Catch ex As Exception ' ================================

            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)


        Finally ' ===================================================

            oCnn.Close()
            oCnn = Nothing


        End Try

    End Sub


    'Process ======================================================================================
    Public Sub DB_READ_EXCEL_JET(ByRef Field_Name() As String, ByRef DB_Data() As String)
        '==========================================================================================
        ' Function : Open Dataset - Get Field name & Data
        ' Data     : Nothing
        ' Return   : Field name(), data()
        ' Remarks  :
        '==========================================================================================

        Dim dtTable As DataTable
        Dim dtCol As DataColumn
        Dim oRow As DataRow

        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Dim i As Integer = 0

        mErrorComment = ""

        ReDim Field_Name(0)
        ReDim DB_Data(0)

        Try ' =======================================================

            'DB Cnn =================================================
            oCnn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                       "Extended Properties=Excel 8.0;" & _
                                       "Data Source=" & mdbPath & ";")

            'Set SQL -----------------------
            oCmd = oCnn.CreateCommand
            oCmd.CommandText = mSQL
            Debug.Print(mSQL)

            'Cnn Adapter -------------------
            oAdp = New OleDbDataAdapter(oCmd)

            'Data Set ----------------------
            dts = New DataSet
            oAdp.Fill(dts, mTblName)


            'Get Field name =========================================
            dtTable = dts.Tables(mTblName)
            For Each dtCol In dtTable.Columns
                mdbField &= dtCol.ColumnName & ","
            Next

            'Make return field name --------
            If mdbField <> "" Then
                mdbField = Left(mdbField, Len(mdbField) - 1)
                fName = Split(mdbField, ",", -1, CompareMethod.Binary)
                mFIeldCount = fName.Length
            Else
                ReDim fName(0) : fName(0) = ""
            End If
            Field_Name = DirectCast(fName.Clone(), String())


            'Get Data ===============================================
            mRecordCount = 0
            For Each oRow In dtTable.Rows
                i = 0
                For Each dtCol In dtTable.Columns
                    If Not IsDBNull(oRow(i)) Then
                        mdbData &= CStr(oRow(i)) & ",,"
                    Else
                        mdbData &= "" & ",,"
                    End If
                    i += 1
                Next
                mRecordCount += 1
            Next

            'Make return data --------------
            If mdbData <> "" Then
                mdbData = Left(mdbData, Len(mdbData) - 2)
                sData = Split(mdbData, ",,", -1, CompareMethod.Binary)
                mDataCount = sData.Length
            Else
                ReDim sData(0) : sData(0) = ""
            End If
            DB_Data = DirectCast(sData.Clone(), String())


        Catch ex As Exception ' ================================

            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            mRecordCount = 0 : mFIeldCount = 0 : mdbData = "" : mDataCount = 0
            fName = Split(mdbField, ",", -1, CompareMethod.Binary)
            sData = Split(mdbData, ",,", -1, CompareMethod.Binary)

            Field_Name = DirectCast(fName.Clone(), String())
            DB_Data = DirectCast(sData.Clone(), String())

            Throw New ApplicationException(mErrorComment)


        Finally ' ===================================================

            Debug.Print("Field:" & mdbField)
            Debug.Print("Rec:" & mRecordCount & " Field:" & mFIeldCount & " Data:" & mdbData & " DataNum:" & mDataCount)

            dts = Nothing
            oAdp = Nothing

            oCnn.Close()
            oCnn = Nothing


        End Try

    End Sub

    Public Sub MDB_SAITEKIKA()

        Dim MdbFilePath As String
        Dim MdbFilePath2 As String
        Dim MdbJet As JRO.JetEngine
        Dim MdbCon As String

        Try
            MdbJet = New JRO.JetEngine
            '最適化を行うファイル
            MdbFilePath2 = mdbPath
            'wkファイル
            MdbFilePath = My.Application.Info.DirectoryPath & "\進捗管理_wk.mdb"
            'ファイルをwkファイル名にリネーム
            IO.File.Move(MdbFilePath2, MdbFilePath)
            MdbCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
            '最適化
            Call MdbJet.CompactDatabase(MdbCon & MdbFilePath, MdbCon & MdbFilePath2)

            'wkファイルの削除
            IO.File.Delete(MdbFilePath)

        Catch ex As Exception

        End Try

    End Sub

End Class
