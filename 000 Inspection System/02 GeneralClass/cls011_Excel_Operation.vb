Option Strict On : Option Explicit On

Imports Microsoft.Office.Interop

'============================================================================================================================
'Class Name : cls011_Excel_Operation
'Function   : Excel operation
'Reference  : Microsoft Excel object Library
'           : Microsoft Access BarCode Control
'Memo       :
'
'Issue      : 2008/12/26  Rev.0    T.Yanagawa       New Release
'============================================================================================================================


Public Class cls011_Excel_Operation

    'Excel Object =========================================
    Dim mxlApp As Excel.Application
    Dim mxlBooks As Excel.Workbooks
    Dim mxlBook As Excel.Workbook
    Dim mxlSheets As Excel.Sheets
    Dim mxlSheet As Excel.Worksheet

    'Local ================================================
    Dim mFilePath As String
    Dim mSheetName As String
    Dim mErrorComment As String
    Dim mErrorCode As String

    'Property =====================================================================================
    'Excel File Path ------------------
    Public WriteOnly Property FILE_PATH() As String
        Set(ByVal Value As String)
            mFilePath = Value
        End Set
    End Property

    'Sheet Name -----------------------
    Public WriteOnly Property SHEET_NAME() As String
        Set(ByVal Value As String)
            mSheetName = Value
        End Set
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
    Public Sub EXCEL_START()
        '==========================================================================================
        ' Function : Excel Start
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' Exel start ==================================
            mxlApp = New Excel.Application
            mxlBooks = mxlApp.Workbooks
            mxlBook = mxlBooks.Open(mFilePath)

            mxlApp.Visible = True

            mxlSheets = mxlBook.Worksheets
            mxlSheet = DirectCast(mxlSheets(mSheetName), Excel.Worksheet)

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub CELL_DATA_SET(ByVal CellName As String, ByVal CellData As String)
        '==========================================================================================
        ' Function : Cell Name & Data set
        ' Data     : Cell Name , Cell(Data)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' Setting of Cell range =================================
            Dim mxlRange As Excel.Range
            mxlRange = mxlSheet.Range(CellName)

            Try ' Setting of Cell data ====================
                mxlRange.Value = CellData
            Finally ' =====================================
                If Not mxlRange Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlRange)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub CELL_DATA_CLEAR(ByVal CellName As String, ByVal CellData As String)
        '==========================================================================================
        ' Function : Cell Name & Data set
        ' Data     : Cell Name , Cell(Data)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' Setting of Cell range =================================
            Dim mxlRange As Excel.Range
            mxlRange = mxlSheet.Range(CellName)

            Try ' Setting of Cell data ====================
                mxlRange.Clear()
            Finally ' =====================================
                If Not mxlRange Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlRange)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub CELL_DATA_SETR(ByVal CellName As String, ByVal CellData As String)
        '==========================================================================================
        ' Function : Cell Name & Data set
        ' Data     : Cell Name , Cell(Data)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' Setting of Cell range =================================
            Dim mxlRange As Excel.Range
            mxlRange = mxlSheet.Range(CellName)

            Try ' Setting of Cell data ====================
                mxlRange.Value = CellData
                mxlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            Finally ' =====================================
                If Not mxlRange Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlRange)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub CELL_FONT_COLOR_SET(ByVal CellName As String, ByVal color As Integer)
        '==========================================================================================
        ' Function : Setting of cell font color
        ' Data     : Cell name , color
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' Setting of Cell range =================================
            Dim mxlRange As Excel.Range
            mxlRange = mxlSheet.Range(CellName)

            Try ' Setting of Cell font ====================
                Dim mxlFont As Excel.Font
                mxlFont = mxlRange.Font

                Try ' Setting of font color ======
                    mxlFont.ColorIndex = color
                Finally ' ========================
                    If Not mxlFont Is Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlFont)
                    End If
                End Try

            Finally ' =====================================
                If Not mxlRange Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlRange)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub CELL_BACK_COLOR_SET(ByVal CellName As String, ByVal color As Integer)
        '==========================================================================================
        ' Function : Setting of cell back color
        ' Data     : Cell name , color
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' Setting of Cell range =================================
            Dim mxlRange As Excel.Range
            mxlRange = mxlSheet.Range(CellName)

            Try ' Setting of Cell Interior ================
                Dim mxlBack As Excel.Interior
                mxlBack = mxlRange.Interior

                Try ' Setting of back color =====
                    mxlBack.ColorIndex = color
                Finally ' =======================
                    If Not mxlBack Is Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlBack)
                    End If
                End Try

            Finally ' =====================================
                If Not mxlRange Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlRange)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub PICTURE_VISIBLE_CONTROL(ByVal PicName As String, ByVal disp As Boolean)
        '==========================================================================================
        ' F.Name   : PictureVisible
        ' Function : Picture Visible Control
        ' Data     : Picture neme , Display = True or False
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        Try ' Setting of shapes =====================================
            Dim mxlShapes As Excel.Shapes
            mxlShapes = mxlSheet.Shapes

            Try ' Setting of shape ========================
                Dim mxlShape As Excel.Shape
                mxlShape = mxlShapes._Default(PicName)

                Try ' Control of display ======== 
                    If disp = False Then
                        mxlShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    Else
                        mxlShape.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue
                    End If
                Finally ' =======================
                    If Not mxlShape Is Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlShape)
                    End If
                End Try

            Finally ' =====================================
                If Not mxlShapes Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlShapes)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub PICTURE_QRPASTE_CONTROL()
        '==========================================================================================
        ' F.Name   : PicturePaste
        ' Function : Picture Paste Control from clipborad
        ' Data     : Picture neme , Display = True or False
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================
        Dim CellName As String = "QR"
        mErrorComment = ""

        Try ' Setting of Cell range =================================
            Dim mxlRange As Excel.Range
            mxlRange = mxlSheet.Range(CellName)

            Try ' Setting of Cell data ====================
                mxlSheet.Paste()
                Dim mxlShapes As Excel.Shapes
                mxlShapes = mxlSheet.Shapes
                Dim mxlShape As Excel.Shape
                Dim i As Integer
                For Each mxlShape In mxlShapes
                    i += 1
                    If i = mxlShapes.Count Then
                        'mxlShape.Rotation = 270     '270“x‰ñ“]
                        mxlShape.Height = 25
                        mxlShape.Top = CSng(mxlRange.Top)
                        mxlShape.Left = CSng(mxlRange.Left) - 1
                    End If
                Next

            Finally ' =====================================
                If Not mxlRange Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlRange)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try


    End Sub
    Public Sub PICTURE_QRPASTE_CONTROL2()
        '==========================================================================================
        ' F.Name   : PicturePaste
        ' Function : Picture Paste Control from clipborad
        ' Data     : Picture neme , Display = True or False
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================
        Dim CellName As String = "QR2"
        mErrorComment = ""

        Try ' Setting of Cell range =================================
            Dim mxlRange As Excel.Range
            mxlRange = mxlSheet.Range(CellName)

            Try ' Setting of Cell data ====================
                mxlSheet.Paste()
                Dim mxlShapes As Excel.Shapes
                mxlShapes = mxlSheet.Shapes
                Dim mxlShape As Excel.Shape
                Dim i As Integer
                For Each mxlShape In mxlShapes
                    i += 1
                    If i = mxlShapes.Count Then
                        'mxlShape.Rotation = 270     '270“x‰ñ“]
                        mxlShape.Height = 25
                        mxlShape.Top = CSng(mxlRange.Top)
                        mxlShape.Left = CSng(mxlRange.Left) - 1
                    End If
                Next

            Finally ' =====================================
                If Not mxlRange Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlRange)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try


    End Sub

    Public Sub BARCODE_SHOW(ByVal CellName As String, ByVal t As Integer, ByVal l As Integer, ByVal h As Integer, ByVal w As Integer)
        '==========================================================================================
        ' Function : Barcode Control
        ' Data     : BarCodeValue
        ' Return   : Nothing
        ' Remarks  : 
        '==========================================================================================

        mErrorComment = ""

        Try
            'OLEObjects =========================
            Dim mxlSheetObj As Excel.OLEObjects
            mxlSheetObj = DirectCast(mxlSheet.OLEObjects, Excel.OLEObjects)
            Try
                'OLEObject =========================
                Dim objOLEObject As Excel.OLEObject
                objOLEObject = mxlSheetObj.Add(ClassType:="BARCODE.BarCodeCtrl.1", Link:=False, DisplayAsIcon:=False)

                Try
                    'Barcode =========================
                    Dim objBarCode As BARCODELib.BarCodeCtrl
                    objBarCode = DirectCast(objOLEObject.Object, BARCODELib.BarCodeCtrl)

                    Try
                        'Set Barcode Property ====================
                        With objBarCode
                            .BackColor = &HFFFFFF
                            .Style = 7
                            .Direction = 1
                            .Validation = 0
                            .ShowData = 0
                            .LineWeight = 3
                            .Refresh()
                        End With

                        'Setting of objOLEObject ===================  
                        objOLEObject.Locked = False
                        objOLEObject.LinkedCell = CellName
                        objOLEObject.Width = w : objOLEObject.Height = h : objOLEObject.Top = t : objOLEObject.Left = l
                    Finally
                        If Not objBarCode Is Nothing Then
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objBarCode)
                        End If
                    End Try

                Finally
                    If Not objOLEObject Is Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objOLEObject)
                    End If
                End Try

            Finally
                If Not mxlSheetObj Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlSheetObj)
                End If
            End Try

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub PRINT_EXCEL(ByVal PrinterName As String)
        '==========================================================================================
        ' Function : Printing of Excel file
        ' Data     : Printer name
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        'Printer check ==============================================
        Dim okPrinter As Boolean = False
        For Each InsPrt As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            Debug.Print(InsPrt)
            If InsPrt = PrinterName Then okPrinter = True
        Next
        If okPrinter = False Then
            mErrorComment = "ERROR:PRINTER NOTHING"
            Throw New ApplicationException(mErrorComment)
        End If

        'Print ======================================================
        Try
            mxlSheet.PrintOut(ActivePrinter:=PrinterName)

        Catch ex As Exception ' ================================
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Public Sub CLOSE_EXCEL()
        '==========================================================================================
        ' Function : Close of excel file
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        'App. Message kill ------------
        mxlApp.DisplayAlerts = False

        If Not mxlSheet Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlSheet)
        End If

        If Not mxlSheets Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlSheets)
        End If

        If Not mxlBook Is Nothing Then
            mxlBook.Close()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlBook)
        End If

        If Not mxlBooks Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlBooks)
        End If

        If Not mxlApp Is Nothing Then
            mxlApp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mxlApp)
        End If

    End Sub

End Class