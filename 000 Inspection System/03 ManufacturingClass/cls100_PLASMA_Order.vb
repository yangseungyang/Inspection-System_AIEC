Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls100_PLASMA_Order
'Function   : Get order information from PLASMA server 
'Language   : VB2019　Framework4.8
'Reference  : EtherComN
'Memo       :
'           :
'Issue      : 2021/05/07  Rev.0    JANG DAE IL       New Release
'================================================================================================================================

Public Class cls100_PLASMA_Order

    Private EtherComDll_definst As New EtherComN.Dll

    'Local ================================================
    Dim mPlasmaIndex As String
    Dim mPlasmaResult As String

    'Local Property =======================================
    Dim mBarCodeText As String

    Dim mMSCODE, mSERIAL, mORDER, mSALES, mSHIP, mITEM, mCOMP, mODQTY, mQICYN, mQICRQ, mDCTYP As String
    Dim mXJNO, mSPCYN, mSPCID, mSPCC1, mSPCC2, mPRDNO, mPRDSF, mPRDCR, mITMRV, mINSRV, mSTART, mLINE, mCORPG As String
    Dim mSTASC, mFINSC, mINSFN, mUSRM1, mUSRM2, mENDUS, mINDEX As String
    Dim mLINKAGE As String

    Dim mErrorComment As String
    Dim mErrorCode As String

    'Property =====================================================================================
    'Bar code -------------------------
    Public WriteOnly Property BARCODE_TEXT() As String
        Set(ByVal Value As String)
            mBarCodeText = Value
        End Set
    End Property

    'PLASMA Data ----------------------
    Public ReadOnly Property MSCODE() As String
        Get
            MSCODE = mMSCODE
        End Get
    End Property

    Public ReadOnly Property SERIAL() As String
        Get
            SERIAL = mSERIAL
        End Get
    End Property

    Public ReadOnly Property ORDER() As String
        Get
            ORDER = mORDER
        End Get
    End Property

    Public ReadOnly Property SALSE() As String
        Get
            SALSE = mSALES
        End Get
    End Property

    Public ReadOnly Property SHIP() As String
        Get
            SHIP = mSHIP
        End Get
    End Property

    Public ReadOnly Property ITEM() As String
        Get
            ITEM = mITEM
        End Get
    End Property

    Public ReadOnly Property COMP() As String
        Get
            COMP = mCOMP
        End Get
    End Property

    Public ReadOnly Property ODQTY() As String
        Get
            ODQTY = mODQTY
        End Get
    End Property

    Public ReadOnly Property QICYN() As String
        Get
            QICYN = mQICYN
        End Get
    End Property

    Public ReadOnly Property QICRQ() As String
        Get
            QICRQ = mQICRQ
        End Get
    End Property

    Public ReadOnly Property DCTYP() As String
        Get
            DCTYP = mDCTYP
        End Get
    End Property

    Public ReadOnly Property XJNO() As String
        Get
            XJNO = mXJNO
        End Get
    End Property

    Public ReadOnly Property SPCYN() As String
        Get
            SPCYN = mSPCYN
        End Get
    End Property

    Public ReadOnly Property SPCID() As String
        Get
            SPCID = mSPCID
        End Get
    End Property

    Public ReadOnly Property SPCC1() As String
        Get
            SPCC1 = mSPCC1
        End Get
    End Property

    Public ReadOnly Property SPCC2() As String
        Get
            SPCC2 = mSPCC2
        End Get
    End Property

    Public ReadOnly Property PRDNO() As String
        Get
            PRDNO = mPRDNO
        End Get
    End Property

    Public ReadOnly Property PRDSF() As String
        Get
            PRDSF = mPRDSF
        End Get
    End Property

    Public ReadOnly Property PRDCR() As String
        Get
            PRDCR = mPRDCR
        End Get
    End Property

    Public ReadOnly Property ITMRV() As String
        Get
            ITMRV = mITMRV
        End Get
    End Property

    Public ReadOnly Property INSRV() As String
        Get
            INSRV = mINSRV
        End Get
    End Property

    Public ReadOnly Property START() As String
        Get
            START = mSTART
        End Get
    End Property

    Public ReadOnly Property LINE() As String
        Get
            LINE = mLINE
        End Get
    End Property

    Public ReadOnly Property CORPG() As String
        Get
            CORPG = mCORPG
        End Get
    End Property

    Public ReadOnly Property STASC() As String
        Get
            STASC = mSTASC
        End Get
    End Property

    Public ReadOnly Property FINSC() As String
        Get
            FINSC = mFINSC
        End Get
    End Property

    Public ReadOnly Property INSFN() As String
        Get
            INSFN = mINSFN
        End Get
    End Property

    Public ReadOnly Property USRM1() As String
        Get
            USRM1 = mUSRM1
        End Get
    End Property

    Public ReadOnly Property USRM2() As String
        Get
            USRM2 = mUSRM2
        End Get
    End Property

    Public ReadOnly Property ENDUS() As String
        Get
            ENDUS = mENDUS
        End Get
    End Property

    Public ReadOnly Property LINKAGE() As String
        Get
            LINKAGE = mLINKAGE
        End Get
    End Property

    Public ReadOnly Property Index() As String
        Get
            Index = mINDEX
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
    Public Sub GET_ORDER_INFORMAION()
        '==========================================================================================
        ' Function : Plasma communication
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        mErrorComment = ""

        'Bar code -> Index ==========================================
        Try
            Call GetPlasmaIndex()
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        'PLASMA communication =======================================
        Try
            Call PlasmaComm()
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        'PLASMA Data(Property) ======================================
        Try
            Call PlasmaDataProperty()
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

    End Sub

    Private Sub GetPlasmaIndex()
        '==========================================================================================
        ' Function : Get Index Number
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim i As Integer
        Dim String32Word As String
        Dim String10Word As String
        Dim Trans As Integer
        Dim Order32 As String

        Try
            String32Word = mBarCodeText
            String32Word = Trim(String32Word)
            Order32 = "123456789ABCDEFGHIJKLMNOPQRSTUV"
            Trans = 0
            For i = 1 To Len(String32Word)
                Trans = InStr(1, Order32, Mid(String32Word, i, 1)) + Trans * 32
            Next i

            String10Word = Str(Trans)
            String10Word = Right("0000000000" & Trim(String10Word), 10)

            mPlasmaIndex = String10Word

            If mPlasmaIndex = "0000000000" Then
                mErrorComment = "バーコード（Index）が異常です。：0000000000"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Private Sub PlasmaComm()
        '==========================================================================================
        ' Function : PLASMA srv communication
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim ServerName() As String = {"10.7.15.254", "10.7.15.140"}

        For i = 0 To UBound(ServerName)
            Try
                Dim GetEtherComTime As Date
                Dim PC_TNow As Date
                Dim dt As Long

                GetEtherComTime = CDate("1970/1/1 00:00:00") '初期値
                mPlasmaResult = EtherComDll_definst.ComHost(ServerName(i), "", 2031, "SERV-" & "EthercomN")

                Select Case Left(mPlasmaResult, 8)
                    Case "ERROR:A0"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "生産情報がありません。このバ-コ-ドは無効です。"
                    Case "ERROR:A1"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "生産情報がｷｬﾝｾ?ｳ・ﾄい・ｷ。無効の着工札です。"
                    Case "ERROR:A2" : Case "ERROR:A3" : Case "ERROR:A4"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "プログラムエラーです。"
                    Case "ERROR:A5"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "ＯＲＡＣＬＥエラーです。"
                    Case "ERROR:13"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "タイムアウトエラーです。"
                    Case ""
                        mErrorCode = mPlasmaResult
                        mErrorComment = "データがありません。"
                End Select

                If mErrorCode <> "" And i = UBound(ServerName) Then Throw New ApplicationException(mErrorComment)
                mErrorComment = ""
                If mErrorCode <> "" Then Throw New ApplicationException(mErrorComment)

                If InStr(mPlasmaResult, ",") > 0 Then
                    GetEtherComTime = CDate(Split(mPlasmaResult, ",")(1))
                Else
                    mErrorComment = "EtherComデータ取得エラー," & mPlasmaResult
                    Throw New ApplicationException(mErrorComment)
                End If

                'PC_TNow = Format(Now, "yyyy/mm/dd hh:nn:ss") 'PCの現在時刻
                PC_TNow = Now 'PCの現在時刻

                dt = DateDiff("s", GetEtherComTime, PC_TNow)  '時間差(秒)

                If dt > 60 Or dt < -60 Then
                    mErrorComment = "PC時間にずれが発生しました。　Ethercom:" & GetEtherComTime & ", PC:" & PC_TNow

                    MsgBox(mErrorComment)
                    Exit Sub
                End If

                If mErrorCode <> "" Then Throw New ApplicationException(mErrorComment)
                'mPlasmaResult = EtherComDll_definst.ComHost(ServerName(i), "", 1521, "PLAS-" & mPlasmaIndex)   'VER 0.31 GYS Del
                mPlasmaResult = EtherComDll_definst.ComHost(ServerName(i), "", 2031, "PLAS-" & mPlasmaIndex)    'VER 0.31 GYS ADD
                Debug.Print(mPlasmaResult)

                Select Case Left(mPlasmaResult, 8)
                    Case "ERROR:A0"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "生産情報がありません。このバーコードは無効です。"
                    Case "ERROR:A1"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "生産情報がｷｬﾝｾﾙされています。無効の着工札です。"
                    Case "ERROR:A2" : Case "ERROR:A3" : Case "ERROR:A4"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "プログラムエラーです。"
                    Case "ERROR:A5"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "ＯＲＡＣＬＥエラーです。"
                    Case "ERROR:13"
                        mErrorCode = mPlasmaResult
                        mErrorComment = "タイムアウトエラーです。"
                    Case ""
                        mErrorCode = mPlasmaResult
                        mErrorComment = "データがありません。"
                    Case Else
                        Exit Sub
                End Select

                If mErrorCode <> "" And i = UBound(ServerName) Then Throw New ApplicationException(mErrorComment)
                mErrorComment = ""

            Catch ex As Exception
                mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
                Debug.Print(mErrorComment)

                Throw New ApplicationException(mErrorComment)

            End Try
        Next i

    End Sub

    Private Sub PlasmaDataProperty()
        '==========================================================================================
        ' Function : PLASMA data -> Property
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Try
            mMSCODE = DividePlasma("$MS_CODE", mPlasmaResult)
            mSERIAL = DividePlasma("$SERIAL_NO", mPlasmaResult)

            If mMSCODE = "" And mSERIAL = "" Then
                mMSCODE = DividePlasma("$MODEL", mPlasmaResult)
                mSERIAL = "Fxxxxxxxx"
            End If

            mORDER = DividePlasma("$ORDER_NO", mPlasmaResult)
            mSALES = DividePlasma("$SALES_NO", mPlasmaResult)
            mSHIP = DividePlasma("$SHIPMENT_NO", mPlasmaResult)
            mITEM = DividePlasma("$ITEM_NO", mPlasmaResult)
            mCOMP = DividePlasma("$COMP_NO", mPlasmaResult)
            mODQTY = DividePlasma("$ORD_QTY", mPlasmaResult)
            mQICYN = DividePlasma("$TEST_CERT_YN", mPlasmaResult)
            mQICRQ = DividePlasma("$TEST_CERT_SIGN", mPlasmaResult)
            mDCTYP = DividePlasma("$DOC_LANG_TYPE", mPlasmaResult)

            mXJNO = DividePlasma("$XJ_NO", mPlasmaResult)
            mSPCYN = DividePlasma("$TOKUCHU_SPEC_SIGN", mPlasmaResult)
            mSPCID = DividePlasma("$TOKUCHU_SPEC_ID", mPlasmaResult)
            mSPCC1 = DividePlasma("$TOKUCHU_SPEC11", mPlasmaResult)
            mSPCC2 = DividePlasma("$TOKUCHU_SPEC12", mPlasmaResult)

            mPRDNO = DividePlasma("$PROD_NO", mPlasmaResult)
            mPRDSF = DividePlasma("$PROD_NO_SFIX", mPlasmaResult)
            mPRDCR = DividePlasma("$PROD_CAREER", mPlasmaResult)
            mITMRV = DividePlasma("$PROD_ITEM_REV_NO", mPlasmaResult)
            mINSRV = DividePlasma("$PROD_INST_REV_NO", mPlasmaResult)
            mSTART = DividePlasma("$START_NO", mPlasmaResult)
            mLINE = DividePlasma("$LINE_NO", mPlasmaResult)
            mCORPG = DividePlasma("$CRP_GR_NO", mPlasmaResult)

            mSTASC = DividePlasma("$START_SCHDULE_D", mPlasmaResult)
            mFINSC = DividePlasma("$FINISH_SCHDULE_D", mPlasmaResult)
            mINSFN = DividePlasma("$INST_FINISH_D", mPlasmaResult)

            mUSRM1 = DividePlasma("$CUST_INDICATE_DESC1", mPlasmaResult)
            mUSRM2 = DividePlasma("$CUST_INDICATE_DESC2", mPlasmaResult)
            mENDUS = DividePlasma("$END_USER_CUST_N_J", mPlasmaResult)

            mINDEX = DividePlasma("$INDEX_NO", mPlasmaResult)

            mLINKAGE = DividePlasma("$SAP_LINKAGE_NO", mPlasmaResult)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Sub

    Private Function DividePlasma(ByVal Header As String, ByRef Source As String) As String
        '==========================================================================================
        ' Function : PLASMA data divide
        ' Data     : Item name , Data
        ' Return   : Data
        ' Remarks  :
        '==========================================================================================

        Dim StartPosition As Integer
        Dim EndPosition As Integer
        Dim TargetLen As Integer

        Try
            'Start Position
            StartPosition = InStr(Source, Header)
            If StartPosition <= 0 Then
                DividePlasma = ""
                Exit Function
            End If

            'End Position -> Divide
            EndPosition = InStr(StartPosition + Len(Header), Source, "$")
            If EndPosition <= 0 Then
                DividePlasma = Mid(Source, StartPosition + Len(Header))
            Else
                TargetLen = EndPosition - (StartPosition + Len(Header))
                DividePlasma = Mid(Source, StartPosition + Len(Header), TargetLen)
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Left(mErrorComment, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

    End Function

End Class