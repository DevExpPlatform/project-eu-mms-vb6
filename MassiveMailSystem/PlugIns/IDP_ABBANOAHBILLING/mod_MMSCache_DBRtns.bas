Attribute VB_Name = "mod_MMSCache_DBRtns"
Option Explicit

Public Enum enum_DATASECTION
    CATEGORIE
    FORNITURA_MASTER_PHASE01
    FORNITURA_MASTER_PHASE02_F
    FORNITURA_MASTER_PHASE02_P
    CODICE_IPA
    LOC_UBI
    INDE
    NATIONALITY
    WMS
    FORNITURA_MASTER_QUERIES
    TOTCONG_TICSI
    DEL_547_19_B_CAUSALI
    FLG_DOM
    DYN_ANNXD_UI_CMG = 12
    DYN_ANNXD_UI_IFUR = 13
    DYN_MSG_PDF = 14
    DYN_ANNXD_ONE_SHOT = 40
    DYN_ANNXD_T01 = 50
    DYN_ANNXD_547_2019 = 51
    DYN_ANNXD_T02 = 52
    DYN_ANNXD_T03 = 53
    DYN_REASON = 54
    DYN_ANNXD_TMPLT = 56
End Enum

Public Type strct_DATA
    DATAID                                  As String
    dataDescription                         As String
    EXTRAPARAMS()                           As String
    FLG_EXTRAPARAMS                         As Boolean
End Type

Private WS_DC_CATEGORIE()                   As strct_DATA
Private WS_DC_CODICE_IPA()                  As strct_DATA
Private WS_DC_DEL_547_19_B_CAUSALI()        As strct_DATA
Private WS_DC_DOM()                         As strct_DATA
Private WS_DC_DYN_ANNXD_547_2019            As Collection
Private WS_DC_DYN_ANNXD_ONE_SHOT            As Collection
Private WS_DC_DYN_ANNXD_T01                 As Collection
Private WS_DC_DYN_ANNXD_T02                 As Collection
Private WS_DC_DYN_ANNXD_T03                 As Collection
Private WS_DC_DYN_MSG_PDF()                 As strct_DATA
Private WS_DC_DYN_REASON                    As Collection
Private WS_DC_DYN_TMPLT()                   As strct_DATA
Private WS_DC_FORN_MASTER_PHASE01()         As strct_DATA
Private WS_DC_FORN_MASTER_PHASE02_F()       As strct_DATA
Private WS_DC_FORN_MASTER_PHASE02_P()       As strct_DATA
Private WS_DC_FORN_MASTER_QUERIES()         As strct_DATA
'Private WS_DC_IFUR()                        As strct_DATA
Private WS_DC_INDE()                        As strct_DATA
Private WS_DC_LOC_UBI()                     As strct_DATA
Private WS_DC_NATIONALITY()                 As strct_DATA
Private WS_DC_TOTCONG_TICSI                 As Collection
'Private WS_DC_UI_CUR()                      As strct_DATA
Private WS_DC_UI_CMG()                      As strct_DATA
Private WS_DC_UI_IFUR()                     As strct_DATA
Private WS_DC_WMS()                         As strct_DATA

Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_01    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_02    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_03    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_04    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_05    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_06    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_07    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_08    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_09    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_10    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_11    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_12    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_13    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_14    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_15    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_16    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_17    As strct_DATA
Public WS_CS_DYN_ATTCH_UI_GMG_TBL_DSCR      As strct_DATA
Public WS_CS_DYN_ATTCH_UI_GMC_TBL_FTR       As strct_DATA
Public WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR       As strct_DATA
Public WS_CS_DYN_ATTCH_UI_GMC_TBL_ROW       As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_TBL_DSCR     As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_TBL_FTR      As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR      As strct_DATA
Public WS_CS_DYN_ATTCH_UI_IFUR_TBL_ROW      As strct_DATA
Public WS_CS_P01_MSG_BS                     As strct_DATA
Public WS_CS_P01_MSG_CMA                    As strct_DATA
Public WS_CS_P01_MSG_COM_ARERA              As strct_DATA
Public WS_CS_P01_MSG_INFO_FATTURA_BDS       As strct_DATA
Public WS_CS_P01_MSG_INFO_FATTURA_STD       As strct_DATA
Public WS_CS_P01_MSG_NSO                    As strct_DATA
Public WS_CS_P01_MSG_SERDEP                 As strct_DATA
Public WS_CS_PXX_ADF_CFPI_MSG_OK            As strct_DATA
Public WS_CS_PXX_ADF_CFPI_MSG_KO            As strct_DATA
Public WS_CS_PXX_LBL_ADF_RCA                As strct_DATA
Public WS_CS_PXX_LBL_AUI                    As strct_DATA
Public WS_CS_PXX_LBL_AUI_FTR                As strct_DATA
Public WS_CS_PXX_LBL_AUI_ROW                As strct_DATA
Public WS_CS_PXX_LBL_BONUS_SOCIALE          As strct_DATA
Public WS_CS_PXX_TBL_DELAY_INFO_M01         As strct_DATA
Public WS_CS_PXX_TBL_DELAY_INFO_M02         As strct_DATA
Public WS_CS_PXX_TBL_MSG_AC                 As strct_DATA
Public WS_CS_PXX_TBL_MSG_PAGOPA             As strct_DATA
Public WS_CS_PXX_TBL_MSG_UC                 As strct_DATA
Public WS_CUSTOMERSTYLE()                   As strct_DATA
Public WS_LOGFILEPATH                       As String
Public WS_PXX_FOOTER                        As String
Public WS_PXX_FOOTER_WSM                    As String
Public WS_FLG_WMS                           As Boolean

Public Sub CACHE_CLEAR()
    
    Erase WS_DC_CATEGORIE()
    Erase WS_DC_CODICE_IPA()
    Erase WS_DC_DEL_547_19_B_CAUSALI()
    Erase WS_DC_DOM()
    Erase WS_DC_FORN_MASTER_PHASE01()
    Erase WS_DC_FORN_MASTER_PHASE02_F()
    Erase WS_DC_FORN_MASTER_PHASE02_P()
    Erase WS_DC_FORN_MASTER_QUERIES()
    Erase WS_DC_INDE()
    Erase WS_DC_LOC_UBI()
    Erase WS_DC_NATIONALITY()
    Erase WS_DC_WMS()

    Set WS_DC_DYN_ANNXD_547_2019 = Nothing
    Set WS_DC_DYN_ANNXD_ONE_SHOT = Nothing
    Set WS_DC_DYN_ANNXD_T01 = Nothing
    Set WS_DC_DYN_ANNXD_T02 = Nothing
    Set WS_DC_DYN_ANNXD_T03 = Nothing
    Set WS_DC_TOTCONG_TICSI = Nothing

    Erase WS_CUSTOMERSTYLE()
    
    WS_PXX_FOOTER = ""
    WS_PXX_FOOTER_WSM = ""
    WS_FLG_WMS = False

End Sub

Public Sub CACHE_INIT()
    
    ' EXTERNAL CONFIGS/PARAMS
    '
    If (DLLParams.DOCMODE = "BOL") Then
        GET_EXTERNALDATA DLLParams.EXTRASPATH & "Comunicazioni.STL", WS_CUSTOMERSTYLE, -1
        GET_EXTERNALDATA DLLParams.EXTRASPATH & "Invoice.STL", WS_CUSTOMERSTYLE, UBound(WS_CUSTOMERSTYLE)
        GET_EXTERNALDATA DLLParams.EXTRASPATH & "TXT_DYN_ANNXD_UI.STL", WS_CUSTOMERSTYLE, UBound(WS_CUSTOMERSTYLE)
    
        ' COMMUNICATIONS
        '
        WS_CS_P01_MSG_BS = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "P01_MSG_BS")
        WS_CS_P01_MSG_CMA = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "P01_MSG_CMA")
        WS_CS_P01_MSG_COM_ARERA = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "P01_MSG_COM_ARERA")
        WS_CS_P01_MSG_INFO_FATTURA_BDS = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "P01_MSG_INFO_FATTURA_BDS")
        WS_CS_P01_MSG_INFO_FATTURA_STD = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "P01_MSG_INFO_FATTURA_STD")
        WS_CS_P01_MSG_NSO = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "P01_MSG_NSO")
        WS_CS_P01_MSG_SERDEP = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "P01_MSG_SERDEP")
        WS_CS_PXX_ADF_CFPI_MSG_KO = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_ADF_CFPI_MSG_KO")
        WS_CS_PXX_ADF_CFPI_MSG_OK = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_ADF_CFPI_MSG_OK")
        WS_CS_PXX_LBL_ADF_RCA = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_LBL_ADF_RCA")
        WS_CS_PXX_LBL_AUI = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_LBL_AUI")
        WS_CS_PXX_LBL_AUI_FTR = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_LBL_AUI_FTR")
        WS_CS_PXX_LBL_AUI_ROW = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_LBL_AUI_ROW")
        WS_CS_PXX_LBL_BONUS_SOCIALE = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_LBL_BONUS_SOCIALE")
        WS_CS_PXX_TBL_DELAY_INFO_M01 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_TBL_DELAY_INFO_M01")
        WS_CS_PXX_TBL_DELAY_INFO_M02 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_TBL_DELAY_INFO_M02")
        WS_CS_PXX_TBL_MSG_AC = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_TBL_MSG_AC")
        WS_CS_PXX_TBL_MSG_PAGOPA = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_TBL_MSG_PAGOPA")
        WS_CS_PXX_TBL_MSG_UC = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_TBL_MSG_UC")
        
        ' UI_GMC
        '
        WS_CS_DYN_ATTCH_UI_GMG_TBL_DSCR = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_GMG_TBL_DSCR")
        WS_CS_DYN_ATTCH_UI_GMC_TBL_FTR = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_GMC_TBL_FTR")
        WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_GMC_TBL_HDR")
        WS_CS_DYN_ATTCH_UI_GMC_TBL_ROW = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_GMC_TBL_ROW")
        
        ' UI_IFUR
        '
        WS_CS_DYN_ATTCH_UI_IFUR_TBL_DSCR = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_IFUR_TBL_DSCR")
        WS_CS_DYN_ATTCH_UI_IFUR_TBL_FTR = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_IFUR_TBL_FTR")
        WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_IFUR_TBL_HDR")
        WS_CS_DYN_ATTCH_UI_IFUR_TBL_ROW = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_IFUR_TBL_ROW")
        
        ' UI_MB_XX
        '
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_01 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_01")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_02 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_02")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_03 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_03")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_04 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_04")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_05 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_05")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_06 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_06")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_07 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_07")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_08 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_08")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_09 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_09")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_10 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_10")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_11 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_11")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_12 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_12")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_13 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_13")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_14 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_14")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_15 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_15")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_16 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_16")
        WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_17 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "UI_P02_MB_17")

        ' DATA CACHE
        '
        GET_DATA_CACHE_LOAD WS_DC_CATEGORIE, CATEGORIE                            ' ID_DECODERTYPE = 0
        GET_DATA_CACHE_LOAD WS_DC_INDE, INDE                                      ' ID_DECODERTYPE = 6
        GET_DATA_CACHE_LOAD WS_DC_FORN_MASTER_QUERIES, FORNITURA_MASTER_QUERIES   ' ID_DECODERTYPE = 9
        GET_DATA_CACHE_LOAD WS_DC_DEL_547_19_B_CAUSALI, DEL_547_19_B_CAUSALI      ' ID_DECODERTYPE = 11
        'GET_DATA_COLLECTION_LOAD WS_DC_DYN_ANNXD_ONE_SHOT, DYN_ANNXD_ONE_SHOT    ' ID_DECODERTYPE = 40
        'GET_DATA_COLLECTION_LOAD WS_DC_DYN_ANNXD_T01, DYN_ANNXD_T01              ' ID_DECODERTYPE = 50
        'GET_DATA_COLLECTION_LOAD WS_DC_DYN_ANNXD_547_2019, DYN_ANNXD_547_2019    ' ID_DECODERTYPE = 51
        'GET_DATA_COLLECTION_LOAD WS_DC_DYN_ANNXD_T02, DYN_ANNXD_T02              ' ID_DECODERTYPE = 52
        GET_DATA_CACHE_LOAD WS_DC_DYN_TMPLT, DYN_ANNXD_TMPLT                      ' ID_DECODERTYPE = 56
    End If
    
    GET_DATA_COLLECTION_LOAD WS_DC_DYN_ANNXD_T03, DYN_ANNXD_T03                   ' ID_DECODERTYPE = 53
    GET_DATA_COLLECTION_LOAD WS_DC_DYN_REASON, DYN_REASON                         ' ID_DECODERTYPE = 54
    
End Sub

Public Function GET_ACCESSIBILITÀ_CONTATORE() As String

    Select Case WS_G09.R001.INDICAZIONEACCESSIBILITÀ_218_16
    Case "1"
        GET_ACCESSIBILITÀ_CONTATORE = "accessibile"

    Case "2"
        GET_ACCESSIBILITÀ_CONTATORE = "non accessibile"

    Case "3"
        GET_ACCESSIBILITÀ_CONTATORE = "parzialmente accessibile"

    End Select

End Function

Public Function GET_DATA_CACHE(dataType As enum_DATASECTION, strDataSrch As String) As strct_DATA

    On Error GoTo ErrHandler

    Dim dataCache As cls_DataCache

    strDataSrch = Trim$(strDataSrch)

    Select Case dataType
    Case CATEGORIE
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_CATEGORIE, strDataSrch)
        
    Case CODICE_IPA
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_CODICE_IPA, strDataSrch)
    
    Case DEL_547_19_B_CAUSALI
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_DEL_547_19_B_CAUSALI, strDataSrch)
    
    Case DYN_ANNXD_547_2019
        Set dataCache = WS_DC_DYN_ANNXD_547_2019(strDataSrch)

        GET_DATA_CACHE.dataDescription = dataCache.getDescription

        If (dataCache.getExtraParamsFlag) Then
            GET_DATA_CACHE.EXTRAPARAMS = dataCache.getExtraParams
            GET_DATA_CACHE.FLG_EXTRAPARAMS = True
        End If

        Set dataCache = Nothing
    
    Case DYN_ANNXD_ONE_SHOT
        Set dataCache = WS_DC_DYN_ANNXD_ONE_SHOT(strDataSrch)
    
        GET_DATA_CACHE.dataDescription = dataCache.getDescription
        
        If (dataCache.getExtraParamsFlag) Then
            GET_DATA_CACHE.EXTRAPARAMS = dataCache.getExtraParams
            GET_DATA_CACHE.FLG_EXTRAPARAMS = True
        End If
    
        Set dataCache = Nothing
    
    Case DYN_ANNXD_T01
        Set dataCache = WS_DC_DYN_ANNXD_T01(strDataSrch)
    
        GET_DATA_CACHE.dataDescription = dataCache.getDescription
        
        If (dataCache.getExtraParamsFlag) Then
            GET_DATA_CACHE.EXTRAPARAMS = dataCache.getExtraParams
            GET_DATA_CACHE.FLG_EXTRAPARAMS = True
        End If
    
        Set dataCache = Nothing
    
    Case DYN_ANNXD_T02
        Set dataCache = WS_DC_DYN_ANNXD_T02(strDataSrch)
    
        GET_DATA_CACHE.dataDescription = dataCache.getDescription
        
        If (dataCache.getExtraParamsFlag) Then
            GET_DATA_CACHE.EXTRAPARAMS = dataCache.getExtraParams
            GET_DATA_CACHE.FLG_EXTRAPARAMS = True
        End If
    
        Set dataCache = Nothing
    
    Case DYN_ANNXD_T03
        Set dataCache = WS_DC_DYN_ANNXD_T03(strDataSrch)
    
        GET_DATA_CACHE.dataDescription = dataCache.getDescription
        
        If (dataCache.getExtraParamsFlag) Then
            GET_DATA_CACHE.EXTRAPARAMS = dataCache.getExtraParams
            GET_DATA_CACHE.FLG_EXTRAPARAMS = True
        End If
    
        Set dataCache = Nothing
    
    Case DYN_ANNXD_TMPLT
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_DYN_TMPLT, strDataSrch)
    
    Case DYN_ANNXD_UI_CMG
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_UI_CMG, strDataSrch)
    
    Case DYN_ANNXD_UI_IFUR
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_UI_IFUR, strDataSrch)
    
    Case DYN_MSG_PDF
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_DYN_MSG_PDF, strDataSrch)
    
    Case DYN_REASON
        Set dataCache = WS_DC_DYN_REASON(strDataSrch)
    
        GET_DATA_CACHE.dataDescription = dataCache.getDescription
        
        If (dataCache.getExtraParamsFlag) Then
            GET_DATA_CACHE.EXTRAPARAMS = dataCache.getExtraParams
            GET_DATA_CACHE.FLG_EXTRAPARAMS = True
        End If
    
        Set dataCache = Nothing
    
    Case FLG_DOM
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_DOM, strDataSrch)
    
    Case FORNITURA_MASTER_PHASE01
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_FORN_MASTER_PHASE01, strDataSrch)
        
    Case FORNITURA_MASTER_PHASE02_F
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_FORN_MASTER_PHASE02_F, strDataSrch)
        
    Case FORNITURA_MASTER_PHASE02_P
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_FORN_MASTER_PHASE02_P, strDataSrch)
        
    Case FORNITURA_MASTER_QUERIES
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_FORN_MASTER_QUERIES, strDataSrch)
        
    Case INDE
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_INDE, strDataSrch)
        
    Case LOC_UBI
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_LOC_UBI, strDataSrch)
        
    Case NATIONALITY
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_NATIONALITY, strDataSrch)
        
    Case TOTCONG_TICSI
        GET_DATA_CACHE.dataDescription = WS_DC_TOTCONG_TICSI(strDataSrch)
    
    Case WMS
        GET_DATA_CACHE = RTN_DICOTOMICSEARCH(WS_DC_WMS, strDataSrch)
        
    End Select
    
ErrHandler:
    If (Trim$(GET_DATA_CACHE.dataDescription) = "") Then GET_DATA_CACHE.dataDescription = strDataSrch

    Set dataCache = Nothing

End Function

Private Sub GET_DATA_CACHE_LOAD(WS_DATA_CACHE() As strct_DATA, decodeIdx As enum_DATASECTION)

    Dim I   As Long
    Dim RS  As ADODB.Recordset

    I = -1

    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, STR_DECODER, STR_DECODER_EXTRA" & _
                            " FROM EST_WABBNABDECODER" & _
                            " WHERE (ID_DECODERTYPE = " & decodeIdx & ")" & _
                            " ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")

    If RS.RecordCount > 0 Then
        ReDim WS_DATA_CACHE(RS.RecordCount - 1)

        Do Until RS.EOF
            I = (I + 1)
            WS_DATA_CACHE(I).DATAID = RS("ID_DECODER")
            WS_DATA_CACHE(I).dataDescription = RS("STR_DECODER")

            If (Not IsNull(RS("STR_DECODER_EXTRA"))) Then
                WS_DATA_CACHE(I).FLG_EXTRAPARAMS = True
                WS_DATA_CACHE(I).EXTRAPARAMS = Split(RS("STR_DECODER_EXTRA"), "|")
            End If

            If ((RS.AbsolutePosition Mod 1000) = 0) Then DoEvents

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

End Sub

Private Sub GET_DATA_COLLECTION_LOAD(WS_DATA_COLLECTION As Collection, decodeIdx As enum_DATASECTION)

    Dim RS            As ADODB.Recordset
    Dim WS_COLL_KEY   As String
    Dim WS_DATA()     As String
    Dim WS_DATA_CACHE As cls_DataCache

    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, STR_DECODER, STR_DECODER_EXTRA" & _
                            " FROM EST_WABBNABDECODER" & _
                            " WHERE (ID_DECODERTYPE = " & decodeIdx & ")" & _
                            " ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")

    If RS.RecordCount > 0 Then
        Set WS_DATA_COLLECTION = New Collection

        Do Until RS.EOF
            Set WS_DATA_CACHE = New cls_DataCache

            WS_COLL_KEY = RS("ID_DECODER")
            WS_DATA_CACHE.setDescrition = RS("STR_DECODER")

            If (Not IsNull(RS("STR_DECODER_EXTRA"))) Then
                WS_DATA = Split(RS("STR_DECODER_EXTRA"), "|")
                WS_DATA_CACHE.setExtraParams = WS_DATA

                Erase WS_DATA()
            End If

            WS_DATA_COLLECTION.Add WS_DATA_CACHE, WS_COLL_KEY

            If ((RS.AbsolutePosition Mod 1000) = 0) Then DoEvents

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

End Sub

Public Function GET_EXTERNALDATA(WS_EXTDATAFILE As String, WS_EXTDATA() As strct_DATA, WS_CNTR As Integer) As Boolean

    On Error GoTo ErrHandler

    Dim I                       As Long
    Dim intFileNumber           As Integer
    Dim StrIn                   As String
    Dim WS_DATA()               As String
    Dim WS_EXTDATA_CNTR         As Long

    intFileNumber = FreeFile
    WS_EXTDATA_CNTR = WS_CNTR

    Open WS_EXTDATAFILE For Input As #intFileNumber
        Do Until EOF(intFileNumber)
            Line Input #intFileNumber, StrIn

            If ((Trim$(StrIn) <> "") And (Left$(StrIn, 2) <> "<!")) Then
                If (Left$(StrIn, 1) = "<") Then
                    WS_EXTDATA_CNTR = (WS_EXTDATA_CNTR + 1)
                    ReDim Preserve WS_EXTDATA(WS_EXTDATA_CNTR)

                    If (InStrRev(StrIn, "|") > 0) Then
                        WS_DATA = Split(Mid$(StrIn, 2, Len(StrIn) - 2), "|")

                        WS_EXTDATA(WS_EXTDATA_CNTR).DATAID = WS_DATA(0)

                        ReDim WS_EXTDATA(WS_EXTDATA_CNTR).EXTRAPARAMS(UBound(WS_DATA) - 1)

                        For I = 1 To UBound(WS_DATA)
                            WS_EXTDATA(WS_EXTDATA_CNTR).EXTRAPARAMS(I - 1) = WS_DATA(I)
                        Next I
                    Else
                        WS_EXTDATA(WS_EXTDATA_CNTR).DATAID = Mid$(StrIn, 2, Len(StrIn) - 2)
                    End If
                Else
                    With WS_EXTDATA(WS_EXTDATA_CNTR)
                        .dataDescription = .dataDescription & Replace$(StrIn, vbTab, "")
                    End With
                End If
            End If
        Loop
    Close #intFileNumber

    GET_QS_DATA WS_EXTDATA

    GET_EXTERNALDATA = True

    Exit Function

ErrHandler:
    Erase WS_EXTDATA

    GET_EXTERNALDATA = False

End Function

Public Function GET_EXTERNALINFO(strPath As String) As String
    
    On Error GoTo ErrHandler
    
    Dim I             As Integer
    Dim IDX           As Integer
    Dim intFileNumber As Integer
    Dim WS_DATA()     As String
    
    If Dir$(strPath) = "" Then Exit Function
    
    intFileNumber = FreeFile
    
    Open strPath For Input As #intFileNumber
        GET_EXTERNALINFO = Input(LOF(intFileNumber), #intFileNumber)
    Close #intFileNumber
    
    If (Trim$(GET_EXTERNALINFO) <> "") Then
        GET_EXTERNALINFO = Replace$(GET_EXTERNALINFO, vbTab, "")
        GET_EXTERNALINFO = Replace$(GET_EXTERNALINFO, vbNewLine, "")
    
        WS_DATA = Split(GET_EXTERNALINFO, "<!--")
        
        If (UBound(WS_DATA) > 0) Then
            GET_EXTERNALINFO = ""
            
            For I = 0 To UBound(WS_DATA)
                IDX = InStr(1, WS_DATA(I), "-->")
        
                If (IDX = 0) Then
                    GET_EXTERNALINFO = GET_EXTERNALINFO & WS_DATA(I)
                Else
                    GET_EXTERNALINFO = GET_EXTERNALINFO & Mid$(WS_DATA(I), (IDX + 3))
                End If
            Next I
        End If
    End If
    
    Erase WS_DATA()
    
    Exit Function

ErrHandler:
    Close #intFileNumber
    
    Erase WS_DATA()
    
    GET_EXTERNALINFO = ""

End Function

Public Sub GET_QS_GDF_QF(ByRef WS_ARRAY() As strct_GDF_Data)

    Dim I               As Long
    Dim INDEXLEFT       As Long
    Dim INDEXRIGHT      As Long
    Dim J               As Long
    Dim STACKLEFT(32)   As Long
    Dim STACKPOINTER    As Integer
    Dim STACKRIGHT(32)  As Long
    Dim TEMP            As strct_GDF_Data
    Dim VALUE           As String

    ' INIT POINTERS
    '
    INDEXLEFT = 0
    INDEXRIGHT = UBound(WS_ARRAY)
    STACKPOINTER = 1
    STACKLEFT(STACKPOINTER) = INDEXLEFT
    STACKRIGHT(STACKPOINTER) = INDEXRIGHT

    Do
        If (INDEXRIGHT > INDEXLEFT) Then
            VALUE = WS_ARRAY(INDEXRIGHT).SORT_KEY
            I = (INDEXLEFT - 1)
            J = INDEXRIGHT

            ' FIND THE PIVOT ITEM
            '
            Do
                Do: I = I + 1: Loop Until (WS_ARRAY(I).SORT_KEY >= VALUE)
                Do: J = J - 1: Loop Until ((J = INDEXLEFT) Or (WS_ARRAY(J).SORT_KEY <= VALUE))

                TEMP = WS_ARRAY(I)
                WS_ARRAY(I) = WS_ARRAY(J)
                WS_ARRAY(J) = TEMP
            Loop Until J <= I

            ' SWAP FOUND ITEMS
            '
            TEMP = WS_ARRAY(J)
            WS_ARRAY(J) = WS_ARRAY(I)
            WS_ARRAY(I) = WS_ARRAY(INDEXRIGHT)
            WS_ARRAY(INDEXRIGHT) = TEMP

            ' PUSH ON THE STACK THE PAIR OF POINTERS THAT DIFFER MOST
            '
            STACKPOINTER = (STACKPOINTER + 1)

            If ((I - INDEXLEFT) > (INDEXRIGHT - I)) Then
                STACKLEFT(STACKPOINTER) = INDEXLEFT
                STACKRIGHT(STACKPOINTER) = (I - 1)
                INDEXLEFT = (I + 1)
            Else
                STACKLEFT(STACKPOINTER) = (I + 1)
                STACKRIGHT(STACKPOINTER) = INDEXRIGHT
                INDEXRIGHT = (I - 1)
            End If
        Else
            INDEXLEFT = STACKLEFT(STACKPOINTER)
            INDEXRIGHT = STACKRIGHT(STACKPOINTER)
            STACKPOINTER = (STACKPOINTER - 1)

            If STACKPOINTER = 0 Then Exit Do
        End If
    Loop

End Sub

Public Sub GET_QS_DATA(ByRef WS_ARRAY() As strct_DATA)

    Dim I               As Long
    Dim INDEXLEFT       As Long
    Dim INDEXRIGHT      As Long
    Dim J               As Long
    Dim STACKLEFT(32)   As Long
    Dim STACKPOINTER    As Integer
    Dim STACKRIGHT(32)  As Long
    Dim TEMP            As strct_DATA
    Dim VALUE           As String

    ' INIT POINTERS
    '
    INDEXLEFT = 0
    INDEXRIGHT = UBound(WS_ARRAY)
    STACKPOINTER = 1
    STACKLEFT(STACKPOINTER) = INDEXLEFT
    STACKRIGHT(STACKPOINTER) = INDEXRIGHT

    Do
        If (INDEXRIGHT > INDEXLEFT) Then
            VALUE = WS_ARRAY(INDEXRIGHT).DATAID
            I = (INDEXLEFT - 1)
            J = INDEXRIGHT

            ' FIND THE PIVOT ITEM
            '
            Do
                Do: I = I + 1: Loop Until (WS_ARRAY(I).DATAID >= VALUE)
                Do: J = J - 1: Loop Until ((J = INDEXLEFT) Or (WS_ARRAY(J).DATAID <= VALUE))

                TEMP = WS_ARRAY(I)
                WS_ARRAY(I) = WS_ARRAY(J)
                WS_ARRAY(J) = TEMP
            Loop Until J <= I

            ' SWAP FOUND ITEMS
            '
            TEMP = WS_ARRAY(J)
            WS_ARRAY(J) = WS_ARRAY(I)
            WS_ARRAY(I) = WS_ARRAY(INDEXRIGHT)
            WS_ARRAY(INDEXRIGHT) = TEMP

            ' PUSH ON THE STACK THE PAIR OF POINTERS THAT DIFFER MOST
            '
            STACKPOINTER = (STACKPOINTER + 1)

            If ((I - INDEXLEFT) > (INDEXRIGHT - I)) Then
                STACKLEFT(STACKPOINTER) = INDEXLEFT
                STACKRIGHT(STACKPOINTER) = (I - 1)
                INDEXLEFT = (I + 1)
            Else
                STACKLEFT(STACKPOINTER) = (I + 1)
                STACKRIGHT(STACKPOINTER) = INDEXRIGHT
                INDEXRIGHT = (I - 1)
            End If
        Else
            INDEXLEFT = STACKLEFT(STACKPOINTER)
            INDEXRIGHT = STACKRIGHT(STACKPOINTER)
            STACKPOINTER = (STACKPOINTER - 1)

            If STACKPOINTER = 0 Then Exit Do
        End If
    Loop

End Sub

Public Function GET_SERVIZIODEPURAZIONE() As String
 
    Dim I As Integer
                
    GET_SERVIZIODEPURAZIONE = ""
 
    For I = 0 To UBound(WS_G12)
        If (WS_G12(I).TIPOPER = "2") Then
            Select Case WS_G12(I).INFO_SERVIZIO_DEPURAZIONE
            Case "A"
                GET_SERVIZIODEPURAZIONE = "È servito da un impianto di depurazione attivo"
            
            Case "B"
                GET_SERVIZIODEPURAZIONE = "Non è servito da un impianto di depurazione attivo per il quale sia in corso attività di progettazione, realizzazione, completamento o attivazione come da programma di cui all'articolo 3 del d.m. 30 settembre 2009"
            
            Case "C"
                GET_SERVIZIODEPURAZIONE = "Non è servito perché l'impianto di depurazione risulta temporaneamente inattivo o è stato temporaneamente inattivo"
            
            Case "D"
                GET_SERVIZIODEPURAZIONE = "Non è servito da un impianto di depurazione attivo per il quale non è in corso alcuna attività di progettazione, completamento o attivazione come da programma di cui all'articolo 3 del d.m. 30 settembre 2009"
            
            End Select
            
            If (GET_SERVIZIODEPURAZIONE <> "") Then Exit For
        End If
    Next I
    
    If ((UBound(WS_G12) = 0) And GET_SERVIZIODEPURAZIONE = "") Then GET_SERVIZIODEPURAZIONE = "È servito da un impianto di depurazione attivo"
    
    GET_SERVIZIODEPURAZIONE = GET_SERVIZIODEPURAZIONE & ".<br>Ulteriori informazioni sono disponibili nel sito; si rinvia anche all'art.8 del D.M. 30/09/2009."
    
End Function

Public Function GET_STABO_INFO_BOL() As Boolean

    On Error GoTo ErrHandler
    
    Dim FileLenInfo        As Long
    Dim FileLocInfo        As Long
    Dim myAPB              As cls_APB
    Dim ROWHEADER          As strct_STABODF_Header
    Dim RS                 As ADODB.Recordset
    Dim StrIn              As String
    Dim WS_BDS             As Boolean
    Dim WS_CHK_GSE         As Boolean
    Dim WS_CNTR_ROWS       As Long
    Dim WS_CODSER          As String
    Dim WS_COLL_KEY        As String
    Dim WS_COLL_VALUE      As String
    Dim WS_FLG_NOTACREDITO As Boolean
    Dim WS_FLG_PARTITE     As Boolean
    Dim WS_FLG_PROCESS     As Boolean
    Dim WS_STRING          As String
    Dim WS_TRANSACTION     As Boolean
    Dim WS_WMS_KEY_TMP     As String
    Dim WS_WMS_ROWS        As Long
    
    Set myAPB = New cls_APB
    
    With myAPB
        .APBMode = PBSingle
        .APBCaption = "Get STABO Info:"
        .APBMaxItems = 1
        .APBShow
    End With
        
    ' 00
    '
    WS_STRING = String$(Len(WS_G00), " ")
    CopyMemory ByVal VarPtr(WS_G00), ByVal StrPtr(WS_STRING), Len(WS_G00) * 2
    
    ' 01
    '
    WS_STRING = String$(Len(WS_G01), " ")
    CopyMemory ByVal VarPtr(WS_G01), ByVal StrPtr(WS_STRING), Len(WS_G01) * 2
    
    ' 02
    '
    WS_STRING = String$(Len(WS_G02), " ")
    CopyMemory ByVal VarPtr(WS_G02), ByVal StrPtr(WS_STRING), Len(WS_G02) * 2
    
    ' 09
    '
    WS_STRING = String$(Len(WS_G09), " ")
    CopyMemory ByVal VarPtr(WS_G09), ByVal StrPtr(WS_STRING), Len(WS_G09) * 2
    
    ' EXECUTE
    '
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Caching STABO File Data -> START"
    
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 1"
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 2"
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 3"
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 4"
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 10"
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 12"
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 14"
    
    DBConn.BeginTrans
    
    WS_TRANSACTION = True

    Open DLLParams.INPUTFILENAME For Input As #1
        FileLenInfo = (LOF(1) \ 1024)
    
        myAPB.APBMaxItems = FileLenInfo
    
        Do Until EOF(1)
            Line Input #1, StrIn
    
            FileLocInfo = (Loc(1) \ 8)
            myAPB.APBItemsLabel = IIf((WS_CNTR_ROWS = 0), "", "Items: " & Format$(WS_CNTR_ROWS, "##,##") & " - ") & "Read " & Format$(FileLocInfo, "##,##") & " KB of " & Format$(FileLenInfo, "##,##") & " KB"
    
            If ((FileLocInfo > 0) And (FileLenInfo > 0)) Then myAPB.APBItemsProgress = FileLocInfo
           
            CopyMemory ByVal VarPtr(ROWHEADER), ByVal StrPtr(StrIn), Len(ROWHEADER) * 2
            StrIn = Mid$(StrIn, 6)
               
            Select Case ROWHEADER.GROUP
               Case "00"
                If (ROWHEADER.ROWNUMBER = "000") Then
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G00), " ", False)
                    CopyMemory ByVal VarPtr(WS_G00), ByVal StrPtr(StrIn), Len(WS_G00) * 2
                End If
               
            Case "01"
                Select Case ROWHEADER.ROWNUMBER
                   Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R001), ByVal StrPtr(StrIn), Len(WS_G01.R001) * 2
                               
                    WS_BDS = (WS_G01.R001.TIPOSERVIZIO = "05")
                    WS_FLG_PARTITE = (WS_G01.R001.TIPOBOLLETTAZIONE = "P")
                    WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)
               
                Case "004"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R004), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R004), ByVal StrPtr(StrIn), Len(WS_G01.R004) * 2
                    
                    WS_FLG_NOTACREDITO = (Trim$(WS_G01.R004.DESCRIZIONE) <> "")
               
                Case "006"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R006), ByVal StrPtr(StrIn), Len(WS_G01.R006) * 2
                           
'                    If (WS_WMS_CHK = False) Then
'                        WS_WMS_ROWS = DB_GetValueByID("SELECT COUNT(*) AS NMR_ROWS FROM UTE.W_MESS_STORNO_335 WHERE BOT_PROGFAT = " & Trim$(WS_G01.R006.PROGRESSIVOFATTURAZIONE))
'
'                        WS_FLG_WMS = (WS_WMS_ROWS > 0)
'                        WS_WMS_CHK = True
'                    End If
                       
                Case "007"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R007), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R007), ByVal StrPtr(StrIn), Len(WS_G01.R007) * 2
                       
                    If (DLLParams.PLUGMODE = "SDI") Then
                        WS_FLG_PROCESS = ((WS_G01.R001.TIPONUMERAZIONE = "5") Or (Trim$(WS_G01.R007.MODALITÀ_INVIO) = "XSPDF"))
                    Else
                        WS_FLG_PROCESS = True
                    End If
                       
                End Select
               
            Case "02"
                Select Case ROWHEADER.ROWNUMBER
                   Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G02.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G02.R001), ByVal StrPtr(StrIn), Len(WS_G02.R001) * 2
                       
                    WS_CODSER = Trim$(WS_G02.R001.CODICESERVIZIO)
                       
                    If (WS_FLG_PROCESS) Then
                        On Error Resume Next
                       
                        If (WS_G01.R001.TIPONUMERAZIONE = "5") Then DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES(" & WS_CODSER & ",2,'TRG_PARAM')"
                           
                        DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES(" & WS_CODSER & ",3,'TRG_PARAM')"
                       
                        On Error GoTo ErrHandler
                    End If
                   
                End Select
               
            Case "09"
                Select Case ROWHEADER.ROWNUMBER
                    Case "003"
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G09.R003), " ", False)
                        CopyMemory ByVal VarPtr(WS_G09.R003), ByVal StrPtr(StrIn), Len(WS_G09.R003) * 2
                        
                        If (WS_FLG_PROCESS) Then
                            If ((WS_G01.R001.TIPOBOLLETTAZIONE <> "P") And (WS_BDS = False)) Then
                                If (GET_DATA_CACHE(CATEGORIE, Format$(WS_G09.R003.CODICECATEGORIAUTENZA, "000")).dataDescription = "TRG_CATEGORIA") Then
                                    On Error Resume Next
                                    
                                    DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES(" & WS_CODSER & ",1,'TRG_PARAM')"
                                    
                                    On Error GoTo ErrHandler
                                End If
                            End If
                        End If
    
                        If ((WS_FLG_NOTACREDITO = False) And (WS_FLG_PARTITE = False) And (Trim$(WS_G09.R003.CODICECATEGORIAUTENZA) = "119")) Then
                            On Error Resume Next
                        
                            DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES('" & Format$(WS_CODSER, "0000000000") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000") & "',12,'TRG_PARAM')"
                        
                            On Error GoTo ErrHandler
                        
                            If (WS_CHK_GSE = False) Then WS_CHK_GSE = True
                        End If
                                    
                        On Error Resume Next
                    
                        DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES('" & Format$(WS_CODSER, "0000000000") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000") & "',14,'TRG_PARAM')"
                    
                        On Error GoTo ErrHandler
                                    
                End Select
            Case "BS"
                If ((ROWHEADER.ROWNUMBER = "001") And WS_FLG_PROCESS) Then
                    If (WS_WMS_ROWS > 0) Then
                        On Error Resume Next
                       
                        DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES('" & WS_G00.AZIENDASIU & "_" & Format$(WS_G01.R001.SEZIONALE, "00") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & Format$(WS_G01.R001.CODSERVIZIONUMERAZIONE, "00") & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000") & "_" & WS_G01.R001.RATABOLLETTA & "',4,'TRG_PARAM')"
                       
                        On Error GoTo ErrHandler
                    End If
                   
                    On Error Resume Next
                   
                    DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES('" & WS_G00.AZIENDASIU & "_" & Format$(WS_CODSER, "0000000000") & "_" & Format$(WS_G01.R006.PROGRESSIVOFATTURAZIONE, "000000") & "_" & Format$(WS_G01.R001.SEZIONALE, "00") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & Format$(WS_G01.R001.CODSERVIZIONUMERAZIONE, "00") & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000") & "_" & WS_G01.R001.RATABOLLETTA & "',10,'TRG_PARAM')"
                       
                    On Error GoTo ErrHandler
                End If
                   
            End Select
           
            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents
       Loop
    Close #1
    
    DBConn.CommitTrans
    
    WS_TRANSACTION = False

'    GoTo CONTINUE

    If (WS_BDS = False) Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Cached STABO File Data -> END"

        ' GET DB DATA
        '
        ' WS_DC_FORN_MASTER_PHASE01
        '
        If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
        myAPB.APBItemsLabel = "Querying NET@H2O Data (FORN_MASTER_PHASE01)..."

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (FORN_MASTER_PHASE01) -> START"

        Set RS = DBConn.Execute("SELECT ID_DECODER, CENTRALE, CLIENTE, ISMASTER, PUNTOEROGA, MATRICOLA_CONTATORE FROM MMS.VIEW_WABBNAH2ODIV ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")

        If (RS.RecordCount > 0) Then
            WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

            ReDim WS_DC_FORN_MASTER_PHASE01(RS.RecordCount - 1) As strct_DATA

            myAPB.APBMaxItems = RS.RecordCount

            WS_CNTR_ROWS = 0

            Do Until RS.EOF
                myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
                myAPB.APBItemsProgress = RS.AbsolutePosition

                With WS_DC_FORN_MASTER_PHASE01(WS_CNTR_ROWS)
                    .DATAID = RS("ID_DECODER")
                    .dataDescription = "TRG_PARAM"
                    .EXTRAPARAMS = Split(RS("CENTRALE") & "|" & RS("CLIENTE") & "|" & RS("ISMASTER") & "|" & RS("PUNTOEROGA") & "|" & RS("MATRICOLA_CONTATORE"), "|")
                End With

                If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

                WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

                RS.MoveNext
            Loop
        End If

        RS.Close

        Set RS = Nothing

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (FORN_MASTER_PHASE01) -> END"

        ' WS_DC_FORN_MASTER_PHASE02_F
        '
        If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
        myAPB.APBItemsLabel = "Querying NET@H2O Data (FORN_MASTER_PHASE02_F)..."

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_FORN_MASTER_PHASE02_F) -> START"

        WS_STRING = GET_DATA_CACHE(FORNITURA_MASTER_QUERIES, "WS_DC_FORN_MASTER_PHASE02_F").dataDescription
        WS_STRING = Replace$(WS_STRING, ":DT_EMISS", "'" & Trim$(WS_G01.R001.DATAEMISSIONE & "'"))
        WS_STRING = Replace$(WS_STRING, ":PROGFAT", Trim$(WS_G01.R006.PROGRESSIVOFATTURAZIONE))

        Set RS = DBConn.Execute(WS_STRING)

        If (RS.RecordCount > 0) Then
            WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

            ReDim WS_DC_FORN_MASTER_PHASE02_F(RS.RecordCount - 1) As strct_DATA

            myAPB.APBMaxItems = RS.RecordCount

            WS_CNTR_ROWS = 0

            Do Until RS.EOF
                myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
                myAPB.APBItemsProgress = RS.AbsolutePosition

                With WS_DC_FORN_MASTER_PHASE02_F(WS_CNTR_ROWS)
                    .DATAID = RS("ID_DECODER")
                    .dataDescription = "TRG_PARAM"
                    .EXTRAPARAMS = Split(RS("TIP_TIPORIPARTO") & "|" & RS("DATA_LETTPRECMASTER") & "|" & RS("LETT_PRECMASTER") & "|" & RS("TIPOLETT_PRECMASTER") & "|" & RS("DATA_LETTATTMASTER") & "|" & RS("LETT_ATTMASTER") & "|" & RS("TIPOLETT_ATTMASTER") & "|" & RS("CONSUMO_LETTATTMASTER") & "|" & RS("ECCEDENZA") & "|" & RS("CONSUMI_DIVISIONALI") & "|" & RS("TOTDIVNEG"), "|")
                End With

                If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

                WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

                RS.MoveNext
            Loop
        End If

        RS.Close

        Set RS = Nothing

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_FORN_MASTER_PHASE02_F) -> END"

        ' WS_DC_FORN_MASTER_PHASE02_P
        '
        If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
        myAPB.APBItemsLabel = "Querying NET@H2O Data (FORN_MASTER_PHASE02_P)..."

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_FORN_MASTER_PHASE02_P) -> START"

        WS_STRING = GET_DATA_CACHE(FORNITURA_MASTER_QUERIES, "WS_DC_FORN_MASTER_PHASE02_P").dataDescription
        WS_STRING = Replace$(WS_STRING, ":DT_EMISS", "'" & Trim$(WS_G01.R001.DATAEMISSIONE & "'"))
        WS_STRING = Replace$(WS_STRING, ":PROGFAT", Trim$(WS_G01.R006.PROGRESSIVOFATTURAZIONE))

        Set RS = DBConn.Execute(WS_STRING)

        If (RS.RecordCount > 0) Then
            WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

            ReDim WS_DC_FORN_MASTER_PHASE02_P(RS.RecordCount - 1) As strct_DATA

            myAPB.APBMaxItems = RS.RecordCount

            WS_CNTR_ROWS = 0

            Do Until RS.EOF
                myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
                myAPB.APBItemsProgress = RS.AbsolutePosition

                With WS_DC_FORN_MASTER_PHASE02_P(WS_CNTR_ROWS)
                    .DATAID = RS("ID_DECODER")
                    .dataDescription = "TRG_PARAM"
                    .EXTRAPARAMS = Split(RS("TIP_TIPORIPARTO") & "|" & RS("DATA_LETTPRECMASTER") & "|" & RS("LETT_PRECMASTER") & "|" & RS("TIPOLETT_PRECMASTER") & "|" & RS("DATA_LETTATTMASTER") & "|" & RS("LETT_ATTMASTER") & "|" & RS("TIPOLETT_ATTMASTER") & "|" & RS("CONSUMO_LETTATTMASTER") & "|" & RS("ECCEDENZA") & "|" & RS("CONSUMI_DIVISIONALI") & "|" & RS("TOTDIVNEG"), "|")
                End With

                If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

                WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

                RS.MoveNext
            Loop
        End If

        RS.Close

        Set RS = Nothing

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_FORN_MASTER_PHASE02_P) -> END"
    End If

    ' WS_DC_CODICE_IPA
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_CODICE_IPA)..."

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_CODICE_IPA) -> START"

    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ VGC_PUNTOPRESA, VGC_CODICE_IPA FROM MMS.VIEW_WABBNAH2OBIPA ORDER BY NLSSORT(VGC_PUNTOPRESA, 'NLS_SORT=BINARY')")

    If (RS.RecordCount > 0) Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

        ReDim WS_DC_CODICE_IPA(RS.RecordCount - 1) As strct_DATA

        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = 0

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            With WS_DC_CODICE_IPA(WS_CNTR_ROWS)
                .DATAID = RS("VGC_PUNTOPRESA")
                .dataDescription = RS("VGC_CODICE_IPA")
            End With

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_CODICE_IPA) -> END"

    ' WS_DC_LOC_UBI
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_LOC_UBI)..."

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_LOC_UBI) -> START"

    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, DCO_DES_30 FROM MMS.VIEW_WABBNAH2OBUBI ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")

    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

        ReDim WS_DC_LOC_UBI(RS.RecordCount - 1) As strct_DATA

        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = 0

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            With WS_DC_LOC_UBI(WS_CNTR_ROWS)
                .DATAID = RS("ID_DECODER")
                .dataDescription = RS("DCO_DES_30")
            End With

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_LOC_UBI) -> END"

    ' WS_DC_NATIONALITY
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_NATIONALITY)..."

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_NATIONALITY) -> START"

    Set RS = DBConn.Execute("SELECT DNZ_SIGNAZ, DNZ_DES_30 FROM UTE.DESNAZ WHERE (DNZ_UTE = '01') AND (DNZ_SIGNAZ <> 'IT') AND (DNZ_SIGNAZ <> 'ITA') ORDER BY NLSSORT(DNZ_SIGNAZ, 'NLS_SORT=BINARY')")

    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

        ReDim WS_DC_NATIONALITY(RS.RecordCount - 1) As strct_DATA

        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = 0

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            With WS_DC_NATIONALITY(WS_CNTR_ROWS)
                .DATAID = RS("DNZ_SIGNAZ")
                .dataDescription = RS("DNZ_DES_30")
            End With

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_NATIONALITY) -> END"

    ' WS_DC_WMS
    '
    If (WS_WMS_ROWS > 0) Then
        If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
        myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_WMS)..."

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_WMS) -> START"

        Set RS = DBConn.Execute("SELECT ID_DECODER, BET_ANNO, BET_TRSAP_NUMDOCSAP, BET_DATAEMIS FROM MMS.VIEW_WABBNAH2OWMS335 WHERE (BOT_PROGFAT = " & Trim$(WS_G01.R006.PROGRESSIVOFATTURAZIONE) & ") ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY'), BET_ANNO, BET_TRSAP_NUMDOCSAP, BET_DATAEMIS")

        If RS.RecordCount > 0 Then
            WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

            myAPB.APBMaxItems = RS.RecordCount

            WS_CNTR_ROWS = -1

            Do Until RS.EOF
                myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
                myAPB.APBItemsProgress = RS.AbsolutePosition

                WS_STRING = RS("BET_ANNO") & "/" & RS("BET_TRSAP_NUMDOCSAP") & "|" & Format$(RS("BET_DATAEMIS"), "dd/MM/yyyy")

                If (WS_WMS_KEY_TMP <> RS("ID_DECODER")) Then
                    If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

                    WS_WMS_KEY_TMP = RS("ID_DECODER")
                    WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

                    ReDim Preserve WS_DC_WMS(WS_CNTR_ROWS) As strct_DATA

                    With WS_DC_WMS(WS_CNTR_ROWS)
                        .DATAID = RS("ID_DECODER")
                        .dataDescription = WS_STRING
                    End With
                Else
                    WS_DC_WMS(WS_CNTR_ROWS).dataDescription = WS_DC_WMS(WS_CNTR_ROWS).dataDescription & vbNewLine & WS_STRING
                End If

                RS.MoveNext
            Loop
        End If

        RS.Close

        Set RS = Nothing

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_WMS) -> END"
    End If

    ' WS_DC_TOTCONG_TICSI
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_TOTCONG_TICSI)..."

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_TOTCONG_TICSI) -> START"

    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(8) */ ID_DECODER, BOT_TOTCONG_TICSI FROM MMS.VIEW_WABBNAH2O_TC_TICSI ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")

    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

        Set WS_DC_TOTCONG_TICSI = New Collection

        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = 0

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            WS_COLL_VALUE = RS("BOT_TOTCONG_TICSI")
            WS_COLL_KEY = RS("ID_DECODER")

            WS_DC_TOTCONG_TICSI.Add WS_COLL_VALUE, WS_COLL_KEY

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_TOTCONG_TICSI) -> END"

    ' WS_DC_DOM
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_DOM)..."

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_DOM) -> START"

    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, FLG_DOM FROM MMS.VIEW_WABBNAH2O_DOM ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")

    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

        ReDim WS_DC_DOM(RS.RecordCount - 1) As strct_DATA

        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = 0

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            With WS_DC_DOM(WS_CNTR_ROWS)
                .DATAID = RS("ID_DECODER")
                .dataDescription = RS("FLG_DOM")
            End With

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_TOTCONG_TICSI) -> END"

CONTINUE:

    ' WS_CHK_GSE
    '
    If (WS_CHK_GSE) Then
        ' WS_DC_CMG
        '
        If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
        myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_CMG)..."

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_CMG) -> START"

        Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, USO, PERIODO, NUMERO_UNITA_IMMOBILIARI, COMPONENTI_NUCLEO_FAMILIARE, CONSUMO_USO_LITRI_GIORNO_UNITA FROM MMS.VIEW_WABBNAH2O_CMG")

        If RS.RecordCount > 0 Then
            WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data (WS_DC_CMG) -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

            myAPB.APBMaxItems = RS.RecordCount

            WS_CNTR_ROWS = -1
            WS_WMS_KEY_TMP = ""

            Do Until RS.EOF
                myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
                myAPB.APBItemsProgress = RS.AbsolutePosition

                WS_STRING = RS("USO") & "|" & RS("PERIODO") & "|" & RS("NUMERO_UNITA_IMMOBILIARI") & "|" & RS("COMPONENTI_NUCLEO_FAMILIARE") & "|" & RS("CONSUMO_USO_LITRI_GIORNO_UNITA")

                If (WS_WMS_KEY_TMP <> RS("ID_DECODER")) Then
                    If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

                    WS_WMS_KEY_TMP = RS("ID_DECODER")
                    WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

                    ReDim Preserve WS_DC_UI_CMG(WS_CNTR_ROWS) As strct_DATA

                    With WS_DC_UI_CMG(WS_CNTR_ROWS)
                        .DATAID = RS("ID_DECODER")
                        .dataDescription = WS_STRING
                    End With
                Else
                    WS_DC_UI_CMG(WS_CNTR_ROWS).dataDescription = WS_DC_UI_CMG(WS_CNTR_ROWS).dataDescription & vbNewLine & WS_STRING
                End If

                RS.MoveNext
            Loop
        End If

        RS.Close

        Set RS = Nothing

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_CMG) -> END"

        ' WS_DC_IFUR
        '
        If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
        myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_IFUR)..."

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_IFUR) -> START"

        Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, DATA_EMISSIONE, NUMERO_DOCUMENTO, PERIODO_CONSUMI, TIPO_FATTURA, IMPORTO FROM MMS.VIEW_WABBNAH2O_IFUR")

        If RS.RecordCount > 0 Then
            WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data (WS_DC_IFUR) -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

            myAPB.APBMaxItems = RS.RecordCount

            WS_CNTR_ROWS = -1
            WS_WMS_KEY_TMP = ""

            Do Until RS.EOF
                myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
                myAPB.APBItemsProgress = RS.AbsolutePosition

                WS_STRING = RS("DATA_EMISSIONE") & "|" & RS("NUMERO_DOCUMENTO") & "|" & RS("PERIODO_CONSUMI") & "|" & RS("TIPO_FATTURA") & "|" & RS("IMPORTO")

                If (WS_WMS_KEY_TMP <> RS("ID_DECODER")) Then
                    If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents

                    WS_WMS_KEY_TMP = RS("ID_DECODER")
                    WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

                    ReDim Preserve WS_DC_UI_IFUR(WS_CNTR_ROWS) As strct_DATA

                    With WS_DC_UI_IFUR(WS_CNTR_ROWS)
                        .DATAID = RS("ID_DECODER")
                        .dataDescription = WS_STRING
                    End With
                Else
                    WS_DC_UI_IFUR(WS_CNTR_ROWS).dataDescription = WS_DC_UI_IFUR(WS_CNTR_ROWS).dataDescription & vbNewLine & WS_STRING
                End If

                RS.MoveNext
            Loop
        End If

        RS.Close

        Set RS = Nothing

        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_IFUR) -> END"
    End If
   
    ' WS_DC_CUR
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_DYN_MSG_PDF)..."

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_DYN_MSG_PDF) -> START"

    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, BILL_PERIOD FROM MMS.VIEW_WABBNAH2O_PDF ORDER BY ID_DECODER, BILL_PERIOD")

    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data (WS_DC_DYN_MSG_PDF) -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"

        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = -1
        WS_WMS_KEY_TMP = ""

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents
            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            ReDim Preserve WS_DC_DYN_MSG_PDF(WS_CNTR_ROWS) As strct_DATA

            With WS_DC_DYN_MSG_PDF(WS_CNTR_ROWS)
                .DATAID = RS("ID_DECODER")
                .dataDescription = RS("BILL_PERIOD")
            End With

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_DYN_MSG_PDF) -> END"
        
    '
    '
    myAPB.APBClose
    Set myAPB = Nothing

    GET_STABO_INFO_BOL = True

    Exit Function
    
ErrHandler:
    If (WS_TRANSACTION) Then DBConn.RollbackTrans

    Close #1

    myAPB.APBClose
    Set myAPB = Nothing

    MsgBox Err.Description, vbExclamation, "Guru Meditation:"

End Function

Public Function GET_STABO_INFO_LXX() As Boolean

    On Error GoTo ErrHandler

    Dim FileLenInfo    As Long
    Dim FileLocInfo    As Long
    Dim myAPB          As cls_APB
    Dim ROWHEADER      As strct_STABODF_Header
    Dim RS             As ADODB.Recordset
    Dim StrIn          As String
    Dim WS_CNTR_ROWS   As Long
    Dim WS_CODSER      As String
    Dim WS_FLG_PROCESS As Boolean
    Dim WS_STRING      As String
    Dim WS_TRANSACTION As Boolean
    
    Set myAPB = New cls_APB
    
    With myAPB
        .APBMode = PBSingle
        .APBCaption = "Get STABO Info:"
        .APBMaxItems = 1
        .APBShow
    End With
    
    ' 02
    '
    WS_STRING = String$(Len(WS_G02), " ")
    CopyMemory ByVal VarPtr(WS_G02), ByVal StrPtr(WS_STRING), Len(WS_G02) * 2
    
    ' EXECUTE
    '
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Caching STABO File Data -> START"
    
    DBConn.Execute "DELETE FROM MMS.EST_WABBNABDECODER WHERE ID_DECODERTYPE = 3"
    
    DBConn.BeginTrans

    WS_TRANSACTION = True

    Open DLLParams.INPUTFILENAME For Input As #1
        FileLenInfo = (LOF(1) \ 1024)

        myAPB.APBMaxItems = FileLenInfo

        Do Until EOF(1)
            Line Input #1, StrIn

            FileLocInfo = (Loc(1) \ 8)
            myAPB.APBItemsLabel = IIf((WS_CNTR_ROWS = 0), "", "Items: " & Format$(WS_CNTR_ROWS, "##,##") & " - ") & "Read " & Format$(FileLocInfo, "##,##") & " KB of " & Format$(FileLenInfo, "##,##") & " KB"

            If ((FileLocInfo > 0) And (FileLenInfo > 0)) Then myAPB.APBItemsProgress = FileLocInfo
        
            CopyMemory ByVal VarPtr(ROWHEADER), ByVal StrPtr(StrIn), Len(ROWHEADER) * 2
            StrIn = Mid$(StrIn, 6)
            
            Select Case ROWHEADER.GROUP
            Case "02"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G02.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G02.R001), ByVal StrPtr(StrIn), Len(WS_G02.R001) * 2
                    
                    If (WS_FLG_PROCESS) Then
                        WS_CODSER = Trim$(WS_G02.R001.CODICESERVIZIO)
                        
                        On Error Resume Next
                        
                        DBConn.Execute "INSERT INTO MMS.EST_WABBNABDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES(" & WS_CODSER & ",3,'TRG_PARAM')"
                    
                        On Error GoTo ErrHandler
                    End If
                
                End Select
            
            End Select
        
            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents
        Loop
    Close #1

    DBConn.CommitTrans

    WS_TRANSACTION = False

    ' WS_DC_LOC_UBI
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_LOC_UBI)..."
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_LOC_UBI) -> START"
    
    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ID_DECODER, DCO_DES_30 FROM MMS.VIEW_WABBNAH2OBUBI ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")
    
    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"
        
        ReDim WS_DC_LOC_UBI(RS.RecordCount - 1) As strct_DATA
        
        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = 0

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            With WS_DC_LOC_UBI(WS_CNTR_ROWS)
                .DATAID = RS("ID_DECODER")
                .dataDescription = RS("DCO_DES_30")
            End With

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents
            
            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_LOC_UBI) -> END"

    ' WS_DC_NATIONALITY
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (WS_DC_NATIONALITY)..."
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_DC_NATIONALITY) -> START"
    
    Set RS = DBConn.Execute("SELECT DNZ_SIGNAZ, DNZ_DES_30 FROM UTE.DESNAZ WHERE (DNZ_UTE = '01') AND (DNZ_SIGNAZ <> 'IT') AND (DNZ_SIGNAZ <> 'ITA') ORDER BY NLSSORT(DNZ_SIGNAZ, 'NLS_SORT=BINARY')")
    
    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"
        
        ReDim WS_DC_NATIONALITY(RS.RecordCount - 1) As strct_DATA
        
        myAPB.APBMaxItems = RS.RecordCount

        WS_CNTR_ROWS = 0

        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition

            With WS_DC_NATIONALITY(WS_CNTR_ROWS)
                .DATAID = RS("DNZ_SIGNAZ")
                .dataDescription = RS("DNZ_DES_30")
            End With

            If ((WS_CNTR_ROWS Mod 1000) = 0) Then DoEvents
            
            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)

            RS.MoveNext
        Loop
    End If

    RS.Close

    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_DC_NATIONALITY) -> END"

    ' END
    '
    myAPB.APBClose
    Set myAPB = Nothing

    GET_STABO_INFO_LXX = True

    Exit Function
    
ErrHandler:
    If (WS_TRANSACTION) Then DBConn.RollbackTrans

    Close #1

    myAPB.APBClose
    Set myAPB = Nothing

    MsgBox Err.Description, vbExclamation, "Guru Meditation:"

End Function

Public Function GET_TIPOFATTURA(ID_FATTURA As String) As String

    Select Case ID_FATTURA
    Case "A"
        GET_TIPOFATTURA = "di acconto"
    
    Case "D"
        GET_TIPOFATTURA = "di conguaglio + acconto"
    
    Case "F"
        GET_TIPOFATTURA = "di cessazione"
    
    Case "L"
        GET_TIPOFATTURA = "di conguaglio"
        
    Case "P"
        GET_TIPOFATTURA = "di sole partite varie"
    
    Case "S"
        GET_TIPOFATTURA = "multisito"
        
    End Select

End Function

Private Function RTN_DICOTOMICSEARCH(varARRAY() As strct_DATA, varSRCH As String) As strct_DATA
    
    Dim WS_CENTER   As Long
    Dim WS_COMPARE  As Long
    Dim WS_END      As Long
    Dim WS_START    As Long
    
    WS_END = UBound(varARRAY)
    
    While (WS_START <= WS_END)
        WS_CENTER = ((WS_START + WS_END) / 2)
        WS_COMPARE = StrComp(varSRCH, varARRAY(WS_CENTER).DATAID)
        
        If (WS_COMPARE < 0) Then
            WS_END = (WS_CENTER - 1)
        Else
            If (WS_COMPARE > 0) Then
                WS_START = (WS_CENTER + 1)
            Else
                RTN_DICOTOMICSEARCH = varARRAY(WS_CENTER)
                
                Exit Function
            End If
        End If
    Wend

End Function

Public Sub MMS_LOG_INSERT()

    DBConn.Execute "INSERT INTO MMS.EST_WABBNAH2OLOG(ID_WORKINGLOAD,STR_FILENAME,STR_MODE) VALUES(" & DLLParams.IDWORKINGLOAD & ",'" & DLLParams.INPUTFILENAME & "','B01')"

End Sub
