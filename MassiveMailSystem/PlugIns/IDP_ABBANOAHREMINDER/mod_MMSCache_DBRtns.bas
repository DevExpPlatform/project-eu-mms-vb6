Attribute VB_Name = "mod_MMSCache_DBRtns"
Option Explicit

Public Enum enum_DATASECTION
    EST_MSG_STAMP
    EST_DATA
    EST_DATA_ANA
    EST_DATA_STAMP
    EST_IMP_PRESCR
End Enum

Public Type strct_DATA
    DATAID                              As String
    DATADESCRIPTION                     As String
    EXTRAPARAMS()                       As String
    FLG_EXTRAPARAMS                     As Boolean
End Type

Private WS_EST_DATA()                   As strct_DATA
Private WS_EST_DATA_ANA()               As strct_DATA
Private WS_EST_DATA_STAMP()             As strct_DATA
Private WS_EST_MSG_STAMP()              As strct_DATA
Private WS_EST_IMP_PRESCR()             As strct_DATA

Public WS_CS_PXX_DOCUMENTDETAILS_LEGEND As strct_DATA
Public WS_CS_PXX_MSG_L17                As strct_DATA
Public WS_CUSTOMERSTYLE()               As strct_DATA
Public WS_FLG_EST_IMP_PRESCR            As Boolean
Public WS_GG_DESCR                      As String
Public WS_LOGFILEPATH                   As String
Public WS_PAGE_FOOTER                   As String

Public Sub CACHE_CLEAR()
    
    WS_FLG_EST_IMP_PRESCR = False
    
    Erase WS_CUSTOMERSTYLE()
    Erase WS_EST_DATA()
    Erase WS_EST_DATA_ANA()
    Erase WS_EST_DATA_STAMP()

End Sub

Public Sub CACHE_INIT()
    
    ' EXTERNAL CONFIGS/PARAMS
    '
    GET_EXTERNALDATA DLLParams.EXTRASPATH & "T" & Mid$(DLLParams.LAYOUT, 2) & "_External.STL", WS_CUSTOMERSTYLE, -1
    
    ' COMMUNICATIONS
    '
    WS_CS_PXX_DOCUMENTDETAILS_LEGEND = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_DOCUMENTDETAILS_LEGEND")
    
    ' DATA CACHE
    '
    GET_DATACACHELOAD WS_EST_MSG_STAMP, EST_MSG_STAMP           ' ID_DECODERTYPE = 0
    
    Select Case DLLParams.LAYOUT
    Case "L03"
        GET_DATACACHELOAD WS_EST_IMP_PRESCR, EST_IMP_PRESCR     ' ID_DECODERTYPE = 4
    
        WS_FLG_EST_IMP_PRESCR = CHK_EST_IMP_PRESCR()
    
    Case "L17"
        WS_CS_PXX_MSG_L17 = RTN_DICOTOMICSEARCH(WS_CUSTOMERSTYLE, "PXX_MSG_L17")
        
    End Select

End Sub

Private Function CHK_EST_IMP_PRESCR() As Boolean

    On Error GoTo ErrHandler

    Dim I As Integer

    I = UBound(WS_EST_IMP_PRESCR)

    CHK_EST_IMP_PRESCR = True

ErrHandler:

End Function

Public Function GET_DATACACHE(dataType As enum_DATASECTION, strDataSrch As String) As strct_DATA

    On Error GoTo ErrHandler

    strDataSrch = Trim$(strDataSrch)

    Select Case dataType
    Case EST_DATA
        GET_DATACACHE = RTN_DICOTOMICSEARCH(WS_EST_DATA, strDataSrch)
        
    Case EST_DATA_ANA
        GET_DATACACHE = RTN_DICOTOMICSEARCH(WS_EST_DATA_ANA, strDataSrch)
        
    Case EST_DATA_STAMP
        GET_DATACACHE = RTN_DICOTOMICSEARCH(WS_EST_DATA_STAMP, strDataSrch)
        
    Case EST_IMP_PRESCR
        GET_DATACACHE = RTN_DICOTOMICSEARCH(WS_EST_IMP_PRESCR, strDataSrch)
    
    Case EST_MSG_STAMP
        GET_DATACACHE = RTN_DICOTOMICSEARCH(WS_EST_MSG_STAMP, strDataSrch)
        
    End Select
    
ErrHandler:
    If (Trim$(GET_DATACACHE.DATADESCRIPTION) = "") Then GET_DATACACHE.DATADESCRIPTION = strDataSrch

End Function

Private Sub GET_DATACACHELOAD(WS_DATACACHE() As strct_DATA, decodeIdx As enum_DATASECTION)
    
    Dim I   As Long
    Dim RS  As ADODB.Recordset
    
    I = -1
    
    Set RS = DBConn.Execute("SELECT ID_DECODER, STR_DECODER, STR_DECODER_EXTRA" & _
                            " FROM EST_WABBNARDECODER" & _
                            " WHERE (ID_DECODERTYPE = " & decodeIdx & ")" & _
                            " ORDER BY NLSSORT(ID_DECODER, 'NLS_SORT=BINARY')")
    
    If RS.RecordCount > 0 Then
        ReDim WS_DATACACHE(RS.RecordCount - 1)
        
        Do Until RS.EOF
            I = (I + 1)
            WS_DATACACHE(I).DATAID = RS("ID_DECODER")
            WS_DATACACHE(I).DATADESCRIPTION = RS("STR_DECODER")
            
            If (Not IsNull(RS("STR_DECODER_EXTRA"))) Then
                WS_DATACACHE(I).FLG_EXTRAPARAMS = True
                WS_DATACACHE(I).EXTRAPARAMS = Split(RS("STR_DECODER_EXTRA"), "|")
            End If
            
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
                        .DATADESCRIPTION = .DATADESCRIPTION & Replace$(StrIn, vbTab, "")
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

Public Function GET_GARC_INFO() As Boolean
    
    On Error GoTo ErrHandler
    
    Dim FileLenInfo             As Long
    Dim FileLocInfo             As Long
    Dim myAPB                   As cls_APB
    Dim ROWHEADER               As strct_GARC_HEADER
    Dim RS                      As ADODB.Recordset
    Dim StrIn                   As String
    Dim WS_01S_AN_R001          As strct_01S_AN_R001
    Dim WS_01S_ED_RXXX          As strct_01S_ED_RXXX
    Dim WS_CNTR_CODANA          As Long
    Dim WS_CNTR_ROWS            As Long
    Dim WS_CODANA               As String
    Dim WS_CODSER               As String
    Dim WS_LOTTO                As String
    Dim WS_ROW_DATA_KEY         As String
    Dim WS_ROW_DATA_KEY_TMP     As String
    Dim WS_STRING               As String
    Dim WS_TRANSACTION          As Boolean
    
    Set myAPB = New cls_APB
    
    With myAPB
        .APBMode = PBSingle
        .APBCaption = "Get GARC Info:"
        .APBMaxItems = 1
        .APBShow
    End With

    Erase WS_EST_DATA()
    Erase WS_EST_DATA_ANA()
    Erase WS_EST_DATA_STAMP()

    ReDim WS_01S_ED_RXXX.ED_RXXX_S_D(0)
    ReDim WS_01S_ED_RXXX.ED_RXXX_A_D(0)
    
'    WS_LOTTO = 52152
'    GoTo CONTINUE:
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Caching GARC File Data -> START"
    
    DBConn.Execute "TRUNCATE TABLE MMS.EST_WABBNAH2OREMINDERS DROP STORAGE"
    DBConn.BeginTrans
    
    WS_TRANSACTION = True

    Open DLLParams.INPUTFILENAME For Input As #1
        FileLenInfo = (LOF(1) \ 1024)

        myAPB.APBMaxItems = FileLenInfo

        Do Until EOF(1)
            Line Input #1, StrIn

            FileLocInfo = (Loc(1) \ 8)
            myAPB.APBItemsLabel = "Items: " & Format$(WS_CNTR_ROWS, "##,##") & " - Read " & Format$(FileLocInfo, "##,##") & " KB of " & Format$(FileLenInfo, "##,##") & " KB"

            If ((FileLocInfo > 0) And (FileLenInfo > 0)) Then myAPB.APBItemsProgress = FileLocInfo

            CopyMemory ByVal VarPtr(ROWHEADER), ByVal StrPtr(StrIn), Len(ROWHEADER) * 2
            StrIn = Mid$(StrIn, 9)
            
            Select Case ROWHEADER.SUBGROUP
            Case "AF"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_01S.AF_R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_01S.AF_R001), ByVal StrPtr(StrIn), Len(WS_01S.AF_R001) * 2
                    
                    WS_LOTTO = Trim$(WS_01S.AF_R001.CODICELOTTO)
                End If
            
            Case "AN"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    ReDim WS_01S_ED_RXXX.ED_RXXX_S_D(0)
                    
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_01S_AN_R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_01S_AN_R001), ByVal StrPtr(StrIn), Len(WS_01S_AN_R001) * 2
                
                    WS_CODANA = Trim$(WS_01S_AN_R001.CODICEANAGRAFICO)
                End If
            
            Case "ED"
                WS_STRING = Left$(StrIn, 2)
                StrIn = Mid$(StrIn, 3)
                
                With WS_01S_ED_RXXX
                    Select Case WS_STRING
                    Case "AC"
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_R001_A_C), " ", False)
                        CopyMemory ByVal VarPtr(.ED_R001_A_C), ByVal StrPtr(StrIn), Len(.ED_R001_A_C) * 2
                    
                        WS_CODSER = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)
                    
                    Case "AD"
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_RXXX_A_D(0)), " ", False)
                        CopyMemory ByVal VarPtr(.ED_RXXX_A_D(0)), ByVal StrPtr(StrIn), Len(.ED_RXXX_A_D(0)) * 2
                        
                        WS_ROW_DATA_KEY = WS_CODANA & "_" & WS_CODSER & "_" & .ED_RXXX_A_D(0).TIPODOCUMENTO & "_" & .ED_RXXX_A_D(0).ANNO & Format$(IIf((Trim$(.ED_RXXX_A_D(0).NUMERO) = ""), "0", Trim$(.ED_RXXX_A_D(0).NUMERO)), "00000000")
                    
                        If (WS_ROW_DATA_KEY_TMP <> WS_ROW_DATA_KEY) Then
                            DBConn.Execute "INSERT INTO MMS.EST_WABBNAH2OREMINDERS(ANA_CODANA,BOS_PUNTOPRESA,BOT_TIPODOC,BOT_ANNO,BOT_NUMBOLDOC) VALUES(" & WS_CODANA & "," & WS_CODSER & ",'" & .ED_RXXX_A_D(0).TIPODOCUMENTO & "'," & .ED_RXXX_A_D(0).ANNO & "," & IIf((Trim$(.ED_RXXX_A_D(0).NUMERO) = ""), "NULL", Trim$(.ED_RXXX_A_D(0).NUMERO)) & ")"
                            
                            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)
                            WS_ROW_DATA_KEY_TMP = WS_ROW_DATA_KEY
                        End If
                    
                    Case "SD"
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_RXXX_S_D(0)), " ", False)
                        CopyMemory ByVal VarPtr(.ED_RXXX_S_D(0)), ByVal StrPtr(StrIn), Len(.ED_RXXX_S_D(0)) * 2
                        
                        WS_ROW_DATA_KEY = WS_CODANA & "_" & WS_CODSER & "_" & .ED_RXXX_S_D(0).TIPODOCUMENTO & "_" & .ED_RXXX_S_D(0).ANNO & Format$(IIf((Trim$(.ED_RXXX_S_D(0).NUMERO) = ""), "0", Trim$(.ED_RXXX_S_D(0).NUMERO)), "00000000")
                    
                        If (WS_ROW_DATA_KEY_TMP <> WS_ROW_DATA_KEY) Then
                            DBConn.Execute "INSERT INTO MMS.EST_WABBNAH2OREMINDERS(ANA_CODANA,BOS_PUNTOPRESA,BOT_TIPODOC,BOT_ANNO,BOT_NUMBOLDOC) VALUES(" & WS_CODANA & "," & WS_CODSER & ",'" & .ED_RXXX_S_D(0).TIPODOCUMENTO & "'," & .ED_RXXX_S_D(0).ANNO & "," & IIf((Trim$(.ED_RXXX_S_D(0).NUMERO) = ""), "NULL", Trim$(.ED_RXXX_S_D(0).NUMERO)) & ")"
                            
                            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)
                            WS_ROW_DATA_KEY_TMP = WS_ROW_DATA_KEY
                        End If
                    
                    Case "SE"
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_R001_S_E), " ", False)
                        CopyMemory ByVal VarPtr(.ED_R001_S_E), ByVal StrPtr(StrIn), Len(.ED_R001_S_E) * 2
                        
                        WS_CODSER = Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    
                    End Select
                
                End With
            
            Case "IL"
                If (DLLParams.LAYOUT <> "L07") Then
                    If (ROWHEADER.ROWNUMBER = "001") Then
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_01S.IL_R001), " ", False)
                        CopyMemory ByVal VarPtr(WS_01S.IL_R001), ByVal StrPtr(StrIn), Len(WS_01S.IL_R001) * 2
                        
                        DBConn.Execute "INSERT INTO MMS.EST_WABBNARDECODER(ID_DECODER,ID_DECODERTYPE,STR_DECODER) VALUES('" & DLLParams.IDWORKINGLOAD & "_" & Format$(WS_LOTTO, "000000") & "_" & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000") & "',1,'TRG_PARAM')"
                    End If
                End If
                
            End Select
        Loop
    Close #1

    DBConn.CommitTrans

    WS_TRANSACTION = False

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Cached GARC File Data -> END"
    
    ' GET DB DATA
    '
    If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
    myAPB.APBItemsLabel = "Querying NET@H2O Data (Found " & Format$(WS_CNTR_ROWS, "##,##") & " Items)..."
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_EST_DATA_ANA/WS_EST_DATA) -> START"
    
    Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ ANA_KEY, ANA_CODFIS, ANA_PARIVA, TCA_DES_40, ROW_DATA_KEY, BOT_TRSAP_NUMDOCSAP, BOT_IMPTOTBOL, BOT_DATSCAEFF, STATO_SERV, NMR_SOLLECITI FROM MMS.VIEW_WABBNAH2OREMINDERS")
    
    If RS.RecordCount > 0 Then
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " of " & Format$(WS_CNTR_ROWS, "##,##") & " Elements - INFO"
   
        WS_CODANA = ""
        WS_CNTR_CODANA = -1
        WS_CNTR_ROWS = 0
        
        ReDim WS_EST_DATA(RS.RecordCount - 1) As strct_DATA
        
        myAPB.APBMaxItems = RS.RecordCount
        
        Do Until RS.EOF
            myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
            myAPB.APBItemsProgress = RS.AbsolutePosition
            
            ' WS_EST_DATA_ANA
            '
            If (WS_CODANA <> RS("ANA_KEY")) Then
                WS_CODANA = RS("ANA_KEY")
                WS_STRING = ""
                
                If (IsNull(RS("ANA_PARIVA"))) Then
                    If (IsNull(RS("ANA_CODFIS"))) Then
                        WS_STRING = "-"
                    Else
                        WS_STRING = RS("ANA_CODFIS")
                    End If
                Else
                    WS_STRING = Format$(RS("ANA_PARIVA"), "00000000000")
                End If
                
                WS_STRING = WS_STRING & "|"
                
                If (Not IsNull(RS("TCA_DES_40"))) Then WS_STRING = WS_STRING & RS("TCA_DES_40")
                
                WS_STRING = WS_STRING & "|" & RS("STATO_SERV")
                
                WS_CNTR_CODANA = (WS_CNTR_CODANA + 1)
                ReDim Preserve WS_EST_DATA_ANA(WS_CNTR_CODANA) As strct_DATA
                
                With WS_EST_DATA_ANA(WS_CNTR_CODANA)
                    .DATAID = WS_CODANA
                    .DATADESCRIPTION = "TRG_PARAM"
                    .EXTRAPARAMS = Split(WS_STRING, "|")
                End With
            End If
            
            ' WS_EST_DATA
            '
            WS_STRING = ""
            
            If (Not IsNull(RS("BOT_TRSAP_NUMDOCSAP"))) Then WS_STRING = RS("BOT_TRSAP_NUMDOCSAP")
            
            WS_STRING = WS_STRING & "|"
                
            If (Not IsNull(RS("BOT_IMPTOTBOL"))) Then WS_STRING = WS_STRING & RS("BOT_IMPTOTBOL")
                
            WS_STRING = WS_STRING & "|"
                
            If (Not IsNull(RS("BOT_DATSCAEFF"))) Then WS_STRING = WS_STRING & RS("BOT_DATSCAEFF")
                
            WS_STRING = WS_STRING & "|" & IIf((RS("NMR_SOLLECITI") = 0), "No", "Si")
                
            With WS_EST_DATA(WS_CNTR_ROWS)
                .DATAID = RS("ROW_DATA_KEY")
                .DATADESCRIPTION = "TRG_PARAM"
                .EXTRAPARAMS = Split(WS_STRING, "|")
            End With
            
            WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
    Set RS = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_EST_DATA_ANA/WS_EST_DATA): " & WS_CNTR_ROWS & " Items -> END"
        
    ' GET STAMP DATA
    '
    If ((DLLParams.BYPASS_STAMP = False) And (DLLParams.LAYOUT <> "L07")) Then
        If (FileLenInfo > 0) Then myAPB.APBItemsProgress = 0
        myAPB.APBItemsLabel = "Querying NET@H2O Stamp Data..."
        
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Querying NET@H2O Data (WS_EST_DATA_STAMP) -> START"
        
        WS_CNTR_ROWS = 0
        
        Set RS = DBConn.Execute("SELECT /*+ PARALLEL(4) */ SSQ_SOTTOLOTTO, SSQ_BOL_OPZ FROM MMS.VIEW_WABBNAH2ORMNDRS_STAMP WHERE (ID_WORKINGLOAD = '" & DLLParams.IDWORKINGLOAD & "') AND (SSQ_LOTTO = " & WS_LOTTO & ") ORDER BY SSQ_SOTTOLOTTO")
        
        If RS.RecordCount > 0 Then
            WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Getting NET@H2O Data -> Found " & Format$(RS.RecordCount, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " Elements - INFO"
        
            ReDim WS_EST_DATA_STAMP(RS.RecordCount - 1) As strct_DATA
            
            myAPB.APBMaxItems = RS.RecordCount
            
            Do Until RS.EOF
                myAPB.APBItemsLabel = "Parsing " & Format$(RS.AbsolutePosition, "##,##") & " of " & Format$(RS.RecordCount, "##,##") & " item"
                myAPB.APBItemsProgress = RS.AbsolutePosition
                
                With WS_EST_DATA_STAMP(WS_CNTR_ROWS)
                    .DATAID = Format$(WS_LOTTO, "000000") & Format$(RS("SSQ_SOTTOLOTTO"), "000000")
                    .DATADESCRIPTION = RS("SSQ_BOL_OPZ")
                End With
                
                WS_CNTR_ROWS = (WS_CNTR_ROWS + 1)
                
                RS.MoveNext
            Loop
        End If
        
        RS.Close
        Set RS = Nothing
        
        WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Got NET@H2O Data (WS_EST_DATA_STAMP): " & WS_CNTR_ROWS & " Items -> END"
    End If
        
    ' CLEAN
    '
    DBConn.Execute "DELETE FROM MMS.EST_WABBNARDECODER WHERE (ID_DECODER LIKE '" & DLLParams.IDWORKINGLOAD & "_%')"
    
CONTINUE:
    
    myAPB.APBClose
    Set myAPB = Nothing

    GET_GARC_INFO = True

    Exit Function
    
ErrHandler:
    If (WS_TRANSACTION) Then DBConn.RollbackTrans

    DBConn.RollbackTrans

    Close #1

    myAPB.APBClose
    Set myAPB = Nothing

    MsgBox Err.Description, vbExclamation, "Guru Meditation:"

End Function

Private Sub GET_QS_DATA(ByRef WS_ARRAY() As strct_DATA)

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

Public Sub MMS_LOG_INSERT()

    DBConn.Execute "INSERT INTO MMS.EST_WABBNAH2OLOG(ID_WORKINGLOAD,STR_FILENAME,STR_MODE) VALUES(" & DLLParams.IDWORKINGLOAD & ",'" & GET_BASENAME(DLLParams.INPUTFILENAME, False) & "','" & DLLParams.LAYOUT & "')"

End Sub

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
