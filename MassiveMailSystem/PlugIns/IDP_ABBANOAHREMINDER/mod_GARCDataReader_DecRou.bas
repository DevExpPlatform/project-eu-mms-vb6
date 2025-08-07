Attribute VB_Name = "mod_GARCDataReader_DecRou"
Option Explicit

Public Function NRM_ABBANOAREMINDERS_DataProcessor() As Boolean

    On Error GoTo ErrHandler

    Dim CNTR_ED             As Integer
    Dim CNTR_ED_A_D         As Integer
    Dim CNTR_ED_S_D         As Integer
    Dim WS_ERRORCNTR        As Long
    Dim WS_ITEMS            As Long
    Dim WS_STRING           As String
    Dim WS_STRING_EMPTY     As String

    Dim ErrMsg              As String
    Dim FileLenInfo         As Long
    Dim FileLocInfo         As Long
    Dim myAPB               As cls_APB
    Dim ROWHEADER           As strct_GARC_HEADER
    Dim StrIn               As String

    ' INIT
    '
    If (MMS_Open = False) Then Exit Function
    
    WS_LOGFILEPATH = GET_PATHNAME(DLLParams.INPUTFILENAME) & GET_BASENAME(DLLParams.INPUTFILENAME, True) & ".LOG"
    If (FDEXIST(WS_LOGFILEPATH, False)) Then Kill WS_LOGFILEPATH
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - ABBANOA Reminders " & DLLParams.LAYOUT & " Log System - START"
    
    ' GET INFO
    '
    DBConn.Open
    
    MMS_LOG_INSERT
    CACHE_INIT
    
    GARC_DATA_INIT
    GARC_DATA_CLEAR
    
    If (GET_GARC_INFO = False) Then Exit Function

    GARC_DATA_INIT
    GARC_DATA_CLEAR

    DBConn.Close
        
    ' START
    '
    CNTR_ED = -1
    CNTR_ED_A_D = -1
    CNTR_ED_S_D = -1
    
    TEMPLATES_MANAGER_INIT GET_EXTERNALINFO(DLLParams.EXTRASPATH & "TemplateOrganizer_" & DLLParams.LAYOUT & ".XML")
    
    WS_GG_DESCR = GET_NUM2STRING(DLLParams.PRM_GG)
    WS_PAGE_FOOTER = GET_EXTERNALINFO(DLLParams.EXTRASPATH & "TXT_PXX_FOOTER.XML")

    Set myAPB = New cls_APB

    With myAPB
        .APBMode = PBSingle
        .APBCaption = "Import Data Processor:"
        .APBMaxItems = 100
        .APBItemsProgress = 0
        .APBShow
    End With

    Open DLLParams.INPUTFILENAME For Input As #1
        FileLenInfo = (LOF(1) \ 1024)

        myAPB.APBMaxItems = FileLenInfo

        Do Until EOF(1)
            Line Input #1, StrIn

            CopyMemory ByVal VarPtr(ROWHEADER), ByVal StrPtr(StrIn), Len(ROWHEADER) * 2
            StrIn = Mid$(StrIn, 9)

            With WS_01S
                Select Case ROWHEADER.GROUP
                Case "01S"
                    Select Case ROWHEADER.SUBGROUP
                    Case "AF"
                        If (ROWHEADER.ROWNUMBER = "001") Then
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.AF_R001), " ", False)
                            CopyMemory ByVal VarPtr(.AF_R001), ByVal StrPtr(StrIn), Len(.AF_R001) * 2
                        End If

                    Case "AN"
                        If (ROWHEADER.ROWNUMBER = "001") Then
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.AN_R001), " ", False)
                            CopyMemory ByVal VarPtr(.AN_R001), ByVal StrPtr(StrIn), Len(.AN_R001) * 2
                        End If

                    Case "BO"
                        If (ROWHEADER.ROWNUMBER = "001") Then
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.BO_R001), " ", False)
                            CopyMemory ByVal VarPtr(.BO_R001), ByVal StrPtr(StrIn), Len(.BO_R001) * 2
                        
                            WS_FLG_BO = True
                        End If
                    
                    Case "ED"
                        WS_STRING = Left$(StrIn, 2)
                        StrIn = Mid$(StrIn, 3)

                        If ((WS_STRING = "AC") Or (WS_STRING = "SE")) Then
                            CNTR_ED_A_D = -1
                            CNTR_ED_S_D = -1
                            CNTR_ED = (CNTR_ED + 1)

                            ReDim Preserve .ED_RXXX(CNTR_ED)
                            
                            .ED_RXXX(CNTR_ED).TIPORECORD = WS_STRING
                            
                            If (WS_STRING = "SE") Then
                                WS_STRING_EMPTY = String$(Len(.ED_RXXX(CNTR_ED).ED_R001_S_I), " ")
                                CopyMemory ByVal VarPtr(.ED_RXXX(CNTR_ED).ED_R001_S_I), ByVal StrPtr(WS_STRING_EMPTY), Len(.ED_RXXX(0).ED_R001_S_I) * 2
    
                                WS_STRING_EMPTY = String$(Len(.ED_RXXX(CNTR_ED).ED_R001_S_L), " ")
                                CopyMemory ByVal VarPtr(.ED_RXXX(CNTR_ED).ED_R001_S_L), ByVal StrPtr(WS_STRING_EMPTY), Len(.ED_RXXX(CNTR_ED).ED_R001_S_L) * 2
                            End If
                        End If
                        
                        With .ED_RXXX(CNTR_ED)
                            Select Case WS_STRING
                            Case "AC"
                                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_R001_A_C), " ", False)
                                CopyMemory ByVal VarPtr(.ED_R001_A_C), ByVal StrPtr(StrIn), Len(.ED_R001_A_C) * 2
                            
                            Case "AD"
                                CNTR_ED_A_D = (CNTR_ED_A_D + 1)

                                ReDim Preserve .ED_RXXX_A_D(CNTR_ED_A_D)

                                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_RXXX_A_D(0)), " ", False)
                                CopyMemory ByVal VarPtr(.ED_RXXX_A_D(CNTR_ED_A_D)), ByVal StrPtr(StrIn), Len(.ED_RXXX_A_D(CNTR_ED_A_D)) * 2
                            
                            Case "AT", "ST"
                                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_R001_S_T), " ", False)
                                CopyMemory ByVal VarPtr(.ED_R001_S_T), ByVal StrPtr(StrIn), Len(.ED_R001_S_T) * 2

                            Case "SD"
                                CNTR_ED_S_D = (CNTR_ED_S_D + 1)

                                ReDim Preserve .ED_RXXX_S_D(CNTR_ED_S_D)

                                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_RXXX_S_D(0)), " ", False)
                                CopyMemory ByVal VarPtr(.ED_RXXX_S_D(CNTR_ED_S_D)), ByVal StrPtr(StrIn), Len(.ED_RXXX_S_D(CNTR_ED_S_D)) * 2
                                                                
                            Case "SE"
                                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_R001_S_E), " ", False)
                                CopyMemory ByVal VarPtr(.ED_R001_S_E), ByVal StrPtr(StrIn), Len(.ED_R001_S_E) * 2

                            Case "SI"
                                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_R001_S_I), " ", False)
                                CopyMemory ByVal VarPtr(.ED_R001_S_I), ByVal StrPtr(StrIn), Len(.ED_R001_S_I) * 2
                            
                            Case "SL"
                                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ED_R001_S_L), " ", False)
                                CopyMemory ByVal VarPtr(.ED_R001_S_L), ByVal StrPtr(StrIn), Len(.ED_R001_S_L) * 2

                            End Select
                        End With

                    Case "ES"
                        If (ROWHEADER.ROWNUMBER = "001") Then
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.ES_R001), " ", False)
                            CopyMemory ByVal VarPtr(.ES_R001), ByVal StrPtr(StrIn), Len(.ES_R001) * 2
                        End If

                    Case "IL"
                        If (ROWHEADER.ROWNUMBER = "001") Then
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.IL_R001), " ", False)
                            CopyMemory ByVal VarPtr(.IL_R001), ByVal StrPtr(StrIn), Len(.IL_R001) * 2
                        End If

                    Case "IR"
                        Select Case ROWHEADER.ROWNUMBER
                        Case "001"
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.IR_R001), " ", False)
                            CopyMemory ByVal VarPtr(.IR_R001), ByVal StrPtr(StrIn), Len(.IR_R001) * 2

                        Case "002"
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.IR_R002), " ", False)
                            CopyMemory ByVal VarPtr(.IR_R002), ByVal StrPtr(StrIn), Len(.IR_R002) * 2

                        End Select

                    Case "IS"
                        Select Case ROWHEADER.ROWNUMBER
                        Case "001"
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.IS_R001), " ", False)
                            CopyMemory ByVal VarPtr(.IS_R001), ByVal StrPtr(StrIn), Len(.IS_R001) * 2

                            If (MMS_Insert = False) Then
                                ErrMsg = MMS_GetErrMsg
                                
                                GoTo ErrHandler
                            Else
                                WS_ITEMS = (WS_ITEMS + 1)
                            End If

CONTINUE:
                            CNTR_ED = -1
                            CNTR_ED_A_D = -1
                            CNTR_ED_S_D = -1
        
                            GARC_DATA_CLEAR
        
                            FileLocInfo = (Loc(1) \ 8)
                            myAPB.APBItemsLabel = "Itms: " & WS_ITEMS & " - Errs: " & WS_ERRORCNTR & " - Read " & Format$(FileLocInfo, "##,##") & " KB of " & Format$(FileLenInfo, "##,##") & " KB"
                
                            If ((FileLocInfo > 0) And (FileLenInfo > 0)) Then myAPB.APBItemsProgress = FileLocInfo

                        End Select

                    Case "NF"
                        If (ROWHEADER.ROWNUMBER = "001") Then
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.NF_R001), " ", False)
                            CopyMemory ByVal VarPtr(.NF_R001), ByVal StrPtr(StrIn), Len(.NF_R001) * 2
                        End If

                    Case "PA"
                        If (ROWHEADER.ROWNUMBER = "001") Then
                            WS_FLG_SPA = True
                        
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.PA_R001), " ", False)
                            CopyMemory ByVal VarPtr(.PA_R001), ByVal StrPtr(StrIn), Len(.PA_R001) * 2
                        End If
                    
                    Case "SL"
                        Select Case ROWHEADER.ROWNUMBER
                        Case "001"
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.SL_R001), " ", False)
                            CopyMemory ByVal VarPtr(.SL_R001), ByVal StrPtr(StrIn), Len(.SL_R001) * 2
                        
                        Case "002"
                            StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(.SL_R002), " ", False)
                            CopyMemory ByVal VarPtr(.SL_R002), ByVal StrPtr(StrIn), Len(.SL_R002) * 2
                        
                        End Select

                    End Select

                End Select
            End With
        Loop
    Close #1

    MMS_Close True

    CACHE_CLEAR
    GARC_DATA_INIT
    GARC_DATA_CLEAR

    myAPB.APBClose
    Set myAPB = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - ABBANOA Reminders " & DLLParams.LAYOUT & " Log System - END"
    
    NRM_ABBANOAREMINDERS_DataProcessor = True

    If (WS_ERRORCNTR > 0) Then MsgBox WS_ERRORCNTR & " errori trovati durante l'importazione." & vbNewLine & vbNewLine & "Consultare il log su: " & vbNewLine & WS_LOGFILEPATH, vbExclamation, "Guru Meditation:"

    Exit Function

ErrHandler:
    WS_ERRORCNTR = (WS_ERRORCNTR + 1)
    WS_STRING = MMS_GetErrSctn
    
    If (Err.Description <> "") Then ErrMsg = Err.Description
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
              "COD. ANAGR.: " & IIf((Trim$(WS_01S.AN_R001.CODICEANAGRAFICO) = ""), "NO INFO", Trim$(WS_01S.AN_R001.CODICEANAGRAFICO)) & " - " & _
              "NUM. DOC.: " & IIf((WS_01S.AF_R001.CODICELOTTO & "/" & WS_01S.IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(WS_01S.AF_R001.CODICELOTTO, "000000") & "/" & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
              "CODE SECTION: " & IIf(WS_STRING = "", "NO INFO", WS_STRING) & " - " & _
              "ERR. MESSAGE: " & UCase$(ErrMsg)

    DoEvents

    Resume CONTINUE

End Function
