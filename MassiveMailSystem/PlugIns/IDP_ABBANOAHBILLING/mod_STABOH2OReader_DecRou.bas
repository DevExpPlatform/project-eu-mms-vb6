Attribute VB_Name = "mod_STABOH2OReader_DecRou"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)

Private Sub ADD_ROW(ByRef WS_CNTR_ARRAY As Integer, ByRef WS_GDF As strct_GDF_RXXX, WS_DATA As String)
    
    WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
    ReDim Preserve WS_GDF.RXXX(WS_CNTR_ARRAY)
                                    
    WS_GDF.RXXX(WS_CNTR_ARRAY).ROW = WS_DATA

End Sub

Public Function NRM_ABBANOAHBILLING_DataProcessor() As Boolean

    On Error GoTo ErrHandler
    
    Dim WS_BOOLEAN          As Boolean
    Dim WS_CNTR_ARRAY       As Integer
    Dim WS_CNTR_ARRAY_NV018 As Integer
    Dim WS_CNTR_ARRAY_NV019 As Integer
    Dim WS_CNTR_ARRAY_NV020 As Integer
    Dim WS_CNTR_ARRAY_NV021 As Integer
    Dim WS_CNTR_ARRAY_NV022 As Integer
    Dim WS_CNTR_ARRAY_NV023 As Integer
    Dim WS_CNTR_ARRAY_NV024 As Integer
    Dim WS_CNTR_ARRAY_NV025 As Integer
    Dim WS_CNTR_ARRAY_QF    As Integer
    Dim WS_CNTR_DF          As Integer
    Dim WS_CNTR_ERROR       As Long
    Dim WS_CNTR_EMAIL       As Long
    Dim WS_CNTR_FEPA        As Long
    Dim WS_CNTR_G03R001     As Integer
    Dim WS_CNTR_ITEMS       As Long
    Dim WS_CNTR_NO_PACKAGE  As Long
    Dim WS_CNTR_OP          As Integer
    Dim WS_CNTR_WAIVER      As Long
    Dim WS_FLG_FEPA         As Boolean
    Dim WS_FLG_NOMERGE      As Boolean
    Dim WS_FLG_PROCESS      As Boolean
    Dim WS_FLG_RCA          As Boolean
    Dim WS_G16_DATE_TMP     As Long
    Dim WS_GPB_TMP          As strct_GPB_RXXX
    Dim WS_GPB_FLG_R002     As Boolean
    Dim WS_PERIOD_START     As String
    Dim WS_PERIOD_END       As String
    Dim WS_REC_TYPE         As String
    Dim WS_ROW_TMP_NVR      As String
    Dim WS_ROW_TMP_VAR      As String
    Dim WS_STRING           As String
    
    Dim ErrMsg              As String
    Dim FileLenInfo         As Long
    Dim FileLocInfo         As Long
    Dim myAPB               As cls_APB
    Dim ROWHEADER           As strct_STABODF_Header
    Dim StrIn               As String
    Dim strSuffix           As String
    
    ' INIT
    '
    If (MMS_Open = False) Then Exit Function
    
    Select Case DLLParams.DOCMODE
    Case "L01"
        strSuffix = "LP"
    
    Case "L02"
        strSuffix = "LM"
    
    Case "L03"
        strSuffix = "LPM"
    
    End Select
    
    WS_LOGFILEPATH = GET_PATHNAME(DLLParams.INPUTFILENAME) & GET_BASENAME(DLLParams.INPUTFILENAME, True) & "_" & DLLParams.PLUGMODE & strSuffix & ".LOG"
    If (FDEXIST(WS_LOGFILEPATH, False)) Then Kill WS_LOGFILEPATH
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - ABBANOA H2O Billing " & DLLParams.PLUGMODE & " Log System - START"
    
    ' GET INFO
    '
    DBConn.Open
    
    MMS_LOG_INSERT
    CACHE_INIT
    
    WS_STRING = String$(Len(WS_GPB_TMP), " ")
    CopyMemory ByVal VarPtr(WS_GPB_TMP), ByVal StrPtr(WS_STRING), Len(WS_GPB_TMP) * 2
    
    Select Case DLLParams.DOCMODE
    Case "BOL"
        If (GET_STABO_INFO_BOL = False) Then Exit Function
    
    Case "L01", "L02", "L03"
        If (GET_STABO_INFO_LXX = False) Then Exit Function
    
    End Select
        
    DBConn.Close
        
    H2O_DATA_INIT
    H2O_DATA_CLR
        
    ' START
    '
    WS_CNTR_ARRAY = -1
    WS_CNTR_ARRAY_NV018 = -1
    WS_CNTR_ARRAY_NV019 = -1
    WS_CNTR_ARRAY_NV020 = -1
    WS_CNTR_ARRAY_NV021 = -1
    WS_CNTR_ARRAY_NV022 = -1
    WS_CNTR_ARRAY_NV023 = -1
    WS_CNTR_ARRAY_NV024 = -1
    WS_CNTR_ARRAY_NV025 = -1
    WS_CNTR_ARRAY_QF = -1
    WS_CNTR_G03R001 = -1
    WS_CNTR_DF = -1
    WS_CNTR_OP = -1
    
    TEMPLATES_MANAGER_INIT GET_EXTERNALINFO(DLLParams.EXTRASPATH & Trim$(DLLParams.TEMPLATEORGANIZER))
    
    WS_PXX_FOOTER = GET_EXTERNALINFO(DLLParams.EXTRASPATH & "TXT_PXX_FOOTER.XML")
    WS_PXX_FOOTER_WSM = GET_EXTERNALINFO(DLLParams.EXTRASPATH & "TXT_PXX_FOOTER_MS335.XML")
    
    Set myAPB = New cls_APB
    
    With myAPB
        .APBMode = PBSingle
        .APBCaption = "Import Data Processor:"
        .APBMaxItems = 100
        .APBItemsProgress = 0
        .APBShow
    End With
    
    ' DATA READER
    '
    Open DLLParams.INPUTFILENAME For Input As #1
        FileLenInfo = (LOF(1) \ 1024)

        myAPB.APBMaxItems = FileLenInfo

        Do Until EOF(1)
            Line Input #1, StrIn

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
                    
                Case "002"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R002), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R002), ByVal StrPtr(StrIn), Len(WS_G01.R002) * 2
                    
                Case "004"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R004), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R004), ByVal StrPtr(StrIn), Len(WS_G01.R004) * 2
                    
                Case "005"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R005), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R005), ByVal StrPtr(StrIn), Len(WS_G01.R005) * 2
                    
                Case "006"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R006), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R006), ByVal StrPtr(StrIn), Len(WS_G01.R006) * 2
                    
                Case "007"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G01.R007), " ", False)
                    CopyMemory ByVal VarPtr(WS_G01.R007), ByVal StrPtr(StrIn), Len(WS_G01.R007) * 2
                
                End Select
                
            Case "02"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G02.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G02.R001), ByVal StrPtr(StrIn), Len(WS_G02.R001) * 2
                
                Case "002"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G02.R002), " ", False)
                    CopyMemory ByVal VarPtr(WS_G02.R002), ByVal StrPtr(StrIn), Len(WS_G02.R002) * 2
                
                Case "003"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G02.R003), " ", False)
                    CopyMemory ByVal VarPtr(WS_G02.R003), ByVal StrPtr(StrIn), Len(WS_G02.R003) * 2
                
                Case "004"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G02.R004), " ", False)
                    CopyMemory ByVal VarPtr(WS_G02.R004), ByVal StrPtr(StrIn), Len(WS_G02.R004) * 2
                
                Case "005"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G02.R005), " ", False)
                    CopyMemory ByVal VarPtr(WS_G02.R005), ByVal StrPtr(StrIn), Len(WS_G02.R005) * 2
                
                End Select
                
            Case "03"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    WS_CNTR_G03R001 = (WS_CNTR_G03R001 + 1)
                    ReDim Preserve WS_G03.R001(WS_CNTR_G03R001)

                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G03.R001(0)), " ", False)
                    CopyMemory ByVal VarPtr(WS_G03.R001(WS_CNTR_G03R001)), ByVal StrPtr(StrIn), Len(WS_G03.R001(WS_CNTR_G03R001)) * 2
                
                Case "002"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G03.R002), " ", False)
                    CopyMemory ByVal VarPtr(WS_G03.R002), ByVal StrPtr(StrIn), Len(WS_G03.R002) * 2
                
                Case "003"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G03.R003), " ", False)
                    CopyMemory ByVal VarPtr(WS_G03.R003), ByVal StrPtr(StrIn), Len(WS_G03.R003) * 2
                
                Case "004"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G03.R004), " ", False)
                    CopyMemory ByVal VarPtr(WS_G03.R004), ByVal StrPtr(StrIn), Len(WS_G03.R004) * 2
               
                End Select
            
            Case "05"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G05.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G05.R001), ByVal StrPtr(StrIn), Len(WS_G05.R001) * 2
                
                Case "002"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G05.R002), " ", False)
                    CopyMemory ByVal VarPtr(WS_G05.R002), ByVal StrPtr(StrIn), Len(WS_G05.R002) * 2
                
                End Select

            Case "06"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G06.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G06.R001), ByVal StrPtr(StrIn), Len(WS_G06.R001) * 2
                
                Case "002"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G06.R002), " ", False)
                    CopyMemory ByVal VarPtr(WS_G06.R002), ByVal StrPtr(StrIn), Len(WS_G06.R002) * 2
                
                Case "003"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G06.R003), " ", False)
                    CopyMemory ByVal VarPtr(WS_G06.R003), ByVal StrPtr(StrIn), Len(WS_G06.R003) * 2
              
                End Select
            
            Case "07"
                Select Case ROWHEADER.ROWNUMBER
                Case "003"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G07.R003), " ", False)
                    CopyMemory ByVal VarPtr(WS_G07.R003), ByVal StrPtr(StrIn), Len(WS_G07.R003) * 2
              
                    WS_CHK_G07 = True
                
                End Select
            
            Case "09"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G09.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G09.R001), ByVal StrPtr(StrIn), Len(WS_G09.R001) * 2
              
                Case "002"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G09.R002), " ", False)
                    CopyMemory ByVal VarPtr(WS_G09.R002), ByVal StrPtr(StrIn), Len(WS_G09.R002) * 2
              
                Case "003"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G09.R003), " ", False)
                    CopyMemory ByVal VarPtr(WS_G09.R003), ByVal StrPtr(StrIn), Len(WS_G09.R003) * 2
                
                    WS_CHK_GSE = (Trim$(WS_G09.R003.CODICECATEGORIAUTENZA) = "119")
                
                Case "005"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G09.R005), " ", False)
                    CopyMemory ByVal VarPtr(WS_G09.R005), ByVal StrPtr(StrIn), Len(WS_G09.R005) * 2
                
                Case "008"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G09.R008), " ", False)
                    CopyMemory ByVal VarPtr(WS_G09.R008), ByVal StrPtr(StrIn), Len(WS_G09.R008) * 2
              
                Case "012"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G09.R012), " ", False)
                    CopyMemory ByVal VarPtr(WS_G09.R012), ByVal StrPtr(StrIn), Len(WS_G09.R012) * 2
              
                End Select
            
            Case "11"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G11.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_G11.R001), ByVal StrPtr(StrIn), Len(WS_G11.R001) * 2
                
                Case "002"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G11.R002), " ", False)
                    CopyMemory ByVal VarPtr(WS_G11.R002), ByVal StrPtr(StrIn), Len(WS_G11.R002) * 2
                
                Case "003"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G11.R003), " ", False)
                    CopyMemory ByVal VarPtr(WS_G11.R003), ByVal StrPtr(StrIn), Len(WS_G11.R003) * 2
                
                End Select
            
            Case "12"
                If (ROWHEADER.ROWNUMBER = "001") Then WS_CNTR_ARRAY = -1
                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G12(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G12(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G12(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G12(WS_CNTR_ARRAY)) * 2
                
            Case "13"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    WS_CNTR_ARRAY = -1
                    WS_CHK_G13 = True
                End If
                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G13(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G13(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G13(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G13(WS_CNTR_ARRAY)) * 2
                
            Case "14"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    WS_CNTR_ARRAY = -1
                    WS_CHK_G14 = True
                End If
                                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G14(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G14(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G14(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G14(WS_CNTR_ARRAY)) * 2
                
            Case "16"
                If ((ROWHEADER.ROWNUMBER = "001") Or (WS_G16_DATE_TMP < CLng(Format$(Left$(StrIn, 10), "yyyyMMdd")))) Then
                    WS_CNTR_ARRAY = -1
                    WS_G16_DATE_TMP = CLng(Format$(Left$(StrIn, 10), "yyyyMMdd"))
                    WS_CHK_G16 = True
                End If
                                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G16(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G16(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G16(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G16(WS_CNTR_ARRAY)) * 2
                
            Case "17"
                If (ROWHEADER.ROWNUMBER = "001") Then WS_CNTR_ARRAY = -1
                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G17(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G17(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G17(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G17(WS_CNTR_ARRAY)) * 2
            
            Case "21"
                Select Case ROWHEADER.ROWNUMBER
                Case "999"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G21.R999), " ", False)
                    CopyMemory ByVal VarPtr(WS_G21.R999), ByVal StrPtr(StrIn), Len(WS_G21.R999) * 2
                
                Case Else
                    If (InStr(1, StrIn, "Altre bollette") > 0) Then
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G21.R005), " ", False)
                        CopyMemory ByVal VarPtr(WS_G21.R005), ByVal StrPtr(StrIn), Len(WS_G21.R005) * 2
                    Else
                        If (ROWHEADER.ROWNUMBER = "001") Then WS_CNTR_ARRAY = -1
                        
                        WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                        ReDim Preserve WS_G21.RXXX(WS_CNTR_ARRAY)
                    
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G21.RXXX(0)), " ", False)
                        CopyMemory ByVal VarPtr(WS_G21.RXXX(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G21.RXXX(WS_CNTR_ARRAY)) * 2
                    End If
                
                End Select
            
            Case "22"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    WS_CNTR_ARRAY = -1
                    WS_CHK_G22 = True
                End If
                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G22(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G22(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G22(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G22(WS_CNTR_ARRAY)) * 2
            
            Case "23"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    WS_CNTR_ARRAY = -1
                    WS_CHK_G23 = True
                End If
                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G23(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G23(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G23(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G23(WS_CNTR_ARRAY)) * 2
            
            Case "26"
                If (ROWHEADER.ROWNUMBER = "001") Then WS_CNTR_ARRAY = -1
                                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G26(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G26(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G26(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G26(WS_CNTR_ARRAY)) * 2
            
            Case "34"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    WS_CNTR_ARRAY = -1
                    WS_CHK_G34 = True
                End If
                
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_G34(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_G34(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_G34(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_G34(WS_CNTR_ARRAY)) * 2
                
            Case "BS"
                If (ROWHEADER.ROWNUMBER = "001") Then
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GBS), " ", False)
                    CopyMemory ByVal VarPtr(WS_GBS), ByVal StrPtr(StrIn), Len(WS_GBS) * 2
                    
                    WS_CNTR_ITEMS = (WS_CNTR_ITEMS + 1)
                    
                    Select Case DLLParams.DOCMODE
                    Case "BOL"
                        If (DLLParams.PLUGMODE = "SDI") Then
                            WS_FLG_PROCESS = ((WS_G01.R001.TIPONUMERAZIONE = "5") Or (Trim$(WS_G01.R007.MODALITÀ_INVIO) = "XSPDF"))
                        Else
                            WS_FLG_PROCESS = True
                        End If
                        
                        If (WS_FLG_PROCESS = False) Then Err.Raise vbObjectError + 512, "MMS_INSERT", "Tipo Fattura non gestita in modalità Filtro SDI"
                        
                        WS_BOOLEAN = MMS_Insert_BOL
                        
                    Case "L01"
                        WS_BOOLEAN = MMS_Insert_L01
                        
                    Case "L02"
                        WS_BOOLEAN = MMS_Insert_L02
                    
                    Case "L03"
                        WS_BOOLEAN = MMS_Insert_L03
                    
                    End Select

                    If (WS_BOOLEAN) Then
                        WS_FLG_FEPA = (WS_G01.R001.TIPONUMERAZIONE = "5")
                        WS_FLG_RCA = (WS_G01.R007.RINUNCIA_COPIA_ANALOGICA = "S")
                        
                        Select Case Trim$(WS_G01.R002.CANALEINOLTRO)
                        Case "02"   ' EMAIL
                            WS_FLG_NOMERGE = True
                            WS_CNTR_EMAIL = (WS_CNTR_EMAIL + 1)
                    
                        Case "03"  ' STAMPA + EMAIL
                            ' WS_FLG_NOMERGE = False
                    
                        Case Else
                            WS_FLG_NOMERGE = WS_FLG_RCA
                    
                        End Select
                    
                        WS_CNTR_NO_PACKAGE = (WS_CNTR_NO_PACKAGE + IIf((WS_FLG_FEPA Or WS_FLG_NOMERGE), 1, 0))
                        WS_CNTR_FEPA = (WS_CNTR_FEPA + IIf(WS_FLG_FEPA, 1, 0))
                        WS_CNTR_WAIVER = (WS_CNTR_WAIVER + IIf((WS_FLG_RCA And (WS_FLG_FEPA = False)), 1, 0))
                    Else
                        ErrMsg = MMS_GetErrMsg
                        
                        GoTo ErrHandler
                    End If

CONTINUE:
                    H2O_DATA_CLR
                    
                    WS_CNTR_ARRAY = -1
                    WS_CNTR_ARRAY_QF = -1
                    WS_CNTR_ARRAY_NV018 = -1
                    WS_CNTR_ARRAY_NV019 = -1
                    WS_CNTR_ARRAY_NV020 = -1
                    WS_CNTR_ARRAY_NV021 = -1
                    WS_CNTR_ARRAY_NV022 = -1
                    WS_CNTR_ARRAY_NV023 = -1
                    WS_CNTR_ARRAY_NV024 = -1
                    WS_CNTR_ARRAY_NV025 = -1
                    WS_CNTR_DF = -1
                    WS_CNTR_G03R001 = -1
                    WS_FLG_FEPA = False
                    WS_FLG_NOMERGE = False
                    WS_FLG_RCA = False
                    WS_G16_DATE_TMP = 0
                    WS_GPB_FLG_R002 = False
                    WS_PERIOD_END = ""
                    WS_PERIOD_START = ""
                    WS_ROW_TMP_NVR = ""
                    WS_ROW_TMP_VAR = ""

                    FileLocInfo = (Loc(1) \ 8)
                    myAPB.APBItemsLabel = "Itms: " & WS_CNTR_ITEMS & " - Errs: " & WS_CNTR_ERROR & " - Read " & Format$(FileLocInfo, "##,##") & " KB of " & Format$(FileLenInfo, "##,##") & " KB"
        
                    If ((FileLocInfo > 0) And (FileLenInfo > 0)) Then myAPB.APBItemsProgress = FileLocInfo
                End If
            
            Case "AA", "AR", "AZ", "BO", "DM", "DO", "IV", "NV", "RE", "RM", "SP", "TI", "TM", "TO", "TP", "TS", "VA"
                If (ROWHEADER.GROUP = "SP") Then
                    WS_CNTR_ARRAY = -1
                    WS_CNTR_ARRAY_QF = -1
                    
                    WS_CNTR_DF = (WS_CNTR_DF + 1)
                    
                    ReDim Preserve WS_GDF(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV018(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV019(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV020(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV021(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV022(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV023(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV024(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_NV025(WS_CNTR_DF)
                    ReDim Preserve WS_GDF_QF(WS_CNTR_DF)
                End If
            
                WS_REC_TYPE = ROWHEADER.GROUP & ROWHEADER.TIPOLOGIASEZIONE & ROWHEADER.IDENTIFICATIVO
                
                Select Case WS_REC_TYPE
                Case "NVR000", "NVR001", "NVR002"
                    ADD_ROW WS_CNTR_ARRAY_QF, WS_GDF_QF(WS_CNTR_DF), ROWHEADER.GROUP & StrIn
                    
                    WS_GDF_QF(WS_CNTR_DF).RXXX(WS_CNTR_ARRAY_QF).SORT_KEY = Format$(Mid$(StrIn, 18, 10), "yyyyMMdd") & "_" & Mid$(StrIn, 29, 45)
                
                Case "NVR018", "NVR019", "NVR020", "NVR021", "NVR022", "NVR023", "NVR024", "NVR025"
                    WS_STRING = StrIn
                    
                    If (WS_ROW_TMP_NVR <> "") Then
                        If (CDbl(Mid$(WS_ROW_TMP_NVR, 179, 18)) + CDbl(Mid$(WS_STRING, 177, 18)) = 0) Then
                            WS_ROW_TMP_NVR = ""
                        Else
                            Select Case WS_REC_TYPE
                            Case "NVR018"
                                ADD_ROW WS_CNTR_ARRAY_NV018, WS_GDF_NV018(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            Case "NVR019"
                                ADD_ROW WS_CNTR_ARRAY_NV019, WS_GDF_NV019(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            Case "NVR020"
                                ADD_ROW WS_CNTR_ARRAY_NV020, WS_GDF_NV020(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            Case "NVR021"
                                ADD_ROW WS_CNTR_ARRAY_NV021, WS_GDF_NV021(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            Case "NVR022"
                                ADD_ROW WS_CNTR_ARRAY_NV022, WS_GDF_NV022(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            Case "NVR023"
                                ADD_ROW WS_CNTR_ARRAY_NV023, WS_GDF_NV023(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            Case "NVR024"
                                ADD_ROW WS_CNTR_ARRAY_NV024, WS_GDF_NV024(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            Case "NVR025"
                                ADD_ROW WS_CNTR_ARRAY_NV025, WS_GDF_NV025(WS_CNTR_DF), WS_ROW_TMP_NVR
                            
                            End Select
                            
                            WS_ROW_TMP_NVR = ROWHEADER.GROUP & WS_STRING
                        End If
                    Else
                        WS_ROW_TMP_NVR = ROWHEADER.GROUP & WS_STRING
                    End If
                
                Case "VAR000", "VAR001", "VAR002"
                    WS_PERIOD_START = Mid$(StrIn, 7, 10)
                    WS_PERIOD_END = Mid$(StrIn, 18, 10)
                    
                    If ((Trim$(WS_PERIOD_START) <> "") And (Trim$(WS_PERIOD_END) <> "")) Then
                        If (CDate(WS_PERIOD_START) > CDate(WS_PERIOD_END)) Then
                            Mid$(StrIn, 7, 10) = WS_PERIOD_END
                            Mid$(StrIn, 18, 10) = WS_PERIOD_START
                        End If
                    End If
                    
                    WS_STRING = StrIn
                    
                    If (WS_ROW_TMP_VAR <> "") Then
                        If (CDbl(Mid$(WS_ROW_TMP_VAR, 179, 18)) + CDbl(Mid$(WS_STRING, 177, 18)) = 0) Then
                            WS_ROW_TMP_VAR = ""
                        Else
                            ADD_ROW WS_CNTR_ARRAY_QF, WS_GDF_QF(WS_CNTR_DF), WS_ROW_TMP_VAR
                            
                            WS_GDF_QF(WS_CNTR_DF).RXXX(WS_CNTR_ARRAY_QF).SORT_KEY = Format$(Mid$(StrIn, 18, 10), "yyyyMMdd") & "_" & Mid$(StrIn, 29, 45)
                            WS_ROW_TMP_VAR = ROWHEADER.GROUP & WS_STRING
                        End If
                    Else
                        WS_ROW_TMP_VAR = ROWHEADER.GROUP & WS_STRING
                    End If
                
                Case Else
                    If (WS_ROW_TMP_NVR <> "") Then
                        Select Case Left$(WS_ROW_TMP_NVR, 6)
                        Case "NVR018"
                            ADD_ROW WS_CNTR_ARRAY_NV018, WS_GDF_NV018(WS_CNTR_DF), WS_ROW_TMP_NVR
                        
                        Case "NVR019"
                            ADD_ROW WS_CNTR_ARRAY_NV019, WS_GDF_NV019(WS_CNTR_DF), WS_ROW_TMP_NVR
                        
                        Case "NVR020"
                            ADD_ROW WS_CNTR_ARRAY_NV020, WS_GDF_NV020(WS_CNTR_DF), WS_ROW_TMP_NVR
                        
                        Case "NVR021"
                            ADD_ROW WS_CNTR_ARRAY_NV021, WS_GDF_NV021(WS_CNTR_DF), WS_ROW_TMP_NVR
                        
                        Case "NVR022"
                            ADD_ROW WS_CNTR_ARRAY_NV022, WS_GDF_NV022(WS_CNTR_DF), WS_ROW_TMP_NVR
                        
                        Case "NVR023"
                            ADD_ROW WS_CNTR_ARRAY_NV023, WS_GDF_NV023(WS_CNTR_DF), WS_ROW_TMP_NVR
                        
                        Case "NVR024"
                            ADD_ROW WS_CNTR_ARRAY_NV024, WS_GDF_NV024(WS_CNTR_DF), WS_ROW_TMP_NVR
                        
                        Case "NVR025"
                            ADD_ROW WS_CNTR_ARRAY_NV025, WS_GDF_NV025(WS_CNTR_DF), WS_ROW_TMP_NVR
                    
                        End Select
                        
                        WS_ROW_TMP_NVR = ""
                    End If
                    
                    If (WS_ROW_TMP_VAR <> "") Then
                        ADD_ROW WS_CNTR_ARRAY_QF, WS_GDF_QF(WS_CNTR_DF), WS_ROW_TMP_VAR
                        
                        WS_GDF_QF(WS_CNTR_DF).RXXX(WS_CNTR_ARRAY_QF).SORT_KEY = Format$(Mid$(StrIn, 18, 10), "yyyyMMdd") & "_" & Mid$(StrIn, 29, 45)
                        WS_ROW_TMP_VAR = ""
                    End If
                    
                    ADD_ROW WS_CNTR_ARRAY, WS_GDF(WS_CNTR_DF), ROWHEADER.GROUP & StrIn
                
                End Select
            
            Case "OR"
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GOR), " ", False)
                CopyMemory ByVal VarPtr(WS_GOR), ByVal StrPtr(StrIn), Len(WS_GOR) * 2

                WS_CHK_GOR = True
            
            Case "PB"
                Select Case ROWHEADER.ROWNUMBER
                Case "001"
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GPB_TMP.R001), " ", False)
                    CopyMemory ByVal VarPtr(WS_GPB_TMP.R001), ByVal StrPtr(StrIn), Len(WS_GPB_TMP.R001) * 2
    
                    If (WS_CHK_GPB) Then
                        WS_GPB.R001.POTENZIALE_PRESCRIZIONE = (CDbl(WS_GPB.R001.POTENZIALE_PRESCRIZIONE) + Val(WS_GPB_TMP.R001.POTENZIALE_PRESCRIZIONE))
                    Else
                        WS_GPB = WS_GPB_TMP
                        WS_GPB.R001.POTENZIALE_PRESCRIZIONE = Val(WS_GPB.R001.POTENZIALE_PRESCRIZIONE)
                    End If
                    
                    WS_CHK_GPB = True
                Case "002"
                    If (Mid$(StrIn, 22, 1) = " ") Then
                        If (WS_GPB_FLG_R002 = False) Then
                            WS_STRING = String$(Len(WS_GPB.R002), " ")
                            CopyMemory ByVal VarPtr(WS_GPB.R002), ByVal StrPtr(WS_STRING), Len(WS_GPB.R002) * 2
                        End If
                    Else
                        StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GPB.R002), " ", False)
                        CopyMemory ByVal VarPtr(WS_GPB.R002), ByVal StrPtr(StrIn), Len(WS_GPB.R002) * 2
                    
                        WS_GPB_FLG_R002 = True
                    End If
                
                End Select

'            Case "PR"
'                WS_CHK_GPR = True
'
'                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GPR), " ", False)
'                CopyMemory ByVal VarPtr(WS_GPR), ByVal StrPtr(StrIn), Len(WS_GPR) * 2

            Case "SE"
                If (WS_CHK_GSE) Then
                    If (ROWHEADER.ROWNUMBER = "001") Then WS_CNTR_ARRAY = -1

                    WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                    ReDim Preserve WS_GSE(WS_CNTR_ARRAY)

                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GSE(0)), " ", False)
                    CopyMemory ByVal VarPtr(WS_GSE(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_GSE(WS_CNTR_ARRAY)) * 2
                End If
            
            Case "SF"
                If (WS_CHK_GSE) Then
                    StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GSF), " ", False)
                    CopyMemory ByVal VarPtr(WS_GSF), ByVal StrPtr(StrIn), Len(WS_GSF) * 2
                End If

            Case "SI"
                If (ROWHEADER.ROWNUMBER = "001") Then WS_CNTR_ARRAY = -1
                    
                WS_CNTR_ARRAY = (WS_CNTR_ARRAY + 1)
                ReDim Preserve WS_GSI(WS_CNTR_ARRAY)
                
                StrIn = GET_TEXTPAD(PADLEFT, StrIn, Len(WS_GSI(0)), " ", False)
                CopyMemory ByVal VarPtr(WS_GSI(WS_CNTR_ARRAY)), ByVal StrPtr(StrIn), Len(WS_GSI(WS_CNTR_ARRAY)) * 2
            
            End Select
        Loop
    Close #1

    MMS_Close True

    CACHE_CLEAR
    H2O_DATA_INIT
    H2O_DATA_CLR

    myAPB.APBClose
    Set myAPB = Nothing

    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - ---------------------------------"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " -        Processed: " & Format$(WS_CNTR_ITEMS, "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " -           Errors: " & Format$(WS_CNTR_ERROR, "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " -         Imported: " & Format$((WS_CNTR_ITEMS - WS_CNTR_ERROR), "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Non Postalizzati: " & Format$(WS_CNTR_NO_PACKAGE, "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " -     Postalizzati: " & Format$((WS_CNTR_ITEMS - WS_CNTR_ERROR - WS_CNTR_NO_PACKAGE), "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - ---------------------------------"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " -      Invio eMail: " & Format$(WS_CNTR_EMAIL, "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " -         F.E.P.A.: " & Format$(WS_CNTR_FEPA, "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - Rinuncia Cliente: " & Format$(WS_CNTR_WAIVER, "000000") & " Elements - INFO"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - ---------------------------------"
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - ABBANOA H2O Billing " & DLLParams.PLUGMODE & " Log System - END"

    NRM_ABBANOAHBILLING_DataProcessor = True

    If (WS_CNTR_ERROR > 0) Then MsgBox WS_CNTR_ERROR & " errori trovati durante l'importazione." & vbNewLine & vbNewLine & "Consultare il log su: " & vbNewLine & WS_LOGFILEPATH, vbExclamation, "Guru Meditation:"

    Exit Function

ErrHandler:
    WS_CNTR_ERROR = (WS_CNTR_ERROR + 1)
    WS_STRING = MMS_GetErrSctn
    
    If (Err.Description <> "") Then ErrMsg = Err.Description
    
    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
              "COD. ANAGR.: " & IIf((Trim$(WS_G02.R001.CODICEANAGRAFICO) = ""), "NO INFO", Trim$(WS_G02.R001.CODICEANAGRAFICO)) & " - " & _
              "NUM. DOC.: " & IIf((Trim$(WS_G01.R001.ANNOBOLLETTA & "/" & WS_G01.R001.NUMEROBOLLETTA) = "/"), "NO INFO", WS_G01.R001.ANNOBOLLETTA & "/" & Trim$(WS_G01.R001.NUMEROBOLLETTA)) & " - " & _
              "CODE SECTION: " & IIf(WS_STRING = "", "NO INFO", WS_STRING) & " - " & _
              "ERR. MESSAGE: " & UCase$(ErrMsg)

    Resume CONTINUE

End Function
