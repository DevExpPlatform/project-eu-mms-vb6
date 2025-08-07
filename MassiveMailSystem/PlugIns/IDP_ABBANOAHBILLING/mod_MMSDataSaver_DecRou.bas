Attribute VB_Name = "mod_MMSDataSaver_DecRou"
Option Explicit

Private Const WS_PAGEHEIGHTMAX = 270

Private mySQLImporter           As SQL_Import.PlugIn
Private myXFDFMLTable           As cls_XFDFMLTable
Private myXFDFMLText            As cls_XFDFMLText

Private WS_ANNEXED_DATA         As String
Private WS_BDS                  As Boolean
Private WS_CALC_IV_TOT          As Single
Private WS_CCP                  As Boolean
Private WS_CNTTR_MASTER_MTRCL   As String
Private WS_CNTTR_MASTER_SRVZ    As String
Private WS_CODICE_SERVIZIO_KEY  As String
Private WS_DEL_547_19_B_CAUSALE As String
Private WS_DF_GBO_FLG           As Boolean
Private WS_DF_GDM_FLG           As String
Private WS_DF_GNV_PERIODO       As String
Private WS_DF_GRM_RIFDOC        As String
Private WS_DOCPAGES_DA()        As String
Private WS_DOCPAGES_DA_CNTR     As Integer
Private WS_DOCPAGES_DF()        As String
Private WS_DOCPAGES_DF_CNTR     As Integer
Private WS_DOCPAGES_UI_CNTR     As Integer
Private WS_ERRMSG               As String
Private WS_ERRSCT               As String
Private WS_FATTURANUMERO        As String
Private WS_FLG_ANNXD_547_19     As Boolean
Private WS_FLG_CATEGORIA        As Boolean
Private WS_FLG_CONSUMI          As Boolean
Private WS_FLG_DF_IV            As Boolean
Private WS_FLG_DF_TS            As Boolean
Private WS_FLG_DIV              As Boolean
Private WS_FLG_DOM              As Boolean
Private WS_FLG_FATTELE_PA       As Boolean
Private WS_FLG_INDE_LBL         As Boolean
Private WS_FLG_MASTER           As Boolean
Private WS_FLG_MASTER_ECC_POS   As Boolean
'Private WS_FLG_MSG_BS           As Boolean
Private WS_FLG_NEG_NODOM        As Boolean
Private WS_FLG_NOTACREDITO      As Boolean
Private WS_FLG_PARTITE          As Boolean
Private WS_FORNITURA_MASTER_P02 As strct_DATA
Private WS_IMPORTO_TC_TICSI     As String
Private WS_IMPORTOTOTALE        As String
Private WS_LOCALITY             As String
Private WS_MSG_CA               As Boolean
Private WS_MSG_CI               As Boolean
Private WS_NATIONALITY          As String
Private WS_PAGEHEIGHT           As Single
Private WS_PAGENUM              As String
Private WS_RECIPIENT            As String

Private Sub ADD_PXX_ATTACH_TABLE_CLOSE()

    GET_CHK_DA_ITEMSHEIGHT 11.3

    With myXFDFMLTable
        .setCellColSpan = "4"
        .setCellHeight = "10"
        .addCell = ""

        .setCellColSpan = "4"
        .setCellHeight = "10"
        .addCell = ""

        .setCellAlignH = "right"
        .setCellColSpan = "3"
        .addCell = "Servizio Clienti Abbanoa"
        
        .addCell = " "
    End With

End Sub

Private Sub ADD_PXX_ATTACH_TABLE_HDR()

    With myXFDFMLTable
        .setTableAlignH = "center"
        '.setTableBorders = "0.3"
        .setTableColumns = "4"
        .setTableFontName = "helr45w.ttf"
        .setTableFontSize = "10"
        .setTablePaddingTop = "0"
        .setTableWidths = "4.5,4.5,9,1"
         
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .setCellHeight = "12"
        .addCell = "Numero Fattura"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .addCell = "Data Fattura"
        
        .setCellColSpan = "2"
        .addCell = " "
    End With

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.23)

End Sub

Private Sub ADD_PXX_ATTACH_TABLE_ROW(rowData As String)
    
    GET_CHK_DA_ITEMSHEIGHT 4.23
    
    Dim WS_DATA() As String
    
    WS_DATA = Split(rowData, "|")
    
    With myXFDFMLTable
        .setCellAlignH = "right"
        .addCell = WS_DATA(0)
        
        .addCell = WS_DATA(1)
        
        .setCellColSpan = "2"
        .addCell = " "
    End With

End Sub

Private Sub ADD_PXX_DETTAGLIOBOLLETTA_ROW(I As Integer, RowType As String, rowData As String)
    
    Dim J                  As Integer
    Dim WS_DATA()          As String
    Dim WS_DF_PXX_GBO_FLG  As Boolean
    Dim WS_FLG_INDE        As Boolean
    Dim WS_FLG_INDE_DESCR  As String
    Dim WS_STRCT           As strct_DATA
    
    With myXFDFMLTable
        Select Case RowType
        Case "AA"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GAA), " ", False)
            CopyMemory ByVal VarPtr(WS_GAA), ByVal StrPtr(rowData), Len(WS_GAA) * 2

            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT

            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = "  " & Trim$(WS_GAA.DESCRIZIONE)
            
            .addCell = Trim$(WS_GAA.IMPORTO) & "|" & Trim$(WS_GAA.ALIQUOTA)
        
        Case "AR"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GAR), " ", False)
            CopyMemory ByVal VarPtr(WS_GAR), ByVal StrPtr(rowData), Len(WS_GAR) * 2

            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT

            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = "  " & Trim$(WS_GAR.DESCRIZIONE)
            
            .addCell = Trim$(WS_GAR.IMPORTO) & "|" & Trim$(WS_GAR.ALIQUOTA)
        
        Case "AZ"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GAZ), " ", False)
            CopyMemory ByVal VarPtr(WS_GAZ), ByVal StrPtr(rowData), Len(WS_GAZ) * 2

            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT

            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = "  " & Trim$(WS_GAZ.DESCRIZIONE)
            
            .addCell = Trim$(WS_GAZ.IMPORTO) & "|" & Trim$(WS_GAZ.ALIQUOTA)
        
        Case "BO"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GBO), " ", False)
            CopyMemory ByVal VarPtr(WS_GBO), ByVal StrPtr(rowData), Len(WS_GBO) * 2
            
            If (WS_DF_PXX_GBO_FLG = False) Then
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.88)
                GoSub CHECK_ROWHEIGHT
    
                .setCellColSpan = "9"
                .setCellHeight = "2"
                .addCell = ""
    
                .setCellAlignH = "left"
                .setCellColSpan = "9"
                .setCellFontName = "helr65w.ttf"
                .addCell = "  BOLLO DI QUIETANZA"
        
                WS_DF_GBO_FLG = True
                WS_DF_PXX_GBO_FLG = True
            End If
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT
            
            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = "  " & Replace$(Trim$(WS_GBO.DESCRIZIONE), "  ", " ")
            
            .addCell = Trim$(WS_GBO.IMPORTO) & "|" & IIf(Trim$(WS_GBO.ALIQUOTA) = "", " ", Trim$(WS_GBO.ALIQUOTA))
            
        Case "DM"
            WS_DF_GRM_RIFDOC = ""
        
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GDM), " ", False)
            CopyMemory ByVal VarPtr(WS_GDM), ByVal StrPtr(rowData), Len(WS_GDM) * 2
            
            If (WS_DF_GDM_FLG <> WS_GDM.TIPOLOGIASOTTOTIPO) Then
                WS_DF_GDM_FLG = WS_GDM.TIPOLOGIASOTTOTIPO
                
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
                GoSub CHECK_ROWHEIGHT
    
                .setCellAlignH = "left"
                .setCellColSpan = "9"
                .setCellFontName = "helr65w.ttf"
                .addCell = "  " & IIf(WS_DF_GDM_FLG = "D", "INTERESSI DILATORI", "INTERESSI DI MORA")
            End If
        
        Case "DO"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GDO), " ", False)
            CopyMemory ByVal VarPtr(WS_GDO), ByVal StrPtr(rowData), Len(WS_GDO) * 2

            WS_STRCT = GET_DATA_CACHE(DEL_547_19_B_CAUSALI, WS_GDO.PARCAU)
            If (WS_STRCT.dataDescription = "TRG_PARAM") Then WS_DEL_547_19_B_CAUSALE = WS_STRCT.EXTRAPARAMS(0)
            
            WS_FLG_INDE = ((WS_G09.R005.FLAGINDENNIZZI = "S") And (GET_DATA_CACHE(INDE, "INDE_" & WS_GDO.PARCAU).dataDescription = "TRG_PARAM"))
            WS_FLG_INDE_DESCR = IIf(WS_FLG_INDE, " (*)", "")
            
            If (WS_FLG_INDE_LBL = False) Then WS_FLG_INDE_LBL = WS_FLG_INDE
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT

            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = "  " & Trim$(WS_GDO.DESCRIZIONE) & WS_FLG_INDE_DESCR
            
            .addCell = Trim$(WS_GDO.IMPORTO) & "|" & Trim$(WS_GDO.ALIQUOTA)
        
        Case "IV"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GIV), " ", False)
            CopyMemory ByVal VarPtr(WS_GIV), ByVal StrPtr(rowData), Len(WS_GIV) * 2
            
            If (WS_FLG_DF_IV = False) Then
                WS_CALC_IV_TOT = 0
                WS_FLG_DF_IV = True
                
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
                GoSub CHECK_ROWHEIGHT
                
                .setCellBackColor = "230,230,230"
                .setCellAlignH = "left"
                .setCellColSpan = "9"
                .setCellFontName = "helr66w.ttf"
                .addCell = "Dettaglio IVA"
                
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.88)
                GoSub CHECK_ROWHEIGHT
                
                .setCellColSpan = "9"
                .setCellHeight = "2"
                .addCell = ""
                    
                .setCellAlignH = "left"
                .setCellColSpan = "9"
                .setCellFontName = "helr46w.ttf"
                .addCell = " Dettaglio IVA"
            End If
            
            WS_CALC_IV_TOT = (WS_CALC_IV_TOT + CSng(WS_GIV.IMPORTO))
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT

            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = "  " & GET_DF_IV(Trim$(WS_GIV.DESCRIZIONE))
            
            .addCell = Trim$(WS_GIV.IMPORTO) & "| "
        
        Case "NV"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GNV), " ", False)
            CopyMemory ByVal VarPtr(WS_GNV), ByVal StrPtr(rowData), Len(WS_GNV) * 2
            
            If ((WS_GNV.IDENTIFICATIVO = "004") And (WS_DF_GNV_PERIODO <> WS_GNV.PERIODO)) Then
                If (Trim$(WS_GNV.PERIODO) <> "") Then
                    WS_DF_GNV_PERIODO = WS_GNV.PERIODO
                Else
                    WS_GNV.PERIODO = WS_DF_GNV_PERIODO
                End If
            End If
        
            'If ((Trim$(WS_GNV.QUANTITÀ) <> "") And (Trim$(WS_GNV.QUANTITÀ) <> "0,000000")) Then
            'If ((Trim$(WS_GNV.PREZZO) <> "0,00000000")) Then
            If (Trim$(WS_GNV.IMPORTO) <> "0,00") Then
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
                GoSub CHECK_ROWHEIGHT
                
                .setCellAlignH = "left"
                .addCell = "  " & Trim$(WS_GNV.DESCRIZIONE)
                
                If (Trim$(WS_GNV.PERIODO) = "") Then
                    .setCellColSpan = "2"
                    .addCell = ""
                Else
                     WS_DATA = Split(WS_GNV.PERIODO, "-")
                    
                    .setCellAlignH = "center"
                    .addCell = WS_DATA(0) & "|" & WS_DATA(1)
                End If
                
                .addCell = Trim$(WS_GNV.TEMPO) & " " & Trim$(WS_GNV.UNITÀMISURATEMPO)
                
                If (Trim$(WS_GNV.CONCESSIONI) = "") Then
                    If (Trim$(WS_GNV.QUANTITÀ) = "") Then
                        .setCellColSpan = "2"
                        .addCell = ""
                    Else
                        .addCell = NRM_REMOVEZEROES(WS_GNV.QUANTITÀ, True)
                        
                        .setCellAlignH = "left"
                        .addCell = Trim$(WS_GNV.UNITÀMISURAQUANTITÀ)
                    End If
                Else
                    .addCell = Trim$(WS_GNV.CONCESSIONI)
                        
                    .setCellAlignH = "left"
                    .addCell = "un. imm."
                End If
                
                .addCell = Format$(Trim$(WS_GNV.PREZZO), "##,######0.000000") & "|" & Trim$(WS_GNV.IMPORTO) & "|" & Trim$(WS_GNV.ALIQUOTA)
            End If
        
        Case "RE"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GRE), " ", False)
            CopyMemory ByVal VarPtr(WS_GRE), ByVal StrPtr(rowData), Len(WS_GRE) * 2

            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.94)
            GoSub CHECK_ROWHEIGHT
            
            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = Trim$(WS_GRE.DESCRIZIONE)
            
            .addCell = Trim$(WS_GRE.IMPORTO) & "|" & Trim$(WS_GRE.ALIQUOTA)
        
            .setCellColSpan = "9"
            .setCellHeight = "4"
            .addCell = ""

        Case "RM"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GRM), " ", False)
            CopyMemory ByVal VarPtr(WS_GRM), ByVal StrPtr(rowData), Len(WS_GRM) * 2
        
            WS_DF_GRM_RIFDOC = WS_GRM.ANNODOCUMENTOMOROSO & "-" & Trim$(WS_GRM.NUMERODOCUMENTOMOROSO)
        
        Case "SP"
            WS_DF_PXX_GBO_FLG = False

            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GSP), " ", False)
            CopyMemory ByVal VarPtr(WS_GSP), ByVal StrPtr(rowData), Len(WS_GSP) * 2
        
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.94)
            GoSub CHECK_ROWHEIGHT
            
            .setCellColSpan = "9"
            .setCellHeight = "4"
            .addCell = ""
        
            .setCellAlignH = "left"
            .setCellColSpan = "9"
            .setCellFontName = "helr65w.ttf"
            .setCellFontSize = "8"
            .addCell = Trim$(WS_GSP.INDENTIFICATIVOFATTURAZIONE)
        
        Case "TI"
            WS_DF_GNV_PERIODO = ""

            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GTI), " ", False)
            CopyMemory ByVal VarPtr(WS_GTI), ByVal StrPtr(rowData), Len(WS_GTI) * 2
            
            Select Case WS_GTI.TIPOLOGIASEZIONE
            Case "A"
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
                GoSub CHECK_ROWHEIGHT
                
                .setCellBackColor = "230,230,230"
                .setCellAlignH = "left"
                .setCellColSpan = "9"
                .setCellFontName = "helr66w.ttf"
                .addCell = Trim$(WS_GTI.DESCRIZIONE)
                
                If (WS_GTI.IDENTIFICATIVO = "000") Then
                    If (CHK_GDF_NVXXX(WS_GDF_QF(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_QF(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_QF(I).RXXX(J).ROW, 2), Mid$(WS_GDF_QF(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                End If
            
            Case "R"
                If ((WS_GTI.IDENTIFICATIVO = "000") Or (WS_GTI.IDENTIFICATIVO = "001") Or (WS_GTI.IDENTIFICATIVO = "002")) Then Exit Sub
                
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.88)
                GoSub CHECK_ROWHEIGHT
                
                .setCellColSpan = "9"
                .setCellHeight = "2"
                .addCell = ""
                    
                .setCellAlignH = "left"
                .setCellColSpan = "9"
                .setCellFontName = "helr46w.ttf"
                .addCell = " " & Trim$(WS_GTI.DESCRIZIONE)
            
                Select Case WS_GTI.IDENTIFICATIVO
                Case "018"
                    If (CHK_GDF_NVXXX(WS_GDF_NV018(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV018(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV018(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV018(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                
                Case "019"
                    If (CHK_GDF_NVXXX(WS_GDF_NV019(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV019(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV019(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV019(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                    
                Case "020"
                    If (CHK_GDF_NVXXX(WS_GDF_NV020(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV020(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV020(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV020(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                
                Case "021"
                    If (CHK_GDF_NVXXX(WS_GDF_NV021(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV021(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV021(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV021(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                
                Case "022"
                    If (CHK_GDF_NVXXX(WS_GDF_NV022(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV022(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV022(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV022(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                
                Case "023"
                    If (CHK_GDF_NVXXX(WS_GDF_NV023(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV023(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV023(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV023(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                
                Case "024"
                    If (CHK_GDF_NVXXX(WS_GDF_NV024(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV024(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV024(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV024(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                
                Case "025"
                    If (CHK_GDF_NVXXX(WS_GDF_NV025(I).RXXX)) Then
                        For J = 0 To UBound(WS_GDF_NV025(I).RXXX)
                            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF_NV025(I).RXXX(J).ROW, 2), Mid$(WS_GDF_NV025(I).RXXX(J).ROW, 3)
                        Next J
                    End If
                
                End Select
            
            End Select
            
        Case "TM"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GTM), " ", False)
            CopyMemory ByVal VarPtr(WS_GTM), ByVal StrPtr(rowData), Len(WS_GTM) * 2
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT
        
            .setCellAlignH = "left"
            .setCellColSpan = "9"
            .setCellFontName = "helr45w.ttf"
            
            .addCell = "  " & Replace$(Trim$(WS_GTM.PERIODO), "-", " - ") & " • " & IIf(WS_DF_GDM_FLG = "D", "Interessi Dilatori", "Mora") & " su fattura " & WS_DF_GRM_RIFDOC & " del " & Replace$(Trim$(WS_GTM.TASSOINTERESSEAPPLICATO), ".", ",") & "% su imponibile di € " & NRM_IMPORT(WS_GTM.IMPONIBILE, "##,##0.00", False) & " per un importo di € " & NRM_IMPORT(WS_GTM.IMPORTO, "##,##0.00", False)
            
        Case "TO"
            WS_DF_GDM_FLG = ""
            
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GTO), " ", False)
            CopyMemory ByVal VarPtr(WS_GTO), ByVal StrPtr(rowData), Len(WS_GTO) * 2
        
            If (WS_GTO.TIPOLOGIASEZIONE = "A") Then
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 5.29)
                GoSub CHECK_ROWHEIGHT
                
                .setCellColSpan = "9"
                .setCellHeight = "2"
                .addCell = ""
                
                .setCellAlignH = "left"
                .setCellColSpan = "7"
                .setCellFontName = "helr65w.ttf"
                .addCell = " " & Trim$(WS_GTO.DESCRIZIONE)
                
                .setCellFontName = "helr65w.ttf"
                .addCell = Trim$(WS_GTO.IMPORTO) & "| "
                                
                .setCellColSpan = "9"
                .setCellHeight = "4"
                .addCell = ""
            End If
        
        Case "TP"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GTP), " ", False)
            CopyMemory ByVal VarPtr(WS_GTP), ByVal StrPtr(rowData), Len(WS_GTP) * 2
        
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 5.64)
            GoSub CHECK_ROWHEIGHT
            
            .setCellColSpan = "9"
            .setCellHeight = "2"
            .addCell = ""
            
            .setCellAlignH = "left"
            .setCellBackColor = "230,230,230"
            .setCellBorderTop = "0.55"
            .setCellBorderTopColor = "0,55,110"
            .setCellColSpan = "7"
            .setCellFontName = "helr65w.ttf"
            .setCellFontSize = "8"
            .setCellHeight = "14"
            .addCell = "  " & Trim$(WS_GTP.DESCRIZIONE)
            
            .setCellBackColor = "230,230,230"
            .setCellBorderTop = "0.55"
            .setCellBorderTopColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellFontSize = "8"
            .addCell = Trim$(WS_GTP.IMPORTO) & "| "
        
            If (WS_FLG_INDE_LBL) Then
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 8.47)
                GoSub CHECK_ROWHEIGHT
            
                .setCellAlignH = "justified"
                .setCellColSpan = "9"
                .setCellFontName = "helr46w.ttf"
                .setCellHeight = "24"
                .addCell = "(*) Indennizzo automatico per mancato rispetto dei livelli specifici di qualità contrattuale definiti dall’Autorità per l’energia elettrica il gas e il sistema idrico. La corresponsione dell’indennizzo automatico non esclude la possibilità per il richiedente di richiedere nelle opportune sedi il risarcimento dell’eventuale danno ulteriore subito."
            End If
        
        Case "TS"
            If (WS_FLG_DF_IV) Then
                WS_FLG_DF_IV = False
            
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 5.29)
                GoSub CHECK_ROWHEIGHT
                
                .setCellColSpan = "9"
                .setCellHeight = "2"
                .addCell = ""
                
                .setCellAlignH = "left"
                .setCellColSpan = "7"
                .setCellFontName = "helr65w.ttf"
                .addCell = " Totale IVA"
                
                .setCellFontName = "helr65w.ttf"
                .addCell = NRM_IMPORT(WS_CALC_IV_TOT, "##,##0.00", False) + "| "
                                
                .setCellColSpan = "9"
                .setCellHeight = "4"
                .addCell = ""
            End If
        
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GTS), " ", False)
            CopyMemory ByVal VarPtr(WS_GTS), ByVal StrPtr(rowData), Len(WS_GTS) * 2
            
            If (WS_FLG_DF_TS = False) Then
                WS_FLG_DF_TS = True
                
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.88)
                GoSub CHECK_ROWHEIGHT
                
                .setCellBackColor = "230,230,230"
                .setCellAlignH = "left"
                .setCellColSpan = "9"
                .setCellFontName = "helr66w.ttf"
                .addCell = "Dettaglio totali"
                
                .setCellColSpan = "9"
                .setCellHeight = "2"
                .addCell = ""
            End If
        
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT

            .setCellAlignH = "left"
            .setCellColSpan = "7"
            .addCell = "  " & Replace$(Trim$(WS_GTS.DESCRIZIONE), "Totale Servizio Idrico", "Riepilogo Importi")
            
            .addCell = Trim$(WS_GTS.IMPORTO) & "|" & IIf(Trim$(WS_GTS.ALIQUOTA) = "", " ", Trim$(WS_GTS.ALIQUOTA))
        
        Case "VA"
            rowData = GET_TEXTPAD(PADLEFT, rowData, Len(WS_GVA), " ", False)
            CopyMemory ByVal VarPtr(WS_GVA), ByVal StrPtr(rowData), Len(WS_GVA) * 2

'            Select Case WS_GVA.PARCAU
'            Case "BPAD", "BPDD", "BPFD"
'                If (WS_FLG_MSG_BS = False) Then WS_FLG_MSG_BS = True
'
'            End Select
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            GoSub CHECK_ROWHEIGHT
            
            .setCellAlignH = "left"
            .addCell = "  " & Trim$(WS_GVA.DESCRIZIONE) & IIf(Trim$(WS_GVA.DESCRIZIONEAGGIUNTIVA) = "", "", "<br>  (" & Trim$(WS_GVA.DESCRIZIONEAGGIUNTIVA) & ")")
            
            If (Trim$(WS_GVA.PERIODO) = "") Then
                .setCellColSpan = "2"
                .addCell = ""
            Else
                WS_DATA = Split(WS_GVA.PERIODO, "-")
                
                .setCellAlignH = "center"
                .addCell = WS_DATA(0) & "|" & WS_DATA(1)
            End If
            
            If (Trim$(WS_GVA.UNITÀMISURAQUANTITÀ) = "gg") Then
                WS_GVA.TEMPO = Int(WS_GVA.QUANTITÀ)
                WS_GVA.QUANTITÀ = ""
            
                WS_GVA.UNITÀMISURATEMPO = WS_GVA.UNITÀMISURAQUANTITÀ
                WS_GVA.UNITÀMISURAQUANTITÀ = ""
            End If
            
            .addCell = Trim$(WS_GVA.TEMPO) & " " & Trim$(WS_GVA.UNITÀMISURATEMPO)
            
            If (Trim$(WS_GVA.CONCESSIONI) = "") Then
                If (Trim$(WS_GVA.QUANTITÀ) = "") Then
                    .setCellColSpan = "2"
                    .addCell = ""
                Else
                    .addCell = NRM_REMOVEZEROES(WS_GVA.QUANTITÀ, True)
                    
                    .setCellAlignH = "left"
                    .addCell = Trim$(WS_GVA.UNITÀMISURAQUANTITÀ)
                End If
            Else
                .addCell = Trim$(WS_GVA.CONCESSIONI)
                    
                .setCellAlignH = "left"
                .addCell = "un. imm."
            End If
            
            .addCell = Format$(Trim$(WS_GVA.PREZZO), "##,######0.000000") & "|" & Trim$(WS_GVA.IMPORTO) & "|" & Trim$(WS_GVA.ALIQUOTA)
            
        End Select
    End With
    
    Exit Sub
    
CHECK_ROWHEIGHT:
    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
        ReDim Preserve WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR)
        
        WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) = myXFDFMLTable.getXFDFTableNode
        WS_DOCPAGES_DF_CNTR = (WS_DOCPAGES_DF_CNTR + 1)
        WS_PAGEHEIGHT = 15
        
        ADD_PXX_DETTAGLIOBOLLETTA_TABLEHEADER True
    End If
Return

End Sub

Private Sub ADD_PXX_DETTAGLIOBOLLETTA_TABLEHEADER(addEmptyRow As Boolean)

    With myXFDFMLTable
        .setTableAlignH = "right"
        '.setTableBorders = "0.3"
        .setTableColumns = "10"
        .setTableFontName = "helr45w.ttf"
        .setTableFontSize = "7"
        .setTablePaddingTop = "0"
        .setTableWidths = "10,61,15,15,10,14,14,23.5,17.5,10"
         
        .addCell = ""

        .setCellAlignH = "left"
        .setCellColSpan = "9"
        .setCellFontName = "helr65w.ttf"
        .setCellHeight = "12"
        .addCell = "Dettaglio fattura"
         
        .setCellImage = "icnDetails.jpg"
        .setCellImageScale = "23.96"
        .setCellAlignV = "top"
        .setCellRowSpan = "999"
        .addCell = ""
            
        .setCellAlignH = "left"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .setCellHeight = "12"
        .addCell = "Descrizione"
        
        .setCellAlignH = "center"
        .setCellColSpan = "3"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .addCell = "Periodo"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .addCell = "Quantità"
        
        .setCellAlignH = "left"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .addCell = "UdM"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .addCell = "Prezzo Unitario (€)|Importo (€)|IVA(%)"
        
        If (addEmptyRow) Then
            .setCellColSpan = "9"
            .setCellHeight = "4"
            .addCell = ""
        
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 1.41)
        End If
    End With

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 8.47)

End Sub

Private Function GET_CHK_DA_ITEMSHEIGHT(WS_ITEMHEIGHT As Single) As Boolean

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + WS_ITEMHEIGHT)

    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
        WS_PAGEHEIGHT = 30

        ReDim Preserve WS_DOCPAGES_DA(WS_DOCPAGES_DA_CNTR)
        WS_DOCPAGES_DA(WS_DOCPAGES_DA_CNTR) = WS_DOCPAGES_DA(WS_DOCPAGES_DA_CNTR) & myXFDFMLTable.getXFDFTableNode
        
        ADD_PXX_ATTACH_TABLE_HDR

        WS_DOCPAGES_DA_CNTR = (WS_DOCPAGES_DA_CNTR + 1)
        GET_CHK_DA_ITEMSHEIGHT = True
    End If

End Function

Private Function GET_CHK_LF() As Boolean

    Dim I As Integer

    If (WS_CHK_G13) Then
        For I = 0 To UBound(WS_G13)
            GET_CHK_LF = (WS_G13(I).TIPOLETTURA = "F")
                    
            If (GET_CHK_LF) Then Exit For
        Next I
    End If

End Function

Private Function GET_DF_IV(varStr As String) As String

    Dim WS_INT      As Integer
    Dim WS_STRING   As String
        
    WS_INT = InStrRev(varStr, "€")
    
    If (WS_INT > 0) Then
        WS_STRING = Left$(varStr, WS_INT) & " " & Trim$(Mid$(varStr, (WS_INT + 1)))
    Else
        WS_STRING = varStr
    End If

    GET_DF_IV = WS_STRING

End Function

'Private Function GET_DYN_ANNXD_UI_CHART() As String
'
'    WS_ERRSCT = "GET_DYN_ANNXD_UI_CHART"
'
'    Dim I              As Integer
'    Dim WS_CHRT_DATA() As String
'    Dim WS_DATA        As String
'    Dim WS_ITEMS       As Integer
'    Dim WS_MAX         As Single
'    Dim WS_Y           As String
'
'    If (CHK_ARRAY(WS_ROW_DATA)) Then
'        WS_ITEMS = UBound(WS_ROW_DATA)
'
'        If (WS_ITEMS > 3) Then WS_ITEMS = 3
'
'        For I = 0 To WS_ITEMS
'            WS_CHRT_DATA = Split(WS_ROW_DATA(I), "|")
'
'            If (WS_MAX < CSng(WS_CHRT_DATA(1))) Then WS_MAX = CSng(WS_CHRT_DATA(1))
'
'            WS_DATA = WS_DATA & "<dataset><![CDATA[" & _
'                                Replace$(Format$(WS_CHRT_DATA(1), "0.00"), ",", ".") & _
'                                "|Consumo|" & _
'                                WS_CHRT_DATA(0) & _
'                                "]]></dataset>"
'        Next I
'
'        If (WS_ITEMS < 3) Then
'            For I = 2 To WS_ITEMS Step -1
'                WS_DATA = WS_DATA & "<dataset><![CDATA[0|Consumo|" & Space$(I) & "]]></dataset>"
'            Next I
'        End If
'
'        If (WS_MAX = 0) Then
'            WS_MAX = 0.1
'        Else
'            WS_MAX = (WS_MAX * 1.25)
'        End If
'
'        WS_Y = Replace$(Format$(WS_MAX, "0.00"), ",", ".")
'
'        If (WS_Y = "0.00") Then WS_Y = "0.1"
'
'        GET_DYN_ANNXD_UI_CHART = "<table columns=""2"" widths=""90,100"" padding=""0"" tableshiftY=""5"">" & _
'                                     "<cell chart=""true"" cellheight=""45mm"" chartHeight=""100"" chartWidth=""250"">" & _
'                                         "<chart type=""bars"" chartPadding=""4,0,0,0"">" & _
'                                             "<barRenderer customBarRenderer=""true"" barsColor=""70,150,245"" outLineStroke=""0.5"" outLineColor=""35,75,122"" showBaseItemLabel=""true"" bilAnchorOffset=""0.5""/>" & _
'                                             "<categoryPlot backColor=""255,255,255"" outlineVisible=""false"" rangeGridlinesVisible=""false"" margins=""0,0,0,0"" axisOffset=""0,0,0,0""/>" & _
'                                             "<categoryAxis tickMarksVisible=""false"" catlFontSize=""5.5"" catlAlignment=""right"" catlMargins=""0,0,0,0""/>" & _
'                                             "<valueAxis label=""mc/giorno"" catlMargins=""0,0,0,0"" axisRange=""0," & WS_Y & """/>" & _
'                                             WS_DATA & _
'                                         "</chart>" & _
'                                     "</cell>" & _
'                                     "<cell><![CDATA[]]></cell>" & _
'                                 "</table>"
'    End If
'
'End Function

Private Function GET_DYN_ANNXD_UI_DATA(ByRef HST_P02_CONSUMI As String) As String
    
    Dim I             As Integer
    Dim WS_COL_DATA() As String
    Dim WS_KEY        As String
    Dim WS_ROW_DATA() As String
    Dim WS_STRING     As String
    Dim WS_UI_DATA    As strct_DATA
    
    If (WS_FLG_CONSUMI) Then
        WS_STRING = "<text fontname=""helr66w.ttf"" fontsize=""9"" fontstyle=""underline"" alignment=""justified"">" & _
                        "<chunk><![CDATA[<br><br>Consumi dell’utenza raggruppata, con evidenza delle variazioni dei consumi medi giornalieri di acqua]]></chunk>" & _
                    "</text>" & _
                    "<table columns=""2"" widths=""90,100"" padding=""0"" tableshiftY=""5"">" & _
                        "<cell chart=""true"" cellheight=""45mm"" chartHeight=""120"" chartWidth=""250"">" & _
                        HST_P02_CONSUMI & _
                        "</cell>" & _
                        "<cell><![CDATA[]]></cell>" & _
                    "</table>"
    
        DDS_ADD "[$UI_CUR_CHRT]", WS_STRING
    Else
        DDS_ADD "[$UI_CUR_CHRT]", ""
    End If
    
    ' COMMON KEY/DATA
    '
    Dim dtlPages As New cls_PagesManager
    dtlPages.setMin = 35
    dtlPages.setMax = 270
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_GMG_TBL_DSCR.dataDescription & WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR.dataDescription, (Val(WS_CS_DYN_ATTCH_UI_GMG_TBL_DSCR.EXTRAPARAMS(0)) + Val(WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR.EXTRAPARAMS(0))), "", 0
    
    WS_KEY = Format$(WS_G02.R001.CODICESERVIZIO, "0000000000") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000")
    
    ' WS_GMC_DATA
    '
    WS_UI_DATA = GET_DATA_CACHE(DYN_ANNXD_UI_CMG, WS_KEY)

    If (WS_UI_DATA.dataDescription = WS_KEY) Then
        WS_STRING = Replace$(WS_CS_DYN_ATTCH_UI_GMC_TBL_ROW.dataDescription, "[$UI_GMC_COL_01]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_02]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_03]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_04]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_05]", "-")

        dtlPages.setData WS_CS_DYN_ATTCH_UI_GMC_TBL_FTR.dataDescription, WS_STRING, Val(WS_CS_DYN_ATTCH_UI_GMC_TBL_ROW.EXTRAPARAMS(0)), WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR.dataDescription, Val(WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR.EXTRAPARAMS(0))
    Else
        DDS_ADD "[$TBL_MSG_UC]", WS_CS_PXX_TBL_MSG_UC.dataDescription

        WS_ROW_DATA = Split(WS_UI_DATA.dataDescription, vbNewLine)
        WS_STRING = ""

        For I = 0 To UBound(WS_ROW_DATA)
            WS_COL_DATA = Split(WS_ROW_DATA(I), "|")

            WS_STRING = Replace$(WS_CS_DYN_ATTCH_UI_GMC_TBL_ROW.dataDescription, "[$UI_GMC_COL_01]", WS_COL_DATA(0))
            WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_02]", IIf((WS_COL_DATA(1) = ""), "-", WS_COL_DATA(1)))
            WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_03]", IIf((WS_COL_DATA(2) = ""), "-", WS_COL_DATA(2)))
            WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_04]", IIf((WS_COL_DATA(3) = ""), "-", WS_COL_DATA(3)))
            WS_STRING = Replace$(WS_STRING, "[$UI_GMC_COL_05]", IIf((WS_COL_DATA(4) = ""), "-", WS_COL_DATA(4)))
        
            dtlPages.setData WS_CS_DYN_ATTCH_UI_GMC_TBL_FTR.dataDescription, WS_STRING, Val(WS_CS_DYN_ATTCH_UI_GMC_TBL_ROW.EXTRAPARAMS(0)), WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR.dataDescription, Val(WS_CS_DYN_ATTCH_UI_GMC_TBL_HDR.EXTRAPARAMS(0))
        Next I
    End If
    
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_GMC_TBL_FTR.dataDescription, 0, "", 0
   
    ' WS_IFUR_DATA
    '
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_TBL_DSCR.dataDescription & WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR.dataDescription, (Val(WS_CS_DYN_ATTCH_UI_IFUR_TBL_DSCR.EXTRAPARAMS(0)) + Val(WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR.EXTRAPARAMS(0))), "", 0
    
    WS_UI_DATA = GET_DATA_CACHE(DYN_ANNXD_UI_IFUR, WS_KEY)
    
    If (WS_UI_DATA.dataDescription = WS_KEY) Then
        WS_STRING = Replace$(WS_CS_DYN_ATTCH_UI_IFUR_TBL_ROW.dataDescription, "[$UI_IFUR_COL_01]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_02]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_03]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_04]", "-")
        WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_05]", "-")
            
        dtlPages.setData WS_CS_DYN_ATTCH_UI_IFUR_TBL_FTR.dataDescription, WS_STRING, Val(WS_CS_DYN_ATTCH_UI_IFUR_TBL_ROW.EXTRAPARAMS(0)), WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR.EXTRAPARAMS(0))
    Else
        WS_ROW_DATA = Split(WS_UI_DATA.dataDescription, vbNewLine)
        WS_STRING = ""

        For I = 0 To UBound(WS_ROW_DATA)
            WS_COL_DATA = Split(WS_ROW_DATA(I), "|")

            WS_STRING = Replace$(WS_CS_DYN_ATTCH_UI_IFUR_TBL_ROW.dataDescription, "[$UI_IFUR_COL_01]", WS_COL_DATA(0))
            WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_02]", IIf((WS_COL_DATA(1) = ""), "-", WS_COL_DATA(1)))
            WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_03]", IIf((WS_COL_DATA(2) = ""), "-", WS_COL_DATA(2)))
            WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_04]", IIf((WS_COL_DATA(3) = ""), "-", WS_COL_DATA(3)))
            WS_STRING = Replace$(WS_STRING, "[$UI_IFUR_COL_05]", IIf((WS_COL_DATA(4) = ""), "-", WS_COL_DATA(4)))
            
            dtlPages.setData WS_CS_DYN_ATTCH_UI_IFUR_TBL_FTR.dataDescription, WS_STRING, Val(WS_CS_DYN_ATTCH_UI_IFUR_TBL_ROW.EXTRAPARAMS(0)), WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_TBL_HDR.EXTRAPARAMS(0))
        Next I
    End If
    
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_TBL_FTR.dataDescription, 0, "", 0
    
    Erase WS_COL_DATA()
    Erase WS_ROW_DATA()
    
    ' UI_CA
    '
    WS_STRING = Replace$(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_02.dataDescription, "[$UI_YR]", WS_G01.R001.ANNOBOLLETTA)
    WS_STRING = Replace$(WS_STRING, "[$UI_CA]", Trim$(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AC))
    
    ' MB_XX
    '
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_01.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_01.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_STRING, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_02.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_03.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_03.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_04.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_04.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_05.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_05.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_06.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_06.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_07.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_07.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_08.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_08.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_09.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_09.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_10.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_10.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_11.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_11.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_12.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_12.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_13.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_13.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_14.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_14.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_15.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_15.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_16.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_16.EXTRAPARAMS(0)), "", 0
    dtlPages.setData "", WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_17.dataDescription, Val(WS_CS_DYN_ATTCH_UI_IFUR_P02_MB_17.EXTRAPARAMS(0)), "", 0

    WS_DOCPAGES_UI_CNTR = dtlPages.getPages
    GET_DYN_ANNXD_UI_DATA = dtlPages.getData

End Function

Private Function GET_FLG_INFOPAGE() As Boolean

    WS_FLG_DOM = (Trim$(WS_G06.R002.DESCRIZIONEBANCA) <> "")
    
    If (CSng(WS_G06.R001.TOTALEEURO) <= 0) Then WS_FLG_NEG_NODOM = (WS_FLG_DOM = False)

End Function

Private Function GET_P01_COM_ARERA() As String

    If (WS_BDS = False) Then GET_P01_COM_ARERA = WS_CS_P01_MSG_COM_ARERA.dataDescription

End Function

Private Sub GET_P01_CONSUMI()
    
    Dim WS_STRING_LEFT  As String
    Dim WS_STRING_RIGHT As String
    
    If (Trim$(WS_G09.R008.CONSUMOANNUO) <> "") Then
        WS_STRING_LEFT = "Consumo Annuo"
        WS_STRING_RIGHT = Int(WS_G09.R008.CONSUMOANNUO) & " mc"
    End If

'    If (Trim$(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AC) <> "") Then
'        WS_STRING_LEFT = WS_STRING_LEFT & IIf((WS_STRING_LEFT = ""), "", "<br>") & "Consumo Anno Corrente"
'        WS_STRING_RIGHT = WS_STRING_RIGHT & IIf((WS_STRING_RIGHT = ""), "", "<br>") & Int(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AC) & " mc"
'    End If

'    If (Trim$(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AS) <> "") Then
'        WS_STRING_LEFT = WS_STRING_LEFT & IIf((WS_STRING_LEFT = ""), "", "<br>") & "Consumo Anno Succ."
'        WS_STRING_RIGHT = WS_STRING_RIGHT & IIf((WS_STRING_RIGHT = ""), "", "<br>") & Int(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AS) & " mc"
'    End If

    DDS_ADD "[$LBL_CONSUMI]", IIf((WS_STRING_LEFT = ""), "Consumo Annuo", WS_STRING_LEFT)
    DDS_ADD "[$VAR_CONSUMI]", IIf((WS_STRING_RIGHT = ""), " 0 mc", WS_STRING_RIGHT)

End Sub

Private Function GET_P01_CONSUMOFATTURATO() As String
    
    Select Case WS_G01.R001.TIPOBOLLETTAZIONE
    Case "A"
        GET_P01_CONSUMOFATTURATO = NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True) & " mc (Stimata)"
    
    Case "D"
        GET_P01_CONSUMOFATTURATO = NRM_REMOVEZEROES(WS_G11.R003.TOTALECONSUMOFATTURATO, True) & " mc (Effettiva + Stimata)"
    
    Case "L"
        GET_P01_CONSUMOFATTURATO = NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True) & " mc (Effettiva)"
    
    Case "M"
        GET_P01_CONSUMOFATTURATO = NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True) & " mc (Effettiva)<br>" & _
                                   NRM_REMOVEZEROES(WS_G11.R002.TOTALECONSUMOFATTURATO, True) & " mc (Stimata)"

    End Select

End Function

Private Function GET_P01_DETTAGLIOLETTURE_G13() As String

    Dim I        As Integer
    Dim WS_END   As Integer
    Dim WS_MAX   As Integer
    Dim WS_START As Integer
    
    If (WS_CHK_G13) Then
        WS_ERRSCT = "GET_P01_DETTAGLIOLETTURE_G13"
        WS_MAX = UBound(WS_G13)
    
        With myXFDFMLTable
            '.setTableBorders = "0.3"
            .setTableColumns = "4"
            .setTableFontName = "helr45w.ttf"
            .setTableFontSize = "7"
            .setTablePaddingTop = "0"
            .setTableWidths = "0.4,1,1,1.6"
        
            .addCell = ""
    
            .setCellAlignH = "left"
            .setCellColSpan = "3"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Letture e Consumi" & IIf((WS_G01.R001.TIPOBOLLETTAZIONE = "D"), " - Parte a Conguaglio", "")

            .setCellImage = "icnReadings.jpg"
            .setCellImageScale = "23.96"
            .setCellAlignV = "top"
            .setCellRowSpan = "999"
            .addCell = ""
                
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Data"
            
            .setCellAlignH = "right"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Lettura"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Tipo"
            
            Select Case WS_MAX
            Case 0
                .addCell = WS_G13(0).DATALETTURA
                
                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G13(0).LETTURA, True) & " mc"
                
                .addCell = IIf((WS_G13(0).TIPOLETTURA = "F"), "Finale", Trim$(WS_G13(0).TIPOLETTURAAEEG))
            
            Case 1
                ' START ROW
                '
                .addCell = WS_G13(0).DATALETTURA
                
                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G13(0).LETTURA, True) & " mc"
                
                .addCell = IIf((WS_G13(0).TIPOLETTURA = "F"), "Finale", Trim$(WS_G13(0).TIPOLETTURAAEEG))
                
                ' END ROW
                '
                .addCell = WS_G13(WS_MAX).DATALETTURA
                
                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G13(WS_MAX).LETTURA, True) & " mc"
                
                .addCell = IIf((WS_G13(WS_MAX).TIPOLETTURA = "F"), "Finale", Trim$(WS_G13(WS_MAX).TIPOLETTURAAEEG))
            
            Case Is > 1
                ' START ROW
                '
                .addCell = WS_G13(0).DATALETTURA

                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G13(0).LETTURA, True) & " mc"

                .addCell = IIf((WS_G13(0).TIPOLETTURA = "F"), "Finale", Trim$(WS_G13(0).TIPOLETTURAAEEG))

                ' INTERMEDIATE ROWS
                '
                WS_START = IIf((WS_MAX > 2), (WS_MAX - 2), 1)
                WS_END = (WS_MAX - 1)

                For I = WS_START To WS_END
                    .addCell = WS_G13(I).DATALETTURA

                    .setCellAlignH = "right"
                    .addCell = NRM_REMOVEZEROES(WS_G13(I).LETTURA, True) & " mc"

                    .addCell = IIf((WS_G13(I).TIPOLETTURA = "F"), "Finale", Trim$(WS_G13(I).TIPOLETTURAAEEG))
                Next I

                ' END ROW
                '
                .addCell = WS_G13(WS_MAX).DATALETTURA

                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G13(WS_MAX).LETTURA, True) & " mc"

                .addCell = IIf((WS_G13(WS_MAX).TIPOLETTURA = "F"), "Finale", Trim$(WS_G13(WS_MAX).TIPOLETTURAAEEG))

            End Select
 
            If (WS_G01.R001.TIPOBOLLETTAZIONE <> "D") Then
                .setCellAlignV = "top"
                .setCellAlignH = "justified"
                .setCellChuncked = "true"
                .setCellColSpan = "3"
                .setCellBorderTop = "0.55"
                .setCellBorderTopColor = "0,55,110"
                '.setCellHeight = "25"
                .addCell = GET_XML("Le eventuali letture intermedie sono disponibili in seconda pagina nella sezione Dettaglio Letture.<br>|<b>Modalità per comunicare l’autolettura: vedi box informativo nella pagina seguente.", "left", "helr45w.ttf", "7", "7", "110,0,0")
            End If
            
            GET_P01_DETTAGLIOLETTURE_G13 = .getXFDFTableNode
        End With
    End If

End Function

Private Function GET_P01_DETTAGLIOLETTURE_G14() As String

    Dim I        As Integer
    Dim WS_END   As Integer
    Dim WS_MAX   As Integer
    Dim WS_START As Integer

    If (WS_CHK_G14) Then
        WS_ERRSCT = "GET_P01_DETTAGLIOLETTURE_G14"
        WS_MAX = UBound(WS_G14)
    
        With myXFDFMLTable
            '.setTableBorders = "0.3"
            .setTableColumns = "4"
            .setTableFontName = "helr45w.ttf"
            .setTableFontSize = "7"
            .setTablePaddingTop = "0"
            .setTableWidths = "0.4,1,1,1.6"
            .setTableShiftY = "4"
        
            .addCell = ""
    
            .setCellAlignH = "left"
            .setCellColSpan = "3"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Letture e Consumi" & IIf((WS_G01.R001.TIPOBOLLETTAZIONE = "D"), " - Parte in Acconto", "")

            .setCellRowSpan = "999"
            .addCell = ""
                
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Data"
            
            .setCellAlignH = "right"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Lettura"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Tipo"
            
            Select Case WS_MAX
            Case 0
                .addCell = WS_G14(0).DATALETTURA
                
                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G14(0).LETTURA, True) & " mc"
                
                .addCell = IIf((WS_G14(0).TIPOLETTURA = "F"), "Finale", Trim$(WS_G14(0).TIPOLETTURAAEEG))
            
            Case 1
                ' START ROW
                '
                .addCell = WS_G14(0).DATALETTURA
                
                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G14(0).LETTURA, True) & " mc"
                
                .addCell = IIf((WS_G14(0).TIPOLETTURA = "F"), "Finale", Trim$(WS_G14(0).TIPOLETTURAAEEG))
                
                ' END ROW
                '
                .addCell = WS_G14(WS_MAX).DATALETTURA
                
                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G14(WS_MAX).LETTURA, True) & " mc"
                
                .addCell = IIf((WS_G14(WS_MAX).TIPOLETTURA = "F"), "Finale", Trim$(WS_G14(WS_MAX).TIPOLETTURAAEEG))
            
            Case Is > 1
                ' START ROW
                '
                .addCell = WS_G14(0).DATALETTURA

                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G14(0).LETTURA, True) & " mc"

                .addCell = IIf((WS_G14(0).TIPOLETTURA = "F"), "Finale", Trim$(WS_G14(0).TIPOLETTURAAEEG))

                ' INTERMEDIATE ROWS
                '
                WS_START = IIf((WS_MAX > 2), (WS_MAX - 2), 1)
                WS_END = (WS_MAX - 1)

                For I = WS_START To WS_END
                    .addCell = WS_G14(I).DATALETTURA

                    .setCellAlignH = "right"
                    .addCell = NRM_REMOVEZEROES(WS_G14(I).LETTURA, True) & " mc"

                    .addCell = IIf((WS_G14(I).TIPOLETTURA = "F"), "Finale", Trim$(WS_G14(I).TIPOLETTURAAEEG))
                Next I

                ' END ROW
                '
                .addCell = WS_G14(WS_MAX).DATALETTURA

                .setCellAlignH = "right"
                .addCell = NRM_REMOVEZEROES(WS_G14(WS_MAX).LETTURA, True) & " mc"

                .addCell = IIf((WS_G14(WS_MAX).TIPOLETTURA = "F"), "Finale", Trim$(WS_G14(WS_MAX).TIPOLETTURAAEEG))

            End Select

            .setCellAlignV = "top"
            .setCellAlignH = "justified"
            .setCellChuncked = "true"
            .setCellColSpan = "3"
            .setCellBorderTop = "0.55"
            .setCellBorderTopColor = "0,55,110"
            '.setCellHeight = "25"
            .addCell = GET_XML("Le eventuali letture intermedie sono disponibili in seconda pagina nella sezione Dettaglio Letture.<br>|<b>Modalità per comunicare l’autolettura: vedi box informativo nella pagina seguente.", "left", "helr45w.ttf", "7", "7", "110,0,0")
            
            GET_P01_DETTAGLIOLETTURE_G14 = .getXFDFTableNode
        End With
    End If

End Function

Private Function GET_P01_INFOBOLLETTA() As String
    
    Dim I                   As Integer
    Dim WS_BOOLEAN          As Boolean
    Dim WS_INT              As Integer
    Dim WS_STRING           As String
    
    WS_ERRSCT = "GET_P01_INFOBOLLETTA"
    
    Select Case WS_G01.R001.TIPOBOLLETTAZIONE
    Case "A"
        If (WS_FLG_NOTACREDITO) Then
            WS_STRING = "Questa Nota di Credito si riferisce al consumo di mc. " & NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True) & " restituito per il periodo dal " & WS_G11.R001.DATALETTURAPRECEDENTE & " al " & WS_G01.R001.DATALETTURAATTUALE & "."
        Else
            WS_STRING = "In questa bolletta di acconto è stato attribuito un consumo di mc. " & NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True) & " per il periodo dal " & WS_G11.R001.DATALETTURAPRECEDENTE & " al " & WS_G01.R001.DATALETTURAATTUALE & " tenendo conto delle informazioni disponibili circa i suoi consumi abituali."
        End If
    
    Case "L", "D"
        WS_INT = -1

        For I = UBound(WS_G13) To 0 Step -1
            WS_INT = I
            WS_BOOLEAN = (Trim$(WS_G13(WS_INT).AUTOLETTURA) = "S")

            Exit For
        Next I

        If (WS_BOOLEAN = False) Then WS_INT = UBound(WS_G13)
        
        If (WS_FLG_NOTACREDITO) Then
            WS_STRING = "Questa Nota di Credito si riferisce alla lettura di mc. " & NRM_REMOVEZEROES(WS_G13(WS_INT).LETTURA, True) & IIf(WS_BOOLEAN, " da Lei comunicata", " rilevata") & " in data " & WS_G13(WS_INT).DATALETTURA & ". " & _
                        "Il consumo restituito per il periodo dal " & WS_G11.R001.DATALETTURAPRECEDENTE & " al " & WS_G11.R001.DATALETTURAATTUALE & " è di mc. "
        Else
            WS_STRING = "Questa bolletta è stata calcolata utilizzando la lettura di mc. " & NRM_REMOVEZEROES(WS_G13(WS_INT).LETTURA, True) & IIf(WS_BOOLEAN, " da Lei comunicata", " rilevata") & " in data " & WS_G13(WS_INT).DATALETTURA & ". " & _
                        "Il consumo rilevato per il periodo dal " & WS_G11.R001.DATALETTURAPRECEDENTE & " al " & WS_G11.R001.DATALETTURAATTUALE & " è di mc. "
        End If
        
        If (WS_FLG_MASTER) Then
            WS_STRING = WS_STRING & NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMORILEVATO, True) & ". Il consumo totale dei divisionali è di mc. "
            
            If (WS_FLG_MASTER_ECC_POS) Then
                WS_STRING = WS_STRING & (NRM_REMOVEZEROES(WS_G11.R001.CONSUMO_SOMMARE_DETRARRE, True) * -1) & " pertanto il valore fatturato ammonta a mc. " & NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True)
            Else
                If (WS_FORNITURA_MASTER_P02.dataDescription = "TRG_PARAM") Then
                    If (WS_FORNITURA_MASTER_P02.EXTRAPARAMS(10) = "") Then
                        WS_STRING = WS_STRING & "0"
                    Else
                        WS_STRING = WS_STRING & Format$(Val(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(10) * -1), "#,##")
                    End If
                Else
                    WS_STRING = WS_STRING & "0"
                End If

                WS_STRING = WS_STRING & " pertanto il valore fatturato ammonta a mc. 0"
            End If
        Else
            If (Trim$(WS_G11.R001.CONSUMO_SOMMARE_DETRARRE) = "0,000000") Then
                WS_STRING = WS_STRING & NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True)
            Else
                WS_STRING = WS_STRING & NRM_REMOVEZEROES(WS_G11.R001.CONSUMO_ATT_PREC, True) & ". La quota relativa al centrale è di mc. " & NRM_REMOVEZEROES(WS_G11.R001.CONSUMO_SOMMARE_DETRARRE, True) & " pertanto il valore fatturato ammonta a mc. " & NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True)
            End If
        End If
        
        If (Trim$(WS_G11.R001.MC_ACCONTIFATTURATI) <> "0,000000") Then WS_STRING = WS_STRING & " e sono stati detratti gli importi relativi a mc. " & NRM_REMOVEZEROES(WS_G11.R001.MC_ACCONTIFATTURATI, True) & " fatturati nelle precedenti bollette relativi al periodo dal " & WS_G11.R001.DATAINIZIO_CONSUMISTIMATI_BOLPREC & " al " & WS_G11.R001.DATAFINE_CONSUMISTIMATI_BOLPREC
        
        If (WS_G01.R001.TIPOBOLLETTAZIONE = "D") Then WS_STRING = WS_STRING & ".<br>Inoltre è stato addebitato un consumo in acconto di mc. " & NRM_REMOVEZEROES(WS_G11.R002.TOTALECONSUMOFATTURATO, True) & " per il periodo dal " & WS_G11.R002.DATALETTURAPRECEDENTE & " al " & WS_G01.R001.DATALETTURAATTUALEACCONTO & " tenendo conto delle informazioni disponibili circa i suoi consumi abituali"

        WS_STRING = WS_STRING & "."
        
    Case "P"
        WS_STRING = "In questa bolletta non sono fatturati consumi."
    
    End Select

    With myXFDFMLTable
        '.setTableBorders = "0.3"
        .setTableColumns = "1"
        .setTableFontName = "helr45w.ttf"
        .setTableFontSize = "7"
        .setTableWidths = "1"
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.65"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "4"
        .setCellRoundedBorders = "0.7,3.0,1,1,0,0"
        .addCell = "Info Bolletta"
        
        .setCellAlignH = "justified"
        .setCellPaddingBottom = "3"
        .setCellRoundedBorders = "0.7,3.0,0,0,1,1"
        .addCell = WS_STRING
        
        GET_P01_INFOBOLLETTA = .getXFDFTableNode
    End With
    
End Function

Private Function GET_P01_PAGAMENTIPRECEDENTI() As String

    If (WS_FLG_WMS = False) Then
        WS_ERRSCT = "GET_P01_PAGAMENTIPRECEDENTI"
                
        If (Trim$(WS_G21.R999.IMPORTOTOTALE) = "") Then
            GET_P01_PAGAMENTIPRECEDENTI = GET_XML(String$(14, " ") & "<br>|<b>Pagamenti Precedenti<br>|" & String$(14, " ") & "I pagamenti delle precedenti fatture sono regolari. Grazie.", "left", "helr45w.ttf", "8", "8")
        Else
            If (CSng(WS_G21.R999.IMPORTOTOTALE) > 0) Then
                With myXFDFMLTable
                    '.setTableBorders = "0.3"
                    .setTableColumns = "2"
                    .setTableFontName = "helr45w.ttf"
                    .setTableFontSize = "7"
                    .setTablePaddingTop = "0"
                    .setTableShiftY = "6"
                    .setTableWidths = GET_TABLECOLUMNS_SIZE(9, "0.9,8.1")
        
                    .setCellAlignV = "top"
                    .setCellAlignH = "center"
                    .setCellImage = "icnWarning_Red.jpg"
                    .setCellImageScale = "23.96"
                    .setCellRowSpan = "99"
                    .addCell = ""
                    
                    .setCellBackColor = "240,220,220"
                    .setCellBorderBottom = "0.55"
                    .setCellBorderBottomColor = "110,0,0"
                    .setCellFontName = "helr65w.ttf"
                    .setCellForeColor = "110,0,0"
                    .setCellHeight = "10"
                    .addCell = "DA PAGARE (Fatture Scadute e a Scadere)"
                    
                    .addCell = ""
                    
                    .setCellAlignH = "justified"
                    .setCellChuncked = True
                    .addCell = GET_XML("Risultano importi non pagati per un totale di |<b>Euro " & NRM_IMPORT(WS_G21.R999.IMPORTOTOTALE, "#,##0.00", False) & "| (incassi aggiornati indicativamente a 30 gg precedenti l'emissione della fattura se il pagamento è avvenuto con PagoPA, carta di credito e delega bancaria di pagamento SDD; se effettuati con modalità differenti, gli incassi potrebbero essere aggiornati con un ritardo superiore ai 120 giorni).", "left", "helr45w.ttf", "7", "7", "110,0,0")
                    
                    GET_P01_PAGAMENTIPRECEDENTI = .getXFDFTableNode
                End With
            Else
                GET_P01_PAGAMENTIPRECEDENTI = GET_XML(String$(14, " ") & "|<b>Pagamenti Precedenti<br>|" & String$(14, " ") & "I pagamenti delle precedenti fatture sono regolari. Grazie.", "left", "helr45w.ttf", "8", "8")
            End If
        End If
    Else
        DDS_ADD "[$TXT_PXX_FOOTER_MS335]", WS_PXX_FOOTER_WSM
    End If

End Function

Private Function GET_P01_PAYMODE() As String

    WS_ERRSCT = "GET_P01_PAYMODE"

    Dim WS_ADDEBITOACCREDITO As Single
    Dim WS_FONTSIZE          As String
    Dim I                    As Integer

    WS_FONTSIZE = 6

    If (WS_BDS) Then
        GET_P01_PAYMODE = "Per le modalità di riscossione della presente bolletta a Suo credito consulti la pagina delle Note Informative."
        WS_FONTSIZE = "6.5"
    Else
        Select Case CSng(WS_G06.R001.TOTALEEURO)
        Case 0
            For I = 0 To UBound(WS_GSI)
                If (WS_GSI(I).TIPOLOGIASEZIONE = "Y") Then
                    WS_ADDEBITOACCREDITO = CSng(WS_GSI(I).IMPORTO_IMPOSTA)
                                
                    Exit For
                End If
            Next I
            
            Select Case WS_ADDEBITOACCREDITO
            Case 0
                GET_P01_PAYMODE = "In questa bolletta non c'è nulla da pagare."
            
            Case Is > 0
                GET_P01_PAYMODE = "L'importo di € " & NRM_IMPORT(WS_ADDEBITOACCREDITO, "#,##0.00", False) & " Le sarà accreditato sulla prossima bolletta."
                WS_FONTSIZE = "7.5"
            
            Case Is < 0
                GET_P01_PAYMODE = "L'importo di € " & NRM_IMPORT(Abs(WS_ADDEBITOACCREDITO), "#,##0.00", False) & " Le sarà addebitato sulla prossima bolletta."
                WS_FONTSIZE = "7.5"
            
            End Select
            
        Case Is > 0
            'If (WS_FLG_FATTELE_PA = False) Then
                If (WS_FLG_DOM) Then
                    If (Trim$(WS_G09.R005.FLAGRATEIZZABILITÀ) = "S") Then
                        GET_P01_PAYMODE = "La bolletta, come richiesto, sarà addebitata salvo buon fine, il giorno della scadenza sul conto corrente da Lei indicato.<br>L’importo supera del 80% il valore dell’addebito medio riferito alle bollette emesse nel corso degli ultimi 12 mesi. Per rateizzarla contatta il nostro Servizio Clienti."
                    Else
                        GET_P01_PAYMODE = "La bolletta, come richiesto, sarà addebitata salvo buon fine, il giorno della scadenza sul conto corrente da Lei indicato."
                    End If
                Else
                    If (WS_CCP) Then
                        If ((Trim$(WS_G09.R005.FLAGRATEIZZABILITÀ) = "N") And (Trim$(WS_G09.R005.RATEIZZATO_NORMATIVA) <> "S")) Then
                            GET_P01_PAYMODE = "Per il pagamento della presente fattura è possibile utilizzare l’Avviso di Pagamento PagoPA allegato o le altre modalità indicate nella sezione finale Modalità di pagamento."
                        ElseIf ((Trim$(WS_G09.R005.FLAGRATEIZZABILITÀ) = "S") And (Trim$(WS_G09.R005.RATEIZZATO_NORMATIVA) <> "S")) Then
                            GET_P01_PAYMODE = "L’importo di questa fattura supera del 80% il valore dell’addebito medio riferito alle bollette emesse nel corso degli ultimi 12 mesi, per cui è dilazionabile con le modalità indicate nell’apposito box.<br>Per il pagamento in un’unica soluzione è possibile utilizzare il è possibile utilizzare l’Avviso di Pagamento PagoPA allegato o le altre modalità indicate nella sezione finale Modalità di pagamento."
                        ElseIf (Trim$(WS_G09.R005.RATEIZZATO_NORMATIVA) = "S") Then
                            GET_P01_PAYMODE = "L’importo di questa fattura supera del 150% il valore dell’addebito medio riferito alle bollette emesse nel corso degli ultimi 12 mesi; al documento di fatturazione sono allegati gli Avvisi di Pagamento PagoPA per il pagamento rateale dell'importo dovuto. La fattura potrà essere pagata in un’unica soluzione o ratealmente utilizzando gli Avvisi allegati."
                        Else
                            If (UBound(WS_G23) = 0) Then
                                GET_P01_PAYMODE = "Per il pagamento della presente fattura è possibile utilizzare l’Avviso di Pagamento PagoPA allegato o le altre modalità indicate nella sezione finale Modalità di pagamento."
                            Else
                                GET_P01_PAYMODE = "Pagamento da effettuare entro la scadenza indicata nella rata unica o delle " & UBound(WS_G23) & " rate di Avviso di Pagamento PagoPA allegato."
                            End If
                        End If
                    End If
                End If
            'End If
        
        Case Is < 0
            GET_P01_PAYMODE = "Per le modalità di riscossione della presente bolletta a Suo credito consulti la pagina delle Note Informative."
        
        End Select
    End If

    DDS_ADD "[$VAR_PM_FS]", WS_FONTSIZE

End Function

Private Function GET_P01_PAYTYPE() As String

    If (WS_FLG_FATTELE_PA) Then
        WS_ERRSCT = "GET_P01_PAYTYPE"
        
        With myXFDFMLTable
            '.setTableBorders = "0.3"
            .setTableColumns = "2"
            .setTableFontName = "helr45w.ttf"
            .setTableFontSize = "7"
            .setTablePaddingTop = "0"
            .setTableShiftY = "5"
            .setTableWidths = GET_TABLECOLUMNS_SIZE(9, "1,8")
    
            .addCell = ""
    
            .setCellAlignH = "left"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellColSpan = "4"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Pagamenti"
    
            .addCell = ""
            
            .setCellAlignH = "justified"
            .addCell = "Pagamento da effettuare tramite bonifico utilizzando le coordinate internazionali riportate nella sezione MODALITA’ DI PAGAMENTO."
            
            GET_P01_PAYTYPE = .getXFDFTableNode
        End With
    End If

End Function

Private Sub GET_P01_PUNTO_EROGAZIONE()
    
    Dim WS_FORNITURA_MASTER_P01 As strct_DATA
    Dim WS_STRING_LEFT      As String
    Dim WS_STRING_RIGHT     As String
    
    WS_ERRSCT = "GET_P01_PUNTO_EROGAZIONE"

    If (WS_FLG_CATEGORIA) Then
        WS_FORNITURA_MASTER_P01 = GET_DATA_CACHE(FORNITURA_MASTER_PHASE01, WS_CODICE_SERVIZIO_KEY)
        
        If (WS_FORNITURA_MASTER_P01.dataDescription = "TRG_PARAM") Then
            WS_FLG_DIV = True
        
            WS_STRING_LEFT = "<br>——————————<br>Dati Fornitura Master<br>Matricola Contatore<br>Codice Servizio<br>Codice cliente<br>Punto Erogazione"

            With WS_FORNITURA_MASTER_P01
                WS_CNTTR_MASTER_MTRCL = .EXTRAPARAMS(4)
                WS_CNTTR_MASTER_SRVZ = .EXTRAPARAMS(0)
                WS_STRING_RIGHT = "<br>——————————————————<br><br>" & .EXTRAPARAMS(4) & "<br>" & .EXTRAPARAMS(0) & "<br>" & .EXTRAPARAMS(1) & "<br>" & .EXTRAPARAMS(3)
            End With

            If (WS_G01.R001.TIPOBOLLETTAZIONE <> "P") Then
                WS_FLG_MASTER = (Trim$(WS_G02.R001.CODICESERVIZIO) = WS_CNTTR_MASTER_SRVZ)
                
                If (WS_FLG_MASTER) Then WS_FLG_MASTER_ECC_POS = ((CSng(WS_G11.R001.TOTALECONSUMOFATTURATO) > 0) Or ((CSng(WS_G11.R001.TOTALECONSUMOFATTURATO) = 0) And (CSng(WS_G11.R001.CONSUMO_SOMMARE_DETRARRE) = 0)))
            
                WS_FORNITURA_MASTER_P02 = GET_DATA_CACHE(FORNITURA_MASTER_PHASE02_F, WS_CODICE_SERVIZIO_KEY)
                
                If (WS_FORNITURA_MASTER_P02.dataDescription = "TRG_PARAM") Then If (WS_FORNITURA_MASTER_P02.EXTRAPARAMS(0) = "P") Then WS_FORNITURA_MASTER_P02 = GET_DATA_CACHE(FORNITURA_MASTER_PHASE02_P, WS_CODICE_SERVIZIO_KEY)
            End If
        End If
    End If

    DDS_ADD "[$LBL_PNTERO]", "Punto Erogazione" & WS_STRING_LEFT
    DDS_ADD "[$VAR_PNTERO]", Trim$(WS_G09.R002.CODICEPUNTORICONSEGNA) & WS_STRING_RIGHT

End Sub

Private Function GET_P01_QUADROSINTETICO() As String

    Dim I             As Integer
    Dim WS_BOOLEAN_OP As Boolean
    Dim WS_CALC       As Single
    Dim WS_INT        As Integer
    Dim WS_STRING     As String

    WS_ERRSCT = "GET_P01_QUADROSINTETICO"
    
    For I = 0 To UBound(WS_GSI)
        If (InStr(1, WS_GSI(I).DESCRIZIONESINTETICO, "Oneri perequazione") > 0) Then
            If (WS_BOOLEAN_OP = False) Then WS_BOOLEAN_OP = True
            
            WS_CALC = (WS_CALC + CSng(WS_GSI(I).IMPORTO_IMPOSTA))
        End If
    Next I

    With myXFDFMLTable
        '.setTableBorders = "0.3"
        .setTableColumns = "3"
        .setTableFontName = "helr45w.ttf"
        .setTableFontSize = "7"
        .setTablePaddingTop = "0"
        .setTableWidths = "0.33,1.87,0.8"
        
        .addCell = ""

        .setCellAlignH = "left"
        .setCellColSpan = "2"
        .setCellFontName = "helr65w.ttf"
        .setCellHeight = "10"
        .addCell = "Quadro di Sintesi"
        
        .setCellAlignH = "center"
        .setCellImage = "icnCurrency.jpg"
        .setCellImageScale = "23.96"
        .setCellAlignV = "top"
        .setCellRowSpan = "99"
        .addCell = ""
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .setCellHeight = "10"
        .addCell = "Descrizione"
        
        .setCellAlignH = "right"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .addCell = "Importo"
        
        For I = 0 To UBound(WS_GSI)
            Select Case WS_GSI(I).TIPOLOGIASEZIONE
            Case "P"
                .setCellBackColor = "230,230,230"
                .setCellBorderTop = "0.55"
                .setCellBorderTopColor = "0,55,110"
                .setCellFontName = "helr65w.ttf"
                .addCell = Trim$(WS_GSI(I).DESCRIZIONESINTETICO)
                
                .setCellAlignH = "right"
                .setCellBackColor = "230,230,230"
                .setCellBorderTop = "0.55"
                .setCellBorderTopColor = "0,55,110"
                .setCellFontName = "helr65w.ttf"
                .addCell = WS_GSI(I).UNITAMISURA & " " & Trim$(WS_GSI(I).IMPORTO_IMPOSTA)
            
            Case Else
                'If ((WS_GSI(I).TIPOLOGIASEZIONE <> "B") And (WS_GSI(I).TIPOLOGIASEZIONE <> "Y")) Then
                If (WS_GSI(I).TIPOLOGIASEZIONE <> "B") Then
                    If (WS_GSI(I).TIPOLOGIASEZIONE = "X") Then
                        WS_STRING = "Azzeramento bolletta"
                    Else
                        WS_STRING = Trim$(WS_GSI(I).DESCRIZIONESINTETICO)
                        
                        WS_INT = InStrRev(WS_STRING, "€")
                        If (WS_INT > 0) Then WS_STRING = Left$(WS_STRING, WS_INT) & " " & NRM_IMPORT(Mid$(WS_STRING, WS_INT + 1), "##,##0.00", False)
                    End If
                                        
                    If ((InStr(1, WS_STRING, "Oneri perequazione") > 0) And WS_BOOLEAN_OP) Then
                        WS_STRING = "Oneri perequazione"
                        
                        .addCell = WS_STRING
                        
                        .setCellAlignH = "right"
                        .addCell = WS_GSI(I).UNITAMISURA & " " & NRM_IMPORT(WS_CALC, "#,##0.00", False)
                        
                        WS_BOOLEAN_OP = False
                    Else
                        If (InStr(1, WS_STRING, "Oneri perequazione") = 0) Then
                            .addCell = WS_STRING
                            
                            .setCellAlignH = "right"
                            .addCell = WS_GSI(I).UNITAMISURA & " " & Trim$(WS_GSI(I).IMPORTO_IMPOSTA)
                            
                            'If ((WS_MSG_CI = False) And (WS_FLG_NOTACREDITO = False) And (WS_STRING = "Depurazione") And (WS_IMPORTO_TC_TICSI <> "")) Then
                            '    .addCell = "Adeguamenti tariffari TICSI"
                            '
                            '    .setCellAlignH = "right"
                            '    .addCell = "€ " & NRM_IMPORT(WS_IMPORTO_TC_TICSI, "#,##0.00", False)
                            'End If
                        End If
                    End If
                End If
            
            End Select
        Next I
            
        GET_P01_QUADROSINTETICO = .getXFDFTableNode
    End With

End Function

Private Function GET_P01_TICSI() As String

    If (WS_FLG_PARTITE And (WS_CHK_G16 = False)) Then Exit Function
    
    WS_ERRSCT = "GET_P01_TICSI"
    
    Dim I As Integer
    
    With myXFDFMLTable
        '.setTableBorders = "0.3"
        .setTableColumns = "6"
        .setTableFontName = "helr45w.ttf"
        .setTableFontSize = "7"
        .setTablePaddingTop = "0"
        .setTableShiftY = "5"
        .setTableWidths = "1.1,3.7,0.5,1,0.7,3.5"
    
        ' HEADER
        '
        .addCell = ""

        .setCellColSpan = "5"
        .setCellHeight = "10"
        
        If (WS_BDS) Then
            .addCell = ""
        Else
            .setCellAlignH = "left"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Numero Totale delle Unità Immobiliari: " & Trim$(WS_G09.R003.NUMEROTOTALECONCESSIONI)
        End If

        .setCellImage = "icnHouse.jpg"
        .setCellImageScale = "23.96"
        .setCellAlignV = "top"
        .setCellRowSpan = "20"
        .addCell = ""
            
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .setCellHeight = "12"
        .addCell = "Tipologia"
        
        If (WS_BDS) Then
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .addCell = ""
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .addCell = ""
    
            .setCellAlignH = "right"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .addCell = ""
        Else
            .setCellAlignH = "center"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .addCell = "R"
            
            .setCellAlignH = "right"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .addCell = "C.N.F."
    
            .setCellAlignH = "right"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .addCell = "U.I."
        End If
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellFontName = "helr65w.ttf"
        .addCell = "Tariffa"

        ' DATA
        '
        If (WS_FLG_NOTACREDITO And (WS_CHK_G16 = False)) Then
            .setCellFontSize = "6.5"
            .addCell = Trim$(WS_G17(0).DESCRIZIONE_TIPOLOGIA_RIGA)
            
            .setCellAlignH = "center"
            .addCell = WS_G17(0).RESIDENTE
            
            .setCellAlignH = "right"
            .addCell = WS_G17(0).NUCLEO_FAMILIARE
            
            .setCellAlignH = "right"
            .addCell = Trim$(WS_G17(0).NUMERO_CONCESSIONI)
            
            .addCell = Trim$(WS_G17(0).DESCRIZIONE_TARIFFA_APPLICATA)
        Else
            For I = 0 To UBound(WS_G16)
                .setCellFontSize = "6.5"
                .addCell = Trim$(WS_G16(I).DESCRIZIONE_TIPOLOGIA_RIGA)
                
                .setCellAlignH = "center"
                .addCell = WS_G16(I).RESIDENTE
                
                .setCellAlignH = "right"
                .addCell = Trim$(WS_G16(I).NUCLEO_FAMILIARE)
                
                .setCellAlignH = "right"
                .addCell = Trim$(WS_G16(I).NUMERO_CONCESSIONI)
                
                .addCell = Trim$(WS_G16(I).DESCRIZIONE_TARIFFA_APPLICATA)
            Next I
        End If
        
        ' FOOTER
        '
        .setCellAlignV = "top"
        .setCellAlignH = "justified"
        .setCellColSpan = "5"
        .setCellBorderTop = "0.55"
        .setCellBorderTopColor = "0,55,110"
        
        If (WS_BDS) Then
            .addCell = ""
        Else
            .setCellChuncked = True
            .addCell = GET_XML("<b>R| = Residente, |<b>C.N.F.| = Comp. Nucleo Familiare, |<b>U.I.| = Unità Immobiliari.", "justified", "helr45w.ttf", "6", "6")
        End If
        
        GET_P01_TICSI = .getXFDFTableNode
    End With

End Function

Private Function GET_P02_DETTAGLIOCONSUMI_CHART() As String

    WS_ERRSCT = "GET_P02_DETTAGLIOCONSUMI_CHART"

    Dim I        As Integer
    Dim WS_DATA  As String
    Dim WS_ITEMS As Integer
    Dim WS_MAX   As Single
    Dim WS_Y     As String
    
    WS_ITEMS = UBound(WS_G22)
    
    If (WS_ITEMS > 3) Then WS_ITEMS = 3
    
    For I = WS_ITEMS To 0 Step -1
        If (WS_MAX < CSng(WS_G22(I).CONSUMOMEDIOGIORNALIEROPERIODO)) Then WS_MAX = CSng(WS_G22(I).CONSUMOMEDIOGIORNALIEROPERIODO)
    
        WS_DATA = WS_DATA & "<dataset><![CDATA[" & _
                            Replace$(Format$(WS_G22(I).CONSUMOMEDIOGIORNALIEROPERIODO, "0.00"), ",", ".") & _
                            "|Consumo|" & _
                            "Dal " & Format$(WS_G22(I).DATAINIZIALEPERIODO, "dd/MM/yy") & "<br>al " & Format$(WS_G22(I).DATAFINALEPERIODO, "dd/MM/yy") & _
                            "]]></dataset>"
    Next I
    
    If (WS_ITEMS < 3) Then
        For I = 2 To WS_ITEMS Step -1
            WS_DATA = WS_DATA & "<dataset><![CDATA[0|Consumo|" & Space$(I) & "]]></dataset>"
        Next I
    End If
    
    If (WS_MAX = 0) Then
        WS_MAX = 0.1
    Else
        WS_MAX = (WS_MAX * 1.25)
    End If
    
    WS_Y = Replace$(Format$(WS_MAX, "0.00"), ",", ".")
    
    If (WS_Y = "0.00") Then WS_Y = "0.1"
        
    If (WS_FLG_NOTACREDITO) Then GET_P02_DETTAGLIOCONSUMI_CHART = "<textTitle label=""Il grafico si riferisce ai consumi restituiti""/>"

    GET_P02_DETTAGLIOCONSUMI_CHART = "<chart type=""bars"" chartPadding=""4,0,0,0"">" & _
                                          GET_P02_DETTAGLIOCONSUMI_CHART & _
                                          "<barRenderer customBarRenderer=""true"" barsColor=""70,150,245"" outLineStroke=""0.5"" outLineColor=""35,75,122"" showBaseItemLabel=""true"" bilAnchorOffset=""0.5""/>" & _
                                          "<categoryPlot backColor=""255,255,255"" outlineVisible=""false"" rangeGridlinesVisible=""false"" margins=""0,0,0,0"" axisOffset=""0,0,0,0""/>" & _
                                          "<categoryAxis tickMarksVisible=""false"" catlFontSize=""5.5"" catlAlignment=""right"" catlMargins=""0,0,0,0""/>" & _
                                          "<valueAxis label=""mc/giorno"" catlMargins=""0,0,0,0"" axisRange=""0," & WS_Y & """/>" & _
                                          WS_DATA & _
                                     "</chart>"

End Function

Private Function GET_P02_DETTAGLIOLETTURE_G13() As String
    
    Dim I          As Integer
    Dim WS_ADDCLMN As Boolean

    If (WS_CHK_G13) Then
        WS_ERRSCT = "GET_P02_DETTAGLIOLETTURE_G13"
        WS_ADDCLMN = ((WS_FLG_MASTER = False) And (WS_CNTTR_MASTER_SRVZ <> ""))
    
        With myXFDFMLTable
            .setTableAlignH = "right"
            '.setTableBorders = "0.3"
            
            If (WS_ADDCLMN) Then
                .setTableColumns = "8"
                .setTableFontName = "helr45w.ttf"
                .setTableFontSize = "7"
                .setTablePaddingTop = "0"
                .setTableWidths = "1,3,3,2,3,2,2.5,2.5"
            Else
                .setTableColumns = "7"
                .setTableFontName = "helr45w.ttf"
                .setTableFontSize = "7"
                .setTablePaddingTop = "0"
                .setTableWidths = "1,3,3,3,3,2,4"
            End If
            
            .addCell = ""
    
            .setCellAlignH = "left"
            .setCellColSpan = IIf(WS_ADDCLMN, "7", "6")
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Dettaglio letture" & IIf((WS_G01.R001.TIPOBOLLETTAZIONE = "D"), " - Parte a Conguaglio", "")
    
            .setCellImage = "icnReadings.jpg"
            .setCellImageScale = "23.96"
            .setCellAlignV = "top"
            .setCellRowSpan = "999"
            .addCell = ""
                
            .setCellAlignH = "left"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Contatore"
            
            .setCellAlignH = "center"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Data lettura"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Giorni|Lettura|Consumo"
            
            .setCellAlignH = "left"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Tipo lettura"
            
            If (WS_ADDCLMN) Then
                .setCellBackColor = "230,230,230"
                .setCellBorderBottom = "0.55"
                .setCellBorderBottomColor = "0,55,110"
                .setCellFontName = "helr65w.ttf"
                .setCellHeight = "10"
                .addCell = "Ecc. Divisionale"
            End If
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 8.47)
            
            For I = 0 To UBound(WS_G13)
                .setCellAlignH = "left"
                .addCell = Trim$(WS_G13(I).MATRICOLACONTATORE)
                
                .setCellAlignH = "center"
                .addCell = WS_G13(I).DATALETTURA
                
                .addCell = IIf(Trim$(WS_G13(I).GIORNI) = "", " ", Trim$(WS_G13(I).GIORNI)) & "|" & _
                           NRM_REMOVEZEROES(WS_G13(I).LETTURA, True) & " mc|" & _
                           IIf(Trim$(WS_G13(I).CONSUMOFATTURATO) = "", " ", NRM_REMOVEZEROES(WS_G13(I).CONSUMOFATTURATO, True) & " mc")
    
                .setCellAlignH = "left"
                .addCell = IIf((WS_G13(I).TIPOLETTURA = "F"), "Finale", Trim$(WS_G13(I).TIPOLETTURAAEEG))
                
                If (WS_ADDCLMN) Then
                    If (I = UBound(WS_G13)) Then
                        .addCell = NRM_REMOVEZEROES(WS_G11.R001.CONSUMO_SOMMARE_DETRARRE, True) & " mc"
                    Else
                        .addCell = " "
                    End If
                End If
            
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            Next I
            
            .setCellColSpan = IIf(WS_ADDCLMN, "7", "6")
            .setCellBorderTop = "0.55"
            .setCellBorderTopColor = "0,55,110"
            .setCellHeight = "10"
            .addCell = ""
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.23)
            
            GET_P02_DETTAGLIOLETTURE_G13 = .getXFDFTableNode
        End With
    End If
    
End Function

Private Function GET_P02_DETTAGLIOLETTURE_G14() As String
    
    Dim I As Integer

    If (WS_CHK_G14) Then
        WS_ERRSCT = "GET_P02_DETTAGLIOLETTURE_G14"

        With myXFDFMLTable
            .setTableAlignH = "right"
            '.setTableBorders = "0.3"
            .setTableColumns = "7"
            .setTableFontName = "helr45w.ttf"
            .setTableFontSize = "7"
            .setTablePaddingTop = "0"
            .setTableWidths = GET_TABLECOLUMNS_SIZE(19, "1,3,3,3,3,2,4")
    
            .addCell = ""
    
            .setCellAlignH = "left"
            .setCellColSpan = "6"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Dettaglio letture" & IIf((WS_G01.R001.TIPOBOLLETTAZIONE = "D"), " - Parte in Acconto", "")
    
            .setCellRowSpan = "999"
            .addCell = ""
                
            .setCellAlignH = "left"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Contatore"
            
            .setCellAlignH = "center"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Data lettura"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Giorni|Lettura|Consumo"
            
            .setCellAlignH = "left"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Tipo lettura"
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 8.47)
            
            For I = 0 To UBound(WS_G14)
                .setCellAlignH = "left"
                .addCell = Trim$(WS_G14(I).MATRICOLACONTATORE)
                
                .setCellAlignH = "center"
                .addCell = WS_G14(I).DATALETTURA
                
                .addCell = IIf(Trim$(WS_G14(I).GIORNI) = "", " ", Trim$(WS_G14(I).GIORNI)) & "|" & _
                           NRM_REMOVEZEROES(WS_G14(I).LETTURA, True) & " mc|" & _
                           IIf(Trim$(WS_G14(I).CONSUMOFATTURATO) = "", " ", NRM_REMOVEZEROES(WS_G14(I).CONSUMOFATTURATO, True) & " mc")
    
                .setCellAlignH = "left"
                .addCell = IIf((WS_G14(I).TIPOLETTURA = "F"), "Finale", Trim$(WS_G14(I).TIPOLETTURAAEEG))
            
                WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            Next I
            
            .setCellColSpan = "6"
            .setCellBorderTop = "0.55"
            .setCellBorderTopColor = "0,55,110"
            .setCellHeight = "10"
            .addCell = ""
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.23)
            
            GET_P02_DETTAGLIOLETTURE_G14 = .getXFDFTableNode
        End With
    End If
    
End Function

Private Function GET_P02_DETTAGLIOLETTURE_MASTER() As String

    WS_ERRSCT = "GET_P02_DETTAGLIOLETTURE_MASTER"

    If (WS_FORNITURA_MASTER_P02.dataDescription = "TRG_PARAM") Then
        With myXFDFMLTable
            .setTableAlignH = "center"
            '.setTableBorders = "0.3"
            .setTableColumns = "12"
            .setTableFontName = "helr45w.ttf"
            .setTableFontSize = "7"
            .setTablePaddingTop = "0"
            .setTableWidths = "1,1.64,1.9,1.6,1.6,1.6,1.6,1.6,1.6,1.6,1.6,1.6"
    
            ' HEADER
            '
            .addCell = ""
    
            .setCellBackColor = "230,230,230"
            .setCellColSpan = "2"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Dati Master"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderLeft = "0.3"
            .setCellColSpan = "3"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Lettura Precedente|Lettura Attuale| "
            
            .addCell = ""
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Codice<br>Servizio|Matricola"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellBorderLeft = "0.3"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Data"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Lettura|Tipo"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellBorderLeft = "0.3"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Data"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Lettura|Tipo"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellBorderLeft = "0.3"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Consumo"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .addCell = "Consumi<br>Divisionali|Eccedenza<br>Master"
            
            ' ROW DATA
            '
            .addCell = ""
            
            .addCell = WS_CNTTR_MASTER_SRVZ & "|" & WS_CNTTR_MASTER_MTRCL
            
            .setCellBorderLeft = "0.3"
            .addCell = IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(1) = ""), "-", Format$(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(1), "dd/MM/yyyy"))
            
            .addCell = IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(2) = ""), "-", Format$(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(2), "#,##")) & "|" & IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(3) = ""), "-", WS_FORNITURA_MASTER_P02.EXTRAPARAMS(3))
            
            .setCellBorderLeft = "0.3"
            .addCell = IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(4) = ""), "-", Format$(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(4), "dd/MM/yyyy"))
            
            .addCell = IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(5) = ""), "-", Format$(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(5), "#,##")) & "|" & IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(6) = ""), "-", WS_FORNITURA_MASTER_P02.EXTRAPARAMS(6))
                
            If (WS_FLG_MASTER) Then
                .setCellBorderLeft = "0.3"
                .addCell = NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMORILEVATO, True)
                
                If (WS_FLG_MASTER_ECC_POS) Then
                    .addCell = (NRM_REMOVEZEROES(WS_G11.R001.CONSUMO_SOMMARE_DETRARRE, True) * -1) & "|" & NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True)
                Else
                    .addCell = IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(10) = ""), "-", Format$((Val(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(10)) * -1), "#,##")) & "|" & Format$((Val(WS_G11.R001.TOTALECONSUMORILEVATO) + Val(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(10))), "#,##")
                End If
            Else
                .setCellBorderLeft = "0.3"
                .addCell = IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(7) = ""), "-", Format$(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(7), "#,##"))
                
                .addCell = IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(9) = ""), "-", Format$(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(9), "#,##")) & "|" & IIf((WS_FORNITURA_MASTER_P02.EXTRAPARAMS(8) = ""), "-", Format$(WS_FORNITURA_MASTER_P02.EXTRAPARAMS(8), "#,##"))
            End If

            ' FOOTER
            '
            .addCell = ""

            .setCellColSpan = "11"
            .setCellBorderTop = "0.55"
            .setCellBorderTopColor = "0,55,110"
            .setCellHeight = "12"
            .addCell = ""
    
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 17.3)
            
            GET_P02_DETTAGLIOLETTURE_MASTER = .getXFDFTableNode
        End With
    End If
    
End Function

Private Function GET_P02_DETTAGLIOLETTURE_SCARTATE() As String
    
    If (WS_CHK_G07) Then
        WS_ERRSCT = "GET_P02_DETTAGLIOLETTURE_SCARTATE"
    
        With myXFDFMLTable
            .setTableAlignH = "right"
            '.setTableBorders = "0.3"
            
            .setTableColumns = "6"
            .setTableFontName = "helr45w.ttf"
            .setTableFontSize = "7"
            .setTablePaddingTop = "0"
            .setTableWidths = "1,3,3,3,2,7"
            
            .addCell = ""
    
            .setCellAlignH = "left"
            .setCellColSpan = 5
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "12"
            .addCell = "Dettaglio Autoletture Scartate"
    
            .setCellImage = "icnReadings.jpg"
            .setCellImageScale = "23.96"
            .setCellAlignV = "top"
            .setCellRowSpan = "999"
            .addCell = ""
                
            .setCellAlignH = "left"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Esecutore"
            
            .setCellAlignH = "center"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Data lettura"
            
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Lettura|Consumo"
            
            .setCellAlignH = "left"
            .setCellBackColor = "230,230,230"
            .setCellBorderBottom = "0.55"
            .setCellBorderBottomColor = "0,55,110"
            .setCellFontName = "helr65w.ttf"
            .setCellHeight = "10"
            .addCell = "Motivazione"
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 8.47)
            
            .setCellAlignH = "left"
            .addCell = Trim$(WS_G07.R003.DESCRIZIONE_ESECUTORE)
            
            .setCellAlignH = "center"
            .addCell = WS_G07.R003.DATA_LETTURA
                
            .addCell = NRM_REMOVEZEROES(WS_G07.R003.LETTURA, True) & " mc|" & NRM_REMOVEZEROES(WS_G07.R003.CONSUMO, True) & " mc"
    
            .setCellAlignH = "left"
            .addCell = IIf((Trim$(WS_G07.R003.DESCRIZIONE_NON_VALIDAZIONE) = ""), "Autolettura non valida", Trim$(WS_G07.R003.DESCRIZIONE_NON_VALIDAZIONE))
                
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 3.17)
            
            .setCellColSpan = 5
            .setCellBorderTop = "0.55"
            .setCellBorderTopColor = "0,55,110"
            .setCellHeight = "10"
            .addCell = ""
            
            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.23)
            
            GET_P02_DETTAGLIOLETTURE_SCARTATE = .getXFDFTableNode
        End With
    End If
    
End Function

Private Function GET_PXX_ATTACH_DETAILS() As String

    WS_ERRSCT = "GET_PXX_ATTACH_DETAILS"

    Dim I             As Integer
    Dim WS_DATA()     As String
    Dim WS_KEY        As String
    Dim WS_WMS_DATA   As strct_DATA

    WS_KEY = WS_G00.AZIENDASIU & "_" & Format$(WS_G01.R001.SEZIONALE, "00") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & Format$(WS_G01.R001.CODSERVIZIONUMERAZIONE, "00") & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000") & "_" & WS_G01.R001.RATABOLLETTA
    WS_WMS_DATA = GET_DATA_CACHE(WMS, WS_KEY)
    
    If (WS_WMS_DATA.dataDescription <> WS_KEY) Then
        WS_DATA = Split(WS_WMS_DATA.dataDescription, vbNewLine)
    
        If (CHK_ARRAY(WS_DATA)) Then
            WS_PAGEHEIGHT = 130
            
            DDS_ADD "[$VAR_DOCNUM_02]", WS_G01.R001.ANNOBOLLETTA & "/" & Trim$(WS_G01.R001.NUMEROBOLLETTA)
        
            ' SECTION 01
            '
            ReDim WS_DOCPAGES_DA(0)
            WS_DOCPAGES_DA(0) = "<text fontname='helr45w.ttf' fontsize='10' alignment='justified' leading='2'>" & _
                                    "<chunk><![CDATA[Gentile cliente,<br><br>       trasmettiamo in allegato il documento di rettifica della quota di depurazione e relativa quota fissa precedentemente addebitata e non dovuta ai sensi della Sentenza della Corte Costituzionale n. 335 del 10/10/2008.<br>       Precisiamo che sono state interamente rettificate le quote di depurazione fatturate a saldo e le relative quote fisse, mentre le quote di depurazione emesse in acconto saranno oggetto di rettifica con la prima fattura di conguaglio consumi.<br>       A breve saranno inoltre emesse le bollette dei consumi relative ai periodi successivi a quelli rettificati. Queste ultime bollette saranno accompagnate da una nota esplicativa che indicherà tutte le forme di pagamento che potranno essere utilizzate, inclusa la rateizzazione sino a 60 rate.<br>       Per maggiore chiarezza, evidenziamo di seguito le fatture oggetto del documento di rettifica appena emesso:<br><br>]]></chunk>" & _
                                "</text>"
            ' SECTION 02
            '
            ADD_PXX_ATTACH_TABLE_HDR
            
            For I = 0 To UBound(WS_DATA)
                ADD_PXX_ATTACH_TABLE_ROW WS_DATA(I)
            Next I
            
            ADD_PXX_ATTACH_TABLE_CLOSE
            
            ReDim Preserve WS_DOCPAGES_DA(WS_DOCPAGES_DA_CNTR)
            WS_DOCPAGES_DA(WS_DOCPAGES_DA_CNTR) = WS_DOCPAGES_DA(WS_DOCPAGES_DA_CNTR) & myXFDFMLTable.getXFDFTableNode
            
            ' PAGES LOADER
            '
            For I = 0 To WS_DOCPAGES_DA_CNTR
                GET_PXX_ATTACH_DETAILS = GET_PXX_ATTACH_DETAILS & WS_DOCPAGES_DA(I) & IIf((I = WS_DOCPAGES_DA_CNTR), "", "[EP]")
            Next I
        End If
    End If
    
End Function

Private Sub GET_PXX_AUI()
    
    WS_ERRSCT = "GET_PXX_AUI"

    Dim I         As Integer
    Dim WS_STRING As String

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + Val(WS_CS_PXX_LBL_AUI.EXTRAPARAMS(0)))
    GoSub CHECK_ROWHEIGHT
    
    For I = 0 To UBound(WS_GSE)
        WS_PAGEHEIGHT = (WS_PAGEHEIGHT + Val(WS_CS_PXX_LBL_AUI_ROW.EXTRAPARAMS(0)))
        GoSub CHECK_ROWHEIGHT
        
        With WS_GSE(I)
            WS_STRING = WS_STRING & Replace$(WS_CS_PXX_LBL_AUI_ROW.dataDescription, "[$LBL_AUI_DSCR]", Trim$(.DESCRIZIONE))
            WS_STRING = Replace$(WS_STRING, "[$LBL_AUI_VAL]", CLng(.CONSUMO_MEDIO_LITRI_UI) & "/" & Trim$(.GIORNI))
        End With
    Next I
    
    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + Val(WS_CS_PXX_LBL_AUI_FTR.EXTRAPARAMS(0)))
    GoSub CHECK_ROWHEIGHT
    
    WS_STRING = WS_STRING & Replace$(WS_CS_PXX_LBL_AUI_FTR.dataDescription, "[$LBL_AUI_PRD]", WS_GSF.DATA_EMISSIONE_AP & " - " & WS_GSF.DATA_EMISSIONE_AC)
    WS_STRING = Replace$(WS_STRING, "[$LBL_AUI_IMPRT]", NRM_IMPORT(Trim$(WS_GSF.TOTALE), "#,##0.00", False))
    
    WS_STRING = Replace$(WS_CS_PXX_LBL_AUI.dataDescription, "[$PXX_LBL_AUI_ROWS]", WS_STRING)
    
    ReDim Preserve WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR)
    WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) = WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) & WS_STRING

    Exit Sub

CHECK_ROWHEIGHT:
    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
        WS_DOCPAGES_DF_CNTR = (WS_DOCPAGES_DF_CNTR + 1)
        WS_PAGEHEIGHT = 31
    End If
Return
        
End Sub

Private Sub GET_PXX_BOLLO_QUIETANZA()

    WS_ERRSCT = "GET_PXX_BOLLO_QUIETANZA"

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.94)
    GoSub CHECK_ROWHEIGHT

    With myXFDFMLTable
        '.setTableBorders = "0.3"
        .setTableColumns = "1"
        .setTableFontName = "helr45w.ttf"
        .setTableFontName = "helr66w.ttf"
        .setTableFontSize = "7"
        .setTableShiftY = "3"
        .setTableWidths = "1"
        
        .addCell = "              " & IIf(WS_G01.R001.TIPONUMERAZIONE = "0", "Imposta di bollo assolta in modo virtuale (Ag. Entrate di Nuoro n. 1/2015)", "Imposta di bollo assolta ai sensi dell’art. 6, comma 2, del D.M. 17.6.2014")
    End With

    ReDim Preserve WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR)
    WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) = WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) & myXFDFMLTable.getXFDFTableNode

    Exit Sub

CHECK_ROWHEIGHT:
    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
        WS_DOCPAGES_DF_CNTR = (WS_DOCPAGES_DF_CNTR + 1)
        WS_PAGEHEIGHT = 31
    End If
Return

End Sub

Private Sub GET_PXX_BONUS_SOCIALE()
    
    WS_ERRSCT = "GET_PXX_BONUS_SOCIALE"

    Dim WS_STRING As String

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + Val(WS_CS_PXX_LBL_BONUS_SOCIALE.EXTRAPARAMS(0)))
    GoSub CHECK_ROWHEIGHT
    
    WS_STRING = Replace$(WS_CS_PXX_LBL_BONUS_SOCIALE.dataDescription, "[$DTA_BS_STRT]", WS_GBS.DATA_INIZIO_PERIODO)
    WS_STRING = Replace$(WS_STRING, "[$DTA_BS_END]", WS_GBS.DATA_FINE_PERIODO)
    
    ReDim Preserve WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR)
    WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) = WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) & WS_STRING

    Exit Sub

CHECK_ROWHEIGHT:
    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
        WS_DOCPAGES_DF_CNTR = (WS_DOCPAGES_DF_CNTR + 1)
        WS_PAGEHEIGHT = 31
    End If
Return
        
End Sub

Private Sub GET_PXX_INFO_PAGE()

    Dim WS_KEY    As String
    Dim WS_STRING As String

    ' AGGIORNAMENTO DATI FISCALI
    '
    WS_STRING = Trim$(IIf((Trim$(WS_G02.R005.PARTITAIVA) = ""), WS_G02.R005.CODICEFISCALE, WS_G02.R005.PARTITAIVA))
    
    If (WS_STRING = "") Then
        DDS_ADD "[$LBL_ADF_01]", WS_CS_PXX_ADF_CFPI_MSG_KO.dataDescription
    Else
        DDS_ADD "[$LBL_ADF_01]", Replace$(WS_CS_PXX_ADF_CFPI_MSG_OK.dataDescription, "[$VAR_ADF_01]", WS_STRING)
    End If

    WS_STRING = ""

    If (WS_FLG_FATTELE_PA) Then
        WS_STRING = GET_DATA_CACHE(CODICE_IPA, WS_CODICE_SERVIZIO_KEY).dataDescription

        If (WS_STRING = WS_CODICE_SERVIZIO_KEY) Then WS_STRING = "-"
    End If

    DDS_ADD "[$VAR_ADF_02]", IIf((Trim$(WS_G01.R007.CODICE_DESTINATARIO) = ""), "-", Trim$(WS_G01.R007.CODICE_DESTINATARIO))
    DDS_ADD "[$VAR_ADF_03]", IIf((Trim$(WS_G01.R007.INDIRIZZO_PEC) = ""), "-", Trim$(WS_G01.R007.INDIRIZZO_PEC))
    DDS_ADD "[$VAR_ADF_04]", WS_STRING
    DDS_ADD "[$VAR_ADF_05]", IIf((Trim$(WS_G01.R007.CODICE_CUP) = ""), "-", Trim$(WS_G01.R007.CODICE_CUP))
    DDS_ADD "[$VAR_ADF_06]", IIf((Trim$(WS_G01.R007.CODICE_CIG) = ""), "-", Trim$(WS_G01.R007.CODICE_CIG))
    
    If (WS_FLG_FATTELE_PA) Then
        DDS_ADD "[$LBL_ADF_RCA]", ""
    Else
        DDS_ADD "[$LBL_ADF_RCA]", Replace$(WS_CS_PXX_LBL_ADF_RCA.dataDescription, "[$VAR_ADF_07]", IIf((Trim$(WS_G01.R007.RINUNCIA_COPIA_ANALOGICA) = "S"), "Si", "No"))
    End If

    ' INFO SUI CONSUMI, ACCESSIBILITÀ CONTATORE, TENTATIVI DI LETTURA, AUTOLETTURA
    '
    DDS_ADD "[$VAC_ISC_01]", Right$(WS_G01.R001.DATAEMISSIONE, 4)
    DDS_ADD "[$VAC_ISC_02]", IIf((Trim$(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AC) = ""), "-", NRM_REMOVEZEROES(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AC, True))
    DDS_ADD "[$VAC_ISC_03]", (Val(Right$(WS_G01.R001.DATAEMISSIONE, 4)) + 1)
    DDS_ADD "[$VAC_ISC_04]", IIf((Trim$(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AS) = "" Or WS_FLG_NOTACREDITO), "-", NRM_REMOVEZEROES(WS_G09.R012.CONSUMO_MEDIO_ANNUO_AS, True))
    DDS_ADD "[$VAR_AC]", GET_ACCESSIBILITÀ_CONTATORE
    DDS_ADD "[$VAR_TDL]", IIf((Trim$(WS_G09.R001.NUMERO_LETTURE_ANNUE_218_16) = ""), "-", Val(WS_G09.R001.NUMERO_LETTURE_ANNUE_218_16))

    WS_KEY = Format$(WS_G02.R001.CODICESERVIZIO, "0000000000") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000")
    WS_STRING = GET_DATA_CACHE(DYN_MSG_PDF, WS_KEY).dataDescription
    
    If (WS_STRING = WS_KEY) Then
        DDS_ADD "[$VAR_PDF]", "N/A"
    Else
        DDS_ADD "[$VAR_PDF]", LCase$(WS_STRING)
    End If

    If (WS_FLG_DIV = False) Then
        If (WS_BDS) Then
            DDS_ADD "[$P01_MSG_CMA]", ""
        Else
            DDS_ADD "[$P01_MSG_CMA]", WS_CS_P01_MSG_CMA.dataDescription
        End If
    End If
    
    If (WS_MSG_CA) Then
        DDS_ADD "[$TBL_MSG_AC]", WS_CS_PXX_TBL_MSG_AC.dataDescription
    Else
        DDS_ADD "[$TBL_MSG_AC]", ""
    End If
    
    If (WS_FLG_DOM) Then
        DDS_ADD "[$TBL_MSG_PAGOPA]", ""
    Else
        DDS_ADD "[$TBL_MSG_PAGOPA]", WS_CS_PXX_TBL_MSG_PAGOPA.dataDescription
    End If
    
    ' RATEIZZAZIONE DI QUESTA FATTURA (ART. 3 REMSI)
    '
    If ((Trim$(WS_G09.R005.FLAGRATEIZZABILITÀ) <> "N") Or (Trim$(WS_G09.R005.RATEIZZATO_NORMATIVA) <> "N")) Then
        WS_STRING = ""
    
        If (Trim$(WS_G09.R005.FLAGRATEIZZABILITÀ) = "S") Then
            WS_STRING = WS_CS_PXX_TBL_DELAY_INFO_M01.dataDescription
        ElseIf (Trim$(WS_G09.R005.RATEIZZATO_NORMATIVA) = "S") Then
            WS_STRING = WS_CS_PXX_TBL_DELAY_INFO_M02.dataDescription
        End If
        
        If (WS_STRING = "") Then
            DDS_ADD "[$TBL_DELAY_INFO]", ""
        Else
            DDS_ADD "[$TBL_DELAY_INFO]", Replace$(WS_STRING, "[$VAR_PID]", Trim$(WS_G09.R005.PERCENTUALE_INTERESSI_DILATORI))
        End If
    End If

End Sub

Private Function GET_PXX_INVOICE_DETAILS() As String

    WS_ERRSCT = "GET_PXX_INVOICE_DETAILS"

    Dim I         As Integer
    Dim J         As Integer
    
    ADD_PXX_DETTAGLIOBOLLETTA_TABLEHEADER False
    
    For I = 0 To UBound(WS_GDF)
        If (CHK_GDF_NVXXX(WS_GDF_QF(I).RXXX)) Then GET_QS_GDF_QF WS_GDF_QF(I).RXXX
        
        For J = 0 To UBound(WS_GDF(I).RXXX)
            ADD_PXX_DETTAGLIOBOLLETTA_ROW I, Left$(WS_GDF(I).RXXX(J).ROW, 2), Mid$(WS_GDF(I).RXXX(J).ROW, 3)
        Next J
    Next I
    
    ReDim Preserve WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR)
    WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) = myXFDFMLTable.getXFDFTableNode
    
    If (WS_DF_GBO_FLG) Then GET_PXX_BOLLO_QUIETANZA
    'If ((WS_MSG_CI = False) And (WS_FLG_NOTACREDITO = False) And (WS_IMPORTO_TC_TICSI <> "")) Then GET_PXX_TICSI
    If (WS_MSG_CI) Then GET_PXX_MSG_CI
    If (Trim$(WS_GBS.DATA_INIZIO_PERIODO & WS_GBS.DATA_FINE_PERIODO) <> "") Then GET_PXX_BONUS_SOCIALE
    If ((WS_FLG_NOTACREDITO = False) And (WS_FLG_PARTITE = False) And WS_CHK_GSE) Then GET_PXX_AUI

    ' PAGES LOADER
    '
    For I = 0 To WS_DOCPAGES_DF_CNTR
        GET_PXX_INVOICE_DETAILS = GET_PXX_INVOICE_DETAILS & WS_DOCPAGES_DF(I) & IIf((I = WS_DOCPAGES_DF_CNTR), "", "[EP]")
    Next I

End Function

Private Sub GET_PXX_MSG_CI()

    WS_ERRSCT = "GET_PXX_MSG_CI"

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 18)
    GoSub CHECK_ROWHEIGHT

    With myXFDFMLTable
        '.setTableBorders = "0.3"
        .setTableColumns = "2"
        .setTableFontName = "helr45w.ttf"
        .setTableFontSize = "8"
        .setTableShiftY = "10"
        .setTableWidths = "1,18"
        
        .setCellPaddingTop = "0"
        .addCell = ""

        .setCellBorderBottom = "0.55"
        .setCellBorderBottomColor = "0,55,110"
        .setCellColSpan = "2"
        .setCellFontName = "helr65w.ttf"
        .setCellHeight = "12"
        .setCellPaddingTop = "0"
        .addCell = "Adeguamenti tariffari TICSI"

        .setCellImage = "icnInfo.jpg"
        .setCellImageScale = "23.96"
        .setCellAlignH = "center"
        .setCellAlignV = "top"
        .setCellPaddingTop = "0"
        .setCellRowSpan = "999"
        .addCell = ""
        
        .setCellAlignH = "justified"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "4"
        .addCell = "Questa fattura contiene importi pari a € " & NRM_IMPORT(WS_IMPORTO_TC_TICSI, "#,##0.00", False) & " derivanti da un riallineamento del nuovo sistema di calcolo adottato in base alle disposizioni sul metodo tariffario previste dall’Autorità di regolazione nazionale “ARERA”. Per maggiori informazioni è possibile consultare l’informativa presente nella sua Cartella Cliente, a cui può accedere dallo Sportello on line."
        
        .setCellAlignV = "top"
        .setCellColSpan = "2"
        .setCellBorderTop = "0.55"
        .setCellBorderTopColor = "0,55,110"
        .setCellHeight = "10"
        .setCellPaddingTop = "0"
        .addCell = ""
    End With

    ReDim Preserve WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR)
    WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) = WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) & myXFDFMLTable.getXFDFTableNode

    Exit Sub

CHECK_ROWHEIGHT:
    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
        WS_DOCPAGES_DF_CNTR = (WS_DOCPAGES_DF_CNTR + 1)
        WS_PAGEHEIGHT = 31
    End If
Return

End Sub

'Private Sub GET_PXX_TICSI()
'
'    WS_ERRSCT = "GET_PXX_TICSI"
'
'    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 18)
'    GoSub CHECK_ROWHEIGHT
'
'    With myXFDFMLTable
        '.setTableBorders = "0.3"
'        .setTableColumns = "2"
'        .setTableFontName = "helr45w.ttf"
'        .setTableFontSize = "8"
'        .setTableShiftY = "10"
'        .setTableWidths = "1,18"
'
'        .setCellPaddingTop = "0"
'        .addCell = ""
'
'        .setCellBorderBottom = "0.55"
'        .setCellBorderBottomColor = "0,55,110"
'        .setCellColSpan = "2"
'        .setCellFontName = "helr65w.ttf"
'        .setCellHeight = "12"
'        .setCellPaddingTop = "0"
'        .addCell = "Adeguamenti tariffari TICSI"
'
'        .setCellImage = "icnInfo.jpg"
'        .setCellImageScale = "23.96"
'        .setCellAlignH = "center"
'        .setCellAlignV = "top"
'        .setCellPaddingTop = "0"
'        .setCellRowSpan = "999"
'        .addCell = ""
'
'        .setCellAlignH = "justified"
'        .setCellFontSize = "7"
'        .setCellPaddingBottom = "4"
'        .addCell = "La fattura contiene importi per adeguamenti tariffari indicati nel Quadro di Sintesi (prima pagina della presente fattura). L’adeguamento tariffario effettuato è il TICSI (Testo Integrato Corrispettivi Servizi Idrici) ai sensi della Delibera ARERA n. 665/2017/R/idr e Delibera EGAS n. 345/2019 del 16.10.2019."
'
'        .setCellAlignV = "top"
'        .setCellColSpan = "2"
'        .setCellBorderTop = "0.55"
'        .setCellBorderTopColor = "0,55,110"
'        .setCellHeight = "10"
'        .setCellPaddingTop = "0"
'        .addCell = ""
'    End With
'
'    ReDim Preserve WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR)
'    WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) = WS_DOCPAGES_DF(WS_DOCPAGES_DF_CNTR) & myXFDFMLTable.getXFDFTableNode
'
'    Exit Sub
'
'CHECK_ROWHEIGHT:
'    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
'        WS_DOCPAGES_DF_CNTR = (WS_DOCPAGES_DF_CNTR + 1)
'        WS_PAGEHEIGHT = 31
'    End If
'Return
'
'End Sub

Private Function GET_TABLECOLUMNS_SIZE(WS_TABLEWIDTH As Integer, WS_TABLEWIDTHS As String) As String

    Dim I           As Integer
    Dim WS_COLUMNS  As Integer
    Dim WS_WIDTHS() As String
    
    WS_WIDTHS = Split(WS_TABLEWIDTHS, ",")
    WS_COLUMNS = (UBound(WS_WIDTHS) + 1)
    
    For I = 0 To UBound(WS_WIDTHS)
        GET_TABLECOLUMNS_SIZE = GET_TABLECOLUMNS_SIZE & "," & Replace$(Round(((Val(WS_WIDTHS(I)) * WS_COLUMNS) / WS_TABLEWIDTH), 2), ",", ".")
    Next I
    
    GET_TABLECOLUMNS_SIZE = Mid$(GET_TABLECOLUMNS_SIZE, 2)

End Function

Public Function GET_TEMPLATEINFO() As String
    
    Dim I                    As Integer
    Dim J                    As Integer
    Dim K                    As Integer
    Dim myXFDFMLTemplate     As cls_XFDFMLTemplate
    Dim WS_BILLS_POS         As Integer
    Dim WS_BILLS_MAX         As Integer
    Dim WS_DF_HEIGHT         As Single
    Dim WS_DF_YSHIFT         As Integer
    Dim WS_DYN_ANNXD         As strct_DATA
    Dim WS_DYN_KEY           As String
    Dim WS_EXTRAFIELDS       As String
    Dim WS_FIELDPAGECNTR     As Integer
    Dim WS_FIELDID           As String
    Dim WS_FIELDIDCNTR       As Integer
    Dim WS_FIELDPAGEID       As String
    Dim WS_FIELDS            As String
    Dim WS_FILENAME          As String
    Dim WS_INT               As Integer
    Dim WS_PAGEBILL          As Integer
    Dim WS_PAGEBILLS         As String
    Dim WS_PAGENUMBER        As String
    Dim WS_PAGESINDXS        As String
    Dim WS_PAGESNUM          As Integer
    Dim WS_SPLIT()           As String
    Dim WS_SPLIT_VLS()       As String
    Dim WS_STRING            As String
    Dim WS_TEMP_UNIQUEID     As Currency
    Dim WS_TMPL_ID           As Integer

    Dim WS_PPA_CNTR As Integer

    WS_ERRSCT = "GET_TEMPLATEINFO"
    
    WS_ANNEXED_DATA = ""
    WS_FILENAME = "ABN_"
    WS_PAGESNUM = 0
    WS_TEMP_UNIQUEID = 0
    
    Set myXFDFMLTemplate = New cls_XFDFMLTemplate
    
    For I = 0 To UBound(WS_SECTIONS)
        With WS_SECTIONS(I)
            Select Case .SECTIONDESCRIPTION
            Case "ATTACH"
                If (WS_FLG_WMS) Then
                    For J = 0 To WS_DOCPAGES_DA_CNTR
                        WS_DF_HEIGHT = 255
                        WS_DF_YSHIFT = 30
                        
                        If (J = 0) Then
                            WS_DF_HEIGHT = 210
                            WS_DF_YSHIFT = 75
                            
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.TEMPLATES(0).FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .TEMPLATES(0).FIELDSDATA
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMPLATES(0).TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMPLATES(0).TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(0).TEMP_BITWISE)
                        Else
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.TEMPLATES(1).FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .TEMPLATES(1).FIELDSDATA
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMPLATES(1).TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMPLATES(1).TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(1).TEMP_BITWISE)
                        End If
                        
                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        
                        myXFDFMLTemplate.setFieldId = "TXT_ATTACH_P" & Format$((J + 1), "000")
                        myXFDFMLTemplate.setPropertyPageId = WS_PAGESNUM
                        myXFDFMLTemplate.setPropertyCoords = "10," & WS_DF_YSHIFT & ",190," & WS_DF_HEIGHT
                        myXFDFMLTemplate.closeProperty
                        myXFDFMLTemplate.closeField
                    Next J
                
                    If (WS_PAGESNUM And 1) Then
                        With .TEMPLATES(2)
                            WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    End If
                End If
                
            Case "BILL"
                If (WS_CCP And (WS_BDS = False)) Then
                    WS_PAGEBILL = (WS_PAGESNUM + 1)
                    WS_BILLS_MAX = UBound(WS_G23)
                    WS_TMPL_ID = 0
            
                    For J = 0 To WS_BILLS_MAX
                        If (WS_G23(J).CL_BOLLETTINO_ID <> "999999999999999999") Then
                            If (J > 0) Then
                                Select Case (WS_BILLS_MAX - J)
                                Case Is >= 1
                                    WS_TMPL_ID = 1
                                    WS_BILLS_POS = 1
    
                                Case 0
                                    WS_TMPL_ID = 2
                                    WS_BILLS_POS = 0
    
                                End Select
                            End If
                            
                            With .TEMPLATES(WS_TMPL_ID)
                                If (Trim$(.FIELDS) <> "") Then
                                    .FIELDS = Replace$(.FIELDS, "[$NI_X]", 110)
                                    .FIELDS = Replace$(.FIELDS, "[$NI_Y]", 6)
                                    .FIELDS = Replace$(.FIELDS, "[$NI_W]", 90)
                                
                                    DDS_ADD "[$NI_WC01]", "4"
                                    DDS_ADD "[$NI_WC02]", "86"
                                    
                                    WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                                    WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                                End If

                                WS_PAGENUMBER = WS_PAGENUMBER & ";0;0"
                                WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                            End With

                            WS_FIELDPAGEID = (WS_PAGEBILL + (WS_FIELDPAGECNTR * 2))
                            WS_PAGEBILLS = WS_PAGEBILLS & ";" & WS_FIELDPAGEID
                            WS_FIELDPAGECNTR = (WS_FIELDPAGECNTR + 1)
                            
                            For K = 0 To WS_BILLS_POS
                                WS_FIELDIDCNTR = (WS_FIELDIDCNTR + 1)
                                WS_FIELDID = Format$(WS_FIELDIDCNTR, "000")
                                
                                If (K = 0) Then
                                    WS_EXTRAFIELDS = WS_EXTRAFIELDS & "<field id='NMR_BOLLETTINOIMPORTO_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='87.7,94.2,4.5,35.5' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='87.7,248,4.5,46' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='NMR_BOLLETTINOFATTURA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='36.4,31.2,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='270'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='42.4,216.4,2.8,55' fontname='helr65w.ttf' fontsize='8' rotation='270'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='DTA_BOLLETTINOSCADENZA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='10.6,22.1,4.6,26.6' fontname='helr65w.ttf' fontsize='12' alignment='center' rotation='270'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC00'><properties pageid='" & WS_FIELDPAGEID & "' coords='26.3,190.3,12,93' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BCT00'><properties pageid='" & WS_FIELDPAGEID & "' coords='23.9,190.3,2.5,93' alignment='center' fontname='helr65w.ttf' fontsize='6' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC01'><properties pageid='" & WS_FIELDPAGEID & "' coords='2.35,81.06,14.97,45' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOIDCLIENTE_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='56,139.5,4.8,52.2' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOID_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='6.67,138.6,5.5,152.8' fontname='ocrb10n.ttf' fontsize='11' rotation='270' comb='60'/></field>"
                                Else
                                    WS_EXTRAFIELDS = WS_EXTRAFIELDS & "<field id='NMR_BOLLETTINOIMPORTO_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='117.8,167.3,4.5,35.5' fontname='ocrb10n.ttf' fontsize='11' rotation='90'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='117.8,3.1,4.5,46' fontname='ocrb10n.ttf' fontsize='11' rotation='90'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='NMR_BOLLETTINOFATTURA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='169.7,209.4,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='90'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='164.8,25.6,2.8,55' fontname='helr65w.ttf' fontsize='8' rotation='90'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='DTA_BOLLETTINOSCADENZA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='194.8,248.3,4.6,26.6' fontname='helr65w.ttf' fontsize='12' alignment='center' rotation='90'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC00'><properties pageid='" & WS_FIELDPAGEID & "' coords='171.7,13.7,12.1,93' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BCT00'><properties pageid='" & WS_FIELDPAGEID & "' coords='183.6,13.7,2.5,93' alignment='center' fontname='helr65w.ttf' fontsize='6' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC01'><properties pageid='" & WS_FIELDPAGEID & "' coords='192.7,170.9,15,45' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOIDCLIENTE_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='149.2,105.3,4.8,52.2' fontname='ocrb10n.ttf' fontsize='11' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOID_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='197.8,5.5,5.5,152.8' fontname='ocrb10n.ttf' fontsize='11' rotation='90' comb='60'/></field>"
                                End If
                                
                                J = (J + K)
                            Next K
                        End If
                    Next J
                
                    WS_PAGEBILLS = Mid$(WS_PAGEBILLS, 2)
                End If
            
            Case "BILL_BDS"
                If (WS_CCP And WS_BDS) Then
                    WS_PAGEBILL = (WS_PAGESNUM + 1)
                    WS_BILLS_MAX = UBound(WS_G23)
                    WS_TMPL_ID = 0

                    For J = 0 To WS_BILLS_MAX
                        If (WS_G23(J).CL_BOLLETTINO_ID <> "999999999999999999") Then
                            If (J > 0) Then
                                Select Case (WS_BILLS_MAX - J)
                                Case Is >= 1
                                    WS_TMPL_ID = 1
                                    WS_BILLS_POS = 1

                                Case 0
                                    WS_TMPL_ID = 2
                                    WS_BILLS_POS = 0

                                End Select
                            End If

                            With .TEMPLATES(WS_TMPL_ID)
                                If (Trim$(.FIELDS) <> "") Then
                                    .FIELDS = Replace$(.FIELDS, "[$NI_X]", 110)
                                    .FIELDS = Replace$(.FIELDS, "[$NI_Y]", 6)
                                    .FIELDS = Replace$(.FIELDS, "[$NI_W]", 90)

                                    DDS_ADD "[$NI_WC01]", "4"
                                    DDS_ADD "[$NI_WC02]", "86"

                                    WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                                    WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                                End If

                                WS_PAGENUMBER = WS_PAGENUMBER & ";0;0"
                                WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                            End With

                            WS_FIELDPAGEID = (WS_PAGEBILL + (WS_FIELDPAGECNTR * 2))
                            WS_PAGEBILLS = WS_PAGEBILLS & ";" & WS_FIELDPAGEID
                            WS_FIELDPAGECNTR = (WS_FIELDPAGECNTR + 1)
                            
                            For K = 0 To WS_BILLS_POS
                                WS_FIELDIDCNTR = (WS_FIELDIDCNTR + 1)
                                WS_FIELDID = Format$(WS_FIELDIDCNTR, "000")
                                
                                If (K = 0) Then
                                    WS_EXTRAFIELDS = WS_EXTRAFIELDS & "<field id='NMR_BOLLETTINOIMPORTO_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='87.7,94.2,4.5,35.5' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='87.7,248,4.5,46' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='NMR_BOLLETTINOFATTURA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='36.4,31.2,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='270'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='42.4,216.4,2.8,45' fontname='helr65w.ttf' fontsize='8' rotation='270'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='DTA_BOLLETTINOSCADENZA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='10.6,22.1,4.6,26.6' fontname='helr65w.ttf' fontsize='12' alignment='center' rotation='270'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC00'><properties pageid='" & WS_FIELDPAGEID & "' coords='26.3,190.3,12,93' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BCT00'><properties pageid='" & WS_FIELDPAGEID & "' coords='23.9,190.3,2.5,93' alignment='center' fontname='helr65w.ttf' fontsize='6' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC01'><properties pageid='" & WS_FIELDPAGEID & "' coords='2.35,81.06,14.97,45' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOIDCLIENTE_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='56,139.5,4.8,52.2' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOID_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='6.67,138.6,5.5,152.8' fontname='ocrb10n.ttf' fontsize='11' rotation='270' comb='60'/></field>"
                                Else
                                    WS_EXTRAFIELDS = WS_EXTRAFIELDS & "<field id='NMR_BOLLETTINOIMPORTO_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='117.8,167.3,4.5,35.5' fontname='ocrb10n.ttf' fontsize='11' rotation='90'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='117.8,3.1,4.5,46' fontname='ocrb10n.ttf' fontsize='11' rotation='90'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='NMR_BOLLETTINOFATTURA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='169.7,209.4,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='90'/>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='164.8,35.6,2.8,45' fontname='helr65w.ttf' fontsize='8' rotation='90'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='DTA_BOLLETTINOSCADENZA_B" & WS_FIELDID & "'>" & _
                                                                          "<properties pageid='" & WS_FIELDPAGEID & "' coords='194.8,248.3,4.6,26.6' fontname='helr65w.ttf' fontsize='12' alignment='center' rotation='90'/>" & _
                                                                      "</field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC00'><properties pageid='" & WS_FIELDPAGEID & "' coords='171.7,13.7,12.1,93' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BCT00'><properties pageid='" & WS_FIELDPAGEID & "' coords='183.6,13.7,2.5,93' alignment='center' fontname='helr65w.ttf' fontsize='6' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOBC_B" & WS_FIELDID & "_BC01'><properties pageid='" & WS_FIELDPAGEID & "' coords='192.7,170.9,15,45' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOIDCLIENTE_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='149.2,105.3,4.8,52.2' fontname='ocrb10n.ttf' fontsize='11' rotation='90'/></field>" & _
                                                                      "<field id='TXT_BOLLETTINOID_B" & WS_FIELDID & "'><properties pageid='" & WS_FIELDPAGEID & "' coords='197.8,5.5,5.5,152.8' fontname='ocrb10n.ttf' fontsize='11' rotation='90' comb='60'/></field>"
                                End If
                                
                                J = (J + K)
                            Next K
                        End If
                    Next J
                
                    WS_PAGEBILLS = Mid$(WS_PAGEBILLS, 2)
                End If

            Case "COMMUNICATIONS"
                If (WS_BDS = False) Then
                    For J = 0 To UBound(WS_SECTIONS(I).TEMPLATES)
                        With .TEMPLATES(J)
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                            
                            WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    Next J
                End If
            
            Case "COMMUNICATIONS_BDS"
                If (WS_BDS) Then
                    For J = 0 To UBound(WS_SECTIONS(I).TEMPLATES)
                        With .TEMPLATES(J)
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                            
                            WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    Next J
                End If
            
            Case "DETAIL"
                For J = 0 To WS_DOCPAGES_DF_CNTR
                    WS_DF_HEIGHT = 260
                    WS_DF_YSHIFT = 15
                    
                    If (J = 0) Then
                        If (WS_FLG_CONSUMI) Then
                            WS_DF_HEIGHT = 220
                            WS_DF_YSHIFT = 59
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMPLATES(1).TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMPLATES(1).TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(1).TEMP_BITWISE)
                            
                            myXFDFMLTemplate.setFieldId = "HST_P02_CONSUMI"
                            myXFDFMLTemplate.setPropertyPageId = WS_PAGESNUM
                            myXFDFMLTemplate.setPropertyFontName = "helr45w.ttf"
                            myXFDFMLTemplate.setPropertyFontSize = "7"
                            myXFDFMLTemplate.setPropertyCoords = "110,15,90,40"
                            myXFDFMLTemplate.closeProperty
                            myXFDFMLTemplate.closeField
                        Else
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMPLATES(0).TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMPLATES(0).TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(0).TEMP_BITWISE)
                        End If
                    Else
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMPLATES(0).TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMPLATES(0).TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(0).TEMP_BITWISE)
                    End If
                    
                    WS_PAGENUMBER = WS_PAGENUMBER & ";1"

                    myXFDFMLTemplate.setFieldId = "TXT_DETAILS_P" & Format$((J + 1), "000")
                    myXFDFMLTemplate.setPropertyPageId = WS_PAGESNUM
                    myXFDFMLTemplate.setPropertyCoords = "10," & WS_DF_YSHIFT & ",190," & WS_DF_HEIGHT
                    myXFDFMLTemplate.closeProperty
                    myXFDFMLTemplate.closeField
                Next J
            
                If (WS_PAGESNUM And 1) Then
                    With .TEMPLATES(2)
                        WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                End If
            
            Case "DYN_ANNXD_GNRC"
                WS_DYN_KEY = WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000")
                WS_DYN_ANNXD = GET_DATA_CACHE(DYN_ANNXD_TMPLT, WS_DYN_KEY)

                If (WS_DYN_ANNXD.dataDescription <> WS_DYN_KEY) Then
                    WS_SPLIT = Split(WS_DYN_ANNXD.dataDescription, "|")

                    For J = 0 To (UBound(.TEMPLATES) - 1)
                        With .TEMPLATES(J)
                            If (.ID_DESCR = WS_SPLIT(0)) Then
                                WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                                WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                                
                                WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                                WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                            
                                WS_SPLIT = Split(WS_SPLIT(1), ";")
                            
                                For K = 0 To UBound(WS_SPLIT)
                                    WS_SPLIT_VLS = Split(WS_SPLIT(K), "=")
                                    
                                    DDS_ADD "[$" & UCase$(WS_SPLIT_VLS(0)) + "]", WS_SPLIT_VLS(1)
                                Next K
                            End If
                        End With
                    Next J

                    If (WS_PAGESNUM And 1) Then
                        With .TEMPLATES(UBound(.TEMPLATES))
                            WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    End If
                End If

            Case "DYN_ANNXD_UI"
                If (((WS_FLG_NOTACREDITO = False) And (WS_FLG_PARTITE = False)) And WS_CHK_GSE) Then
                    With .TEMPLATES(0)
                        WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                        WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        
                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                
                    For J = 1 To WS_DOCPAGES_UI_CNTR
                        With .TEMPLATES(1)
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                            WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
    
                        myXFDFMLTemplate.setFieldId = "TXT_AUI_DTLS_P" & Format$(J, "000")
                        myXFDFMLTemplate.setPropertyPageId = WS_PAGESNUM
                        myXFDFMLTemplate.setPropertyCoords = "10,35,190,180"
                        myXFDFMLTemplate.closeProperty
                        myXFDFMLTemplate.closeField
                    Next J
                    
                    If (WS_PAGESNUM And 1) Then
                        With .TEMPLATES(2)
                            WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    End If
                End If
            
            Case "DYN_ANNXD_547_2019"    ' ALLEGATO DINAMICO - ARERA 547/2019
                WS_DYN_ANNXD = GET_DATA_CACHE(DYN_ANNXD_547_2019, WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000"))
                
                If (WS_DYN_ANNXD.dataDescription = "TRG_PARAM") Then
                    ' OLD
                    '
                    WS_FLG_ANNXD_547_19 = True
                    
                    With .TEMPLATES(0)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If
                    
                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                    
                    If (WS_PAGESNUM And 1) Then
                        With .TEMPLATES(5)
                            WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    End If
                ElseIf (WS_CHK_GPB And (WS_FLG_NOTACREDITO = False)) Then
                    ' NEW
                    '
                    If (CDbl(WS_GPB.R001.POTENZIALE_PRESCRIZIONE) > 0) Then
                        If (Trim$(WS_GPB.R002.MOTIVAZIONE) <> "") Then
                            WS_FLG_ANNXD_547_19 = True
                            
                            With .TEMPLATES(4)
                                If (Trim$(.FIELDS) <> "") Then
                                    WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                                    WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                                End If
    
                                WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                                WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                            End With
                            
                            If (WS_PAGESNUM And 1) Then
                                With .TEMPLATES(5)
                                    WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                                    WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                    WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                    WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                                End With
                            End If
                        ElseIf (Trim$(WS_GPB.R001.DATA_PRESCRIZIONE) <> "") Then
                            WS_FLG_ANNXD_547_19 = True
                            
                            If (GET_DATA_CACHE(FLG_DOM, WS_CODICE_SERVIZIO_KEY).dataDescription = "1") Then
                                WS_TMPL_ID = 3
                            Else
                                If (WS_G01.R001.TIPOBOLLETTAZIONE = "D") Then
                                    WS_TMPL_ID = 1
                                Else
                                    WS_TMPL_ID = 2
                                End If
                            End If
    
                            With .TEMPLATES(WS_TMPL_ID)
                                If (Trim$(.FIELDS) <> "") Then
                                    WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                                    WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                                End If
    
                                WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                                WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                            End With
                        
                            If (WS_PAGESNUM And 1) Then
                                With .TEMPLATES(5)
                                    WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                                    WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                    WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                    WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                                End With
                            End If
                        End If
                    End If
                End If
            
            Case "DYN_ANNXD_547_2019_MODCLI032R0"
                If (WS_FLG_ANNXD_547_19) Then
                    If (Trim$(WS_GPB.R002.MOTIVAZIONE) = "") Then
                        WS_TMPL_ID = 0
                    Else
                        WS_TMPL_ID = 1
                    End If
                
                    With .TEMPLATES(WS_TMPL_ID)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If

                        WS_PAGENUMBER = WS_PAGENUMBER & ";0;0"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                End If
                
            Case "DYN_ANNXD_ONE_SHOT"   ' ALLEGATO GENERICO DINAMICO ANNO/NUMERO_FATTURA (ONE SHOT)
                WS_DYN_ANNXD = GET_DATA_CACHE(DYN_ANNXD_ONE_SHOT, WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000"))

                If (WS_DYN_ANNXD.dataDescription = "TRG_PARAM") Then
                    With .TEMPLATES(0)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If

                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With

                    If (WS_PAGESNUM And 1) Then
                        With .TEMPLATES(1)
                            WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    End If
                End If

            Case "DYN_ANNXD_T01"    ' ALLEGATO DINAMICO VOLTURE MASSIVE ANNO/NUMERO_FATTURA (ONE SHOT) - T01
                WS_DYN_ANNXD = GET_DATA_CACHE(DYN_ANNXD_T01, WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000"))

                If (WS_DYN_ANNXD.dataDescription = "TRG_PARAM") Then
                    With .TEMPLATES(0)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If

                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With

                    If (WS_PAGESNUM And 1) Then
                        With .TEMPLATES(1)
                            WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    End If
                End If

            Case "INFO"
                If ((WS_BDS = False) And (WS_CCP = False) And (WS_FLG_FATTELE_PA Or WS_FLG_NEG_NODOM Or WS_FLG_DOM)) Then
                    With .TEMPLATES(IIf(WS_FLG_DOM, 0, 1))
                        If (Trim$(.FIELDS) <> "") Then
                            If (WS_FLG_DOM) Then
                                .FIELDS = Replace$(.FIELDS, "[$NI_X]", "10")
                                .FIELDS = Replace$(.FIELDS, "[$NI_Y]", "35")
                                .FIELDS = Replace$(.FIELDS, "[$NI_W]", "190")
                                
                                DDS_ADD "[$NI_WC01]", "6"
                                DDS_ADD "[$NI_WC02]", "184"
                            Else
                                .FIELDS = Replace$(.FIELDS, "[$NI_X]", "110")
                                .FIELDS = Replace$(.FIELDS, "[$NI_Y]", "6")
                                .FIELDS = Replace$(.FIELDS, "[$NI_W]", "90")
                                
                                DDS_ADD "[$NI_WC01]", "4"
                                DDS_ADD "[$NI_WC02]", "86"
                            End If

                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If
                        
                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                    
                    With .TEMPLATES(2)
                        WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                End If
            
            Case "INFO_BDS"
                If (WS_BDS And (WS_CCP = False) And (WS_FLG_FATTELE_PA Or WS_FLG_NEG_NODOM Or WS_FLG_DOM)) Then
                    With .TEMPLATES(IIf(WS_FLG_DOM, 0, 1))
                        If (Trim$(.FIELDS) <> "") Then
                            If (WS_FLG_DOM) Then
                                .FIELDS = Replace$(.FIELDS, "[$NI_X]", "10")
                                .FIELDS = Replace$(.FIELDS, "[$NI_Y]", "35")
                                .FIELDS = Replace$(.FIELDS, "[$NI_W]", "190")
                                
                                DDS_ADD "[$NI_WC01]", "6"
                                DDS_ADD "[$NI_WC02]", "184"
                            Else
                                .FIELDS = Replace$(.FIELDS, "[$NI_X]", "110")
                                .FIELDS = Replace$(.FIELDS, "[$NI_Y]", "6")
                                .FIELDS = Replace$(.FIELDS, "[$NI_W]", "90")
                                
                                DDS_ADD "[$NI_WC01]", "4"
                                DDS_ADD "[$NI_WC02]", "86"
                            End If

                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If
                        
                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                    
                    With .TEMPLATES(2)
                        WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                End If

            Case "INVOICE"
                With .TEMPLATES(IIf(WS_FLG_DIV, 1, 0))
                    If (WS_FLG_DIV) Then
                        WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                    Else
                        If (WS_BDS) Then
                            WS_FIELDS = WS_FIELDS & Replace$(GET_FIELDS(.FIELDS, WS_PAGESNUM), "[$QS_PP_Y]", "150")
                        Else
                            WS_FIELDS = WS_FIELDS & Replace$(GET_FIELDS(.FIELDS, WS_PAGESNUM), "[$QS_PP_Y]", "163")
                        End If
                    End If

                    WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                    WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                    WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                    WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                    WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                End With
            
            Case "PAGO_PA"
                If (WS_CCP And WS_CHK_G34) Then
                    If (WS_PAGEBILL = 0) Then WS_PAGEBILL = (WS_PAGESNUM + 1)

                    For J = 0 To UBound(WS_G34)
                        If (WS_G23(J).CL_BOLLETTINO_ID <> "999999999999999999") Then
                            WS_INT = IIf((WS_G23(J).NUMERORATA = "00"), 0, 1)
                            WS_PPA_CNTR = (WS_PPA_CNTR + 1)
                            
                            With .TEMPLATES(WS_INT)
                                WS_FIELDS = WS_FIELDS & Replace$(GET_FIELDS(.FIELDS, WS_PAGESNUM), "XXX", Format$(WS_PPA_CNTR, "000"))
                    
                                If (WS_G23(J).NUMERORATA = "00") Then
                                    WS_STRING = Replace$(.FIELDSDATA, "[$CODICE_AVVISO]", GET_PAGOPACODE(Trim$(WS_G34(J).CODICE_NAV)))
                                Else
                                    WS_STRING = Replace$(.FIELDSDATA, "[$VAR_AMOUNT]", Trim$(WS_G23(J).IMPORTO))
                                    WS_STRING = Replace$(WS_STRING, "[$VAR_DTASCADENZA]", WS_G23(J).SCADENZA)
                                    WS_STRING = Replace$(WS_STRING, "[$VAR_DELAYNUM]", WS_G23(J).NUMERORATA)
                                    WS_STRING = Replace$(WS_STRING, "[$CODICE_AVVISO]", GET_PAGOPACODE(Trim$(WS_G34(J).CODICE_NAV)))
                                End If
                        
                                WS_ANNEXED_DATA = WS_ANNEXED_DATA & WS_STRING
                                WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                                WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                                WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                                WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                            End With
                            
                            WS_PAGEBILLS = WS_PAGEBILLS & ";" & WS_PAGESNUM
                        End If
                    Next J
                    
                    With .TEMPLATES(2)
                        .FIELDS = Replace$(.FIELDS, "[$NI_X]", "110")
                        .FIELDS = Replace$(.FIELDS, "[$NI_Y]", "6")
                        .FIELDS = Replace$(.FIELDS, "[$NI_W]", "90")

                        DDS_ADD "[$NI_WC01]", "4"
                        DDS_ADD "[$NI_WC02]", "86"

                        WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                        WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA

                        WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                    
                    If (WS_PAGESNUM And 1) Then
                        With .TEMPLATES(3)
                            WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                            WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                            WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                        End With
                    End If
                    
                    If (Left$(WS_PAGEBILLS, 1) = ";") Then WS_PAGEBILLS = Mid$(WS_PAGEBILLS, 2)
                End If
            
            Case Else
                For J = 0 To UBound(WS_SECTIONS(I).TEMPLATES)
                    With .TEMPLATES(J)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            
                            For K = 0 To .TEMP_PAGES
                                WS_PAGENUMBER = WS_PAGENUMBER & ";1"
                            Next K
                        Else
                            For K = 0 To .TEMP_PAGES
                                WS_PAGENUMBER = WS_PAGENUMBER & ";0"
                            Next K
                        End If
                        
                        WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                Next J
            
            End Select
        End With
    Next I
    
    With myXFDFMLTemplate
        .setTemplateVersion = DLLParams.TEMPLATEVERSION
        .setTemplateFileName = WS_FILENAME & GET_TEXTPAD(PADRIGHT, GET_VAL2HEX(WS_TEMP_UNIQUEID), 12, "0", False) & "_P" & Format$(WS_PAGESNUM, "000")
        .setTemplateIndexes = Left$(WS_PAGESINDXS, Len(WS_PAGESINDXS) - 1)
        .setExtraFields = WS_FIELDS & WS_EXTRAFIELDS
        
        ' PAGE NUMBER
        '
        WS_SPLIT = Split(Mid$(WS_PAGENUMBER, 2), ";")
        
        For I = 1 To WS_PAGESNUM
            If (WS_SPLIT(I - 1) = "1") Then
                .setFieldId = "NMR_PAGE_R" & Format$(I, "000")
                .setPropertyPageId = I
                .setPropertyCoords = "175,282.5,25,4"
                .setPropertyFontName = "helr45w.ttf"
                .setPropertyFontSize = "7"
                .setPropertyAlignment = "right"
                .closeProperty
                .closeField
                
                WS_PAGENUM = WS_PAGENUM & "Pag. " & I & " di " & WS_PAGESNUM & "|"
            Else
                WS_PAGENUM = WS_PAGENUM & "|"
            End If
        Next I
        
        WS_PAGENUM = Left$(WS_PAGENUM, Len(WS_PAGENUM) - 1)
        
        Erase WS_SPLIT()
        Erase WS_SPLIT_VLS()
        
        ' EXTRA DATA
        '
        .setExtraInfo = "pages=" & Chr$(34) & WS_PAGESNUM & Chr$(34)
        .setExtraInfo = "billpage=" & Chr$(34) & IIf(WS_CCP, WS_PAGEBILLS, "") & Chr$(34)
        
        GET_TEMPLATEINFO = .getXFDFTemplateNode
    End With

End Function

Private Function GET_XML(strData As String, Justify As String, FontName As String, FontSize As String, FontLeading As String, Optional FontColor = "0,0,0") As String
    
    Dim I           As Integer
    Dim J           As Byte
    Dim SplitData() As String
    Dim WS_TXTDATA  As String
    Dim WS_TXTFLAG  As String
    Dim WS_TXTSTYLE As String
    
    SplitData = Split(strData, "|")
    
    With myXFDFMLText
        .setNodeChunkRGBColor = FontColor
        .setNodeTextFontName = FontName
        .setNodeTextFontSize = FontSize
        .setNodeTextFontLeadingFixed = FontLeading
        .setNodeTextTextAlign = Justify
        
        For I = 0 To UBound(SplitData)
            If (SplitData(I) = "") Then
                WS_TXTDATA = "<br>"
            Else
                WS_TXTSTYLE = ""
                
                Select Case Left$(SplitData(I), 4)
                Case "<br>"
                    WS_TXTDATA = SplitData(I)
                
                Case "<bs>"
                    WS_TXTDATA = Replace$(SplitData(I), "<bs>", "")
                    
                Case Else
                    If (Left$(SplitData(I), 1) = "<") Then
                        J = 2
                        
                        Do
                            WS_TXTFLAG = Mid$(SplitData(I), J, 1)
                                    
                            Select Case WS_TXTFLAG
                            Case "b"
                                WS_TXTSTYLE = WS_TXTSTYLE & IIf(Trim$(WS_TXTSTYLE) = "", "", ",") & "bold"
                            
                            Case "i"
                                WS_TXTSTYLE = WS_TXTSTYLE & IIf(Trim$(WS_TXTSTYLE) = "", "", ",") & "italic"
                            
                            Case "u"
                                WS_TXTSTYLE = WS_TXTSTYLE & IIf(Trim$(WS_TXTSTYLE) = "", "", ",") & "underline"
                            
                            End Select
                            
                            J = J + 1
                        Loop Until (WS_TXTFLAG = ">")
                        
                        WS_TXTDATA = Mid$(SplitData(I), J)
                    Else
                        WS_TXTDATA = SplitData(I)
                    End If
                
                End Select
            End If
            
            .setNodeChunkFontStyle = WS_TXTSTYLE
            .addChunk = WS_TXTDATA
        Next I
                
        GET_XML = .getXFDFTextNode
    End With
    
    Erase SplitData

End Function

Private Function GET_XMLMETADATA()
    
    Dim myXMLMD   As cls_XMLMetaData
    
    Set myXMLMD = New cls_XMLMetaData

    With myXMLMD
        .setMetaData("TXT_CAP") = Left$(WS_G03.R004.LOCALITÀ, 5)
        .setMetaData("TXT_DESTINATARIO") = Replace$(WS_RECIPIENT, "<br>", " - ")
        .setMetaData("TXT_INDIRIZZO_RECAPITO") = UCase$(Trim$(IIf((Trim$(WS_G03.R003.DESCRIZIONEESTESARECAPITO) = ""), WS_G03.R003.INDIRIZZO, WS_G03.R003.DESCRIZIONEESTESARECAPITO)))
        .setMetaData("TXT_CLP_RECAPITO") = Replace$(WS_LOCALITY, "<br>", " - ")
        .setMetaData("TXT_NATIONALITY") = WS_NATIONALITY
        
        .setMetaData("IDENTIFICATIVO_DOCUMENTO") = WS_G01.R001.ANNOBOLLETTA & Trim$(WS_G01.R001.NUMEROBOLLETTA)
        .setMetaData("TIPO_DOCUMENTO") = IIf(WS_G01.R001.TIPONUMERAZIONE = "0", "BOC", "BOE")
        .setMetaData("CODICE_CLIENTE") = Trim$(WS_G02.R001.CODICEANAGRAFICO)
        .setMetaData("CODICE_SERVIZIO") = Trim$(WS_G02.R001.CODICESERVIZIO)
        .setMetaData("PROGRESSIVO_LOTTO") = Trim$(WS_G00.SEQUENZASTAMPA) & Trim$(WS_G00.BASEDATI)
                    
        GET_XMLMETADATA = .getXMLMetaData
    End With

    Set myXMLMD = Nothing

End Function

Public Sub MMS_Close(isOk As Boolean)
    
    mySQLImporter.setExtError = (Not isOk)
    mySQLImporter.EndJob
    Set mySQLImporter = Nothing
    
    Set myXFDFMLTable = Nothing
    Set myXFDFMLText = Nothing

End Sub

Public Function MMS_GetErrMsg() As String
    
    MMS_GetErrMsg = WS_ERRMSG

End Function

Public Function MMS_GetErrSctn() As String

    MMS_GetErrSctn = WS_ERRSCT

End Function

Public Function MMS_Insert_BOL() As Boolean
    
    WS_ERRSCT = "MMS_INSERT_BOL"
    
    Dim I                           As Integer
    Dim WS_BILL_ID                  As String
    Dim WS_FLG_LETTURA_FINALE       As Boolean
    Dim WS_FLG_NOMERGE              As Boolean
    Dim WS_ATTACHDETAILS            As String
    Dim WS_AUI_DETAILS              As String
    Dim WS_INVOICEDETAILS           As String
    Dim WS_PERIODOFATTURA           As String
    Dim WS_STRING                   As String
    
    Dim XML_TEMPLATE                As String
    Dim TXT_INDOOR                  As String
    Dim TXT_ADDRESS                 As String
    Dim TXT_PXX_HEADER              As String
    Dim TXT_P02_CONSUMI             As String
    Dim HST_P02_CONSUMI             As String
    Dim TXT_INTESTATARIO            As String
    Dim TXT_INDIRIZZOFORNITURA      As String
    Dim TXT_CLPFORNITURA            As String
    Dim NMR_BOLLETTINOIMPORTO       As String
    Dim NMR_BOLLETTINOFATTURA       As String
    Dim DTA_BOLLETTINOSCADENZA      As String
    Dim TXT_BOLLETTINOIDCLIENTE     As String
    Dim TXT_BOLLETTINOID            As String
    Dim TXT_BOLLETTINOBC            As String
    Dim TXT_FEPAINFO_RXXX           As String
    Dim TXT_BC_PPA_QRCODE           As String
    Dim TXT_BC_PPA_DATAMATRIX       As String
    Dim TXT_EMAIL                   As String
    Dim TXT_DOCFILEPATH             As String
    Dim TXT_DOCFILENAME             As String
    Dim XML_METADATA                As String
    
    ' INIT
    '
    Set myXFDFMLTable = New cls_XFDFMLTable
    Set myXFDFMLText = New cls_XFDFMLText
    
    DDS_INIT
    
    If (WS_G09.R003.TIPOPRESA = "0") Then Err.Raise vbObjectError + 512, "MMS_INSERT", "Tipo Fattura con TIPOPRESA = 0 non gestita"
    
    If (WS_FLG_WMS) Then
        WS_ATTACHDETAILS = GET_PXX_ATTACH_DETAILS
    
        If (WS_ATTACHDETAILS = "") Then Err.Raise vbObjectError + 513, "MMS_INSERT", "Impossibile Integrare Allegato. Dati Impaginazione Assenti"
    End If
    
    With WS_FORNITURA_MASTER_P02
        .dataDescription = ""
        .DATAID = ""
        .FLG_EXTRAPARAMS = False
        
        Erase .EXTRAPARAMS()
    End With
    
    GET_FLG_INFOPAGE
    
    WS_BDS = (WS_G01.R001.TIPOSERVIZIO = "05")
    WS_CCP = False
    WS_DEL_547_19_B_CAUSALE = ""
    WS_CNTTR_MASTER_MTRCL = ""
    WS_CNTTR_MASTER_SRVZ = ""
    WS_CODICE_SERVIZIO_KEY = Format$(Trim$(WS_G02.R001.CODICESERVIZIO), "0000000000")
    WS_FLG_CATEGORIA = GET_DATA_CACHE(CATEGORIE, Format$(WS_G09.R003.CODICECATEGORIAUTENZA, "000")).dataDescription = "TRG_CATEGORIA"
    WS_DF_GBO_FLG = False
    WS_DOCPAGES_DA_CNTR = 0
    WS_DOCPAGES_DF_CNTR = 0
    WS_FLG_ANNXD_547_19 = False
    WS_FLG_CONSUMI = (WS_CHK_G22 And (WS_FLG_WMS = False))
    WS_FLG_DIV = False
    WS_FLG_DF_IV = False
    WS_FLG_DF_TS = False
    WS_FLG_FATTELE_PA = (WS_G01.R001.TIPONUMERAZIONE = "5")
    WS_FLG_INDE_LBL = False
    WS_FLG_LETTURA_FINALE = GET_CHK_LF
    WS_FLG_MASTER = False
    WS_FLG_MASTER_ECC_POS = False
'    WS_FLG_MSG_BS = False
    WS_FLG_NEG_NODOM = False
    WS_FLG_NOTACREDITO = (Trim$(WS_G01.R004.DESCRIZIONE) <> "")
    WS_FLG_PARTITE = (WS_G01.R001.TIPOBOLLETTAZIONE = "P")
    WS_IMPORTOTOTALE = NRM_IMPORT(WS_G06.R001.TOTALEEURO, "##,##0.00", False)
    WS_LOCALITY = ""
    WS_MSG_CI = (GET_DATA_CACHE(DYN_ANNXD_T02, WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000")).dataDescription = "TRG_PARAM")
    WS_NATIONALITY = ""
    WS_PAGENUM = ""
    WS_PERIODOFATTURA = UCase$(Trim$(WS_G01.R001.PERIODO))
    WS_RECIPIENT = ""
    
    WS_STRING = WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000")
    WS_MSG_CA = (GET_DATA_CACHE(DYN_ANNXD_T03, WS_STRING).dataDescription <> WS_STRING)
    
    WS_STRING = WS_G00.AZIENDASIU & "_" & WS_CODICE_SERVIZIO_KEY & "_" & Format$(WS_G01.R006.PROGRESSIVOFATTURAZIONE, "000000") & "_" & Format$(WS_G01.R001.SEZIONALE, "00") & "_" & WS_G01.R001.TIPONUMERAZIONE & "_" & Format$(WS_G01.R001.CODSERVIZIONUMERAZIONE, "00") & "_" & WS_G01.R001.ANNOBOLLETTA & "_" & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000") & "_" & WS_G01.R001.RATABOLLETTA
    WS_IMPORTO_TC_TICSI = GET_DATA_CACHE(TOTCONG_TICSI, WS_STRING).dataDescription
    
    If (WS_IMPORTO_TC_TICSI = WS_STRING) Then WS_IMPORTO_TC_TICSI = ""
    
    If (Trim$(WS_GPB.R002.MOTIVAZIONE) <> "") Then DDS_ADD "[$VAR_REASON]", Trim$(WS_GPB.R002.MOTIVAZIONE)
    
    If (Trim$(WS_GPB.R001.DATA_PRESCRIZIONE) = "") Then
        DDS_ADD "[$DTA_PRESCRIZIONE]", ""
        DDS_ADD "[$IMPORTO_PRESCRIZIONE]", ""
    Else
        If (CDbl(WS_GPB.R001.POTENZIALE_PRESCRIZIONE) > 0) Then
            DDS_ADD "[$DTA_PRESCRIZIONE]", WS_GPB.R001.DATA_PRESCRIZIONE
            DDS_ADD "[$IMPORTO_PRESCRIZIONE]", NRM_IMPORT(WS_GPB.R001.POTENZIALE_PRESCRIZIONE, "##,##0.00", False)
        Else
            DDS_ADD "[$DTA_PRESCRIZIONE]", ""
            DDS_ADD "[$IMPORTO_PRESCRIZIONE]", ""
        End If
    End If
    
    If (WS_CHK_GOR) Then
        If (Trim$(WS_GOR.ORDINE_ACQUISTO_NSO & WS_GOR.DATA_ORDINE_ACQUISTO_NSO & WS_GOR.EMITTENTE_ORDINE_NSO) <> "") Then
            WS_STRING = Replace$(WS_CS_P01_MSG_NSO.dataDescription, "[$VAR_NUM_NSO]", Trim$(WS_GOR.ORDINE_ACQUISTO_NSO))
            WS_STRING = Replace$(WS_STRING, "[$VAR_DTA_NSO]", WS_GOR.DATA_ORDINE_ACQUISTO_NSO)
            WS_STRING = Replace$(WS_STRING, "[$VAR_EMI_NSO]", Trim$(WS_GOR.EMITTENTE_ORDINE_NSO))
        
            DDS_ADD "[$P01_MSG_NSO]", WS_STRING
        Else
            DDS_ADD "[$P01_MSG_NSO]", ""
        End If
    Else
        DDS_ADD "[$P01_MSG_NSO]", ""
    End If
    
    WS_FATTURANUMERO = Mid$(Trim$(WS_G01.R005.NOME_FE), 5)
    TXT_DOCFILENAME = Trim$(WS_G01.R005.NOME_FE)
    
    If (WS_FLG_FATTELE_PA) Then
        TXT_DOCFILEPATH = "FEL"
    
        WS_STRING = GET_DATA_CACHE(CODICE_IPA, WS_CODICE_SERVIZIO_KEY).dataDescription
        
        If (WS_STRING <> WS_CODICE_SERVIZIO_KEY) Then
            DDS_ADD "[$LBL_CODIPA]", "<br>Codice IPA"
            DDS_ADD "[$VAR_CODIPA]", "<br>" & WS_STRING
        Else
            DDS_ADD "[$LBL_CODIPA]", ""
            DDS_ADD "[$VAR_CODIPA]", ""
        End If
    Else
        If (Trim$(WS_G01.R007.MODALITÀ_INVIO) = "XSPDF") Then
            TXT_DOCFILEPATH = "FEPR"
        Else
            TXT_DOCFILEPATH = "BOL"
        End If
        
        DDS_ADD "[$LBL_CODIPA]", ""
        DDS_ADD "[$VAR_CODIPA]", ""
    End If

    If (WS_CHK_G23 And (WS_G01.R001.CCP = "S")) Then
        For I = 0 To UBound(WS_G23)
            If (WS_G23(I).CL_BOLLETTINO_ID <> "999999999999999999") Then
                WS_CCP = True

                Exit For
            End If
        Next I
    End If

    ' P01
    '
    WS_ERRSCT = "ERRORE REPERIMENTO DESTINATARIO - GRUPPO 03"
    
    With WS_G03
        For I = 0 To UBound(.R001)
            WS_RECIPIENT = WS_RECIPIENT & Trim$(.R001(I).NOME_RAGIONESOCIALE) & IIf((I = UBound(.R001)), "", "<br>")
        Next I
        
        TXT_INDOOR = Trim$(IIf(Trim$(.R002.INTERNO) = "", "", "INT. " & Trim$(.R002.INTERNO)) & IIf(Trim$(.R002.SCALA) = "", "", " SCALA " & Trim$(.R002.SCALA)) & IIf(Trim$(.R002.PIANO) = "", "", " PIANO: " & Trim$(.R002.PIANO)))
        TXT_ADDRESS = Trim$(UCase$(Trim$(IIf((Trim$(.R003.DESCRIZIONEESTESARECAPITO) = ""), .R003.INDIRIZZO, .R003.DESCRIZIONEESTESARECAPITO))))
        WS_LOCALITY = Trim$(UCase$(Trim$(IIf(Left$(.R004.LOCALITÀ, 5) = "00000", Mid$(.R004.LOCALITÀ, 7), .R004.LOCALITÀ))))
        
        If ((Trim$(.R004.SIGLA_NAZIONE) <> "") And (.R004.SIGLA_NAZIONE <> "IT")) Then
            If (Right$(WS_LOCALITY, 2) = "EE") Then WS_LOCALITY = Trim$(Left$(WS_LOCALITY, Len(WS_LOCALITY) - 2))
            
            WS_NATIONALITY = GET_DATA_CACHE(NATIONALITY, .R004.SIGLA_NAZIONE).dataDescription
            
            If (WS_NATIONALITY <> .R004.SIGLA_NAZIONE) Then WS_LOCALITY = WS_LOCALITY & "<br>" & WS_NATIONALITY
        End If
    End With
    
    WS_ERRSCT = "ERRORE REPERIMENTO INTESTATARIO - GRUPPO 02"
    
    TXT_INTESTATARIO = Trim$(WS_G02.R002.RAGIONESOCIALEINTESTATARIO)
    TXT_INDIRIZZOFORNITURA = Trim$(WS_G05.R001.INDIRIZZOFORNITURA)
    
    TXT_CLPFORNITURA = GET_DATA_CACHE(LOC_UBI, WS_CODICE_SERVIZIO_KEY).dataDescription
    If (TXT_CLPFORNITURA = WS_CODICE_SERVIZIO_KEY) Then TXT_CLPFORNITURA = Trim$(WS_G05.R002.LOCALITÀFORNITURA)

    GET_P01_PUNTO_EROGAZIONE
    GET_P01_CONSUMI
    
    If (WS_FLG_DIV = False) Then
        If (WS_BDS) Then
            DDS_ADD "[$P01_MSG_SERDEP]", ""
            DDS_ADD "[$MSG_INFO_FATTURA]", WS_CS_P01_MSG_INFO_FATTURA_BDS.dataDescription
        Else
            DDS_ADD "[$P01_MSG_SERDEP]", WS_CS_P01_MSG_SERDEP.dataDescription
            DDS_ADD "[$MSG_INFO_FATTURA]", WS_CS_P01_MSG_INFO_FATTURA_STD.dataDescription
        End If
    End If

    ' PXX - DETAILS
    '
    TXT_PXX_HEADER = GET_XML("Fattura num. |<b>" & WS_FATTURANUMERO & "| del |<b>" & WS_G01.R001.DATAEMISSIONE & "|<br>" & _
                             "Intestata a: |<b>" & TXT_INTESTATARIO & " - " & TXT_INDIRIZZOFORNITURA & " - " & TXT_CLPFORNITURA, "right", "helr45w.ttf", "8", "8")
    
    WS_PAGEHEIGHT = 15
    
    If (WS_FLG_CONSUMI) Then
        HST_P02_CONSUMI = GET_P02_DETTAGLIOCONSUMI_CHART
        WS_PAGEHEIGHT = 59
    End If
    
    WS_INVOICEDETAILS = GET_P02_DETTAGLIOLETTURE_MASTER & GET_P02_DETTAGLIOLETTURE_G13 & GET_P02_DETTAGLIOLETTURE_G14 & GET_P02_DETTAGLIOLETTURE_SCARTATE
    WS_INVOICEDETAILS = WS_INVOICEDETAILS & GET_PXX_INVOICE_DETAILS

    ' BILL DATA
    '
    If (WS_CCP) Then
        WS_ERRSCT = "GET_CCP"

        ' P.I. BILL
        '
        For I = 0 To UBound(WS_G23)
            With WS_G23(I)
                If (.CL_BOLLETTINO_ID <> "999999999999999999") Then
                    NMR_BOLLETTINOIMPORTO = NMR_BOLLETTINOIMPORTO & Trim$(.IMPORTO) & "[EB]"
                    NMR_BOLLETTINOFATTURA = NMR_BOLLETTINOFATTURA & WS_FATTURANUMERO & " - " & IIf(.NUMERORATA = "00", "Rata Unica", "Rata " & .NUMERORATA) & "[EB]"
                    DTA_BOLLETTINOSCADENZA = DTA_BOLLETTINOSCADENZA & .SCADENZA & "[EB]"
                    TXT_BOLLETTINOIDCLIENTE = TXT_BOLLETTINOIDCLIENTE & .CL_BOLLETTINO_ID & "[EB]"
                    TXT_BOLLETTINOID = TXT_BOLLETTINOID & "<" & .CL_BOLLETTINO_ID & ">" & GET_TEXTPAD(PADRIGHT, .CL_IMPORTO & ">", 18, " ", True) & GET_TEXTPAD(PADRIGHT, Format$(DLLParams.CCP_BILL, String$(12, "0")), 14, " ", True) & "<  896>[EB]"
                    TXT_BOLLETTINOBC = TXT_BOLLETTINOBC & "18" & .CL_BOLLETTINO_ID & "12" & Format$(DLLParams.CCP_BILL, String$(12, "0")) & "10" & Replace$(.CL_IMPORTO, "+", "") & "3896[EB]"
                End If
            End With
        Next I
        
        NMR_BOLLETTINOIMPORTO = Left$(NMR_BOLLETTINOIMPORTO, Len(NMR_BOLLETTINOIMPORTO) - 4)
        NMR_BOLLETTINOFATTURA = Left$(NMR_BOLLETTINOFATTURA, Len(NMR_BOLLETTINOFATTURA) - 4)
        DTA_BOLLETTINOSCADENZA = Left$(DTA_BOLLETTINOSCADENZA, Len(DTA_BOLLETTINOSCADENZA) - 4)
        TXT_BOLLETTINOIDCLIENTE = Left$(TXT_BOLLETTINOIDCLIENTE, Len(TXT_BOLLETTINOIDCLIENTE) - 4)
        TXT_BOLLETTINOID = Left$(TXT_BOLLETTINOID, Len(TXT_BOLLETTINOID) - 4)
        TXT_BOLLETTINOBC = Left$(TXT_BOLLETTINOBC, Len(TXT_BOLLETTINOBC) - 4)
        
        ' FEPA BOXES
        '
        TXT_FEPAINFO_RXXX = Trim$(WS_G02.R001.CODICESERVIZIO) & "|" & TXT_INTESTATARIO & "|" & Trim$(WS_G02.R001.CODICEANAGRAFICO) & "|" & Trim$(WS_G09.R002.CODICEPUNTORICONSEGNA)
    
        ' PPA
        '
        If (WS_CHK_G34) Then
            For I = 0 To UBound(WS_G34)
                If (WS_G23(I).CL_BOLLETTINO_ID <> "999999999999999999") Then
                    WS_STRING = Trim$(WS_G34(I).CODICE_NAV)
                    WS_BILL_ID = "18" & WS_STRING & "12" & Format$(DLLParams.CCP_PPA, String$(12, "0")) & "10" & Format$(Replace$(Trim$(WS_G23(I).IMPORTO), ",", ""), String(10, "0")) & "3896"
            
                    TXT_BC_PPA_QRCODE = TXT_BC_PPA_QRCODE & "PAGOPA|002|" & WS_STRING & "|" & DLLParams.CF_ENTE & "|" & Mid$(WS_BILL_ID, 37, 10) & "[EB]"
                    TXT_BC_PPA_DATAMATRIX = TXT_BC_PPA_DATAMATRIX & "codfase=NBPA;" & WS_BILL_ID & "1P1" & DLLParams.CF_ENTE & GET_TEXTPAD(PADLEFT, IIf(Trim$(WS_G02.R005.PARTITAIVA) = "", WS_G02.R005.CODICEFISCALE, WS_G02.R005.PARTITAIVA), 16, " ", True) & GET_TEXTPAD(PADLEFT, TXT_INTESTATARIO, 40, " ", False) + GET_TEXTPAD(PADLEFT, UCase$("Fattura nr. " & WS_FATTURANUMERO & " del " & Replace$(WS_G01.R001.DATAEMISSIONE, "/", "-")), 110, " ", False) + "            A[EB]"
                End If
            Next I
    
            TXT_BC_PPA_QRCODE = Left$(TXT_BC_PPA_QRCODE, Len(TXT_BC_PPA_QRCODE) - 4)
            TXT_BC_PPA_DATAMATRIX = Left$(TXT_BC_PPA_DATAMATRIX, Len(TXT_BC_PPA_DATAMATRIX) - 4)
        End If
    End If
    
    ' TEMPLATE BUILDER
    '
    If (((WS_FLG_NOTACREDITO = False) And (WS_FLG_PARTITE = False)) And WS_CHK_GSE) Then WS_AUI_DETAILS = GET_DYN_ANNXD_UI_DATA(HST_P02_CONSUMI)
    
    XML_TEMPLATE = GET_TEMPLATEINFO

    ' P01 - SECTION 01
    '
    If (WS_FLG_NOTACREDITO) Then
        DDS_ADD "[$VAR_DOCTYPE]", "Nota di Credito"
        DDS_ADD "[$TXT_STORNO]", GET_XML("<br>" & Trim$(WS_G01.R004.DESCRIZIONE) & " |<b>" & WS_G01.R004.ANNO & "/" & WS_G01.R004.NUMERO, "left", "helr45w.ttf", "9", "9")
        DDS_ADD "[$VAR_TIPOFATT]", "Nota di Credito"
        DDS_ADD "[$VAR_CNSMFATT]", NRM_REMOVEZEROES(WS_G11.R001.TOTALECONSUMOFATTURATO, True) & " mc (Restituzione)"
    Else
        DDS_ADD "[$VAR_DOCTYPE]", "Fattura"
        DDS_ADD "[$TXT_STORNO]", ""
        DDS_ADD "[$VAR_TIPOFATT]", "Fattura " & GET_TIPOFATTURA(IIf(WS_FLG_LETTURA_FINALE, "F", WS_G01.R001.TIPOBOLLETTAZIONE))
        DDS_ADD "[$VAR_CNSMFATT]", GET_P01_CONSUMOFATTURATO
    End If

    DDS_ADD "[$VAR_DOCNUM]", WS_FATTURANUMERO
    
    DDS_ADD "[$RECIPIENT]", WS_RECIPIENT
    DDS_ADD "[$PRESSO]", IIf((Trim$(WS_G03.R003.DESCRIZIONERECAPITO) = ""), "", "<br>" & Trim$(WS_G03.R003.DESCRIZIONERECAPITO))
    DDS_ADD "[$ADDRESS]", IIf(TXT_INDOOR = "", "", TXT_INDOOR & "<br>") & TXT_ADDRESS & "<br>" & WS_LOCALITY

    ' P01 - SECTION 02
    '
    DDS_ADD "[$VAR_DTAEMISSIONE]", WS_G01.R001.DATAEMISSIONE
    DDS_ADD "[$VAR_PERIODO]", WS_PERIODOFATTURA
    DDS_ADD "[$VAR_PAYMODE]", GET_P01_PAYMODE
    DDS_ADD "[$VAR_AMOUNT]", WS_IMPORTOTOTALE
    DDS_ADD "[$VAR_DTASCADENZA]", WS_G06.R001.DATASCADENZA

    ' P01 - SECTION 03
    '
    DDS_ADD "[$VAR_CFPI]", IIf(Trim$(WS_G02.R005.PARTITAIVA) = "", WS_G02.R005.CODICEFISCALE, WS_G02.R005.PARTITAIVA)
    DDS_ADD "[$VAR_MATCON]", Trim$(WS_G09.R001.MATRICOLACONTATORE)
    DDS_ADD "[$VAR_COCLI]", Trim$(WS_G02.R001.CODICEANAGRAFICO)
    DDS_ADD "[$VAR_TIPMIS]", Trim$(WS_G09.R001.TIPOLOGIAMISURATORE) & " " & Trim$(WS_G09.R001.DESCRIZIONEMODELLOCONTATORE)
    DDS_ADD "[$VAR_CODSER]", Trim$(WS_G02.R001.CODICESERVIZIO)
    DDS_ADD "[$VAR_INTESTATARIO]", TXT_INTESTATARIO
    DDS_ADD "[$LBL_DEPCAU]", IIf((WS_G09.R001.TIPO = "D" Or WS_BDS), "Deposito versato", "Anticipo fornitura")
    
    If (Trim$(WS_G09.R001.IMPORTO_PAGATO_DEPOSITO_ANTICIPO) = "") Then WS_G09.R001.IMPORTO_PAGATO_DEPOSITO_ANTICIPO = "0"
    
    If (CSng(WS_G09.R001.IMPORTO_PAGATO_DEPOSITO_ANTICIPO) <= 0) Then
        DDS_ADD "[$VAR_DEPCAU]", "€ 0,00"
    Else
        DDS_ADD "[$VAR_DEPCAU]", "€ " & NRM_IMPORT(WS_G09.R001.IMPORTO_PAGATO_DEPOSITO_ANTICIPO, "##,##0.00", False)
    End If
    
    DDS_ADD "[$VAR_UBIC]", TXT_INDIRIZZOFORNITURA & "<br>" & TXT_CLPFORNITURA
    DDS_ADD "[$VAR_SERDEP]", GET_SERVIZIODEPURAZIONE
    
    ' P01 - SECTION 04
    '
    DDS_ADD "[$TBL_SUMMARY]", GET_P01_QUADROSINTETICO
    DDS_ADD "[$TBL_SD_BD]", GET_P01_DETTAGLIOLETTURE_G13 & GET_P01_DETTAGLIOLETTURE_G14 & GET_P01_TICSI & GET_P01_PAGAMENTIPRECEDENTI & GET_P01_PAYTYPE
    DDS_ADD "[$TBL_INVOICEINFO]", GET_P01_INFOBOLLETTA & GET_P01_COM_ARERA
    
    ' P01 - MSG BS
    '
    'If (WS_FLG_MSG_BS) Then
    '    DDS_ADD "[$P01_MSG_BS]", WS_CS_P01_MSG_BS.dataDescription
    'Else
        DDS_ADD "[$P01_MSG_BS]", ""
    'End If
    
    ' INFO PAGE
    '
    GET_PXX_INFO_PAGE
    
    ' GET DATA
    '
    WS_ANNEXED_DATA = GET_FIELDSDATA(WS_ANNEXED_DATA)

    ' OUTPUT PLUGIN SUPPORT
    '
    XML_METADATA = GET_XMLMETADATA

    ' EMAIL MANAGEMENT
    '
    Select Case Trim$(WS_G01.R002.CANALEINOLTRO)
    Case "02"   ' EMAIL
        WS_FLG_NOMERGE = True

    Case "03"  ' STAMPA + EMAIL
        ' WS_FLG_NOMERGE = False

    Case Else
        WS_FLG_NOMERGE = (WS_G01.R007.RINUNCIA_COPIA_ANALOGICA = "S")

    End Select

    ' RECORD
    '
    WS_STRING = XML_TEMPLATE & "§" & _
                WS_RECIPIENT & "§" & TXT_ADDRESS & "§" & WS_LOCALITY & "§" & _
                WS_ANNEXED_DATA & "§" & _
                TXT_PXX_HEADER & "§" & _
                TXT_P02_CONSUMI & "§" & HST_P02_CONSUMI & "§" & _
                WS_AUI_DETAILS & "§" & _
                WS_INVOICEDETAILS & "§" & _
                WS_ATTACHDETAILS & "§" & _
                WS_PXX_FOOTER & "§" & _
                WS_PAGENUM & "§" & _
                WS_IMPORTOTOTALE & "§" & _
                WS_FATTURANUMERO & "§" & _
                WS_G01.R001.DATAEMISSIONE & "§" & _
                WS_G06.R001.DATASCADENZA & "§" & _
                Trim$(WS_G02.R001.CODICESERVIZIO) & "§" & _
                WS_PERIODOFATTURA & "§" & _
                TXT_INTESTATARIO & "§" & TXT_INDIRIZZOFORNITURA & "§" & TXT_CLPFORNITURA & "§" & _
                NMR_BOLLETTINOIMPORTO & "§" & NMR_BOLLETTINOFATTURA & "§" & DTA_BOLLETTINOSCADENZA & "§" & TXT_BOLLETTINOIDCLIENTE & "§" & TXT_BOLLETTINOID & "§" & TXT_BOLLETTINOBC & "§" & _
                TXT_FEPAINFO_RXXX & "§" & _
                TXT_BC_PPA_QRCODE & "§" & TXT_BC_PPA_DATAMATRIX & "§" & _
                TXT_EMAIL & "§" & IIf((WS_FLG_FATTELE_PA Or WS_FLG_NOMERGE Or (DLLParams.PLUGMODE = "SDI")), "1", "") & "§" & _
                TXT_DOCFILEPATH & "§" & _
                TXT_DOCFILENAME & "§" & _
                XML_METADATA

    ' CLEAN
    '
    Erase WS_DOCPAGES_DA()
    Erase WS_DOCPAGES_DF()
    
    Set myXFDFMLTable = Nothing
    Set myXFDFMLText = Nothing
    
    ' INSERT DATA
    '
    MMS_Insert_BOL = mySQLImporter.SQLInsert(WS_STRING)
    WS_ERRMSG = mySQLImporter.GetUMErrorMessage
    WS_ERRSCT = "MMS_INSERT"
    
    DoEvents

End Function

Public Function MMS_Insert_L01() As Boolean ' LP
    
    WS_ERRSCT = "MMS_INSERT_L01"

    Dim WS_FLG_NOMERGE              As Boolean
    Dim WS_STRING                   As String

    Dim I                           As Integer
    Dim TXT_ADDRESS                 As String
    Dim TXT_EMAIL                   As String
    Dim TXT_INDOOR                  As String
    Dim XML_TEMPLATE                As String
    Dim XML_METADATA                As String

    ' INIT
    '
    WS_FATTURANUMERO = Mid$(Trim$(WS_G01.R005.NOME_FE), 5)
    WS_FLG_FATTELE_PA = (WS_G01.R001.TIPONUMERAZIONE = "5")
    WS_PAGENUM = ""
    WS_RECIPIENT = ""

    ' P01
    '
    WS_ERRSCT = "ERRORE REPERIMENTO DESTINATARIO - GRUPPO 03"
    
    With WS_G03
        For I = 0 To UBound(.R001)
            WS_RECIPIENT = WS_RECIPIENT & Trim$(.R001(I).NOME_RAGIONESOCIALE) & IIf((I = UBound(.R001)), "", "<br>")
        Next I
        
        TXT_INDOOR = Trim$(IIf(Trim$(.R002.INTERNO) = "", "", "INT. " & Trim$(.R002.INTERNO)) & IIf(Trim$(.R002.SCALA) = "", "", " SCALA " & Trim$(.R002.SCALA)) & IIf(Trim$(.R002.PIANO) = "", "", " PIANO: " & Trim$(.R002.PIANO)))
        TXT_ADDRESS = Trim$(UCase$(Trim$(IIf((Trim$(.R003.DESCRIZIONEESTESARECAPITO) = ""), .R003.INDIRIZZO, .R003.DESCRIZIONEESTESARECAPITO))))
        WS_LOCALITY = Trim$(UCase$(Trim$(IIf(Left$(.R004.LOCALITÀ, 5) = "00000", Mid$(.R004.LOCALITÀ, 7), .R004.LOCALITÀ))))
        
        If ((Trim$(.R004.SIGLA_NAZIONE) <> "") And (.R004.SIGLA_NAZIONE <> "IT")) Then
            If (Right$(WS_LOCALITY, 2) = "EE") Then WS_LOCALITY = Trim$(Left$(WS_LOCALITY, Len(WS_LOCALITY) - 2))
            
            WS_NATIONALITY = GET_DATA_CACHE(NATIONALITY, .R004.SIGLA_NAZIONE).dataDescription
            
            If (WS_NATIONALITY <> .R004.SIGLA_NAZIONE) Then WS_LOCALITY = WS_LOCALITY & "<br>" & WS_NATIONALITY
        End If
    End With

    DDS_INIT
    DDS_ADD "[$VAR_SYSDATE]", Format$(Now(), "dd/MM/yyyy")
    DDS_ADD "[$RECIPIENT]", WS_RECIPIENT
    DDS_ADD "[$PRESSO]", IIf((Trim$(WS_G03.R003.DESCRIZIONERECAPITO) = ""), "", "<br>" & Trim$(WS_G03.R003.DESCRIZIONERECAPITO))
    DDS_ADD "[$ADDRESS]", IIf(TXT_INDOOR = "", "", TXT_INDOOR & "<br>") & TXT_ADDRESS & "<br>" & WS_LOCALITY
    DDS_ADD "[$VAR_UBIC]", Trim$(WS_G05.R001.INDIRIZZOFORNITURA)
    DDS_ADD "[$VAR_CODSER]", Trim$(WS_G02.R001.CODICESERVIZIO)
    DDS_ADD "[$VAR_MATCON]", Trim$(WS_G09.R001.MATRICOLACONTATORE)
    DDS_ADD "[$VAR_DOCNUM]", WS_FATTURANUMERO
    DDS_ADD "[$VAR_DTAEMISSIONE]", WS_G01.R001.DATAEMISSIONE
    DDS_ADD "[$DTA_PRESCRIZIONE]", WS_GPB.R001.DATA_PRESCRIZIONE
    DDS_ADD "[$IMPORTO_PRESCRIZIONE]", NRM_IMPORT(WS_GPB.R001.POTENZIALE_PRESCRIZIONE, "##,##0.00", False)
    
    ' TEMPLATE BUILDER
    '
    XML_TEMPLATE = GET_TEMPLATEINFO

    ' GET DATA
    '
    WS_ANNEXED_DATA = GET_FIELDSDATA(WS_ANNEXED_DATA)

    ' OUTPUT PLUGIN SUPPORT
    '
    XML_METADATA = GET_XMLMETADATA

    ' EMAIL MANAGEMENT
    '
    Select Case Trim$(WS_G01.R002.CANALEINOLTRO)
    Case "02"   ' EMAIL
        WS_FLG_NOMERGE = True

    Case "03"  ' STAMPA + EMAIL
        ' WS_FLG_NOMERGE = False

    Case Else
        WS_FLG_NOMERGE = (WS_G01.R007.RINUNCIA_COPIA_ANALOGICA = "S")

    End Select

    ' RECORD
    '
    WS_STRING = XML_TEMPLATE & "§" & _
                WS_RECIPIENT & "§" & TXT_ADDRESS & "§" & WS_LOCALITY & "§" & _
                WS_ANNEXED_DATA & "§§§§§§" & _
                WS_PXX_FOOTER & "§" & _
                WS_PAGENUM & "§§" & _
                WS_FATTURANUMERO & "§" & _
                WS_G01.R001.DATAEMISSIONE & "§" & _
                WS_G06.R001.DATASCADENZA & "§" & _
                Trim$(WS_G02.R001.CODICESERVIZIO) & "§§§§§§§§§§§§§§" & _
                TXT_EMAIL & "§" & IIf((WS_FLG_FATTELE_PA Or WS_FLG_NOMERGE), "1", "") & "§" & _
                "LET" & "§" & _
                Trim$(WS_G01.R005.NOME_FE) & "_LP" & "§" & _
                XML_METADATA
    
    ' INSERT DATA
    '
    MMS_Insert_L01 = mySQLImporter.SQLInsert(WS_STRING)
    WS_ERRMSG = mySQLImporter.GetUMErrorMessage
    WS_ERRSCT = "MMS_INSERT"
    
    DoEvents

End Function

Public Function MMS_Insert_L02() As Boolean ' LM
    
    WS_ERRSCT = "MMS_INSERT_L02"

    Dim WS_DATA()                   As String
    Dim WS_FLG_NOMERGE              As Boolean
    Dim WS_MSG_CA_DATA              As strct_DATA
    Dim WS_STRING                   As String

    Dim I                           As Integer
    Dim TXT_ADDRESS                 As String
    Dim TXT_EMAIL                   As String
    Dim TXT_INDOOR                  As String
    Dim XML_TEMPLATE                As String
    Dim XML_METADATA                As String

    ' INIT
    '
    WS_STRING = WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000")
    WS_MSG_CA_DATA = GET_DATA_CACHE(DYN_ANNXD_T03, WS_STRING)

    If (WS_MSG_CA_DATA.dataDescription <> WS_STRING) Then
        WS_FATTURANUMERO = Mid$(Trim$(WS_G01.R005.NOME_FE), 5)
        WS_FLG_FATTELE_PA = (WS_G01.R001.TIPONUMERAZIONE = "5")
        WS_PAGENUM = ""
        WS_RECIPIENT = ""
    
        ' P01
        '
        WS_ERRSCT = "ERRORE REPERIMENTO DESTINATARIO - GRUPPO 03"
        
        With WS_G03
            For I = 0 To UBound(.R001)
                WS_RECIPIENT = WS_RECIPIENT & Trim$(.R001(I).NOME_RAGIONESOCIALE) & IIf((I = UBound(.R001)), "", "<br>")
            Next I
            
            TXT_INDOOR = Trim$(IIf(Trim$(.R002.INTERNO) = "", "", "INT. " & Trim$(.R002.INTERNO)) & IIf(Trim$(.R002.SCALA) = "", "", " SCALA " & Trim$(.R002.SCALA)) & IIf(Trim$(.R002.PIANO) = "", "", " PIANO: " & Trim$(.R002.PIANO)))
            TXT_ADDRESS = Trim$(UCase$(Trim$(IIf((Trim$(.R003.DESCRIZIONEESTESARECAPITO) = ""), .R003.INDIRIZZO, .R003.DESCRIZIONEESTESARECAPITO))))
            WS_LOCALITY = Trim$(UCase$(Trim$(IIf(Left$(.R004.LOCALITÀ, 5) = "00000", Mid$(.R004.LOCALITÀ, 7), .R004.LOCALITÀ))))
            
            If ((Trim$(.R004.SIGLA_NAZIONE) <> "") And (.R004.SIGLA_NAZIONE <> "IT")) Then
                If (Right$(WS_LOCALITY, 2) = "EE") Then WS_LOCALITY = Trim$(Left$(WS_LOCALITY, Len(WS_LOCALITY) - 2))

                WS_NATIONALITY = GET_DATA_CACHE(NATIONALITY, .R004.SIGLA_NAZIONE).dataDescription

                If (WS_NATIONALITY <> .R004.SIGLA_NAZIONE) Then WS_LOCALITY = WS_LOCALITY & "<br>" & WS_NATIONALITY
            End If
        End With
    
        DDS_INIT
        DDS_ADD "[$VAR_SYSDATE]", Format$(Now(), "dd/MM/yyyy")
        DDS_ADD "[$RECIPIENT]", WS_RECIPIENT
        DDS_ADD "[$PRESSO]", IIf((Trim$(WS_G03.R003.DESCRIZIONERECAPITO) = ""), "", "<br>" & Trim$(WS_G03.R003.DESCRIZIONERECAPITO))
        DDS_ADD "[$ADDRESS]", IIf(TXT_INDOOR = "", "", TXT_INDOOR & "<br>") & TXT_ADDRESS & "<br>" & WS_LOCALITY
        DDS_ADD "[$VAR_UBIC]", Trim$(WS_G05.R001.INDIRIZZOFORNITURA)
        DDS_ADD "[$VAR_CODSER]", Trim$(WS_G02.R001.CODICESERVIZIO)
        DDS_ADD "[$VAR_MATCON]", Trim$(WS_G09.R001.MATRICOLACONTATORE)
        DDS_ADD "[$VAR_DOCNUM]", WS_FATTURANUMERO
        DDS_ADD "[$VAR_DTAEMISSIONE]", WS_G01.R001.DATAEMISSIONE
        
        ' DATA ROWS
        '
        WS_DATA = Split(WS_MSG_CA_DATA.dataDescription, "|")
        WS_STRING = ""
        
        DDS_ADD "[$VAL_LBL_DOCS]", IIf((UBound(WS_DATA) = 4), "nella bolletta", "nelle bollette")
        
        For I = 0 To UBound(WS_DATA) Step 5
            WS_STRING = WS_STRING & "<chunk fontsize=""8.2""><![CDATA[ • Bolletta n. ]]></chunk>" & _
                                    "<chunk fontname=""helr65w.ttf"" fontsize=""8.2""><![CDATA[" & WS_DATA(I) & "]]></chunk>" & _
                                    IIf(WS_DATA(I + 1) = "0", "", "<chunk fontsize=""8.2""><![CDATA[ rata ]]></chunk><chunk fontname=""helr65w.ttf"" fontsize=""8.2""><![CDATA[" & WS_DATA(I + 1) & "]]></chunk>") & _
                                    "<chunk fontsize=""8.2""><![CDATA[ del ]]></chunk>" & _
                                    "<chunk fontname=""helr65w.ttf"" fontsize=""8.2""><![CDATA[" & WS_DATA(I + 2) & "]]></chunk>" & _
                                    "<chunk fontsize=""8.2""><![CDATA[ e il relativo addebito degli interessi nella bolletta n. ]]></chunk>" & _
                                    "<chunk fontname=""helr65w.ttf"" fontsize=""8.2""><![CDATA[" & WS_DATA(I + 3) & "]]></chunk>" & _
                                    "<chunk fontsize=""8.2""><![CDATA[ del ]]></chunk>" & _
                                    "<chunk fontname=""helr65w.ttf"" fontsize=""8.2""><![CDATA[" & WS_DATA(I + 4) & "<br>]]></chunk>"
        Next I
        
        DDS_ADD "[$WS_BILLS_DATA_ROWS]", WS_STRING
        
        ' TEMPLATE BUILDER
        '
        XML_TEMPLATE = GET_TEMPLATEINFO
    
        ' GET DATA
        '
        WS_ANNEXED_DATA = GET_FIELDSDATA(WS_ANNEXED_DATA)
    
        ' OUTPUT PLUGIN SUPPORT
        '
        XML_METADATA = GET_XMLMETADATA
    
        ' EMAIL MANAGEMENT
        '
        Select Case Trim$(WS_G01.R002.CANALEINOLTRO)
        Case "02"   ' EMAIL
            WS_FLG_NOMERGE = True
    
        Case "03"  ' STAMPA + EMAIL
            ' WS_FLG_NOMERGE = False
    
        Case Else
            WS_FLG_NOMERGE = (WS_G01.R007.RINUNCIA_COPIA_ANALOGICA = "S")
    
        End Select
    
        ' RECORD
        '
        WS_STRING = XML_TEMPLATE & "§" & _
                    WS_RECIPIENT & "§" & TXT_ADDRESS & "§" & WS_LOCALITY & "§" & _
                    WS_ANNEXED_DATA & "§§§§§§" & _
                    WS_PXX_FOOTER & "§" & _
                    WS_PAGENUM & "§§" & _
                    WS_FATTURANUMERO & "§" & _
                    WS_G01.R001.DATAEMISSIONE & "§" & _
                    WS_G06.R001.DATASCADENZA & "§" & _
                    Trim$(WS_G02.R001.CODICESERVIZIO) & "§§§§§§§§§§§§§§" & _
                    TXT_EMAIL & "§" & IIf((WS_FLG_FATTELE_PA Or WS_FLG_NOMERGE), "1", "") & "§" & _
                    "LET" & "§" & _
                    Trim$(WS_G01.R005.NOME_FE) & "_LM" & "§" & _
                    XML_METADATA
        
        ' INSERT DATA
        '
        MMS_Insert_L02 = mySQLImporter.SQLInsert(WS_STRING)
        WS_ERRMSG = mySQLImporter.GetUMErrorMessage
        WS_ERRSCT = "MMS_INSERT"
    Else
        MMS_Insert_L02 = True
    End If
    
    DoEvents

End Function

Public Function MMS_Insert_L03() As Boolean ' LPM
    
    WS_ERRSCT = "MMS_INSERT_L03"

    Dim WS_FLG_NOMERGE              As Boolean
    Dim WS_STRING                   As String

    Dim I                           As Integer
    Dim TXT_ADDRESS                 As String
    Dim TXT_EMAIL                   As String
    Dim TXT_INDOOR                  As String
    Dim XML_TEMPLATE                As String
    Dim XML_METADATA                As String

    ' INIT
    '
    WS_FATTURANUMERO = Mid$(Trim$(WS_G01.R005.NOME_FE), 5)
    WS_FLG_FATTELE_PA = (WS_G01.R001.TIPONUMERAZIONE = "5")
    WS_PAGENUM = ""
    WS_RECIPIENT = ""

    ' P01
    '
    WS_ERRSCT = "ERRORE REPERIMENTO DESTINATARIO - GRUPPO 03"
    
    With WS_G03
        For I = 0 To UBound(.R001)
            WS_RECIPIENT = WS_RECIPIENT & Trim$(.R001(I).NOME_RAGIONESOCIALE) & IIf((I = UBound(.R001)), "", "<br>")
        Next I
        
        TXT_INDOOR = Trim$(IIf(Trim$(.R002.INTERNO) = "", "", "INT. " & Trim$(.R002.INTERNO)) & IIf(Trim$(.R002.SCALA) = "", "", " SCALA " & Trim$(.R002.SCALA)) & IIf(Trim$(.R002.PIANO) = "", "", " PIANO: " & Trim$(.R002.PIANO)))
        TXT_ADDRESS = Trim$(UCase$(Trim$(IIf((Trim$(.R003.DESCRIZIONEESTESARECAPITO) = ""), .R003.INDIRIZZO, .R003.DESCRIZIONEESTESARECAPITO))))
        WS_LOCALITY = Trim$(UCase$(Trim$(IIf(Left$(.R004.LOCALITÀ, 5) = "00000", Mid$(.R004.LOCALITÀ, 7), .R004.LOCALITÀ))))
        
        If ((Trim$(.R004.SIGLA_NAZIONE) <> "") And (.R004.SIGLA_NAZIONE <> "IT")) Then
            If (Right$(WS_LOCALITY, 2) = "EE") Then WS_LOCALITY = Trim$(Left$(WS_LOCALITY, Len(WS_LOCALITY) - 2))
            
            WS_NATIONALITY = GET_DATA_CACHE(NATIONALITY, .R004.SIGLA_NAZIONE).dataDescription
            
            If (WS_NATIONALITY <> .R004.SIGLA_NAZIONE) Then WS_LOCALITY = WS_LOCALITY & "<br>" & WS_NATIONALITY
        End If
    End With

    DDS_INIT
    DDS_ADD "[$VAR_SYSDATE]", Format$(Now(), "dd/MM/yyyy")
    DDS_ADD "[$RECIPIENT]", WS_RECIPIENT
    DDS_ADD "[$PRESSO]", IIf((Trim$(WS_G03.R003.DESCRIZIONERECAPITO) = ""), "", "<br>" & Trim$(WS_G03.R003.DESCRIZIONERECAPITO))
    DDS_ADD "[$ADDRESS]", IIf(TXT_INDOOR = "", "", TXT_INDOOR & "<br>") & TXT_ADDRESS & "<br>" & WS_LOCALITY
    DDS_ADD "[$VAR_UBIC]", Trim$(WS_G05.R001.INDIRIZZOFORNITURA)
    DDS_ADD "[$VAR_CODSER]", Trim$(WS_G02.R001.CODICESERVIZIO)
    DDS_ADD "[$VAR_MATCON]", Trim$(WS_G09.R001.MATRICOLACONTATORE)
    DDS_ADD "[$VAR_DOCNUM]", WS_FATTURANUMERO
    DDS_ADD "[$VAR_DTAEMISSIONE]", WS_G01.R001.DATAEMISSIONE
    DDS_ADD "[$DTA_PRESCRIZIONE]", WS_GPB.R001.DATA_PRESCRIZIONE
    DDS_ADD "[$IMPORTO_PRESCRIZIONE]", NRM_IMPORT(WS_GPB.R001.POTENZIALE_PRESCRIZIONE, "##,##0.00", False)
    DDS_ADD "[$VAR_REASON]", GET_DATA_CACHE(DYN_REASON, WS_G00.AZIENDASIU & Val(WS_G01.R001.SEZIONALE) & WS_G01.R001.TIPONUMERAZIONE & Val(WS_G01.R001.CODSERVIZIONUMERAZIONE) & WS_G01.R001.ANNOBOLLETTA & Format$(WS_G01.R001.NUMEROBOLLETTA, "00000000")).dataDescription
    
    ' TEMPLATE BUILDER
    '
    XML_TEMPLATE = GET_TEMPLATEINFO

    ' GET DATA
    '
    WS_ANNEXED_DATA = GET_FIELDSDATA(WS_ANNEXED_DATA)

    ' OUTPUT PLUGIN SUPPORT
    '
    XML_METADATA = GET_XMLMETADATA

    ' EMAIL MANAGEMENT
    '
    Select Case Trim$(WS_G01.R002.CANALEINOLTRO)
    Case "02"   ' EMAIL
        WS_FLG_NOMERGE = True

    Case "03"  ' STAMPA + EMAIL
        ' WS_FLG_NOMERGE = False

    Case Else
        WS_FLG_NOMERGE = (WS_G01.R007.RINUNCIA_COPIA_ANALOGICA = "S")

    End Select

    ' RECORD
    '
    WS_STRING = XML_TEMPLATE & "§" & _
                WS_RECIPIENT & "§" & TXT_ADDRESS & "§" & WS_LOCALITY & "§" & _
                WS_ANNEXED_DATA & "§§§§§§" & _
                WS_PXX_FOOTER & "§" & _
                WS_PAGENUM & "§§" & _
                WS_FATTURANUMERO & "§" & _
                WS_G01.R001.DATAEMISSIONE & "§" & _
                WS_G06.R001.DATASCADENZA & "§" & _
                Trim$(WS_G02.R001.CODICESERVIZIO) & "§§§§§§§§§§§§§§" & _
                TXT_EMAIL & "§" & IIf((WS_FLG_FATTELE_PA Or WS_FLG_NOMERGE), "1", "") & "§" & _
                "LET" & "§" & _
                Trim$(WS_G01.R005.NOME_FE) & "_LPM" & "§" & _
                XML_METADATA
    
    ' INSERT DATA
    '
    MMS_Insert_L03 = mySQLImporter.SQLInsert(WS_STRING)
    WS_ERRMSG = mySQLImporter.GetUMErrorMessage
    WS_ERRSCT = "MMS_INSERT"
    
    DoEvents

End Function

Public Function MMS_Open() As Boolean

    Dim myData()    As String
    ReDim myData(0) As String
    
    myData(0) = "§"

    Set mySQLImporter = New SQL_Import.PlugIn

    With mySQLImporter
        .DSN = DLLParams.DSN
        .EXTRAPARAMS = myData
        .IDDATACUTTER = DLLParams.IDDATACUTTER
        .TABLENAME = DLLParams.TABLENAME
        .TNS = DLLParams.TNS
    
        MMS_Open = .StartJob

        If MMS_Open Then
            DLLParams.IDWORKINGLOAD = .GetIdWorkingLoad
        Else
            MsgBox .GetUMErrorMessage, vbExclamation, "Guru Meditation:"
        End If
    End With

End Function
