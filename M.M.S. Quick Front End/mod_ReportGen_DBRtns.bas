Attribute VB_Name = "mod_ReportGen_DBRtns"
Option Explicit

Private Type strct_Costi
    BillingTitle                As String
    CostoBusta                  As Single
    CostoFoglioAgg              As Single
    CostoSupportoOttico         As Single
    CostoSupportoOtticoSingolo  As Single
    MaxFogli                    As Byte
    TemplFieldFilter            As String
End Type

Private Const MaxRows           As Byte = 15

Private Function DB_GetInfoCosti(ByRef InfoCosti As strct_Costi) As Boolean

    On Error GoTo ErrHandler

    Dim RS              As ADODB.Recordset
    
    DBConn.Open

    Set RS = DBConn.Execute("SELECT * FROM edt_ProjectsBilling WHERE id_Project = " & frm_Main.SelectedPrj)
    
    If RS.RecordCount = 1 Then
        With InfoCosti
            .BillingTitle = RS("str_BillingTitle")
            .CostoBusta = RS("nmr_CostoBusta")
            .MaxFogli = RS("nmr_MaxFogli")
            
            If Not IsNull(RS("nmr_CostoFoglioAgg")) Then
                .CostoFoglioAgg = RS("nmr_CostoFoglioAgg")
            End If
            
            If Not IsNull(RS("nmr_CostoSupportoOttico")) Then
                .CostoSupportoOttico = RS("nmr_CostoSupportoOttico")
            End If
            
            If Not IsNull(RS("nmr_CostoSupportoOtticoSingolo")) Then
                .CostoSupportoOtticoSingolo = RS("nmr_CostoSupportoOtticoSingolo")
            End If
            
            If Not IsNull(RS("str_TemplFieldFilter")) Then
                .TemplFieldFilter = RS("str_TemplFieldFilter")
            End If
        End With
    
        DB_GetInfoCosti = True
    End If

    GoSub CleanUp

    Exit Function

CleanUp:
    Set RS = Nothing

    DBConn.Close
Return

ErrHandler:
    GoSub CleanUp
    
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Attenzione:"

End Function

Public Sub DB_PRNTPRVWReportGen(ByVal PrnPrvwMode As Byte)

    On Error GoTo ErrHandler

    Dim Costi           As strct_Costi
    
    If DB_GetInfoCosti(Costi) = False Then Exit Sub
    
    Dim Buste           As Long
    Dim CostoBuste      As Single
    Dim CostoDVD        As Single
    Dim CostoFogliAgg   As Single
    Dim CostoTotale     As Single
    Dim CurrentWorking  As String
    Dim Fogli           As Long
    Dim FogliAgg        As Long
    Dim I               As Byte
    Dim id_FirstSubPrj  As Long
    Dim J               As Byte
    Dim myCntr          As Byte
    Dim RS              As ADODB.Recordset
    Dim Sheets          As Byte
    Dim SQLString       As String
    Dim tmp_Buste       As Long
    Dim tmp_CostoBuste  As Single
    Dim tmp_CostoDVD    As Single
    Dim tmp_CostoFogli  As Single
    Dim tmp_CostoTotale As Single
    Dim tmp_Fogli       As Long
    Dim tmp_FogliAgg    As Long
    Dim tmp_Sheets      As Byte
    Dim YShift          As Single
    
    id_FirstSubPrj = DB_GetValueByID("SELECT id_SubProject FROM edt_SubProjects WHERE id_Project = " & frm_Main.SelectedPrj)
    
    If id_FirstSubPrj > 0 Then
        CurrentWorking = myMMS.GetCurrentWorking
        
        DBConn.Open
                    
        If Costi.TemplFieldFilter = "" Then
            SQLString = "SELECT COUNT(id_WorkCntr) AS Buste, id_WorkingLoad" & _
                        " FROM " & myMMS.GetCurrentRefTable & _
                        " GROUP BY id_WorkingLoad" & _
                        " HAVING id_WorkingLoad = " & CurrentWorking & _
                        " ORDER BY id_WorkingLoad"
        Else
            SQLString = "SELECT COUNT(id_WorkCntr) AS Buste, id_WorkingLoad, " & Costi.TemplFieldFilter & _
                        " FROM " & myMMS.GetCurrentRefTable & _
                        " GROUP BY id_WorkingLoad, " & Costi.TemplFieldFilter & _
                        " HAVING id_WorkingLoad = " & CurrentWorking & _
                        " ORDER BY id_WorkingLoad, " & Costi.TemplFieldFilter
        End If
        
        Set RS = DBConn.Execute(SQLString)
    
        If RS.RecordCount > 0 Then
            If PrnPrvwMode = 0 Then
                frm_SDViewer.Show vbModeless, frm_Main
                frm_SDViewer.Tag = "PMode"
            End If
            
            With mySDF
                .AddPage
            
                If PrnPrvw_LayOut = False Then
                    GoSub CleanUp
                
                    Exit Sub
                End If
                
                YShift = 3.5

                With .PrintText
                    .Caption Costi.BillingTitle
                    .Font "Tahoma", 13, True
                    .Justify CenterCenter
                    .Move 0, YShift, mySDF.PrintWidth
                    .DrawObject
                
                    YShift = YShift + 2.5
            
                    .Caption " Dettaglio:"
                    .Font "Arial", 10, True, True, , RGB(128, 0, 0)
                    .Justify LeftCenter
                    .Move 0.1, YShift, mySDF.PrintWidth - 0.2
                    .BorderBottom
                    .DrawObject
                
                    YShift = YShift + 0.6
                End With
                
                With .PrintGrid
                    .GridLeft = mySDF.PrintMarginLeft + 0.25
                    .GridTop = YShift
                
                    .NumCols = 8
        
                    .ColWidth(1) = 2.5
                    .ColWidth(2) = 1.3
                    .ColWidth(3) = 2.5
                    .ColWidth(4) = 2.5
                    .ColWidth(5) = 2.5
                    .ColWidth(6) = 2.5
                    .ColWidth(7) = 2.5
                    .ColWidth(8) = 2.5
        
                    ' Header
                    '
                    For I = 1 To 8
                        .ColFont I, "Tahoma", 9, True, , , , RGB(249, 241, 207)
                        .ColJustify I, CenterCenter
                        .CellBorder I, , 2
                    Next I
        
                    .ColCaption 1, "Buste"
                    .ColCaption 2, "Fogli" & vbNewLine & "Templ.", True
                    .ColCaption 3, "Totale" & vbNewLine & "Fogli", True
                    .ColCaption 4, "Fogli" & vbNewLine & "Aggiuntivi", True
                    .ColCaption 5, "Importo" & vbNewLine & "Buste", True
                    .ColCaption 6, "Importo" & vbNewLine & "Fogli Agg.", True
                    .ColCaption 7, "Importo" & vbNewLine & "Supporto", True
                    .ColCaption 8, "Importo" & vbNewLine & "Totale", True
                    
                    .RowHeight = 1
                    .RowDraw
        
                    For I = 1 To 8
                        .ColFont I, "Tahoma", 9
                        .ColJustify I, CenterCenter
                        
                        Select Case I
                            Case 1
                                .CellBorderLeft I, , 2
                            
                            Case 8
                                .CellBorderLeft I, , 1
                                .CellBorderRight I, , 2
                            
                            Case Else
                                .CellBorderLeft I, , 1
                        
                        End Select
                        
                        .CellBorderBottom I, vbDot, 1
                    Next I
                    
                    .RowHeight = 0.5
                
                    For I = 1 To 8
                        .ColJustify I, RightCenter
                    Next I
                    
                    ' Rows
                    '
                    For I = 1 To MaxRows
                        CostoFogliAgg = 0
                        
                        If RS.EOF Then
                            Sheets = 0
                        Else
                            If Costi.TemplFieldFilter = "" Then
                                SQLString = ")"
                            Else
                                SQLString = " AND str_QValue = " & RS(Costi.TemplFieldFilter) & ")"
                            End If
                            
                            Sheets = DB_GetValueByID("SELECT nmr_Sheets" & _
                                                     " FROM ref_Templates" & _
                                                     " INNER JOIN edt_Templates ON ref_Templates.id_Template = edt_Templates.id_Template" & _
                                                     " WHERE (id_SubProject = " & id_FirstSubPrj & SQLString)
                        End If
                        
                        If (tmp_Sheets <> Sheets) Then
                            If I > 1 Then
                                myCntr = myCntr + 1
                                
                                CostoBuste = Buste * Costi.CostoBusta
                                
                                If Costi.CostoSupportoOtticoSingolo = 0 Then CostoDVD = Buste * Costi.CostoSupportoOttico
                                
                                CostoTotale = CostoBuste + CostoDVD
                                
                                Fogli = Buste * tmp_Sheets
                                
                                If (tmp_Sheets > Costi.MaxFogli) And (Costi.CostoFoglioAgg > 0) Then
                                    FogliAgg = Fogli - Buste
                                    CostoFogliAgg = FogliAgg * Costi.CostoFoglioAgg
                                    CostoTotale = CostoTotale + CostoFogliAgg
                                End If
                                
                                .ColCaption 1, Format$(Buste, "##,##") & " "
                                .ColCaption 2, tmp_Sheets & " "
                                .ColCaption 3, Format$(Fogli, "##,##") & " "
                                .ColCaption 4, IIf(FogliAgg > 0, Format$(FogliAgg, "##,##"), "0") & " "
                                .ColCaption 5, Format$(CostoBuste, "##,##0.00") & " "
                                .ColCaption 6, Format$(CostoFogliAgg, "##,##0.00") & " "
                                
                                If Costi.CostoSupportoOtticoSingolo = 0 Then
                                    .ColCaption 7, Format$(CostoDVD, "##,##0.00") & " "
                                Else
                                    .ColCaption 7, " - "
                                End If
                                
                                .ColCaption 8, Format$(CostoTotale, "##,##0.00") & " "
                                .RowDraw
                            
                                tmp_Buste = tmp_Buste + Buste
                                tmp_CostoBuste = tmp_CostoBuste + CostoBuste
                                tmp_CostoDVD = tmp_CostoDVD + CostoDVD
                                tmp_CostoFogli = tmp_CostoFogli + CostoFogliAgg
                                tmp_CostoTotale = tmp_CostoTotale + CostoTotale
                                tmp_Fogli = tmp_Fogli + Fogli
                                tmp_FogliAgg = tmp_FogliAgg + FogliAgg
                            End If
                            
                            If RS.EOF = False Then
                                Buste = RS("Buste")
                                tmp_Sheets = Sheets
                            End If
                        Else
                            Buste = Buste + RS("Buste")
                        End If
                        
                        If RS.EOF Then Exit For
                        
                        RS.MoveNext
                    Next I
                    
                    myCntr = myCntr + 1
                    
                    If Costi.CostoSupportoOtticoSingolo > 0 Then
                        tmp_CostoDVD = Costi.CostoSupportoOtticoSingolo
                        tmp_CostoTotale = tmp_CostoTotale + tmp_CostoDVD
                    End If
                    
                    For I = myCntr To MaxRows
                        If I = 15 Then
                            For J = 1 To 8
                                .CellBorderBottom J, , 2
                            Next J
                        End If
            
                        .RowDraw
                    Next I
                    
                    ' Totali
                    '
                    For I = 1 To 8
                        .ColFont I, "Tahoma", 9, True, , , , RGB(246, 234, 176)
                        .CellBorder I, , 2
                    Next I
        
                    .ColCaption 1, Format$(tmp_Buste, "##,##") & " "
                    .ColCaption 2, "-"
                    .ColJustify 2, CenterCenter
                    .ColCaption 3, Format$(tmp_Fogli, "##,##") & " "
                    .ColCaption 4, IIf(tmp_FogliAgg > 0, Format$(tmp_FogliAgg, "##,##"), "0") & " "
                    .ColCaption 5, "€ " & Format$(tmp_CostoBuste, "##,##0.00") & " "
                    .ColCaption 6, "€ " & Format$(tmp_CostoFogli, "##,##0.00") & " "
                    .ColCaption 7, "€ " & Format$(tmp_CostoDVD, "##,##0.00") & " "
                    .ColCaption 8, "€ " & Format$(tmp_CostoTotale, "##,##0.00") & " "
                    
                    .RowHeight = 0.8
                    .RowDraw
                    
                    YShift = .GridHeight + 0.2
                End With
                
                With .PrintText
                    YShift = YShift + 0.6
                
                    .Caption " Info Flusso Dati:"
                    .Font "Arial", 10, True, True, , RGB(128, 0, 0)
                    .Justify LeftCenter
                    .Move 0.1, YShift, mySDF.PrintWidth - 0.2
                    .BorderBottom
                    .DrawObject
                
                    YShift = YShift + 0.5
                
                    .Caption "Flusso Num.:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption CurrentWorking
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                
                    YShift = YShift + 0.5
                    
                    .Caption "del:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption Mid$(CurrentWorking, 7, 2) & "/" & Mid$(CurrentWorking, 5, 2) & "/" & Mid$(CurrentWorking, 1, 4)
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                    
                    YShift = YShift + 0.5
                    
                    .Caption "Tipologia:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption myMMS.ProjectName
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                
                    YShift = YShift + 1
                     
                    ' Riassunto Costi
                    '
                    .Caption " Riassunto Costi:"
                    .Font "Arial", 10, True, True, , RGB(128, 0, 0)
                    .Justify LeftCenter
                    .Move 0.1, YShift, mySDF.PrintWidth - 0.2
                    .BorderBottom
                    .DrawObject
                    
                    YShift = YShift + 0.5
                    
                    .Caption "Num. Buste:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption Format$(tmp_Buste, "##,##")
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                
                    YShift = YShift + 0.5
                
                    .Caption "Num. Fogli:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption Format$(tmp_Fogli, "##,##")
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                
                    YShift = YShift + 0.5
                
                    .Caption "Num. Fogli Agg.:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption IIf(tmp_FogliAgg > 0, Format$(tmp_FogliAgg, "##,##"), "0")
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                    
                    YShift = YShift + 0.5
                    
                    .Caption "Imp. Tot. Buste/Fogli:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption "€ " & Format$((tmp_CostoBuste + tmp_CostoFogli), "##,##0.00")
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                
                    YShift = YShift + 0.5
                    
                    .Caption "Imp. Supporto Ottico:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption "€ " & Format$(tmp_CostoDVD, "##,##0.00")
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                
                    YShift = YShift + 0.5
                
                    .Caption "Importo Totale:"
                    .Font "Arial", 9, , True
                    .Justify RightCenter
                    .Move 0.2, YShift, 3.5
                    .DrawObject
                
                    .Caption "€ " & Format$(tmp_CostoTotale, "##,##0.00")
                    .Font "Arial", 9, True, , , , RGB(240, 240, 220)
                    .Justify LeftCenter
                    .Move 3.8, YShift, mySDF.PrintWidth - 2.9, , 0.02
                    .DrawObject
                End With
                
                .EndDoc
            End With
            
            If (PrnPrvwMode = 0) Then Get_NumPages
        Else
            MsgBox "Nessun record trovato per la lavorazione corrente.", vbExclamation, "Attenzione:"
        End If
        
        GoSub CleanUp
    End If
    
    Exit Sub

CleanUp:
    Set RS = Nothing

    DBConn.Close
Return

ErrHandler:
    GoSub CleanUp

    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Massive Report:"

End Sub

Private Function PrnPrvw_LayOut() As Boolean

    PrnPrvw_LayOut = True

End Function


