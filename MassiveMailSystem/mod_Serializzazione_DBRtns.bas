Attribute VB_Name = "mod_Serializzazione_DBRtns"
Option Explicit
    
Private Type strct_FreeCodes
    RangeEnd    As Double
    RangeId     As Double
    RangeStart  As Double
End Type

Private Function DB_BarCodeSerialize() As Boolean

    On Error GoTo ErrHandler
    
    Dim ErrMsg      As String
    Dim FreeCodes() As strct_FreeCodes
    Dim CodesCntr   As Integer
    Dim RangeCntr   As Double
    Dim retValue    As Boolean
    Dim RS          As ADODB.Recordset
    
    ' Get Free Codes
    '
    Set RS = DBConn.Execute("SELECT NMR_RANGESTART, NMR_RANGEEND, NMR_LASTRANGEID FROM REF_" & ProjectInfo.BARCODETABLE)

    If RS.RecordCount > 0 Then
        ReDim FreeCodes(RS.RecordCount - 1)
                
        Do Until RS.EOF
            With FreeCodes(RS.AbsolutePosition - 1)
                .RangeStart = RS("NMR_RANGESTART")
                
                If IsNull(RS("NMR_LASTRANGEID")) Then
                    .RangeId = RS("NMR_RANGESTART")
                Else
                    .RangeId = RS("NMR_LASTRANGEID")
                End If
                
                .RangeEnd = RS("NMR_RANGEEND")
            End With

            RS.MoveNext
        Loop
    End If
    
    RS.Close
    
    ' Start Process
    '
    DBConn.Execute "UPDATE " & ProjectInfo.REF_TABLE & " SET NMR_" & ProjectInfo.BARCODETYPE & "CODE = NULL WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING, retValue
    
    If retValue Then
        RS.Open "SELECT ID_WORKCNTR, NMR_" & ProjectInfo.BARCODETYPE & "CODE FROM " & ProjectInfo.REF_TABLE & _
                " WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING & _
                " ORDER BY ID_PACCO, ID_POSIZIONE", _
                DBConn, adOpenForwardOnly, adLockOptimistic
        
        If RS.RecordCount > 0 Then
            Set myAPB = New cls_APB
            
            With myAPB
                .APBMode = PBSingle
                .APBCaption = "BarCode Serializer:"
                .APBMaxItems = RS.RecordCount
                .tmrMode = Total
                .APBShow
            End With
            
            RangeCntr = FreeCodes(0).RangeId + 1
            
            Do Until RS.EOF
                myAPB.APBItemsLabel = "Item: " & Format$(RS.AbsolutePosition, "00000") & " of " & Format$(RS.RecordCount, "00000")
                myAPB.APBItemsProgress = RS.AbsolutePosition
                
                RS("NMR_" & ProjectInfo.BARCODETYPE & "CODE") = RangeCntr
                
                If RangeCntr = FreeCodes(CodesCntr).RangeEnd Then
                    DBConn.Execute "UPDATE REF_" & ProjectInfo.BARCODETABLE & " SET NMR_LASTRANGEID = " & RangeCntr & " WHERE NMR_RANGESTART = " & FreeCodes(CodesCntr).RangeStart, retValue
                    
                    If retValue Then
                        CodesCntr = CodesCntr + 1
                        RangeCntr = FreeCodes(CodesCntr).RangeId
                    Else
                        ErrMsg = "Errore durante l'aggiornamento della tabella ref_" & ProjectInfo.BARCODETYPE & "Range."
                        
                        GoTo ErrHandler
                    End If
                Else
                    RangeCntr = RangeCntr + 1
                End If
                
                RS.MoveNext
            Loop
            
            RS.Close
            
            DBConn.Execute "UPDATE REF_" & ProjectInfo.BARCODETABLE & " SET NMR_LASTRANGEID = " & (RangeCntr - 1) & " WHERE NMR_RANGESTART = " & FreeCodes(CodesCntr).RangeStart, retValue
            
            If retValue = False Then
                ErrMsg = "Errore durante l'aggiornamento della tabella ref_" & ProjectInfo.BARCODETYPE & "Range."
                
                GoTo ErrHandler
            End If

            myAPB.APBClose
        Else
            ErrMsg = "Nessun record trovato per la serializzazione."
            
            GoTo ErrHandler
        End If
    Else
        ErrMsg = "Errore durante lo svuotamento dei campi per la tabella " & ProjectInfo.REF_TABLE
        
        GoTo ErrHandler
    End If
    
    GoSub CleanUp

    DB_BarCodeSerialize = True

    Exit Function

CleanUp:
    Set myAPB = Nothing
    
    Set RS = Nothing
Return
    
ErrHandler:
    DB_BarCodeSerialize = False
    
    GoSub CleanUp
    
    MsgBox IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg), vbExclamation, "BarCode Serializer:"

End Function

Public Function DB_PackagesLabels(ByVal idSubProject As Integer) As Boolean

    On Error GoTo ErrHandler
    
    Dim ErrMsg          As String
    
    If idSubProject > -1 Then
        If DB_SubProjectInfo_SELECT(idSubProject) = False Then
            ErrMsg = "Errore durante la lettura dei parametri del sottoprogetto."
    
            GoTo ShowErrMsg
        End If
    End If
    
    With WorkPaths
        If FDExist(.WORKINGDIR, True) = False Then
            ErrMsg = "Directory di lavorazione inesistente."
            
            GoTo ShowErrMsg
        End If
        
        chk_Directory .REPORTSDIR, 3
        chk_Directory .TEMPORARYDIR, 3
    End With
    
    Dim RS              As ADODB.Recordset
    
    DBConn.Open
    
    With ProjectInfo
        Set RS = DBConn.Execute("SELECT MAX(id_Pacco) AS numPacchi FROM " & .REF_TABLE & " WHERE id_WorkingLoad = '" & ProjectInfo.IDWORKING & "'")
        
        If RS.RecordCount = 1 Then
            Dim CollEnd             As Long
            Dim CollStart           As Long
            Dim DataWork            As String
            Dim I                   As Integer
            Dim J                   As Byte
            Dim myXFDF              As cls_GenXFDF
            Dim NumPacchi           As Integer
            Dim NumBuste            As Integer
            Dim ScaglionePeso       As Byte
            Dim SplitData()         As String
            Dim tmpBacino           As String
            Dim tmpCAP              As String
            Dim tmpDest             As String
            Dim tmpInt              As Integer
            Dim tmpPeso()           As String
            Dim tmpString           As String
            Dim tmpTariffa          As String
            Dim XFDF_FName          As String
            
            Set myXFDF = New cls_GenXFDF
            
            DataWork = Format$(Now, "dd/MM/yyyy")
            ScaglionePeso = Calc_ScaglionePeso()
            NumPacchi = RS("numPacchi")
            
            Set myAPB = New cls_APB
            
            With myAPB
                .APBMode = PBSingle
                .APBCaption = "Labels Maker:"
                .APBMaxItems = NumPacchi
                .tmrMode = Total
                .APBShow
            End With
            
            For I = 1 To NumPacchi
                myAPB.APBItemsLabel = "Label " & I & " of " & NumPacchi
                myAPB.APBItemsProgress = I
                
                NumBuste = 0
                tmpCAP = ""
                tmpInt = 0
                tmpString = ""
                
                Set RS = DBConn.Execute("SELECT " & .REF_TABLE & ".id_Pacco, " & .REF_TABLE & ".flg_Mix, COUNT(*) AS nmr_Buste, view_Bacini.descr_Bacino, view_Bacini.id_Tariffa, view_Bacini.descr_Provincia, view_Bacini.nmr_RangeCAP_From FROM " & .REF_TABLE & _
                                        " INNER JOIN view_Bacini ON " & .REF_TABLE & ".id_Provincia = view_Bacini.id_Provincia" & _
                                        " WHERE id_WorkingLoad = '" & .IDWORKING & "' AND " & .REF_TABLE & ".id_Pacco = " & I & _
                                        " GROUP BY " & .REF_TABLE & ".id_Pacco, " & .REF_TABLE & ".flg_Mix, view_Bacini.descr_Bacino, view_Bacini.id_Tariffa, view_Bacini.descr_Provincia, view_Bacini.nmr_RangeCAP_From" & _
                                        " ORDER BY " & .REF_TABLE & ".id_Pacco, view_Bacini.descr_Bacino, view_Bacini.id_Tariffa, view_Bacini.nmr_RangeCAP_From")
                
                If RS.RecordCount > 0 Then
                    If (RS("flg_Mix") = 2) And tmpBacino <> "MIX" Then
                        tmpBacino = "MIX"
                        tmpTariffa = "MIX"
                        tmpDest = "MIX"
                    End If
                                    
                    If RS.RecordCount = 1 Then
                        NumBuste = RS("nmr_Buste")
                        Push_Array tmpPeso, IIf(RS("id_Tariffa") = "ZZ", "EU", RS("id_Tariffa")) & "|" & RS("nmr_Buste")
                        
                        If tmpBacino <> "MIX" Then
                            tmpBacino = RS("descr_Bacino")
                            tmpCAP = RS("nmr_RangeCAP_From")
                            
                            If Mid$(tmpCAP, 3, 1) = "0" Then
                                tmpDest = RS("descr_Provincia") & " Provincia"
                            Else
                                tmpDest = RS("descr_Provincia") & " Citta'"
                            End If
                            
                            tmpTariffa = RS("id_Tariffa")
                        End If
                    Else
                        Do Until RS.EOF
                            NumBuste = NumBuste + RS("nmr_Buste")
                            
                            Push_Array tmpPeso, IIf(RS("id_Tariffa") = "ZZ", "EU", RS("id_Tariffa")) & "|" & RS("nmr_Buste")
                            
                            If tmpBacino <> "MIX" Then
                                If RS.AbsolutePosition = 1 Then
                                    tmpBacino = RS("descr_Bacino")
                                    tmpTariffa = RS("id_Tariffa")
                                Else
                                    If RS("id_Tariffa") <> tmpTariffa Then
                                        tmpTariffa = "MIX"
                                    End If
                                    
                                    tmpDest = "MIX"
                                End If
                            End If
                            
                            RS.MoveNext
                        Loop
                    End If
                
                    CollEnd = CollEnd + NumBuste
                    CollStart = CollEnd - NumBuste + 1
                    
                    XFDF_FName = WorkPaths.TEMPORARYDIR & ProjectInfo.PSTLIDJOB & "_S" & Format(I, "000") & ".xfdf"
                    
                    With myXFDF
                        .XFDF_Open XFDF_FName, AppPath & "Reports\Posta Massiva\Etichetta.PDF", ""
                    
                        .XFDF_FieldText "txt_DataImpostazione", DataWork, False
                    
                        If ProjectInfo.PSTLIDHOMOLOGATION <> "" Then
                            .XFDF_FieldText "txt_CodiceOmologazione", ProjectInfo.PSTLIDHOMOLOGATION, False
                            .XFDF_FieldText "txt_Omologazione", "SI", False
                        Else
                            .XFDF_FieldText "txt_Omologazione", "NO", False
                        End If
                        
                        .XFDF_FieldText "txt_Formato", ProjectInfo.PSTLENVTYPE, False
                    
                        If NumBuste = ProjectInfo.NUMBUSTEMAX Then
                            .XFDF_FieldText "txt_InfoScatolaPiena", "SI", False
                        Else
                            .XFDF_FieldText "txt_InfoScatolaPiena", "NO", False
                        End If
                        
                        .XFDF_FieldText "txt_Bacino", tmpBacino, False
                        .XFDF_FieldText "txt_Tariffa", tmpTariffa, False
                        .XFDF_FieldText "txt_Destinazione", tmpDest, False
                        .XFDF_FieldText "txt_CAP", tmpCAP, False
                        
                        If UBound(tmpPeso) > 0 Then Sort_Quick tmpPeso
                        
                        For J = 0 To UBound(tmpPeso)
                            SplitData = Split(tmpPeso(J), "|")
                            
                            If tmpString = SplitData(0) Then
                                tmpInt = tmpInt + SplitData(1)
                            Else
                                If tmpString <> "" Then
                                    .XFDF_FieldText "txt_" & tmpString & Format$(ScaglionePeso, "00"), tmpInt, False
                                End If
                                
                                tmpInt = SplitData(1)
                                tmpString = SplitData(0)
                            End If
                        Next J
                        
                        .XFDF_FieldText "txt_" & tmpString & Format$(ScaglionePeso, "00"), tmpInt, False
                        
                        .XFDF_FieldText "txt_JobId", ProjectInfo.PSTLIDJOB, False
                        .XFDF_FieldText "txt_PackFrom", I, False
                        .XFDF_FieldText "txt_PackTo", NumPacchi, False
                        
                        .XFDF_FieldText "txt_CollezioneFrom", CollStart, False
                        .XFDF_FieldText "txt_CollezioneTo", CollEnd, False
                        
                        .XFDF_FieldText "txt_NumDocs", NumBuste, False
                        
                        .XFDF_Close
                    End With
                    
                    Erase tmpPeso
                        
                    If XFDF_Merger(XFDF_FName, False, WorkPaths.TEMPORARYDIR & "\Etichette_P" & Format$(I, "000") & ".PDF") = False Then
                    'If XFDF_Merger(AppPath & "Reports\Posta Massiva\Etichetta.PDF", XFDF_FName, False, WorkPaths.TemporaryDir & "\Etichette_P" & Format$(I, "000") & ".PDF") = False Then
                        ErrMsg = "Errore durante la generazione di: " & WorkPaths.TEMPORARYDIR & "\Etichette_P" & Format$(I, "000") & ".PDF"

                        GoTo ErrHandler
                    End If
                Else
                    ErrMsg = "Nessun record trovato per il pacco numero " & I

                    GoTo ErrHandler
                End If
            Next I
            
            myAPB.APBItemsLabel = "Terminating... Wait Please."
            
            If PDF_DirMerger(WorkPaths.TEMPORARYDIR, WorkPaths.REPORTSDIR & "Etichette.PDF") = False Then
                ErrMsg = "Errore durante la generazione di: " & WorkPaths.REPORTSDIR & "Etichette.PDF"

                GoTo ErrHandler
            End If
            
            RS.Close
                        
            GoSub CleanUp
        Else
            '
        End If
    End With
        
    chk_Directory WorkPaths.TEMPORARYDIR, 1

    DB_PackagesLabels = True
    
    Exit Function
    
CleanUp:
    myAPB.APBClose
    
    Set myAPB = Nothing
    
    Set RS = Nothing
        
    DBConn.Close
    
    Erase tmpPeso
    Erase SplitData
Return

ShowErrMsg:
    MsgBox IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg), vbExclamation, "Labels Maker:"
    
    Exit Function

ErrHandler:
    GoSub CleanUp
    
    MsgBox IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg), vbExclamation, "Labels Maker:"
    
End Function

Public Function DB_PackagesManagement() As Boolean

    Dim ErrMsg      As String
    Dim retValue    As Boolean
    
    ' Serializzazione
    '
    DBConn.Open
    DBConn.BeginTrans
    
    Select Case ProjectInfo.IDSERIALIZEMODE
        Case 0  ' Standard
            retValue = DB_PackagesSerialize_STD
                
            If retValue = False Then GoTo ErrHandler
        
        Case 1  ' Postalizzazione
            DBConn.Execute "UPDATE " & ProjectInfo.REF_TABLE & " SET ID_PROVINCIA = NULL, ID_PACCO = NULL, ID_POSIZIONE = NULL WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING, retValue
            
            If retValue Then
                Set myAPB = New cls_APB
                
                With myAPB
                    .APBMode = PBDouble
                    .APBCaption = "Postalizzazione:"
                    .APBMaxItems = 3
                    .tmrMode = Total
                    .APBShow
                End With
                
                retValue = DB_PackagesSortKey
                
                If retValue = False Then GoTo ErrHandler
                        
                retValue = DB_PackagesSerialize_PSTL
            
                If retValue = False Then GoTo ErrHandler
                
                myAPB.APBClose
                
                Set myAPB = Nothing
            Else
                ErrMsg = "Errore durante lo svuotamento dei campi per la tabella " & ProjectInfo.REF_TABLE
                
                GoTo ErrHandler
            End If
    
    End Select
    
    If ProjectInfo.BARCODETYPE <> "" Then
        retValue = DB_BarCodeSerialize
    
        If retValue = False Then GoTo ErrHandler
    End If
    
    DBConn.CommitTrans
    DBConn.Close
    
    DB_PackagesManagement = True
    
    Exit Function

ErrHandler:
    DB_PackagesManagement = False
    
    DBConn.RollbackTrans
    DBConn.Close
    
    If ProjectInfo.IDSERIALIZEMODE = 1 Then
        myAPB.APBClose
        
        Set myAPB = Nothing
    End If
    
    If ErrMsg <> "" Then MsgBox ErrMsg, vbExclamation, "Serializzazione:"

End Function

Private Function DB_PackagesSerialize_PSTL() As Boolean
    
    On Error GoTo ErrHandler
    
    Dim ErrMsg          As String
    Dim ExtraEnvelopes  As Integer
    Dim I               As Integer
    Dim idPacco         As Integer
    Dim NumPacchi       As Integer
    Dim RS              As ADODB.Recordset
    Dim RS_SP           As ADODB.Recordset
    
    With ProjectInfo
        ' Controllo Bacini (B.T.C. Packaging Builder)
        '
        Set RS = DBConn.Execute("SELECT COUNT(*) AS nmr_Buste, view_Bacini.id_Bacino FROM " & .REF_TABLE & _
                                " INNER JOIN dbo.view_Bacini ON " & .REF_TABLE & ".id_Provincia = view_Bacini.id_Provincia" & _
                                " WHERE (view_Bacini.id_Tariffa <> 'ZZ') AND id_WorkingLoad = '" & .IDWORKING & "'" & _
                                " GROUP BY view_Bacini.descr_Bacino, view_Bacini.id_Bacino" & _
                                " ORDER BY view_Bacini.descr_Bacino ASC")
        
        If RS.RecordCount > 0 Then
            Do Until RS.EOF
                ExtraEnvelopes = 0
                
                If RS("nmr_Buste") >= .NUMBUSTEMIN Then
                    NumPacchi = Fix(RS("nmr_Buste") / .NUMBUSTEMAX)
                    
                    If NumPacchi = 0 Then
                        NumPacchi = 1
                    Else
                        ExtraEnvelopes = (RS("nmr_Buste") Mod .NUMBUSTEMAX)
                        
                        If ExtraEnvelopes >= .NUMBUSTEMIN Then
                            NumPacchi = NumPacchi + 1
                        Else
                            ExtraEnvelopes = 0
                        End If
                    End If
                    
                    Set RS_SP = New ADODB.Recordset
                    
                    For I = 1 To NumPacchi
                        idPacco = idPacco + 1
                        
                        myAPB.APBItemsLabel = "Creating Pack " & Format$(idPacco, "000") & " - Std Packages"
                        myAPB.APBItemsProgress = 2
                        
                        ' Serializzazione
                        '
                        RS_SP.Open "SELECT TOP " & IIf((I = NumPacchi) And (ExtraEnvelopes > 0), ExtraEnvelopes, .NUMBUSTEMAX) & " " & .REF_TABLE & ".id_WorkCntr, " & .REF_TABLE & ".id_Pacco, " & .REF_TABLE & ".id_Posizione FROM " & .REF_TABLE & _
                                   " INNER JOIN view_Bacini ON " & .REF_TABLE & ".id_Provincia = view_Bacini.id_Provincia" & _
                                   " WHERE id_Bacino = " & RS("Id_Bacino") & " AND id_Pacco IS NULL  AND id_WorkingLoad = '" & .IDWORKING & "'" & _
                                   " ORDER BY view_Bacini.descr_Bacino, view_Bacini.id_Tariffa, " & .REF_TABLE & "." & .PSTLFIELD & " ASC" & IIf(.PSTLORDERBY = "", "", ", " & .PSTLORDERBY), _
                                   DBConn, adOpenDynamic, adLockOptimistic
                        
                        If RS_SP.RecordCount > 0 Then
                            myAPB.APBMaxItem = RS_SP.RecordCount
                            
                            Do Until RS_SP.EOF
                                myAPB.APBItemLabel = "Item " & RS_SP.AbsolutePosition & " of " & RS_SP.RecordCount
                                myAPB.APBItemProgress = RS_SP.AbsolutePosition
                                
                                RS_SP("id_Pacco") = idPacco
                                RS_SP("id_Posizione") = RS_SP.AbsolutePosition
                                
                                RS_SP.MoveNext
                            Loop
                        End If
                        
                        RS_SP.Close
                    Next I
                    
                    Set RS_SP = Nothing
                End If
                
                RS.MoveNext
            Loop
            
            RS.Close
            
            ' Controllo Bacino Mix (Mix Packaging Builder)
            '
            Set RS = DBConn.Execute("SELECT COUNT(*) AS nmr_Buste, view_Bacini.id_Bacino FROM " & .REF_TABLE & _
                                    " INNER JOIN dbo.view_Bacini ON " & .REF_TABLE & ".id_Provincia = view_Bacini.id_Provincia" & _
                                    " WHERE id_Pacco IS NULL AND (view_Bacini.id_Tariffa <> 'ZZ') AND id_WorkingLoad = '" & .IDWORKING & "'" & _
                                    " GROUP BY view_Bacini.descr_Bacino, view_Bacini.id_Bacino" & _
                                    " ORDER BY view_Bacini.descr_Bacino ASC")

            If RS.RecordCount > 0 Then
                Do Until RS.EOF
                    ExtraEnvelopes = 0
                    
                    If RS("nmr_Buste") >= .NUMBUSTEMIXMIN Then
                        NumPacchi = Fix(RS("nmr_Buste") / .NUMBUSTEMAX)
                        
                        If NumPacchi = 0 Then
                            NumPacchi = 1
                        Else
                            ExtraEnvelopes = (RS("nmr_Buste") Mod .NUMBUSTEMAX)
                            
                            If ExtraEnvelopes >= .NUMBUSTEMIXMIN Then
                                NumPacchi = NumPacchi + 1
                            Else
                                ExtraEnvelopes = 0
                            End If
                        End If
                    
                        Set RS_SP = New ADODB.Recordset
                        
                        For I = 1 To NumPacchi
                            idPacco = idPacco + 1
                            
                            myAPB.APBItemsLabel = "Creating Pack " & Format$(idPacco, "000") & " - Mix Packages"
                            myAPB.APBItemsProgress = 2
                        
                            ' Serializzazione
                            '
                            RS_SP.Open "SELECT TOP " & IIf((I = NumPacchi) And (ExtraEnvelopes > 0), ExtraEnvelopes, .NUMBUSTEMAX) & " " & .REF_TABLE & ".id_WorkCntr, " & .REF_TABLE & ".id_Pacco, " & .REF_TABLE & ".id_Posizione, " & .REF_TABLE & ".flg_Mix FROM " & .REF_TABLE & _
                                       " INNER JOIN view_Bacini ON " & .REF_TABLE & ".id_Provincia = view_Bacini.id_Provincia" & _
                                       " WHERE id_Bacino = " & RS("Id_Bacino") & " AND id_Pacco IS NULL  AND id_WorkingLoad = '" & .IDWORKING & "'" & _
                                       " ORDER BY view_Bacini.descr_Bacino, view_Bacini.id_Tariffa, " & .REF_TABLE & "." & .PSTLFIELD & " ASC" & IIf(.PSTLORDERBY = "", "", ", " & .PSTLORDERBY), _
                                       DBConn, adOpenDynamic, adLockOptimistic
                        
                            If RS_SP.RecordCount > 0 Then
                                myAPB.APBMaxItem = RS_SP.RecordCount
                                
                                Do Until RS_SP.EOF
                                    myAPB.APBItemLabel = "Item " & RS_SP.AbsolutePosition & " of " & RS_SP.RecordCount
                                    myAPB.APBItemProgress = RS_SP.AbsolutePosition
                                    
                                    RS_SP("id_Pacco") = idPacco
                                    RS_SP("id_Posizione") = RS_SP.AbsolutePosition
                                    RS_SP("flg_Mix") = 1
                                    
                                    RS_SP.MoveNext
                                Loop
                            End If
                            
                            RS_SP.Close
                        Next I
                                        
                        Set RS_SP = Nothing
                    End If
                    
                    RS.MoveNext
                Loop
            End If
            
            RS.Close
            
            ' Controllo Bacino Mix (Mix Packaging Builder)
            '
            Set RS = DBConn.Execute("SELECT COUNT(*) AS nmr_Buste FROM " & ProjectInfo.REF_TABLE & " WHERE id_Pacco IS NULL AND id_WorkingLoad = '" & .IDWORKING & "'")
            
            If RS.RecordCount > 0 Then
                NumPacchi = Fix(RS("nmr_Buste") / .NUMBUSTEMAX)
                
                If NumPacchi = 0 Then
                    NumPacchi = 1
                Else
                    ExtraEnvelopes = (RS("nmr_Buste") Mod .NUMBUSTEMAX)
                
                    If ExtraEnvelopes > 0 Then NumPacchi = NumPacchi + 1
                End If
                        
                RS.Close
                
                Set RS = Nothing
                Set RS_SP = New ADODB.Recordset
                        
                For I = 1 To NumPacchi
                    idPacco = idPacco + 1
                
                    myAPB.APBItemsLabel = "Creating Pack " & Format$(idPacco, "000") & " - Mix Packages"
                    myAPB.APBItemsProgress = 3
                    
                    ' Serializzazione
                    '
                    RS_SP.Open "SELECT TOP " & IIf((I = NumPacchi) And (ExtraEnvelopes > 0), ExtraEnvelopes, .NUMBUSTEMAX) & " " & .REF_TABLE & ".id_WorkCntr, " & .REF_TABLE & ".id_Pacco, " & .REF_TABLE & ".id_Posizione, " & .REF_TABLE & ".flg_Mix FROM " & .REF_TABLE & _
                               " INNER JOIN view_Bacini ON " & .REF_TABLE & ".id_Provincia = view_Bacini.id_Provincia" & _
                               " WHERE id_Pacco IS NULL AND id_WorkingLoad = " & .IDWORKING & _
                               " ORDER BY view_Bacini.descr_Bacino, view_Bacini.id_Tariffa, " & .REF_TABLE & "." & .PSTLFIELD & " ASC" & IIf(.PSTLORDERBY = "", "", ", " & .PSTLORDERBY), _
                               DBConn, adOpenDynamic, adLockOptimistic
                
                    If RS_SP.RecordCount > 0 Then
                        myAPB.APBMaxItem = RS_SP.RecordCount
             
                        Do Until RS_SP.EOF
                            myAPB.APBItemLabel = "Item " & RS_SP.AbsolutePosition & " of " & RS_SP.RecordCount
                            myAPB.APBItemProgress = RS_SP.AbsolutePosition
             
                            RS_SP("id_Pacco") = idPacco
                            RS_SP("id_Posizione") = RS_SP.AbsolutePosition
                            RS_SP("flg_Mix") = 2
             
                            RS_SP.MoveNext
                        Loop
                    End If
             
                    RS_SP.Close
                Next I
             
                Set RS_SP = Nothing
            End If
            
            DB_PackagesSerialize_PSTL = True
        Else
            ErrMsg = "Nessun record trovato per la serializzazione dei campi"
        
            GoTo ErrHandler
        End If
    End With
    
    Exit Function
        
ErrHandler:
    myAPB.APBClose
    
    Set RS = Nothing
    Set RS_SP = Nothing
    
    MsgBox IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg), vbExclamation, "Packages Serializer:"

End Function

Private Function DB_PackagesSerialize_STD() As Boolean
    
    On Error GoTo ErrHandler
    
    Dim ErrMsg      As String
    Dim PackCntr    As Integer
    Dim PosCntr     As Integer
    Dim retValue    As Boolean
    Dim RS          As ADODB.Recordset
    Dim SortString  As String
    
    ' Get Sort Mode
    '
    Set RS = DBConn.Execute("SELECT STR_ORDERFIELDS FROM EDT_PROJECTS WHERE ID_PROJECT = " & ProjectInfo.IDPROJECT)
    
    If RS.RecordCount > 0 Then
        If Not IsNull(RS("str_OrderFields")) Then
            Dim SplitData()  As String
            
            SplitData = Split(RS("STR_ORDERFIELDS"), "|")
            
            If chk_Array(SplitData) Then
                Dim I               As Integer
                Dim SubSplitData()  As String
                
                For I = 0 To UBound(SplitData)
                    SubSplitData = Split(SplitData(I), ";")
                    
                    SortString = SortString & IIf(SortString = "", "", ", ") & SubSplitData(0) & IIf(SubSplitData(1) = "", " ASC", " DESC")
                Next I
            End If
            
            Erase SplitData
            Erase SubSplitData
        End If
    End If
    
    RS.Close
    
    ' Start Mode
    '
    DBConn.Execute "UPDATE " & ProjectInfo.REF_TABLE & " SET ID_PACCO = NULL, ID_POSIZIONE = NULL WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING, retValue
    
    If retValue Then
        RS.Open "SELECT ID_WORKCNTR, ID_PACCO, ID_POSIZIONE FROM " & ProjectInfo.REF_TABLE & _
                " WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING & _
                IIf(SortString = "", "", " ORDER BY " & SortString), _
                DBConn, adOpenForwardOnly, adLockOptimistic
        
        If RS.RecordCount > 0 Then
            Set myAPB = New cls_APB
            
            With myAPB
                .APBMode = PBSingle
                .APBCaption = "Serializing:"
                .APBMaxItems = RS.RecordCount
                .tmrMode = Total
                .APBShow
            End With
            
            PackCntr = 1
            
            Do Until RS.EOF
                PosCntr = PosCntr + 1
                
                If PosCntr > ProjectInfo.NUMBUSTEMAX Then
                    PackCntr = PackCntr + 1
                    PosCntr = 1
                End If
                
                myAPB.APBItemsLabel = "Package: " & Format$(PackCntr, "000") & " - Pos.: " & Format$(PosCntr, "000") & " - Rec.: " & Format$(RS.AbsolutePosition, "00000")
                myAPB.APBItemsProgress = RS.AbsolutePosition
                
                RS("ID_PACCO") = PackCntr
                RS("ID_POSIZIONE") = PosCntr
                
                RS.MoveNext
            Loop
            
            RS.Close
            
            myAPB.APBClose
        Else
            ErrMsg = "Nessun record trovato per la serializzazione."
            
            GoTo ErrHandler
        End If
    Else
        ErrMsg = "Errore durante lo svuotamento dei campi per la tabella " & ProjectInfo.REF_TABLE
        
        GoTo ErrHandler
    End If
    
    GoSub CleanUp

    DB_PackagesSerialize_STD = True

    Exit Function

CleanUp:
    myAPB.APBClose

    Set myAPB = Nothing
    
    Set RS = Nothing
Return
    
ErrHandler:
    GoSub CleanUp
    
    MsgBox IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg), vbExclamation, "Packages Serializer:"

End Function

Private Function DB_PackagesSortKey() As Boolean

    On Error GoTo ErrHandler

    Dim ErrMsg      As String
    Dim RS          As ADODB.Recordset
    Dim UKProvincia As Integer
    
    Set RS = New ADODB.Recordset
    
    myAPB.APBItemsLabel = "Generating Sort Keys..."
    myAPB.APBItemsProgress = 1
    
    UKProvincia = Val(DB_GetValueByID("SELECT id_Provincia FROM view_Bacini " & _
                                      "WHERE (nmr_RangeCAP_From >= '99999') AND (nmr_RangeCAP_To <= '99999') " & _
                                      "OR (nmr_RangeCAP_From <= '99999') AND (nmr_RangeCAP_To >= '99999')"))
    
    If UKProvincia <> 0 Then
        RS.Open "SELECT id_WorkCntr, id_Provincia, " & ProjectInfo.PSTLFIELD & " FROM " & ProjectInfo.REF_TABLE & " WHERE id_WorkingLoad = " & ProjectInfo.IDWORKING & " ORDER BY " & ProjectInfo.PSTLFIELD, DBConn, adOpenForwardOnly, adLockOptimistic
        
        If RS.RecordCount > 0 Then
            Dim idProvincia As Integer
            
            myAPB.APBMaxItem = RS.RecordCount
            
            Do Until RS.EOF
                myAPB.APBItemLabel = "Item " & RS.AbsolutePosition & " of " & RS.RecordCount
                myAPB.APBItemProgress = RS.AbsolutePosition
                
                idProvincia = Val(DB_GetValueByID("SELECT id_Provincia FROM view_Bacini " & _
                                                  "WHERE (nmr_RangeCAP_From >= '" & RS(ProjectInfo.PSTLFIELD) & "') AND (nmr_RangeCAP_To <= '" & RS(ProjectInfo.PSTLFIELD) & "') " & _
                                                  "OR (nmr_RangeCAP_From <= '" & RS(ProjectInfo.PSTLFIELD) & "') AND (nmr_RangeCAP_To >= '" & RS(ProjectInfo.PSTLFIELD) & "')"))
                
                If idProvincia = 0 Then idProvincia = UKProvincia
                
                RS("id_Provincia") = idProvincia
                RS.MoveNext
            Loop
        Else
            ErrMsg = "Nessun record trovato per la lavorazione delle provincie"
        
            GoTo ErrHandler
        End If
        
        RS.Close
    
        DB_PackagesSortKey = True
    Else
        ErrMsg = "Impossibile trovare la provincia di riferimento"
    
        GoTo ErrHandler
    End If
    
    Set RS = Nothing
    
    Exit Function

ErrHandler:
    myAPB.APBClose
    
    Set RS = Nothing
    
    MsgBox IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg), vbExclamation, "Packages Sort Key:"

End Function
