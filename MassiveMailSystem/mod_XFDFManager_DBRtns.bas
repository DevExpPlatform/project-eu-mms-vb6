Attribute VB_Name = "mod_XFDFManager_DBRtns"
Option Explicit

Public Type strct_FV
    Field               As String
    Value               As String
End Type

Public ExtFieldsCntr    As Integer
Public XFDF_ExtFields() As strct_FV

Private Function DB_GetXFDFFields() As String()

    Dim RS                As ADODB.Recordset
    Dim XFDF_FieldNames() As String

    Set RS = DBConn.Execute("SELECT * FROM edt_DataCutter WHERE (descr_FieldName <> N'$BLANK') AND id_DataCutter = " & ProjectInfo.IDDATACUTTER & " ORDER BY str_FieldMerger DESC, nmr_FieldOrder")
    
    If RS.RecordCount > 0 Then
        Dim CurPos As Integer
        
        ReDim XFDF_FieldNames(RS.RecordCount - 1, 7)
        
        Do Until RS.EOF
            CurPos = (RS.AbsolutePosition - 1)
            
            If Left$(RS("descr_FieldName"), 6) = "$FIELD" Then
                XFDF_FieldNames(CurPos, 0) = Right$(RS("descr_FieldName"), Len(RS("descr_FieldName")) - InStr(1, RS("descr_FieldName"), "|"))
                XFDF_FieldNames(CurPos, 6) = RS("str_DBFieldType")
            Else
                XFDF_FieldNames(CurPos, 0) = RS("descr_FieldName")
            End If
            
            If IsNull(RS("nmr_Repeats")) Then
                XFDF_FieldNames(CurPos, 1) = 0
            Else
                XFDF_FieldNames(CurPos, 1) = RS("nmr_Repeats")
            End If
            
            If Not IsNull(RS("str_FieldMerger")) Then XFDF_FieldNames(CurPos, 2) = RS("str_FieldMerger")
            If Not IsNull(RS("flg_Splitter")) Then XFDF_FieldNames(CurPos, 3) = RS("flg_Splitter")
            If Not IsNull(RS("str_BarCodeType")) Then XFDF_FieldNames(CurPos, 4) = RS("str_BarCodeType")
            If Not IsNull(RS("flg_IsImage")) Then XFDF_FieldNames(CurPos, 5) = IIf(RS("flg_IsImage"), "1", "0")
            If Not IsNull(RS("flg_IsML")) Then XFDF_FieldNames(CurPos, 7) = RS("flg_IsML")
            
            RS.MoveNext
        Loop
        
        RS.Close
    
        DB_GetXFDFFields = XFDF_FieldNames
        
        Erase XFDF_FieldNames
    End If

End Function

Public Function DB_XFDFExport(ByVal PackStart As Integer, ByVal PackEnd As Integer, Optional ByVal IdDoc As String) As String
    
    ' On Error GoTo ErrHandler

    DBConn.Open

    Dim BarCodeFName        As String
    Dim CanGenerate         As Boolean
    Dim GlobalAbsolutePos   As Long
    Dim GlobalRecordCount   As Long
    Dim idOMRCntr           As Byte
    Dim idPacco             As Integer
    Dim idSubQuery          As Byte
    Dim J                   As Integer
    Dim K                   As Integer
    Dim L                   As Integer
    Dim MaxSubQueries       As Byte
    Dim myBarCode           As cls_BarCodeGen
    Dim myOMR               As cls_BarCodeGen
    Dim myXFDF              As cls_GenXFDF
    Dim NumPacchi           As Integer
    Dim OMRCntr             As Byte
    Dim OMRCntrMax          As Byte
    Dim OMREncParam         As String
    Dim OMREnvMark          As Boolean
    Dim OMRExtra            As String
    Dim OMRFName            As String
    Dim OMRLastCntr         As Byte
    Dim OMRSheets           As Byte
    Dim OMRTYPE             As Byte
    Dim PDF_DName           As String
    Dim PDF_FName           As String
    Dim RS                  As ADODB.Recordset
    Dim tmp_FieldName       As String
    Dim tmp_ImagePathName   As String
    Dim tmp_TemplSheets     As Byte
    Dim XFDF_FieldNames()   As String
    Dim XFDF_FName          As String
    Dim XFDF_SplitData()    As String

    Set myAPB = New cls_APB
    Set myBarCode = New cls_BarCodeGen
    Set myOMR = New cls_BarCodeGen
    Set myXFDF = New cls_GenXFDF
    
    XFDF_FieldNames = DB_GetXFDFFields
        
    If chk_Array(XFDF_FieldNames) Then
        GlobalRecordCount = DB_GetValueByID("SELECT COUNT(*) AS GRC FROM " & ProjectInfo.REF_TABLE & " WHERE id_WorkingLoad = '" & ProjectInfo.IDWORKING & "'")
        NumPacchi = DB_GetValueByID("SELECT MAX(id_Pacco) AS numPacchi FROM " & ProjectInfo.REF_TABLE & " WHERE id_WorkingLoad = '" & ProjectInfo.IDWORKING & "'")
        
        If PackStart <= 0 Then PackStart = 1
        If PackEnd <= 0 Then PackEnd = NumPacchi
        
        With myAPB
            .APBCaption = "Generating PDFs:"
            .APBMaxItems = NumPacchi
            .APBMode = PBDouble
            .tmrMode = Total
            .APBItemsLabel = "Working on: " & SubProjectInfo.NAME
            .APBShow
        End With
            
        If SubProjectInfo.OMRGEN Then
            Select Case SubProjectInfo.OMRTYPE
                Case "OMR"
                    OMRCntrMax = 8
                    OMRTYPE = OMR
                
                Case "OMRPBB"
                    OMRTYPE = OMRPBB
                
                Case "OMRPBE"
                    OMRCntrMax = 9
                    OMRTYPE = OMRPBE
                
                Case "KERN"
                    OMRCntrMax = 16
                    OMRTYPE = KERN
                    
            End Select
        End If
            
        For idPacco = 1 To PackEnd
            myAPB.APBItemLabel = "Querying Package " & Format$(idPacco, "000") & ". Wait Please..."
            
            Set RS = DBConn.Execute("SELECT * FROM " & ProjectInfo.REF_TABLE & _
                                    " WHERE ID_PACCO = " & idPacco & _
                                    " AND ID_WORKINGLOAD = " & ProjectInfo.IDWORKING & _
                                    " ORDER BY ID_PACCO, ID_POSIZIONE")

            If RS.RecordCount > 0 Then
                PDF_DName = WorkPaths.PDFDIR & "Package_" & Format$(RS("id_Pacco"), "000") & "\"
                If idPacco >= PackStart Then chk_Directory PDF_DName, 3
                
                MaxSubQueries = UBound(SubProjectInfo.TEMPLATES)
            
                ' SubQueries
                '
                For idSubQuery = 0 To MaxSubQueries
                    myAPB.APBItemsLabel = SubProjectInfo.NAME & " - Pack: " & Format$(idPacco, "000") & "/" & Format$(NumPacchi, "000")
                    myAPB.APBItemsProgress = idPacco
                    
                    ' Template Info
                    '
                    DB_XFDFExport = DB_GetTemplatesInfo_SELECT(SubProjectInfo.TEMPLATES(idSubQuery))
                    
                    If DB_XFDFExport = "" Then
                        DB_XFDFExport = PDF_TemplatesMerger
                    Else
                        GoSub CleanUp
                        
                        Exit Function
                    End If
                
                    OMRSheets = (TemplateInfo.SHEETS - 1)
                    SubProjectInfo.QUERYFILTER = SubProjectInfo.TEMPQFILTER(idSubQuery)
                    
                    ' Cycle
                    '
                    myAPB.APBMaxItem = RS.RecordCount
                    OMRCntr = OMRLastCntr
                    
                    RS.MoveFirst
                    
                    Do Until RS.EOF
                        If idSubQuery = 0 Then GlobalAbsolutePos = GlobalAbsolutePos + 1
                        
                        myAPB.APBItemLabel = "Doc. Pos.: " & Format$(RS("id_Posizione"), "000") & "/" & Format$(RS.RecordCount, "000") & " - Part: " & Format$((idSubQuery + 1), "00") & "/" & Format$((MaxSubQueries + 1), "00") & " - Item: " & Format$(GlobalAbsolutePos, "00000") & "/" & Format$(GlobalRecordCount, "00000")
                        myAPB.APBItemProgress = RS.AbsolutePosition
                        
                        If (SubProjectInfo.QUERYFILTER.Field <> "") Then
                            CanGenerate = (RS(SubProjectInfo.QUERYFILTER.Field) = SubProjectInfo.QUERYFILTER.Value)
                        
                            If (SubProjectInfo.OMRGEN And (CanGenerate = False) And (OMRCntrMax > 0)) Then
                                tmp_TemplSheets = DB_GetValueByID("SELECT NMR_SHEETS FROM VIEW_TEMPLATESDETAILS WHERE ID_SUBPROJECT = " & SubProjectInfo.SUBPRJID & " AND STR_QFIELD = '" & SubProjectInfo.QUERYFILTER.Field & "' AND STR_QVALUE = '" & RS(SubProjectInfo.QUERYFILTER.Field) & "'")
                                
                                For idOMRCntr = 1 To tmp_TemplSheets
                                    OMRCntr = OMRCntr + 1
                                    
                                    If OMRCntr = OMRCntrMax Then OMRCntr = 1
                                Next idOMRCntr
                            End If
                        Else
                            CanGenerate = True
                        End If
                        
                        If CanGenerate Then
                            If idPacco >= PackStart And IIf(IdDoc = "", True, (RS("ID_WORKCNTR") = IdDoc)) Then
                                XFDF_FName = WorkPaths.TEMPORARYDIR & SubProjectInfo.BASEFILENAME & "_X" & Format$(RS.AbsolutePosition, "000") & ".XFDF"
                                
                                With myXFDF
                                    .XFDF_Open XFDF_FName, SubProjectInfo.MERGEDDOCNAME, WorkPaths.TEMPLATESDIR
                                    
                                    ' Write O.M.R. Fields
                                    '
                                    If SubProjectInfo.OMRGEN Then
                                        With myOMR
                                            .SetCode = OMRTYPE
                                            .SetImageFormat = JPEG
                                            .SetZoom = 4
                                        End With
                                        
                                        OMREnvMark = (RS.AbsolutePosition = RS.RecordCount)
                                        
                                        Select Case OMRTYPE
                                            Case OMR
                                                OMRExtra = "0"
                                                
                                                If OMREnvMark Then OMRExtra = "2"
                                            
                                            Case KERN
                                                OMRExtra = ""
                                                
                                        End Select
                                        
                                        For idOMRCntr = 0 To OMRSheets
                                            If (OMRCntrMax > 0) Then
                                                OMRCntr = OMRCntr + 1
                                                
                                                If OMRCntr = OMRCntrMax Then OMRCntr = 1
                                            End If
                                            
                                            Select Case OMRTYPE
                                            Case OMR
                                                If (idOMRCntr = OMRSheets) Then
                                                    If OMREnvMark Then
                                                        OMRExtra = "3"
                                                    Else
                                                        OMRExtra = "1"
                                                    End If
                                                End If
                                            
                                                OMREncParam = OMRExtra & OMRCntr
                                            
                                            Case OMRPBB
                                                If (OMRSheets = 0) Then
                                                    OMREncParam = "0"
                                                Else
                                                    Select Case idOMRCntr
                                                    Case 0
                                                        OMREncParam = "2"
                                                    
                                                    Case 1
                                                        OMREncParam = "1"
                                                       
'                                                   Case Is > 0
'                                                       OMREncParam = "3"
                                                    
                                                    End Select
                                                End If
                                            
                                            Case OMRPBE
                                                If (OMRSheets = 0) Then
                                                    OMREncParam = "01"
                                                Else
                                                    Select Case idOMRCntr
                                                        Case 0
                                                            OMREncParam = "20"
                                                        
                                                        Case Is = OMRSheets
                                                            OMREncParam = "10"
                                                        
                                                        Case Is > 0
                                                            OMREncParam = "30"
                                                    
                                                    End Select
                                                End If
                                                
                                                OMREncParam = OMREncParam & (OMRCntr - 1)
                                            
                                            Case KERN
                                                If (idOMRCntr = OMRSheets) Then
                                                    OMRExtra = "10"
                                                Else
                                                    OMRExtra = "01"
                                                End If
                                                    
                                                OMREncParam = "0" & OMRExtra & Hex$(OMRCntr) & IIf(OMREnvMark, "1", "0")
                                        
                                            End Select
                                            
                                            If myOMR.Encode(OMREncParam) Then
                                                OMRFName = IIf(idOMRCntr > 0, idOMRCntr & ".", "") & "OMR_00"
                                            
                                                If myOMR.SaveImage(WorkPaths.TEMPORARYDIR & OMRFName & ".JPG") Then
                                                    .XFDF_FieldImage OMRFName, WorkPaths.TEMPORARYDIR & OMRFName & ".JPG", ""
                                                Else
                                                    DB_XFDFExport = "Errore durante il salvataggio file temporaneo dell'OMR."
                                               
                                                    GoSub CleanUp
                                               
                                                    Exit Function
                                                End If
                                            Else
                                                DB_XFDFExport = "Errore durante l'encoding dell'OMR."
                                                
                                                GoSub CleanUp
                                                
                                                Exit Function
                                            End If
                                        Next idOMRCntr
                                    End If
                                    
                                    ' Write BarCode
                                    '
                                    If ProjectInfo.BARCODETYPE <> "" Then
                                        With myBarCode
                                            Select Case ProjectInfo.BARCODETYPE
                                                Case "ITF"
                                                    myBarCode.SetCode = ITF

                                            End Select

                                            .DrawText = ProjectInfo.BARCODETEXT
                                            .SetHeight = 50
                                            .SetImageFormat = JPEG
                                            .SetWidthNarrow = 2
                                            .SetWidthWide = 4
                                            .SetZoom = 4
                                        End With
                                        
                                        If myBarCode.Encode(RS("NMR_" & ProjectInfo.BARCODETYPE & "CODE")) Then
                                            BarCodeFName = ProjectInfo.BARCODETYPE & "_" & RS("NMR_" & ProjectInfo.BARCODETYPE & "CODE")

                                            If myBarCode.SaveImage(WorkPaths.TEMPORARYDIR & BarCodeFName & ".JPG") Then
                                                .XFDF_FieldImage "NMR_" & ProjectInfo.BARCODETYPE & "CODE", WorkPaths.TEMPORARYDIR & BarCodeFName & ".JPG", ""
                                                .XFDF_FieldText "TXT_" & ProjectInfo.BARCODETYPE & "CODE", RS("NMR_" & ProjectInfo.BARCODETYPE & "CODE"), False
                                                .XFDF_FieldText "TXT_" & ProjectInfo.BARCODETYPE & "CHECKDIGIT", myBarCode.GetCheckDigit, False
                                            Else
                                                DB_XFDFExport = "Errore durante l'encoding del BarCode."

                                                GoSub CleanUp

                                                Exit Function
                                            End If
                                        End If
                                    End If
                                    
                                    ' Write JobId Fields
                                    '
                                    If ProjectInfo.PSTLIDJOB <> "" Then
                                        .XFDF_FieldText "txt_IdOmologazione", ProjectInfo.PSTLIDHOMOLOGATION, False
                                        .XFDF_FieldText "txt_LSP", ProjectInfo.PSTLIDJOB & "_" & ProjectInfo.IDWORKING & "_S" & Format$(RS("id_Pacco"), "000") & "_P" & Format$(RS("id_Posizione"), "000"), False
                                    End If
                                            
                                    ' Write Fields
                                    '
                                    For J = 0 To UBound(XFDF_FieldNames)
                                        tmp_ImagePathName = ""
                                        
                                        Select Case XFDF_FieldNames(J, 0)
                                            Case "#SERIALBARCODE"
                                                tmp_FieldName = "#" & Trim$(Left$(ProjectInfo.PSTLIDJOB, 4)) & ProjectInfo.IDWORKING & Format$(RS("id_Pacco"), "000") & Format$(RS("id_Posizione"), "000")
                                                        
                                                .XFDF_FieldBarCode "SERIALBARCODE", XFDF_FieldNames(J, 6), tmp_FieldName
                                                .XFDF_FieldText "txt_SERIALBARCODE", tmp_FieldName, False
                                            
                                                tmp_FieldName = ""
                                            
                                            Case Else
                                                If Not IsNull(RS(XFDF_FieldNames(J, 0))) Then
                                                    For K = 0 To XFDF_FieldNames(J, 1)
                                                        If (XFDF_FieldNames(J, 3) <> "") Then
                                                            ' Write Splitted Data
                                                            '
                                                            XFDF_SplitData = Split(RS(XFDF_FieldNames(J, 0)), XFDF_FieldNames(J, 3))
                                                            
                                                            If chk_Array(XFDF_SplitData) Then
                                                                For L = 0 To UBound(XFDF_SplitData)
                                                                    If Trim$(XFDF_SplitData(L)) <> "" Then
                                                                        .XFDF_FieldText XFDF_FieldNames(J, 0) & Format$((L + 1), "000"), XFDF_SplitData(L), False
                                                                    End If
                                                                Next L
                                                            End If
                                                        Else
                                                            If (Val(XFDF_FieldNames(J, 5)) = 1) Then
                                                                ' Write Image
                                                                '
                                                                tmp_ImagePathName = Get_PathName(RS(XFDF_FieldNames(J, 0)))
        
                                                                .XFDF_FieldImage XFDF_FieldNames(J, 0), IIf(tmp_ImagePathName = "", WorkPaths.TEMPLATESDIR & "Images\", "") & RS(XFDF_FieldNames(J, 0)), ""
                                                            Else
                                                                ' Write Text
                                                                '
                                                                .XFDF_FieldText XFDF_FieldNames(J, 0), RS(XFDF_FieldNames(J, 0)), IIf(XFDF_FieldNames(J, 7) = "1", True, False)
                                                            
                                                                ' Write BarCodes
                                                                '
                                                                If XFDF_FieldNames(J, 4) <> "" And XFDF_FieldNames(J, 2) = "" Then
                                                                    XFDF_SplitData = Split(XFDF_FieldNames(J, 4), "|")
                                                                    
                                                                    If (UBound(XFDF_SplitData) > 0) Then
                                                                        For L = 0 To UBound(XFDF_SplitData)
                                                                            .XFDF_FieldBarCode XFDF_FieldNames(J, 0) & "_BC" & Format$(L, "00"), XFDF_SplitData(L), RS(XFDF_FieldNames(J, 0))
                                                                            .XFDF_FieldText XFDF_FieldNames(J, 0) & "_BCT" & Format$(L, "00"), RS(XFDF_FieldNames(J, 0)), False
                                                                        Next L
                                                                    Else
                                                                        .XFDF_FieldBarCode XFDF_FieldNames(J, 0) & "_BC", XFDF_FieldNames(J, 4), RS(XFDF_FieldNames(J, 0))
                                                                        .XFDF_FieldText XFDF_FieldNames(J, 0) & "_BCT", RS(XFDF_FieldNames(J, 0)), False
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next K
                                                    
                                                    ' Write Merged Fields
                                                    '
                                                    If XFDF_FieldNames(J, 2) <> "" Then
                                                        tmp_FieldName = RS(XFDF_FieldNames(J, 0)) & IIf(tmp_FieldName = "", "", " ") & tmp_FieldName
                                                        
                                                        If Val(Right$(XFDF_FieldNames(J, 2), 2)) = 1 Then
                                                            For K = 0 To XFDF_FieldNames(J, 1)
                                                                .XFDF_FieldText Left$(XFDF_FieldNames(J, 2), Len(XFDF_FieldNames(J, 2)) - 3), tmp_FieldName, False
                                                            
                                                                If XFDF_FieldNames(J, 4) <> "" Then
                                                                    .XFDF_FieldBarCode Left$(XFDF_FieldNames(J, 2), Len(XFDF_FieldNames(J, 2)) - 3), XFDF_FieldNames(J, 4), tmp_FieldName
                                                                    .XFDF_FieldText Left$(XFDF_FieldNames(J, 2), Len(XFDF_FieldNames(J, 2)) - 3) & "_BCT", tmp_FieldName, False
                                                                End If
                                                            Next K
                
                                                            tmp_FieldName = ""
                                                        End If
                                                    End If
                                                End If
                                        
                                        End Select
                                    Next J
                                            
                                    ' Write ExtFields
                                    '
                                    If ExtFieldsCntr > 0 Then
                                        For J = 0 To UBound(XFDF_ExtFields)
                                            .XFDF_FieldText XFDF_ExtFields(J).Field, XFDF_ExtFields(J).Value, False
                                        Next J
                                    End If
                                    
                                    .XFDF_Close
                                End With
                                
                                PDF_FName = SubProjectInfo.BASEFILENAME & "_D" & Format$(RS("id_Posizione"), "000")
                                
                                If (IdDoc <> "") Then PDFInfo.SINGLEPDFFILENAME = PDF_DName & PDF_FName & ".PDF"
                                
                                If XFDF_Merger(XFDF_FName, True, PDF_DName & PDF_FName & ".PDF") Then
                                'If XFDF_Merger(SubProjectInfo.MergedDocName, XFDF_FName, True, PDF_DName & PDF_FName & ".PDF") Then
                                    If (SubProjectInfo.SINGLEPDFPACKING) Then
                                        If (PDF_Packer(PDF_DName & PDF_FName & ".PDF") = False) Then
                                            DB_XFDFExport = "Errore durante la compressione di:" & PDF_FName
                                            
                                            GoTo ErrHandler
                                        End If
                                    End If
                                Else
                                    DB_XFDFExport = "Errore durante la generazione di: " & PDF_FName
                               
                                    GoSub CleanUp
                               
                                    Exit Function
                                End If
                            Else
                                If SubProjectInfo.OMRGEN Then
                                    For idOMRCntr = 1 To TemplateInfo.SHEETS
                                        OMRCntr = OMRCntr
                                        
                                        If OMRCntr = OMRCntrMax Then OMRCntr = 1
                                    Next idOMRCntr
                                End If
                            End If
                        End If
                        
                        RS.MoveNext
                    Loop
                    
                    If (idPacco >= PackStart) Then
                        If idSubQuery = MaxSubQueries Then OMRLastCntr = OMRCntr
                    Else
                        OMRLastCntr = OMRCntr
                        
                        Exit For
                    End If
                Next idSubQuery
            Else
                DB_XFDFExport = "Nessun record sulla tabella " & ProjectInfo.REF_TABLE & "."
                
                GoSub CleanUp
                
                Exit Function
            End If
        Next idPacco
    Else
        DB_XFDFExport = "Errore definizione campi per la tabella " & ProjectInfo.REF_TABLE & "."
    End If

    GoSub CleanUp
    
    Exit Function

CleanUp:
    myAPB.APBClose
    
    Erase XFDF_FieldNames
    Erase XFDF_SplitData
        
    Set myAPB = Nothing
    Set myBarCode = Nothing
    Set myOMR = Nothing
    Set myXFDF = Nothing
    Set RS = Nothing
    
    If DBConn.State Then DBConn.Close
Return

ErrHandler:
    GoSub CleanUp

    DB_XFDFExport = Err.Description

End Function
