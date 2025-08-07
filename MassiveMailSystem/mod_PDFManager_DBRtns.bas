Attribute VB_Name = "mod_PDFManager_DBRtns"
Option Explicit

Private Type strct_FileNaming
    Alias       As String
    FieldName   As String
End Type

Public Function DB_PDFsOrganizer(ByRef CFOFIXEDPATH As String) As Boolean

    On Error GoTo ErrHandler

    Dim DirNaming()         As strct_FileNaming
    Dim DirNamingCntr       As Integer
    Dim ErrMsg              As String
    Dim FileDst             As String
    Dim FileNaming()        As strct_FileNaming
    Dim FileNamingCntr      As Integer
    Dim FileSrc             As String
    Dim I                   As Integer
    Dim J                   As Integer
    Dim myAbsolutePosition  As Long
    Dim RS                  As ADODB.Recordset
    Dim RSFiles             As ADODB.Recordset
    Dim SortString          As String
    Dim SplitData()         As String
    Dim SQLFields           As String
    Dim SQLWhere            As String
    Dim SubSplitData()      As String
    Dim tmp_Dir             As String
    Dim tmp_FileName        As String
    Dim tmp_Path            As String
    Dim tmp_String          As String
    Dim tmpCntr             As Integer
    Dim tmpIdPacco          As Integer
    Dim WS_CFOPATH          As String
    Dim WS_FLG_EMPTY_CFO    As Boolean
    
    DBConn.Open
    
    Set RS = DBConn.Execute("SELECT ID_SUBPROJECT, STR_CUSTOMERFILEORGANIZER, FLG_CFOMODE, STR_CUSTOMERFILEORGANIZERPATH, FLG_CFOFIXEDPATH, FLG_EMPTYCFO, STR_CFOPFIELDS, FLG_CFOPMODE, STR_CFOPPATH FROM EDT_SUBPROJECTS WHERE (ID_PROJECT = " & ProjectInfo.IDPROJECT & ")")

    If RS.RecordCount > 0 Then
        If (DLLParams.UnattendedMode = False) Then
            Set myAPB = New cls_APB
    
            With myAPB
                .APBMode = PBDouble
                .APBCaption = "Customer File Organizer:"
                .APBMaxItems = RS.RecordCount
                .tmrMode = Total
                .APBShow
            End With
        End If
    
        Do Until RS.EOF
            If DB_SubProjectInfo_SELECT(RS("ID_SUBPROJECT")) Then
                If (DLLParams.UnattendedMode = False) Then
                    myAPB.APBItemsLabel = "Working on: " & SubProjectInfo.NAME
                    myAPB.APBItemsProgress = myAPB.APBItemsProgressValue + 1
                End If
                
                If (Not IsNull(RS("STR_CUSTOMERFILEORGANIZER"))) Then
                    WS_FLG_EMPTY_CFO = True
                
                    If (Not IsNull(RS("FLG_CFOMODE"))) Then SubProjectInfo.CFOMODE = RS("FLG_CFOMODE")
                    If (Not IsNull(RS("FLG_CFOFIXEDPATH"))) Then SubProjectInfo.CFOFIXEDPATH = RS("FLG_CFOFIXEDPATH")
                    If (Not IsNull(RS("STR_CUSTOMERFILEORGANIZERPATH"))) Then WS_CFOPATH = Fix_Paths(RS("STR_CUSTOMERFILEORGANIZERPATH"))
                    If (Not IsNull(RS("FLG_EMPTYCFO"))) Then WS_FLG_EMPTY_CFO = (RS("FLG_EMPTYCFO") = 1)
                    
                    WorkPaths.CFODIR = IIf(WS_CFOPATH = "", WorkPaths.WORKINGDIR & "\CFO_Docs\", WS_CFOPATH)

                    SplitData = Split(RS("STR_CUSTOMERFILEORGANIZER"), "|")
                
                    If chk_Array(SplitData) Then
                        Erase DirNaming
                        Erase FileNaming
                
                        DirNamingCntr = 0
                        FileNamingCntr = 0
                        
                        SortString = ""
                        SQLFields = ""
                        SQLWhere = ""
                
                        For I = 0 To UBound(SplitData)
                            SubSplitData = Split(SplitData(I), ";")
                                     
                            If (SubSplitData(0) <> "") Then
                                SortString = SortString & IIf(SortString = "", "", ", ") & SubSplitData(0) & IIf(SubSplitData(1) = "", " ASC", " DESC")
                                
                                If ((SubSplitData(0) <> "ID_PACCO") And (SubSplitData(0) <> "ID_POSIZIONE")) Then
                                    SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & SubSplitData(0)
                                    SQLWhere = SQLWhere & IIf(SQLWhere <> "", " AND ", "") & "(" & SubSplitData(0) & " IS NOT NULL)"
                                End If
                            End If
                            
                            If SubSplitData(3) = 1 Then
                                ReDim Preserve DirNaming(DirNamingCntr)
                    
                                With DirNaming(DirNamingCntr)
                                    .Alias = SubSplitData(2)
                                    .FieldName = SubSplitData(0)
                                End With
                    
                                DirNamingCntr = DirNamingCntr + 1
                            Else
                                ReDim Preserve FileNaming(FileNamingCntr)
                    
                                With FileNaming(FileNamingCntr)
                                    .Alias = SubSplitData(2)
                                    .FieldName = SubSplitData(0)
                                End With
                    
                                FileNamingCntr = FileNamingCntr + 1
                            End If
                        Next I
                    Else
                        ErrMsg = "Errore parametri di Customer File Organization."
                        
                        GoTo ErrHandler
                    End If
                    
                    ' START PROCESS
                    '
                    Set RSFiles = DBConn.Execute("SELECT ID_PACCO, ID_POSIZIONE, " & SQLFields & " FROM " & ProjectInfo.REF_TABLE & " WHERE (ID_WORKINGLOAD = " & ProjectInfo.IDWORKING & ") AND " & SQLWhere & IIf(SortString = "", "", " ORDER BY " & SortString))
                    
                    If RSFiles.RecordCount > 0 Then
                        If (DLLParams.UnattendedMode = False) Then
                            myAPB.APBMaxItem = RSFiles.RecordCount
                            myAPB.APBItemProgress = 0
                        End If
                                     
                        chk_Directory WorkPaths.WORKINGDIR, 2
                        chk_Directory WorkPaths.CFODIR, 2
                        
                        FileNamingCntr = 0
                        myAbsolutePosition = 0
                        
                        Do Until RSFiles.EOF
                            If (DLLParams.UnattendedMode = False) Then
                                myAPB.APBItemLabel = "File: " & RSFiles.AbsolutePosition & " of " & RSFiles.RecordCount
                                myAPB.APBItemProgress = RSFiles.AbsolutePosition
                            End If
                            
                            tmp_Dir = ""
                            tmp_FileName = ""
                            tmpIdPacco = RSFiles("ID_PACCO")
                            tmpCntr = RSFiles("ID_POSIZIONE")
                            
                            If (DirNamingCntr > 0) Then
                                For I = 0 To UBound(DirNaming)
                                    With DirNaming(I)
                                        If (.FieldName = "") Then
                                            tmp_Dir = tmp_Dir & IIf(.Alias = "", "UNKNOWN", .Alias) & "\"
                                        Else
                                            If IsNull(RSFiles(.FieldName)) Then
                                                tmp_Dir = tmp_Dir & .Alias & "UNKNOWN" & "\"
                                            Else
                                                tmp_Dir = tmp_Dir & .Alias & Trim$(RSFiles(.FieldName)) & "\"
                                            End If
                                        End If
                                        
                                        tmp_Dir = Trim$(tmp_Dir)
                                        
                                        chk_Directory WorkPaths.CFODIR & tmp_Dir, 2
                                    End With
                                Next I
                            End If
                                    
                            For I = 0 To UBound(FileNaming)
                                With FileNaming(I)
                                    If IsNull(RSFiles(.FieldName)) Then
                                        tmp_FileName = tmp_FileName & .Alias & "UNKNOWN"
                                    Else
                                        tmp_FileName = tmp_FileName & .Alias & Trim$(RSFiles(.FieldName))
                                    End If
                                End With
                            Next I
                    
                            tmp_FileName = Trim$(tmp_FileName)
                    
                            If tmp_Path <> tmp_Dir Then
                                If (WS_FLG_EMPTY_CFO) Then EmptyDir WorkPaths.CFODIR & tmp_Dir
                                
                                tmp_Path = tmp_Dir
                            End If
                            
                            If tmp_String = tmp_Dir & tmp_FileName Then
                                FileNamingCntr = FileNamingCntr + 1
                            Else
                                tmp_String = tmp_Dir & tmp_FileName
                
                                FileNamingCntr = 0
                            End If
                    
                            ' EXECUTE FILE PROCEDURE (COPY/MOVE)
                            '
                            FileSrc = Fix_Paths(WorkPaths.PDFDIR & "Package_" & Format$(tmpIdPacco, "000")) & SubProjectInfo.BASEFILENAME & "_D" & Format$(tmpCntr, "000") & ".PDF"
                            FileDst = WorkPaths.CFODIR & Trim$(tmp_Dir) & Conv_Name2ConventionalName(tmp_FileName) & IIf(FileNamingCntr = 0, "", "_" & Format$(FileNamingCntr, "000")) & ".PDF"
            
                            If FDExist(FileSrc, False) Then
                                Select Case SubProjectInfo.CFOMODE
                                    Case 0
                                        FileCopy FileSrc, FileDst
                                    
                                    Case 1
                                        Name FileSrc As FileDst
                                    
                                    Case 2, 3
                                        If FileNamingCntr = 0 Then
                                            If SubProjectInfo.CFOMODE = 2 Then
                                                FileCopy FileSrc, FileDst
                                            Else
                                                Name FileSrc As FileDst
                                            End If
                                        Else
                                            FileDst = WorkPaths.CFODIR & tmp_Dir & Conv_Name2ConventionalName(tmp_FileName) & ".PDF"
                                            
                                            If PDF_Merge2File(FileDst, FileSrc, FileDst) Then
                                                If SubProjectInfo.CFOMODE = 3 Then Kill FileSrc
                                            Else
                                                ErrMsg = "Error merging file '" & Get_BaseName(FileSrc) & "'."
                                                
                                                GoTo ErrHandler
                                            End If
                                        End If
                                
                                End Select
                            Else
                                ErrMsg = "File '" & SubProjectInfo.BASEFILENAME & "_D" & Format$(tmpCntr, "000") & ".PDF" & "' not found."
                            
                                GoTo ErrHandler
                            End If
                    
                            RSFiles.MoveNext
                        Loop
                        
                        RSFiles.Close
                    Else
                        ErrMsg = "Errore durante la ricerca su " & ProjectInfo.REF_TABLE & "."
        
                        GoTo ErrHandler
                    End If
                End If
            
                CFOFIXEDPATH = WorkPaths.CFODIR & IIf(SubProjectInfo.CFOFIXEDPATH = 1, tmp_Dir, "")
                
                ' PACKAGE CFO
                '
                If (Not IsNull(RS("STR_CFOPFIELDS"))) Then
                    If (Not IsNull(RS("FLG_CFOPMODE"))) Then SubProjectInfo.CFOMODE = RS("FLG_CFOPMODE")
                    If (Not IsNull(RS("STR_CFOPPATH"))) Then WS_CFOPATH = Fix_Paths(RS("STR_CFOPPATH"))
                
                    WorkPaths.CFODIR = IIf(WS_CFOPATH = "", WorkPaths.WORKINGDIR & "\CFO_Docs\", WS_CFOPATH)

                    SplitData = Split(RS("STR_CFOPFIELDS"), "|")
                
                    If chk_Array(SplitData) Then
                        Erase DirNaming
                        Erase FileNaming
                
                        DirNamingCntr = 0
                        FileNamingCntr = 0
                        
                        SortString = ""
                        SQLFields = ""
                
                        For I = 0 To UBound(SplitData)
                            SubSplitData = Split(SplitData(I), ";")
                                     
                            If (SubSplitData(0) <> "") Then
                                SortString = SortString & IIf(SortString = "", "", ", ") & SubSplitData(0) & IIf(SubSplitData(1) = "", " ASC", " DESC")
                                
                                If ((SubSplitData(0) <> "ID_PACCO") And (SubSplitData(0) <> "ID_POSIZIONE")) Then
                                    SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & SubSplitData(0)
                                End If
                            End If
                            
                            If SubSplitData(3) = 1 Then
                                ReDim Preserve DirNaming(DirNamingCntr)
                    
                                With DirNaming(DirNamingCntr)
                                    .Alias = SubSplitData(2)
                                    .FieldName = SubSplitData(0)
                                End With
                    
                                DirNamingCntr = DirNamingCntr + 1
                            Else
                                ReDim Preserve FileNaming(FileNamingCntr)
                    
                                With FileNaming(FileNamingCntr)
                                    .Alias = SubSplitData(2)
                                    .FieldName = SubSplitData(0)
                                End With
                    
                                FileNamingCntr = FileNamingCntr + 1
                            End If
                        Next I
                    Else
                        ErrMsg = "Errore parametri di Customer File Organization."
                        
                        GoTo ErrHandler
                    End If
                
                    Set RSFiles = DBConn.Execute("SELECT " & SQLFields & ", MAX(ID_PACCO) AS NMR_PACKAGES FROM " & ProjectInfo.REF_TABLE & " WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING & " GROUP BY " & SQLFields)
                
                    If RSFiles.RecordCount = 1 Then
                        tmpIdPacco = RSFiles("NMR_PACKAGES")
                        
                        If (DLLParams.UnattendedMode = False) Then
                            myAPB.APBMaxItem = tmpIdPacco
                            myAPB.APBItemProgress = 0
                        End If
    
                        chk_Directory WorkPaths.WORKINGDIR, 2
                        chk_Directory WorkPaths.CFODIR, 2
                       
                        For J = 1 To tmpIdPacco
                            If (DLLParams.UnattendedMode = False) Then
                                myAPB.APBItemLabel = "Package: " & J & " of " & tmpIdPacco
                                myAPB.APBItemProgress = J
                            End If
                        
                            tmp_Dir = ""
                            tmp_FileName = ""
                            
                            If (DirNamingCntr > 0) Then
                                For I = 0 To UBound(DirNaming)
                                    With DirNaming(I)
                                        If (.FieldName <> "") Then
                                            If IsNull(RSFiles(.FieldName)) Then
                                                tmp_Dir = tmp_Dir & .Alias & "UNKNOWN" & "\"
                                            Else
                                                tmp_Dir = tmp_Dir & .Alias & Trim$(RSFiles(.FieldName)) & "\"
                                            End If
                                        Else
                                            tmp_Dir = tmp_Dir & IIf(.Alias = "", "UNKNOWN", .Alias) & "\"
                                        End If
                                        
                                        tmp_Dir = Trim$(tmp_Dir)
                                        
                                        chk_Directory WorkPaths.CFODIR & tmp_Dir, 2
                                    End With
                                Next I
                            End If
                                    
                            If (FileNamingCntr > 0) Then
                                For I = 0 To UBound(FileNaming)
                                    With FileNaming(I)
                                        If (.FieldName = "") Then
                                            tmp_FileName = tmp_FileName & IIf(.Alias = "", "UNKNOWN", .Alias)
                                        Else
                                            If IsNull(RSFiles(.FieldName)) Then
                                                tmp_FileName = tmp_FileName & .Alias & "UNKNOWN"
                                            Else
                                                tmp_FileName = tmp_FileName & .Alias & Trim$(RSFiles(.FieldName))
                                            End If
                                        End If
                                    End With
                                Next I
                        
                                tmp_FileName = Trim$(tmp_FileName)
                            Else
                                tmp_FileName = SubProjectInfo.BASEFILENAME & "_P"
                            End If
                            
                            ' EXECUTE FILE PROCEDURE (COPY/MOVE)
                            '
                            FileSrc = WorkPaths.PDFPACKSDIR & SubProjectInfo.BASEFILENAME & "_P" & Format$(J, "000") & ".PDF"
                            FileDst = WorkPaths.CFODIR & Trim$(tmp_Dir) & Conv_Name2ConventionalName(tmp_FileName) & Format$(J, "000") & ".PDF"
                            
                            If FDExist(FileSrc, False) Then
                                Select Case SubProjectInfo.CFOMODE
                                    Case 0
                                        FileCopy FileSrc, FileDst
                                    
                                    Case 1
                                        Name FileSrc As FileDst
                                    
                                    Case 2, 3
                                        If FileNamingCntr = 0 Then
                                            If SubProjectInfo.CFOMODE = 2 Then
                                                FileCopy FileSrc, FileDst
                                            Else
                                                Name FileSrc As FileDst
                                            End If
                                        Else
                                            FileDst = WorkPaths.CFODIR & tmp_Dir & Conv_Name2ConventionalName(tmp_FileName) & ".PDF"
                                            
                                            If PDF_Merge2File(FileDst, FileSrc, FileDst) Then
                                                If SubProjectInfo.CFOMODE = 3 Then Kill FileSrc
                                            Else
                                                ErrMsg = "Error merging file '" & Get_BaseName(FileSrc) & "'."
                                                
                                                GoTo ErrHandler
                                            End If
                                        End If
                                
                                End Select
                            Else
                                ErrMsg = "File '" & SubProjectInfo.BASEFILENAME & "_D" & Format$(tmpCntr, "000") & ".PDF" & "' not found."
                            
                                GoTo ErrHandler
                            End If
                        Next J
                    Else
                        ErrMsg = "Errore parametri di Customer Package File Organization."
                        
                        GoTo ErrHandler
                    End If
                End If
            Else
                ErrMsg = "Unable to Open SubProject: " & RS("ID_SUBPROJECT")
            End If
                
            RS.MoveNext
        Loop
        
        RS.Close
    Else
        ErrMsg = "Nessuna criterio di FileName sorting."
 
        GoTo ErrHandler
    End If
    
    GoSub CleanUp
        
    DB_PDFsOrganizer = True
    
    Exit Function

CleanUp:
    Erase DirNaming
    Erase FileNaming
    Erase SplitData
    Erase SubSplitData

    If (DLLParams.UnattendedMode = False) Then
        myAPB.APBClose
        Set myAPB = Nothing
    End If
    
    Set RS = Nothing
    Set RSFiles = Nothing
    
    DBConn.Close
Return

ErrHandler:
    GoSub CleanUp

    If (DLLParams.UnattendedMode) Then
        UMErrMsg = IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg)
    Else
        MsgBox IIf(ErrMsg = "", Purge_ErrDescr(Err.Description), ErrMsg), vbExclamation, "Customer File Organization:"
    End If

End Function

Public Function DB_PDFsPackagesMergerMode01(ByVal PackStart As Integer, ByVal PackEnd As Integer) As Boolean

    On Error GoTo ErrHandler

    Dim MergePackId()   As Long
    Dim PackMailCheck() As Boolean
    Dim RS              As ADODB.Recordset

    DBConn.Open

    Set RS = DBConn.Execute("SELECT ID_SUBPROJECT, FLG_MAKEPACKAGESMAILCHECK FROM EDT_SUBPROJECTS WHERE (FLG_MAKEPACKAGES = 1) AND (ID_PROJECT = " & ProjectInfo.IDPROJECT & ") ORDER BY ID_PROJECT")

    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            ReDim Preserve MergePackId(RS.AbsolutePosition - 1)
            ReDim Preserve PackMailCheck(RS.AbsolutePosition - 1)
            
            MergePackId(RS.AbsolutePosition - 1) = RS("ID_SUBPROJECT")
            PackMailCheck(RS.AbsolutePosition - 1) = RS("FLG_MAKEPACKAGESMAILCHECK")
            
            RS.MoveNext
        Loop
        
        ' Start Process
        '
        Dim ErrMsg          As String
        Dim I               As Byte
        Dim J               As Integer
        Dim PDF_DName       As String
        Dim PDF_FName       As String
        Dim PDF_Number      As String
        Dim NumPacchi       As Integer
        Dim WS_BOOLEAN      As Boolean
        
        WS_BOOLEAN = True
        NumPacchi = DB_GetValueByID("SELECT MAX(ID_PACCO) AS NUMPACCHI FROM " & ProjectInfo.REF_TABLE & " WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING)
        
        If PackStart = -1 Then PackStart = 1
        If PackEnd = -1 Then PackEnd = NumPacchi
    
        If (DLLParams.UnattendedMode = False) Then
            Set myAPB = New cls_APB
            
            With myAPB
                .APBMode = PBDouble
                .APBCaption = "Creating Packages:"
                .APBMaxItem = ((PackEnd - PackStart) + 1)
                .APBMaxItems = UBound(MergePackId) + 1
                .tmrMode = Total
                .APBShow
            End With
        End If
        
        For I = 0 To UBound(MergePackId)
            If DB_SubProjectInfo_SELECT(MergePackId(I)) Then
                If (DLLParams.UnattendedMode = False) Then
                    myAPB.APBItemsLabel = "Working on: " & SubProjectInfo.NAME
                    myAPB.APBItemsProgress = I + 1
                    myAPB.APBItemProgress = 0
                End If
        
                chk_Directory WorkPaths.WORKINGDIR, 2
                chk_Directory WorkPaths.PDFPACKSDIR, 3
        
                For J = PackStart To PackEnd
                    PDF_DName = WorkPaths.PDFDIR & "Package_" & Format$(J, "000") & "\"
                    PDF_FName = SubProjectInfo.BASEFILENAME & "_P" & Format$(J, "000") & ".PDF"
                    PDF_Number = Format$(J, "000") & "/" & Format$(NumPacchi, "000")
               
                    If (DLLParams.UnattendedMode = False) Then
                        myAPB.APBItemLabel = "Merging Package: " & PDF_Number
                        myAPB.APBItemProgress = myAPB.APBItemProgressValue + 1
                    End If
                
                    If (PackMailCheck(I)) Then WS_BOOLEAN = (Val(DB_GetValueByID("SELECT COUNT(*) AS NMR_PDF FROM " & ProjectInfo.REF_TABLE & " WHERE ID_WORKINGLOAD = " & ProjectInfo.IDWORKING & " AND ID_PACCO = " & J & " AND (FLG_NOMERGE IS NULL OR FLG_NOMERGE = 0)")) > 0)
                
                    If (WS_BOOLEAN) Then
                        If (PDF_DBMergerMode01(WorkPaths.BASEDIR, SubProjectInfo.BASEFILENAME, ProjectInfo.REF_TABLE, ProjectInfo.IDWORKING, J, WorkPaths.PDFPACKSDIR & PDF_FName) = False) Then
                            ErrMsg = "Errore durante la generazione di: " & PDF_FName
                    
                            GoTo ErrHandler
                        End If
                        
                        If (DLLParams.UnattendedMode = False) Then myAPB.APBItemLabel = "Packing Package: " & PDF_Number
                        
                        If (SubProjectInfo.PACKAGEPACKING) Then
                            If (PDF_Packer(WorkPaths.PDFPACKSDIR & PDF_FName) = False) Then
                                ErrMsg = "Errore durante la compressione di:" & PDF_FName
                        
                                GoTo ErrHandler
                            End If
                        End If
                    End If
                 
                    DoEvents
                Next J
            End If
        Next I
    End If
    
    RS.Close
    
    GoSub CleanUp
    
    DB_PDFsPackagesMergerMode01 = True

    Exit Function

CleanUp:
    Erase MergePackId
    
    Set RS = Nothing
    
    DBConn.Close
    
    If (DLLParams.UnattendedMode = False) Then
        myAPB.APBClose
        
        Set myAPB = Nothing
    End If
Return

ErrHandler:
    GoSub CleanUp
    
    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox ErrMsg, vbExclamation, "Merging Packages:"
    End If

End Function

Public Function DB_PDFsPackagesMergerMode02(PackStart As String, PackEnd As String, unattended As String) As Boolean

    On Error GoTo ErrHandler

    Dim ErrMsg          As String
    Dim MergePackId()   As Long
    Dim RS              As ADODB.Recordset

    DBConn.Open

    Set RS = DBConn.Execute("SELECT ID_SUBPROJECT, FLG_MAKEPACKAGESMAILCHECK FROM EDT_SUBPROJECTS WHERE (FLG_MAKEPACKAGES = 1) AND (ID_PROJECT = " & ProjectInfo.IDPROJECT & ") ORDER BY ID_PROJECT")

    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            ReDim Preserve MergePackId(RS.AbsolutePosition - 1)
            
            If DB_SubProjectInfo_SELECT(RS("ID_SUBPROJECT")) Then
                DB_PDFsPackagesMergerMode02 = PDF_PackageMerger(PackStart, PackEnd, unattended)
            
                If (DB_PDFsPackagesMergerMode02 = False) Then
                    RS.Close
                                        
                    GoSub CleanUp
                        
                    Exit Function
                End If
            End If
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
    
    GoSub CleanUp
    
    DB_PDFsPackagesMergerMode02 = True

    Exit Function

CleanUp:
    Erase MergePackId
    
    Set RS = Nothing
    
    DBConn.Close
Return

ErrHandler:
    GoSub CleanUp
    
    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox ErrMsg, vbExclamation, "Merging Packages:"
    End If

End Function
