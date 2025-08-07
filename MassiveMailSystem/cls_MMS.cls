VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_MMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
                                                                                                                                                                                                                                           
Dim UnattendedMode As Boolean

Public Property Let AddExtField(ByVal myExtField As String)

    Dim SplitData() As String

    SplitData = Split(myExtField, "|")

    If chk_Array(SplitData) Then
        ReDim XFDF_ExtFields(ExtFieldsCntr)

        With XFDF_ExtFields(ExtFieldsCntr)
            .Field = SplitData(0)
            .Value = SplitData(1)
        End With

        ExtFieldsCntr = ExtFieldsCntr + 1
    End If

    Erase SplitData

End Property

Public Property Let AutoMergePacks(ByVal myAutoMerge As Boolean)

    DLLParams.AutoMergePacks = myAutoMerge

End Property

Property Let BaseWorkDir(ByVal myBWDir As String)

    If Trim$(myBWDir) = "" Then MsgBox "Nessuna directory di lavorazione definita.", vbExclamation, "Attenzione:"
    
    DLLParams.BaseWorkDir = myBWDir

End Property

Private Sub Class_Initialize()
    
    AppPath = Fix_Paths(App.Path)

    ExtFieldsCntr = 0

End Sub

Private Sub Class_Terminate()

    DB_ConnectRelease

    Erase XFDF_ExtFields

End Sub

Public Function ExecuteODPPlugIn(OutPath As String, IDODPPLUGIN As String, IDODPPLUGINPARAM As String) As Boolean
    
    If (chk_ProjectNeededParams And (IDODPPLUGIN <> "")) Then
        Dim myPlugIn As Object
        
        chk_Directory OutPath, 2
        
        If PlugIn_Load("ODP_" & IDODPPLUGIN, myPlugIn) Then
            With myPlugIn
                .DNS = DLLParams.DSN
                .WorkingId = ProjectInfo.IDWORKING
                .workingTable = ProjectInfo.REF_TABLE
                .OutPutFileName = OutPath
            
                If (DLLParams.UnattendedMode) Then .SetUnattendedMode = True
                
                If IDODPPLUGINPARAM <> "" Then
                    Dim SplitData() As String
                    
                    SplitData = Split(IDODPPLUGINPARAM, "|")

                    .ExtraParams = SplitData

                    Erase SplitData
                End If
            
                ExecuteODPPlugIn = .StartJob
            End With
        End If
        
        Set myPlugIn = Nothing
    Else
        ExecuteODPPlugIn = True
    End If
    
End Function

Public Function CustomerOrganize() As Boolean
                        
    If chk_ProjectNeededParams Then
        Dim CanDo               As Boolean
        Dim CFOPath             As String
        
        CanDo = DB_PDFsOrganizer(CFOPath)
            
        DoEvents
        
        If (CanDo) Then CanDo = ExecuteODPPlugIn(CFOPath, SubProjectInfo.CFOODPPLUGIN, SubProjectInfo.CFOODPPLUGINPARAMS)
    End If
    
    CustomerOrganize = CanDo
    
End Function

Public Property Let DSN(ByVal myDSN As String)

    DLLParams.DSN = myDSN

End Property

Public Property Get GetCurrentRefTable() As String
    
    GetCurrentRefTable = ProjectInfo.REF_TABLE

End Property

Public Property Get GetCurrentWorking() As String

    GetCurrentWorking = ProjectInfo.IDWORKING

End Property

Public Property Get GetIDPPlugInParams() As String
    
    GetIDPPlugInParams = ProjectInfo.IDIDPPLUGINPARAM

End Property

Public Function GetNormalizersPlugIns() As String()
    
    GetNormalizersPlugIns = PlugIns_Load("Normalizer")

End Function

Public Property Get GetODPPlugInParams() As String
    
    GetODPPlugInParams = ProjectInfo.IDODPPLUGINPARAM

End Property

Public Property Get GetPDFWorkDir() As String
    
    GetPDFWorkDir = Fix_Paths(WorkPaths.WORKINGDIR)

End Property

Public Function GetSinglePDFFileName()

    GetSinglePDFFileName = PDFInfo.SINGLEPDFFILENAME

End Function

Public Property Get GetUMErrorMessage() As String
    
    GetUMErrorMessage = UMErrMsg

End Property

Public Function GetWorkings() As String()
    
    GetWorkings = DB_GetWorkings

End Function

Public Function ImportData(ByVal ImportFileName As String) As Boolean
    
    Dim byPass          As Boolean
    Dim CanDo           As Boolean
    Dim ExportFilePath  As String
    Dim ExportFileName  As String
    Dim myPlugIn        As Object
    Dim SplitData()     As String
    
    If ImportFileName <> "" Then
        If ProjectInfo.IDIDPPLUGIN <> "" Then
            If PlugIn_Load("IDP_" & ProjectInfo.IDIDPPLUGIN, myPlugIn) Then
                With myPlugIn
                    If ProjectInfo.IDIDPPLUGINPARAM <> "" Then
                        SplitData = Split(ProjectInfo.IDIDPPLUGINPARAM, "|")
    
                        .ExtraParams = SplitData
    
                        Erase SplitData
                    End If
                    
                    .InputFileName = ImportFileName
                    
                    byPass = (ProjectInfo.PlugIn = "SQL")
                    
                    If (byPass) Then
                        .DSN = DLLParams.DSN
                        .IDDATACUTTER = ProjectInfo.IDDATACUTTER
                        .TableName = ProjectInfo.REF_TABLE
                        .TNS = DLLParams.TSN
                    Else
                        ExportFilePath = Get_PathName(ImportFileName)
                        
                        If (InStrRev(ImportFileName, ".") > 0) Then
                            ExportFileName = Get_BaseName(ImportFileName, 4) & "_FXD" & Right$(ImportFileName, 4)
                        Else
                            ExportFileName = Get_BaseName(ImportFileName, 0) & "_FXD"
                        End If
                        
                        ExportFileName = ExportFilePath & ExportFileName
                        
                        .OutPutFileName = ExportFileName
                    End If
                    
                    If (DLLParams.UnattendedMode) Then .SetUnattendedMode = True
                    
                    CanDo = .StartJob
                
                    If CanDo Then
                        If (byPass) Then
                            ProjectInfo.IDWORKING = DB_GetValueByID("SELECT MAX(ID_WORKINGLOAD) AS IDWORKINGLOAD FROM " & ProjectInfo.REF_TABLE)
                            
                            ImportData = True
                        Else
                            ImportFileName = ExportFileName
                        End If
                    Else
                        If (DLLParams.UnattendedMode) Then UMErrMsg = .GetUMErrorMessage
                    End If
                End With
            End If
            
            Set myPlugIn = Nothing
        
            If (byPass) Then Exit Function
        Else
            CanDo = True
        End If
        
        If (CanDo) Then
            If PlugIn_Load(ProjectInfo.PlugIn & "_Import", myPlugIn) Then
                With myPlugIn
                    .AddFields_BarCode = (ProjectInfo.BARCODETYPE <> "")
                    .AddFields_Postalizzazione = (ProjectInfo.PSTLFIELD <> "")
                    .DataFileName = ImportFileName
                    .DSN = DLLParams.DSN
                    .TNS = DLLParams.TSN
                    .IDDATACUTTER = ProjectInfo.IDDATACUTTER
                    .TableName = ProjectInfo.REF_TABLE
    
                    If (DLLParams.UnattendedMode) Then .SetUnattendedMode = True
                    
                    If ProjectInfo.PLUGINPARAMS <> "" Then
                        SplitData = Split(ProjectInfo.PLUGINPARAMS, "|")
    
                        .ExtraParams = SplitData
    
                        Erase SplitData
                    End If
    
                    CanDo = .StartJob
                
                    If (CanDo) Then
                        ProjectInfo.IDWORKING = DB_GetValueByID("SELECT MAX(ID_WORKINGLOAD) AS IDWORKINGLOAD FROM " & ProjectInfo.REF_TABLE)
                        
                        ImportData = True
                    Else
                        If (DLLParams.UnattendedMode) Then UMErrMsg = .GetUMErrorMessage
                    End If
                End With
            End If
    
            Set myPlugIn = Nothing
        End If
    
        If (FDExist(ExportFileName, False)) Then Kill ExportFileName
    End If

End Function

Public Function Init() As Boolean

    Init = DB_ConnectInit

End Function

Public Function MakeDocsMode00(idSubProject As Long, Optional PackFrom As String, Optional PackTo As String, Optional SQLWhere As String) As Boolean

    If (Trim$(SQLWhere) = "") Then If (chk_ProjectNeededParams = False) Then Exit Function
    
    If DB_GetValueByID("SELECT COUNT(*) AS NMR_BUSTE FROM " & ProjectInfo.REF_TABLE & " WHERE ID_PACCO IS NOT NULL") = 0 Then
        MsgBox "Il progetto necessita del calcolo di serializzazione.", vbExclamation, "Generate Docs:"

        Exit Function
    End If

    If FDExist(WorkPaths.TEMPLATESDIR, True) = False Then
        MsgBox "Directory Templates inesistente.", vbExclamation, "Generate Docs:"

        Exit Function
    End If

    ' Start
    '
    If chk_Array(ProjectInfo.SUBPROJECTS) Then
        Dim DocId           As String
        Dim DocSingleData() As String
        Dim ErrMsg          As String
        Dim I               As Byte
        
        ' Merge Data
        '
        If idSubProject = -1 Then
            For I = 0 To UBound(ProjectInfo.SUBPROJECTS)
                idSubProject = ProjectInfo.SUBPROJECTS(I)
            
                GoSub ExecuteProc
            Next I
        Else
            GoSub ExecuteProc
        End If
    
        ' Make Packages
        '
        If DLLParams.AutoMergePacks Then
            If (DB_PDFsPackagesMergerMode01(PackFrom, PackTo)) Then
                If (ExecuteODPPlugIn(WorkPaths.PDFPACKSDIR, ProjectInfo.IDODPPLUGIN, ProjectInfo.IDODPPLUGINPARAM) = False) Then Exit Function
            Else
                Exit Function
            End If

'            If (DB_PDFsPackagesMerger(PackFrom, PackTo)) Then
'                If (ExecuteODPPlugIn(WorkPaths.PDFPACKSDIR, ProjectInfo.IDODPPLUGIN, ProjectInfo.IDODPPLUGINPARAM) = False) Then Exit Function
'            Else
'                Exit Function
'            End If
        End If
    Else
        MsgBox "Nessun sottoprogetto definito. Impossibile proseguire.", vbExclamation, "Generate Docs:"
                
        Exit Function
    End If
    
    MakeDocsMode00 = True

    Exit Function

ExecuteProc:
    If (SQLWhere <> "") Then
        DocSingleData = Split(DB_GetValueByID("SELECT MAX(ID_WORKCNTR), MAX(ID_WORKINGLOAD), MAX(ID_PACCO) FROM " & ProjectInfo.REF_TABLE & " WHERE " & SQLWhere), "|")
        
        If chk_Array(DocSingleData) Then
            DocId = DocSingleData(0)
            PackFrom = DocSingleData(2)
            PackTo = DocSingleData(2)
            ProjectInfo.IDWORKING = DocSingleData(1)
        Else
            MsgBox "Nessun elemento trovato...", vbExclamation, "Generate Docs:"
            
            Exit Function
        End If
    End If

    If DB_SubProjectInfo_SELECT(idSubProject) Then
        chk_Directory ProjectInfo.PRJDIR & "Temporary\", 2
        chk_Directory WorkPaths.WORKINGDIR, 2
        chk_Directory WorkPaths.PDFDIR, 3
                
        ErrMsg = DB_XFDFExport(PackFrom, PackTo, DocId)
            
        If ErrMsg <> "" Then
            MsgBox ErrMsg, vbExclamation, "Generate Docs:"
            
            Exit Function
        End If
 
        chk_Directory WorkPaths.TEMPORARYDIR, 1
    Else
        MsgBox "Errore durante l'apertura del sottoprogetto.", vbExclamation, "Opening SubProject Info:"
        
        Exit Function
    End If
Return

End Function

Public Function MakeDocsMode01() As Boolean
        
    Dim MakePackages As Boolean
        
    If (chk_ProjectNeededParams = False) Then Exit Function
    
    If FDExist(WorkPaths.TEMPLATESDIR, True) = False Then
        MsgBox "Directory Templates inesistente.", vbExclamation, "Generate Docs:"

        Exit Function
    End If

    PDFDB_MergerMode01

    If DLLParams.AutoMergePacks Then If (DB_PDFsPackagesMergerMode01(-1, -1)) Then MakePackages = ExecuteODPPlugIn(WorkPaths.PDFPACKSDIR, ProjectInfo.IDODPPLUGIN, ProjectInfo.IDODPPLUGINPARAM)

    MakeDocsMode01 = True

End Function

Public Function MakeDocsMode02() As Boolean
        
    Dim MakePackages As Boolean
        
    If (chk_ProjectNeededParams = False) Then Exit Function
    
    If FDExist(WorkPaths.TEMPLATESDIR, True) = False Then
        MsgBox "Directory Templates inesistente.", vbExclamation, "Generate Docs:"

        Exit Function
    End If

    If (PDFDB_MergerMode02) Then
        If DLLParams.AutoMergePacks Then
            If (DB_PDFsPackagesMergerMode02("", "", "0")) Then
                MakePackages = ExecuteODPPlugIn(WorkPaths.PDFPACKSDIR, ProjectInfo.IDODPPLUGIN, ProjectInfo.IDODPPLUGINPARAM)
            End If
        Else
            MakePackages = True
        End If
        
        MakeDocsMode02 = MakePackages
    End If

End Function

Public Function MakePackages(mode As Integer, PackFrom As String, PackTo As String, unattended As String) As Boolean
    
    If (chk_ProjectNeededParams) Then
        Select Case mode
        Case 0, 1
            If (PackFrom = "") Then PackFrom = "-1"
            If (PackTo = "") Then PackTo = "-1"
        
            MakePackages = DB_PDFsPackagesMergerMode01(PackFrom, PackTo)

            If (MakePackages) Then MakePackages = ExecuteODPPlugIn(WorkPaths.PDFPACKSDIR, ProjectInfo.IDODPPLUGIN, ProjectInfo.IDODPPLUGINPARAM)
        
        Case 2
            If DB_PDFsPackagesMergerMode02(PackFrom, PackTo, unattended) Then MakePackages = ExecuteODPPlugIn(WorkPaths.PDFPACKSDIR, ProjectInfo.IDODPPLUGIN, ProjectInfo.IDODPPLUGINPARAM)
    
        End Select
    End If

'     MakePackages = ExecuteODPPlugIn("D:\", ProjectInfo.IDODPPLUGIN, ProjectInfo.IDODPPLUGINPARAM)
'     MakePackages = ExecuteODPPlugIn("D:\", SubProjectInfo.CFOODPPLUGIN, SubProjectInfo.CFOODPPLUGINPARAMS)

End Function

Public Function MakeReports(ByVal idSubPrj As Integer) As Boolean

    If chk_ProjectNeededParams = False Then Exit Function
    
    If ProjectInfo.IDSERIALIZEMODE = 1 Then
        If DB_GetValueByID("SELECT COUNT(*) AS nmr_Buste FROM " & ProjectInfo.REF_TABLE & " WHERE id_Pacco IS NOT NULL") = 0 Then
            MsgBox "Il progetto necessita del calcolo di serializzazione.", vbExclamation, "Generate Docs:"

            Exit Function
        End If

        MakeReports = DB_PackagesLabels(idSubPrj)
    Else
        MsgBox "Elaborazione consentita solo per le lavorazioni di postalizzazione.", vbExclamation, "Make Reports:"
    End If

End Function

Public Property Get ProjectName() As String

    ProjectName = ProjectInfo.NAME

End Property

Public Function ProjectOpen(ByVal IDPROJECT As Integer) As Boolean
    
    ProjectOpen = DB_ProjectInfo_SELECT(IDPROJECT)

    If ProjectOpen Then
        ProjectInfo.IDPROJECT = IDPROJECT
    Else
        ProjectInfo.IDPROJECT = -1
    End If

End Function

Public Function Serialize() As Boolean

    If chk_ProjectNeededParams = False Then Exit Function
    
    Serialize = DB_PackagesManagement

End Function

Public Property Let SetIDPPlugInParams(ByVal myIDPPlugInParams As String)

    If DB_ExecuteQuery("UPDATE EDT_PROJECTS SET ID_IDPPLUGINPARAMS = " & Conv_String2SQLString(myIDPPlugInParams) & " WHERE id_Project = " & ProjectInfo.IDPROJECT, True, True, True, "Update IDP ExtraParams:") Then
        ProjectInfo.IDIDPPLUGINPARAM = myIDPPlugInParams
    End If

End Property

Public Property Let SetODPPlugInParams(ByVal myODPPlugInParams As String)

    If DB_ExecuteQuery("UPDATE EDT_PROJECTS SET ID_ODPPLUGINPARAMS = " & Conv_String2SQLString(myODPPlugInParams) & " WHERE id_Project = " & ProjectInfo.IDPROJECT, True, True, True, "Update ODP ExtraParams:") Then
        ProjectInfo.IDODPPLUGINPARAM = myODPPlugInParams
    End If

End Property

Public Property Let SetUnattendedMode(ByVal varBool As Boolean)

    DLLParams.UnattendedMode = varBool

End Property

Public Property Let SetWorking(ByVal myWorking As String)

    ProjectInfo.IDWORKING = myWorking

End Property

Public Property Let TNS(ByVal myTNS As String)

    DLLParams.TSN = myTNS

End Property
