Attribute VB_Name = "mod_Projects_DBRtns"
Option Explicit

Private Type strct_PDFInfo
    SINGLEPDFFILENAME   As String
End Type

Private Type strct_ProjectInfo
    BARCODETABLE        As String
    BARCODETEXT         As Boolean
    BARCODETYPE         As String
    IDDATACUTTER        As Integer
    IDIDPPLUGIN         As String
    IDIDPPLUGINPARAM    As String
    IDODPPLUGIN         As String
    IDODPPLUGINPARAM    As String
    IDPROJECT           As Integer
    IDSERIALIZEMODE     As Byte
    IDWORKING           As String
    NAME                As String
    NUMBUSTEMAX         As Integer
    NUMBUSTEMIN         As Integer
    NUMBUSTEMIXMIN      As Integer
    PESOPRODOTTO        As Byte
    PlugIn              As String
    PLUGINPARAMS        As String
    PRJDIR              As String
    PSTLIDJOB           As String
    PSTLIDHOMOLOGATION  As String
    PSTLFIELD           As String
    PSTLENVTYPE         As String
    PSTLORDERBY         As String
    REF_TABLE           As String
    SUBPROJECTS()       As Long
End Type

Private Type strct_SubProjectInfo
    BASEFILENAME        As String
    CFOMODE             As Byte
    CFOODPPLUGIN        As String
    CFOODPPLUGINPARAMS  As String
    CFOFIXEDPATH        As Integer
    MERGEDDOCNAME       As String
    NAME                As String
    OMRGEN              As Boolean
    OMRTYPE             As String
    PACKAGEMAILCHECK    As Boolean
    SINGLEPDFPACKING    As Boolean
    PACKAGEPACKING      As Boolean
    QUERYFILTER         As strct_FV
    SHEETS              As Byte
    SUBPRJID            As Long
    SUBPRJDIR           As String
    TEMPLATES()         As Long
    TEMPQFILTER()       As strct_FV
End Type

Private Type strct_TemplateInfo
    SHEETS              As Integer
    TEMPLATEFNAME       As String
End Type

Private Type strct_WorkPaths
    BASEDIR             As String
    CFODIR              As String
    PDFDIR              As String
    PDFPACKSDIR         As String
    REPORTSDIR          As String
    TEMPORARYDIR        As String
    TEMPLATESDIR        As String
    WORKINGDIR          As String
End Type

Public PDFInfo          As strct_PDFInfo
Public ProjectInfo      As strct_ProjectInfo
Public SubProjectInfo   As strct_SubProjectInfo
Public TemplateInfo     As strct_TemplateInfo
Public TemplatesInfo()  As strct_TemplateInfo
Public WorkPaths        As strct_WorkPaths

'Private Function DB_Project_GetWeight(ByVal idProject As Integer) As Byte
'
'    Dim RS As ADODB.Recordset
'
'    Set RS = DBConn.Execute("SELECT nmr_Sheets, nmr_FoglioPeso, flg_EmptyDocs, flg_ClosePacks FROM view_TemplatesWeight WHERE id_Project = " & idProject)''
'
'    If RS.RecordCount > 0 Then
'        Dim SubPrjSingle    As Byte
'        Dim SubPrjTemplates As Byte
'
'        Do Until RS.EOF
'            If RS("flg_EmptyDocs") And RS("flg_ClosePacks") Then
'                SubPrjSingle = SubPrjSingle + RS("nmr_Sheets") * RS("nmr_FoglioPeso")
'            Else
'                If (RS("nmr_Sheets") * RS("nmr_FoglioPeso")) > SubPrjTemplates Then
'                    SubPrjTemplates = RS("nmr_Sheets") * RS("nmr_FoglioPeso")
'                End If
'            End If
'
'            RS.MoveNext
'        Loop
'
'        DB_Project_GetWeight = (SubPrjSingle + SubPrjTemplates)
'    End If
'
'    RS.Close
'
'    Set RS = Nothing
'
'End Function

Public Function DB_ProjectInfo_SELECT(ByVal IDPROJECT As Integer) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim I                   As Integer
    Dim ProjectInfo_EMPTY   As strct_ProjectInfo
    Dim RS                  As ADODB.Recordset
    Dim SplitData()         As String
    Dim SubSplitData()      As String
    
    ProjectInfo = ProjectInfo_EMPTY
   
    Erase XFDF_ExtFields
    
    DBConn.Open
    
    ' Load Project Params
    '
    Set RS = DBConn.Execute("SELECT * FROM VIEW_PROJECTS_DLL WHERE ID_PROJECT = " & IDPROJECT)
        
    If RS.RecordCount = 1 Then
        With ProjectInfo
             .NAME = RS("descr_Project")
             
             .PRJDIR = Fix_Paths(DLLParams.BaseWorkDir & RS("str_PrjWorkDir"))
             
             If Not IsNull(RS("id_JobId")) Then .PSTLIDJOB = RS("id_JobId")
             
             ' Ref. Table
             '
             .REF_TABLE = RS("str_RefTableName")
             
             If Not IsNull(RS("id_DataCutter")) Then
                 .IDDATACUTTER = RS("id_DataCutter")
                 
                 If Not IsNull(RS("id_IDPPlugIn")) Then
                    .IDIDPPLUGIN = RS("id_IDPPlugIn")
                    
                    If Not IsNull(RS("id_IDPPlugInParams")) Then .IDIDPPLUGINPARAM = RS("id_IDPPlugInParams")
                 End If
                 
                 If Not IsNull(RS("id_ODPPlugIn")) Then
                    .IDODPPLUGIN = RS("id_ODPPlugIn")
                    
                    If Not IsNull(RS("id_ODPPlugInParams")) Then .IDODPPLUGINPARAM = RS("id_ODPPlugInParams")
                 End If
                 
                 .PlugIn = RS("id_PlugIn")
                 If Not IsNull(RS("str_PlugInParams")) Then .PLUGINPARAMS = RS("str_PlugInParams")
             End If
             
             ' Serialize Mode
             '
             .IDSERIALIZEMODE = RS("id_SerializeMode")
             
             If .IDSERIALIZEMODE = 1 And .IDDATACUTTER > 0 Then
                 .PSTLFIELD = RS("str_PstlField")
                 
                 If RS("flg_PstlExtraSort") Then
                     SplitData = Split(RS("str_OrderFields"), "|")
                     
                     If chk_Array(SplitData) Then
                         For I = 0 To UBound(SplitData)
                             SubSplitData = Split(SplitData(I), ";")
                             
                             .PSTLORDERBY = .PSTLORDERBY & IIf(.PSTLORDERBY = "", "", ", ") & SubSplitData(0) & IIf(SubSplitData(1) = "", " ASC", " DESC")
                         Next I
                     End If
                 End If
             End If
             
             ' BarCode Serialize
             '
             If Not IsNull(RS("FLG_BARCODETYPE")) Then
                 .BARCODETABLE = RS("STR_BARCODETABLE")
                 .BARCODETEXT = RS("FLG_BARCODETXT")
                 .BARCODETYPE = RS("FLG_BARCODETYPE")
             End If
             
             ' Packaging
             '
             If Not IsNull(RS("id_Omologazione")) Then .PSTLIDHOMOLOGATION = RS("id_Omologazione")
             
             .PSTLENVTYPE = RS("id_EnvType")
             
             If Not IsNull(RS("nmr_ProdottoPeso")) Then
                 Dim PesoMaxScatola      As Integer
                 Dim PesoMinMixScatola   As Integer
                 Dim PesoMinScatola      As Integer
                 
                 If RS("flg_ApplyTolerance") Then
                     PesoMinScatola = (RS("nmr_ScatolaPesoMin") * 1000) + (RS("nmr_ScatolaPesoMin") * RS("nmr_ScatolaTollMax") * 10)
                     PesoMaxScatola = (RS("nmr_ScatolaPesoMax") * 1000) + (RS("nmr_ScatolaPesoMax") * RS("nmr_ScatolaTollMax") * 10)
                     PesoMinMixScatola = (RS("nmr_ScatolaMixPesoMin") * 1000) + (RS("nmr_ScatolaPesoMax") * RS("nmr_ScatolaTollMax") * 10)
                 Else
                     PesoMinScatola = RS("nmr_ScatolaPesoMin") * 1000
                     PesoMaxScatola = RS("nmr_ScatolaPesoMax") * 1000
                     PesoMinMixScatola = (RS("nmr_ScatolaMixPesoMin") * 1000)
                 End If
                 
                 .PESOPRODOTTO = RS("nmr_ProdottoPeso") + RS("nmr_PrjWeight")
                 
                 .NUMBUSTEMAX = Fix(PesoMaxScatola / .PESOPRODOTTO)
                 .NUMBUSTEMIXMIN = Fix(PesoMinMixScatola / .PESOPRODOTTO)
                 .NUMBUSTEMIN = Fix(PesoMinScatola / .PESOPRODOTTO)
             End If
             
             ' Get Preliminary SubProjects Info
             '
             Set RS = DBConn.Execute("SELECT ID_SUBPROJECT FROM EDT_SUBPROJECTS WHERE ID_PROJECT = " & IDPROJECT & " ORDER BY ID_PROJECT")
             
             If RS.RecordCount > 0 Then
                 ReDim .SUBPROJECTS(RS.RecordCount - 1)
                             
                 Do Until RS.EOF
                     .SUBPROJECTS(RS.AbsolutePosition - 1) = RS("id_SubProject")
                     
                     RS.MoveNext
                 Loop
             End If
        End With
        
        With WorkPaths
            .TEMPLATESDIR = ProjectInfo.PRJDIR & "Templates\"
            .TEMPORARYDIR = ProjectInfo.PRJDIR & "Temporary\" ' & get_UsrName & "_" & Format$(Now, "YYYYMMDDhhmmss") & "\"
        End With
    End If
    
    RS.Close
        
    GoSub CleanUp
    
    DB_ProjectInfo_SELECT = True
    
    Exit Function
    
CleanUp:
    DBConn.Close

    Set RS = Nothing

    Erase SplitData
    Erase SubSplitData
Return
    
ErrHandler:
    GoSub CleanUp

    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Attenzione:"
    End If

End Function

Public Function DB_SubProjectInfo_SELECT(ByVal idSubProject As Integer) As Boolean

    On Error GoTo ErrHandler

    Dim CanDisconnect        As Boolean
    Dim SubProjectInfo_EMPTY As strct_SubProjectInfo
    Dim RS                   As ADODB.Recordset
    
    SubProjectInfo = SubProjectInfo_EMPTY
    
    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
    
    Set RS = DBConn.Execute("SELECT * FROM EDT_SUBPROJECTS WHERE ID_SUBPROJECT = " & idSubProject)
    
    ' Load Project Params
    '
    If RS.RecordCount = 1 Then
        Dim CurPos      As Integer
        
        With SubProjectInfo
            .SUBPRJID = idSubProject
            .NAME = RS("DESCR_SUBPROJECT")
            .BASEFILENAME = RS("STR_BASEFILENAME")
            .OMRGEN = RS("FLG_OMRGEN")
            .OMRTYPE = RS("STR_OMRTYPE")
            .SINGLEPDFPACKING = RS("FLG_SINGLEPDFPACKING")
            .PACKAGEPACKING = RS("FLG_PACKAGEPACKING")
            .SUBPRJDIR = RS("STR_SUBPRJWORKDIR")

            If (Not IsNull(RS("FLG_CFOFIXEDPATH"))) Then SubProjectInfo.CFOFIXEDPATH = RS("FLG_CFOFIXEDPATH")
            If (Not IsNull(RS("ID_OPDCFOPLUGIN"))) Then SubProjectInfo.CFOODPPLUGIN = RS("ID_OPDCFOPLUGIN")
            If (Not IsNull(RS("ID_OPDCFOPLUGINPARAMS"))) Then SubProjectInfo.CFOODPPLUGINPARAMS = RS("ID_OPDCFOPLUGINPARAMS")
            
            ' Load Templates Info
            '
            Set RS = DBConn.Execute("SELECT ID_TEMPLATE, STR_QFIELD, STR_QVALUE FROM REF_TEMPLATES WHERE ID_SUBPROJECT = " & idSubProject & " ORDER BY STR_QFIELD, STR_QVALUE")

            If RS.RecordCount > 0 Then
                ReDim .TEMPQFILTER(RS.RecordCount - 1)
                ReDim .TEMPLATES(RS.RecordCount - 1)
            
                Do Until RS.EOF
                    CurPos = RS.AbsolutePosition - 1
                            
                    If Not IsNull(RS("str_QField")) Then
                        .TEMPQFILTER(CurPos).Field = RS("str_QField")
                        .TEMPQFILTER(CurPos).Value = RS("str_QValue")
                    End If
                    
                    .TEMPLATES(CurPos) = RS("id_Template")
            
                    RS.MoveNext
                Loop
            Else
                GoSub CleanUp
            
                Exit Function
            End If
        End With
        
        With WorkPaths
            .BASEDIR = ProjectInfo.PRJDIR & SubProjectInfo.SUBPRJDIR & "\"
            .WORKINGDIR = ProjectInfo.PRJDIR & SubProjectInfo.SUBPRJDIR & "\" & ProjectInfo.IDWORKING
            .PDFDIR = .WORKINGDIR & "\Workings\"
            .PDFPACKSDIR = .WORKINGDIR & "\Packages\"
            .REPORTSDIR = .WORKINGDIR & "\Reports\"
            
            chk_Directory .BASEDIR, 2
        End With
    End If

    RS.Close
        
    GoSub CleanUp
        
    DB_SubProjectInfo_SELECT = True
    
    Exit Function
    
CleanUp:
    Set RS = Nothing
    
    If CanDisconnect Then DBConn.Close
Return
    
ErrHandler:
    GoSub CleanUp
    
End Function

Public Function DB_GetTemplatesInfo_SELECT(ByVal idTemplate As Long) As String
            
    On Error GoTo ErrHandler
            
    Dim RS  As ADODB.Recordset
    
    Erase TemplatesInfo
    
    SubProjectInfo.SHEETS = 0
    
'    DBConn.Open
            
    Set RS = DBConn.Execute("SELECT str_TemplateFileName, nmr_Sheets FROM edt_Templates WHERE id_Template = " & idTemplate & " ORDER BY nmr_TemplateOrder ASC")
        
    If RS.RecordCount > 0 Then
        ReDim TemplatesInfo(RS.RecordCount - 1)
        
        Do Until RS.EOF
            With TemplatesInfo(RS.AbsolutePosition - 1)
                .SHEETS = RS("nmr_Sheets")
                .TEMPLATEFNAME = RS("str_TemplateFileName")
            End With
        
            SubProjectInfo.SHEETS = SubProjectInfo.SHEETS + RS("nmr_Sheets")
            
            RS.MoveNext
        Loop
    End If

    RS.Close
        
    GoSub CleanUp

    Exit Function
    
CleanUp:
    Set RS = Nothing

 '   DBConn.Close
Return

ErrHandler:
    GoSub CleanUp
    
    DB_GetTemplatesInfo_SELECT = Purge_ErrDescr(Err.Description)
    
End Function

Public Function DB_GetWorkings() As String()

    On Error GoTo ErrHandler

    Dim RS              As ADODB.Recordset
    Dim tmpWorkings()   As String
       
    DBConn.Open
 
    Set RS = DBConn.Execute("SELECT ID_WORKINGLOAD FROM " & ProjectInfo.REF_TABLE & " GROUP BY ID_WORKINGLOAD ORDER BY ID_WORKINGLOAD")
        
    If RS.RecordCount > 0 Then
        ReDim tmpWorkings(RS.RecordCount - 1)
        
        Do Until RS.EOF
            tmpWorkings(RS.AbsolutePosition - 1) = RS("id_WorkingLoad")
            
            RS.MoveNext
        Loop
    End If

    RS.Close
        
    GoSub CleanUp

    DB_GetWorkings = tmpWorkings

    Exit Function
    
CleanUp:
    Set RS = Nothing

    DBConn.Close
Return

ErrHandler:
    GoSub CleanUp
    
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation

End Function
