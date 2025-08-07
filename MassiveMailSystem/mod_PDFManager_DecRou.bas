Attribute VB_Name = "mod_PDFManager_DecRou"
Option Explicit

Public Function PDF_DBMergerMode01(ByVal BASEDIR As String, ByVal BASEFILENAME As String, ByVal workingTable As String, ByVal IDWORKING As String, ByVal idPackage As String, ByVal OutPutFileName As String) As Boolean

    If FDExist(OutPutFileName, False) Then Kill OutPutFileName
    If (Right$(BASEDIR, 1) = "\") Then BASEDIR = Left$(BASEDIR, Len(BASEDIR) - 1)

    ExecuteAndWait Chr$(34) & AppPath & "Commands\PDFDBMerger.EXE" & Chr$(34) & " " & _
                   Chr$(34) & BASEDIR & Chr$(34) & " " & _
                   BASEFILENAME & " " & _
                   workingTable & " " & _
                   IDWORKING & " " & _
                   idPackage

    PDF_DBMergerMode01 = FDExist(OutPutFileName, False)

    DoEvents

End Function

Public Function PDF_DirMerger(ByVal PDFsDirectory As String, ByVal OutPutFileName As String) As Boolean

    If FDExist(OutPutFileName, False) Then Kill OutPutFileName

    'ExecuteAndWait Chr$(34) & AppPath & "Commands\pdftk.exe" & Chr$(34) & " " & Fix_Paths(PDFsDirectory) & "*.PDF" & _
                   " cat output " & Chr$(34) & OutPutFileName & Chr$(34), 0

    If (Right$(PDFsDirectory, 1) = "\") Then PDFsDirectory = Left$(PDFsDirectory, Len(PDFsDirectory) - 1)

    'ExecuteAndWait Chr$(34) & AppPath & "Commands\PDFDirMerger.EXE" & Chr$(34) & " " & Chr$(34) & PDFsDirectory & Chr$(34) & " " & _
                                                                   Chr$(34) & OutPutFileName & Chr$(34), 0

    PDF_DirMerger = FDExist(OutPutFileName, False)

    DoEvents

End Function

Public Function PDF_Merge2File(ByVal PDFAFName As String, ByVal PDFBFName As String, ByVal PDFCFName As String) As Boolean
    
    Dim KillSrc As Boolean
    
    KillSrc = (PDFCFName = PDFAFName)
    
    If KillSrc Then PDFCFName = Get_PathName(PDFCFName) & Get_BaseName(PDFCFName, 4) & "_TMP.PDF"
    
    'ExecuteAndWait Chr$(34) & AppPath & "Commands\pdftk.exe" & Chr$(34) & _
                   " " & Chr$(34) & PDFAFName & Chr$(34) & _
                   " " & Chr$(34) & PDFBFName & Chr$(34) & _
                   " cat output " & Chr$(34) & PDFCFName & Chr$(34), 0
    
    If KillSrc Then
        Kill PDFAFName
        Name PDFCFName As PDFAFName
    End If
    
    PDF_Merge2File = FDExist(PDFAFName, False)
 
End Function

Public Function PDF_PackageMerger(PackStart As String, PackEnd As String, unattended As String) As Boolean
    
    Dim RValue As String
    
    RValue = StdOutRead(Chr$(34) & AppPath & "Commands\MMSCoreEngine.exe" & Chr$(34) & " " & _
                        Chr$(34) & "2" & Chr$(34) & " " & _
                        Chr$(34) & ProjectInfo.IDPROJECT & Chr$(34) & " " & _
                        Chr$(34) & ProjectInfo.IDWORKING & Chr$(34) & " " & _
                        Chr$(34) & Replace$(DLLParams.BaseWorkDir, "\", "/") & Chr$(34) & " " & _
                        Chr$(34) & PackStart & Chr$(34) & " " & _
                        Chr$(34) & PackEnd & Chr$(34) & " " & _
                        Chr$(34) & unattended & Chr$(34))
    
    PDF_PackageMerger = (RValue = "OK")
    
    DoEvents

End Function

Public Function PDF_Packer(ByVal PDFFName As String) As Boolean

    Dim PDFCName As String

    PDFCName = Left$(PDFFName, Len(PDFFName) - 4) & "-O.PDF"

    'ExecuteAndWait "java -Xms128m -Xmx768m -cp " & Chr$(34) & AppPath & "Commands\Multivalent.jar" & Chr$(34) & " tool.pdf.Compress -compatible -quiet " & Chr$(34) & PDFFName & Chr$(34), 0
    ExecuteAndWait Chr$(34) & AppPath & "Commands\PDFPacker.EXE" & Chr$(34) & " " & Chr$(34) & PDFFName & Chr$(34)
    
    If FDExist(PDFCName, False) Then
        If (FileLen(PDFCName) > 0) Then
            Kill PDFFName
            Name PDFCName As PDFFName
            
            PDF_Packer = FDExist(PDFFName, False)
        End If
    End If

    DoEvents

End Function

Public Function PDF_SeqMerger(ByVal PDFSequence As Variant, ByVal OutPutFileName As String) As Boolean
    
    If FDExist(OutPutFileName, False) Then Kill OutPutFileName

    Dim I           As Integer
    Dim SeqString   As String
    
    For I = 0 To UBound(PDFSequence)
        SeqString = SeqString & " " & Chr$(34) & PDFSequence(I) & Chr$(34)
    Next I
    
    'ExecuteAndWait Chr$(34) & AppPath & "Commands\pdftk.exe" & Chr$(34) & SeqString & _
                   " cat output " & Chr$(34) & OutPutFileName & Chr$(34), 0

    PDF_SeqMerger = FDExist(OutPutFileName, False)
    
    DoEvents

End Function

Public Function PDF_TemplatesMerger() As String
    
    chk_Directory WorkPaths.TEMPORARYDIR, 3

    Dim I               As Integer
    Dim MaxTemplates    As Integer

    MaxTemplates = UBound(TemplatesInfo)
    SubProjectInfo.MERGEDDOCNAME = WorkPaths.TEMPORARYDIR & SubProjectInfo.BASEFILENAME & ".PDF"
    TemplateInfo.TEMPLATEFNAME = SubProjectInfo.MERGEDDOCNAME

    If MaxTemplates = 0 Then
        FileCopy WorkPaths.TEMPLATESDIR & TemplatesInfo(0).TEMPLATEFNAME, SubProjectInfo.MERGEDDOCNAME
    
        TemplateInfo.SHEETS = TemplatesInfo(0).SHEETS
    Else
        Dim PDFSeq() As String

        ReDim PDFSeq(MaxTemplates) As String

        For I = 0 To MaxTemplates
            If FDExist(WorkPaths.TEMPLATESDIR & TemplatesInfo(I).TEMPLATEFNAME, False) Then
                PDFSeq(I) = WorkPaths.TEMPLATESDIR & TemplatesInfo(I).TEMPLATEFNAME
            
                TemplateInfo.SHEETS = TemplateInfo.SHEETS + TemplatesInfo(I).SHEETS
            Else
                Erase PDFSeq

                PDF_TemplatesMerger = "File " & TemplatesInfo(I).TEMPLATEFNAME & " mancante."

                Exit Function
            End If
        Next I

        If PDF_SeqMerger(PDFSeq, SubProjectInfo.MERGEDDOCNAME) = False Then
            Erase PDFSeq
            
            PDF_TemplatesMerger = "Errore durante il merging dei templates."

            Exit Function
        End If
        
        Erase PDFSeq
    End If

End Function
