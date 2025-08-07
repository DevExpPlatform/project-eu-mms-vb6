Attribute VB_Name = "mod_XFDFManager_DecRou"
Option Explicit

'Public Function XFDF_ExtBarCode(ByVal BC_TYPE As String, ByVal Code As String, ByRef OutPutFileName As String) As String
'
'    Dim CmdParams As String
'    Dim RValue    As String
'
'    Select Case BC_TYPE
'        Case "CODE128"
'            CmdParams = "-b 20"
'            RValue = "_C128"
'
'        Case "DATAMATRIX"
'            CmdParams = "-b 71 --vers=30"
'            RValue = "_DM"
'
'    End Select
'
'    OutPutFileName = OutPutFileName & RValue & ".PNG"
'
'    If FDExist(OutPutFileName, False) Then Kill OutPutFileName
'
'    ExecuteAndWait Chr$(34) & AppPath & "Commands\zint.exe" & Chr$(34) & " " & CmdParams & " --notext -o " & Chr$(34) & OutPutFileName & Chr$(34) & " -d " & Code, 0
'
'    If (FDExist(OutPutFileName, False)) Then XFDF_ExtBarCode = RValue
'
'End Function

Public Sub PDFDB_MergerMode01()

    ExecuteAndWait Chr$(34) & AppPath & "Commands\PDFFlatten.EXE" & Chr$(34) & " " & _
                   Chr$(34) & "1" & Chr$(34) & " " & _
                   Chr$(34) & ProjectInfo.IDPROJECT & Chr$(34) & " " & _
                   Chr$(34) & ProjectInfo.IDWORKING & Chr$(34) & " " & _
                   Chr$(34) & Replace$(DLLParams.BaseWorkDir, "\", "/") & Chr$(34) & " " & _
                   Chr$(34) & "0" & Chr$(34)
    
    DoEvents

End Sub

Public Function PDFDB_MergerMode02() As Boolean
    
    Dim RValue As String
    
    RValue = StdOutRead(Chr$(34) & AppPath & "Commands\MMSCoreEngine.exe" & Chr$(34) & " " & _
                   Chr$(34) & "1" & Chr$(34) & " " & _
                   Chr$(34) & ProjectInfo.IDPROJECT & Chr$(34) & " " & _
                   Chr$(34) & ProjectInfo.IDWORKING & Chr$(34) & " " & _
                   Chr$(34) & Replace$(DLLParams.BaseWorkDir, "\", "/") & Chr$(34) & " " & _
                   Chr$(34) & "0" & Chr$(34))

    PDFDB_MergerMode02 = (RValue = "OK")
    
End Function

Public Function XFDF_Merger(ByVal XFDFFileName As String, ByVal XFDFDelete As Boolean, ByVal OutPutFileName As String) As Boolean
    
    If FDExist(OutPutFileName, False) Then Kill OutPutFileName
    
    'ExecuteAndWait Chr$(34) & AppPath & "Commands\xmf.exe" & Chr$(34) & " -f " & Chr$(34) & TemplateFName & Chr$(34) & " " & _
                   Chr$(34) & OutPutFileName & Chr$(34) & " " & _
                   Chr$(34) & XFDFFileName & Chr$(34), 0
    
    'ExecuteAndWait Chr$(34) & AppPath & "Commands\pdftk.exe" & Chr$(34) & " " & Chr$(34) & TemplateFName & Chr$(34) & " fill_form " & _
                   Chr$(34) & XFDFFileName & Chr$(34) & " output " & _
                   Chr$(34) & OutPutFileName & Chr$(34) & " flatten", 0
    
    ExecuteAndWait Chr$(34) & AppPath & "Commands\PDFFlatten.EXE" & Chr$(34) & " " & Chr$(34) & "0" & Chr$(34) & " " & Chr$(34) & XFDFFileName & Chr$(34) & " " & _
                                                                    Chr$(34) & OutPutFileName & Chr$(34)
    
    DoEvents
    
    XFDF_Merger = FDExist(OutPutFileName, False)
    
    If XFDFDelete Then Kill XFDFFileName

End Function
