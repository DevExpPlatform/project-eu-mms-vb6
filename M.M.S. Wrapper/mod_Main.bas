Attribute VB_Name = "mod_Main"
Option Explicit

Public Sub Main()

    On Error GoTo ErrHandler
    
    If (App.PrevInstance) Then Wait 10
    
    Dim WS_ERRCODE          As String
    Dim WS_ERRDESCR         As String
    Dim WS_EXESTATUS        As String
    Dim WS_IDWORKINGLOAD    As String
    Dim WS_LOGDESCR         As String
    Dim WS_LOGFILENAME      As String
    Dim WS_PDFWORKDIR       As String
    Dim WS_TMPSTRING        As String
    
    Dim AppPath             As String
    Dim myLogFile           As cls_LogWrite
    Dim myMMS               As cls_MMS
    Dim myMMSConfig         As cls_MMSConfig
    Dim SplitData()         As String
    
    If (Command$ <> Empty) Then
        AppPath = Fix_Paths(App.Path)
        SplitData = Split(Replace$(Command$, Chr(34), ""), "|")
        
        WS_EXESTATUS = "KO"
        
        Set myMMSConfig = New cls_MMSConfig
        
        myMMSConfig.setAppPath = AppPath
        
        If (myMMSConfig.MMSConfigOpen) Then
            Set myMMS = New cls_MMS
            
            myMMS.DSN = myMMSConfig.getPrjDNS
            myMMS.TNS = myMMSConfig.getPrjTNS
            myMMS.BaseWorkDir = myMMSConfig.getPrjFolder
            myMMS.SetUnattendedMode = True
        
            Set myLogFile = New cls_LogWrite
            
            Select Case SplitData(0)
            Case "1", "a"
                WS_LOGFILENAME = Get_LogFileName(SplitData(2))
            
            Case "3"
                WS_LOGFILENAME = SplitData(3)
            
            End Select
            
            myLogFile.LogOpen WS_LOGFILENAME
            myLogFile.WriteLine "# M.M.S. PHASE_" & Format$(SplitData(0), "00") & " LogFile Start " & Format$(Now, "dd/mm/yyyy hh.MM.ss") & vbNewLine
            
            WS_LOGDESCR = "TOOL INIT"
            
            If (myMMS.Init) Then
                myLogFile.WriteRow WS_LOGDESCR, True
            
                WS_LOGDESCR = "OPEN PROJECT ID.: " & SplitData(1)
                
                If (myMMS.ProjectOpen(SplitData(1))) Then
                    myLogFile.WriteRow WS_LOGDESCR, True
                
                    ' Phase Manager
                    '
                    Select Case SplitData(0)
                    Case "1"    ' M.M.S. Import Data & Generate PDF
                        WS_LOGDESCR = "DATA IMPORT: " & Get_BaseName(SplitData(2))
                        
                        If (myMMS.ImportData(SplitData(2))) Then
                            WS_IDWORKINGLOAD = myMMS.GetCurrentWorking
                            
                            myLogFile.WriteRow WS_LOGDESCR & " - ID.: " & WS_IDWORKINGLOAD, True
                            
                            WS_TMPSTRING = StdOutRead(Chr$(34) & AppPath & "Commands\MMSCoreEngine.exe" & Chr$(34) & " " & _
                                                      Chr$(34) & "1" & Chr$(34) & " " & _
                                                      Chr$(34) & SplitData(1) & Chr$(34) & " " & _
                                                      Chr$(34) & WS_IDWORKINGLOAD & Chr$(34) & " " & _
                                                      Chr$(34) & Replace$(myMMSConfig.getPrjFolder, "\", "/") & Chr$(34) & " " & _
                                                      Chr$(34) & "1" & Chr$(34), False)

                            WS_LOGDESCR = "PDF MAKER"

                            If (Trim$(WS_TMPSTRING) = "OK") Then
                                myLogFile.WriteRow WS_LOGDESCR, True
                
                                WS_LOGDESCR = "PDF PACKAGING:"
                                            
                                If (myMMS.MakePackages(myMMSConfig.getPrjRenderMode, "", "", "1")) Then
                                    WS_EXESTATUS = "OK"
                                    WS_PDFWORKDIR = myMMS.GetPDFWorkDir
                                    
                                    myLogFile.WriteRow WS_LOGDESCR, True
                                Else
                                    WS_ERRCODE = "MMS05"
                                    
                                    myLogFile.WriteRow WS_LOGDESCR, False
                                End If
                            Else
                                WS_ERRCODE = "MMS04"
                                WS_ERRDESCR = WS_TMPSTRING

                                myLogFile.WriteRow WS_LOGDESCR, False
                            End If
                        Else
                            WS_ERRCODE = "MMS03"
                            
                            myLogFile.WriteRow WS_LOGDESCR, False
                        End If
                
                    Case "3"    ' M.M.S. Customer File Organization
                        WS_IDWORKINGLOAD = SplitData(2)
                        myMMS.SetWorking = WS_IDWORKINGLOAD
                        
                        WS_LOGDESCR = "CUSTOMER FILE ORGANIZATION"
                        
                        If (myMMS.CustomerOrganize) Then
                            WS_EXESTATUS = "OK"

                            myLogFile.WriteRow WS_LOGDESCR, True
                        Else
                            WS_ERRCODE = "MMS06"

                            myLogFile.WriteRow WS_LOGDESCR, False
                        End If
                        
                    Case "a"
                        WS_LOGDESCR = "DATA IMPORT: " & Get_BaseName(SplitData(2))
                        
                        If (myMMS.ImportData(SplitData(2))) Then
                            WS_IDWORKINGLOAD = myMMS.GetCurrentWorking
                            
                            myLogFile.WriteRow WS_LOGDESCR & " - ID.: " & WS_IDWORKINGLOAD, True
                            
                            'AppPath = "Z:\Develope\VB Works\SE.TE.SI. Projects\!Common Projects\MassiveMailSystem\"
                            
                            WS_TMPSTRING = StdOutRead(Chr$(34) & AppPath & "Commands\MMSCoreEngine.exe" & Chr$(34) & " " & _
                                                      Chr$(34) & "1" & Chr$(34) & " " & _
                                                      Chr$(34) & SplitData(1) & Chr$(34) & " " & _
                                                      Chr$(34) & WS_IDWORKINGLOAD & Chr$(34) & " " & _
                                                      Chr$(34) & Replace$(myMMSConfig.getPrjFolder, "\", "/") & Chr$(34) & " " & _
                                                      Chr$(34) & "1" & Chr$(34), False)

                            WS_LOGDESCR = "PDF MAKER"

                            If (Trim$(WS_TMPSTRING) = "OK") Then
                                myLogFile.WriteRow WS_LOGDESCR, True
                
                                WS_LOGDESCR = "PDF PACKAGING:"
                                            
                                If (myMMS.MakePackages(myMMSConfig.getPrjRenderMode, "", "", "1")) Then
                                    WS_EXESTATUS = "OK"
                                    WS_PDFWORKDIR = myMMS.GetPDFWorkDir
                                    
                                    myLogFile.WriteRow WS_LOGDESCR, True
                                
                                    WS_LOGDESCR = "CUSTOMER FILE ORGANIZATION"
                                    
                                    If (myMMS.CustomerOrganize) Then
                                        WS_EXESTATUS = "OK"
            
                                        myLogFile.WriteRow WS_LOGDESCR, True
                                    Else
                                        WS_ERRCODE = "MMS06"
            
                                        myLogFile.WriteRow WS_LOGDESCR, False
                                    End If
                                Else
                                    WS_ERRCODE = "MMS05"
                                    
                                    myLogFile.WriteRow WS_LOGDESCR, False
                                End If
                            Else
                                WS_ERRCODE = "MMS04"
                                WS_ERRDESCR = WS_TMPSTRING

                                myLogFile.WriteRow WS_LOGDESCR, False
                            End If
                        Else
                            WS_ERRCODE = "MMS03"
                            
                            myLogFile.WriteRow WS_LOGDESCR, False
                        End If
                        
                    End Select
                Else
                    WS_ERRCODE = "MMS02"
                    
                    myLogFile.WriteRow WS_LOGDESCR, False
                End If
            Else
                WS_ERRCODE = "MMS01"
                
                myLogFile.WriteRow WS_LOGDESCR, False
            End If
            
            If (WS_ERRDESCR = "") Then WS_ERRDESCR = myMMS.GetUMErrorMessage
            
            myLogFile.WriteLine vbNewLine & "# M.M.S. PHASE_" & Format$(SplitData(0), "00") & " LogFile End   " & Format$(Now, "dd/mm/yyyy hh.MM.ss")
            myLogFile.LogClose
        
            Set myLogFile = Nothing
            Set myMMS = Nothing
        Else
            WS_ERRCODE = "MMS00"
            WS_ERRDESCR = myMMSConfig.getErrMsg
        End If
        
        Set myMMSConfig = Nothing
        
        StdOutWrite WS_IDWORKINGLOAD & "|" & WS_EXESTATUS & "|" & WS_ERRCODE & "|" & WS_ERRDESCR & "|" & WS_LOGFILENAME & IIf(WS_PDFWORKDIR = "", "", "|" & WS_PDFWORKDIR)
    Else
        StdOutWrite "M.M.S. Unattended Wrapper - Vers. 0.7.0 [09/04/2024]" & vbNewLine & vbNewLine & _
                    "Usage PHASE_01 (DATA IMPORT + PDF MAKER):" & vbNewLine & _
                    "1|PROJECT ID.|DATA FILE PATH" & vbNewLine & vbNewLine & _
                    "Usage PHASE_02 (CUSTOMER FILE ORGANIZER):" & vbNewLine & _
                    "3|PROJECT ID.|ID. WORKINGLOAD|LOG FILE PATH" & vbNewLine & vbNewLine & _
                    "Usage PHASE_03 (DATA IMPORT + PDF MAKER + CUSTOMER FILE ORGANIZER):" & vbNewLine & _
                    "a|PROJECT ID.|DATA FILE PATH"
    End If

    Exit Sub
    
ErrHandler:
    StdOutWrite Err.Description

End Sub
