Attribute VB_Name = "mod_Localization_DecRou"
Option Explicit

Public LBL_IBXPARAMSTITLE   As String
Public LBL_MSGBOX01         As String
Public LBL_MSGBOX02         As String
Public LBL_MSGBOX03         As String
Public LBL_MSGTTL01         As String
Public LBL_MSGTTL02         As String

Public Sub GUI_GetLocalization()

    On Error GoTo ErrHandler

    Dim FNI             As Integer
    Dim LocaleFileName  As String
    Dim StrIn           As String
    Dim SplitData()     As String
    
    FNI = 1
    LocaleFileName = AppPath & "Locales\" & AppSettings.PrjLocale & ".LCL"
    
    If (FDExist(LocaleFileName, False)) Then
        Open LocaleFileName For Input As #FNI
            Do Until EOF(FNI)
                Line Input #FNI, StrIn
                
                SplitData = Split(StrIn, ";")
            
                With frm_Main
                    Select Case SplitData(0)
                    Case "LBL_PROJECT"
                        .mnu_Prj.Caption = SplitData(1)
                    
                    Case "LBL_MNUIDP"
                        .mnu_SetDLLIDPExtraParams.Caption = SplitData(1)
                    
                    Case "LBL_MNUODP"
                        .mnu_SetDLLODPExtraParams.Caption = SplitData(1)
                    
                    Case "LBL_IBXPARAMSTITLE"
                        LBL_IBXPARAMSTITLE = SplitData(1)
                    
                    Case "LBL_IBXPARAMSDESCR"
                        LBL_IBXPARAMSTITLE = SplitData(1)
                        
                    Case "LBL_FMEBOX1"
                        .fme_AvailableProjects.Caption = SplitData(1)
                        
                    Case "LBL_FMEBOX2"
                        .fme_AvailableImports.Caption = SplitData(1)
                    
                    Case "LBL_FMEBOX3"
                        .fme_CommandsConsolle.Caption = SplitData(1)
                    
                    Case "LBL_CMBBOX1"
                        .lbl_Descr(0).Caption = SplitData(1)
                        
                    Case "LBL_CMBBOX2"
                        .lbl_Descr(1).Caption = SplitData(1)
                    
                    Case "LBL_BTNBOX2AI"
                        .cmd_PrjAddImport.Caption = SplitData(1)
                    
                    Case "LBL_BTNBOX2MS"
                        .cmd_ExecuteProcess(0).Caption = SplitData(1)
                        
                    Case "LBL_BTNBOX3MR"
                        .cmd_ExecuteProcess(2).Caption = SplitData(1)
                    
                    Case "LBL_BTNBOX3GD"
                        .cmd_ExecuteProcess(1).Caption = SplitData(1)
                    
                    Case "LBL_MSGBOX01"
                        LBL_MSGBOX01 = SplitData(1)
                    
                    Case "LBL_MSGBOX02"
                        LBL_MSGBOX02 = SplitData(1)
                    
                    Case "LBL_MSGBOX03"
                        LBL_MSGBOX03 = SplitData(1)
                    
                    Case "LBL_MSGTTL01"
                        LBL_MSGTTL01 = SplitData(1)
                    
                    Case "LBL_MSGTTL02"
                        LBL_MSGTTL02 = SplitData(1)
                    
                    End Select
                End With
            Loop
        Close #FNI
    End If

    Exit Sub

ErrHandler:
    Close #FNI
    
    MsgBox Err.Description, vbExclamation, "Attenzione:"

End Sub
