Attribute VB_Name = "mod_Main"
Option Explicit

Public Sub Main()
    
    On Error GoTo ErrHandler
    
    Dim CanDo       As String
    Dim ErrMsg      As String
    Dim myMMS       As cls_MMS
    Dim SplitData() As String
    
    If Command$ <> Empty Then
        SplitData = Split(Command$, ", ")
        
        If chk_Array(SplitData) Then
            If (UBound(SplitData) = 1) Then
                Set myMMS = New cls_MMS
                
                With myMMS
                    .DSN = GetSetting("MMS_U", "Settings", "PrjDNS", "")
                    .TNS = GetSetting("MMS_U", "Settings", "PrjTNS", "")
                    .BaseWorkDir = GetSetting("MMS_U", "Settings", "PrjFolder", "")
                    
                    If .Init = False Then
                        Set myMMS = Nothing
                
                        End
                    End If
                End With
                
                If myMMS.ProjectOpen(SplitData(0)) Then
                    CanDo = (Trim$(SplitData(1)) <> "")
                    
                    If CanDo Then If myMMS.MakeDocsMode00(-1, -1, -1, Replace$(SplitData(1), Chr$(34), "")) Then PDF_Open myMMS.GetSinglePDFFileName
                End If
            Else
                ErrMsg = "Error Parsing Parameters"
            End If
        Else
            ErrMsg = "Error Parsing Parameters"
        End If
    Else
        ErrMsg = "Usage: MMSU IDPROJECT, SQLWHERECLAUSE"
    End If

    If ErrMsg <> "" Then MsgBox ErrMsg, vbExclamation, "MMS Unattended:"
    
    Set myMMS = Nothing
    
    Exit Sub

ErrHandler:
    Set myMMS = Nothing
    
    MsgBox Err.Description, vbExclamation, "MMS Unattended:"

End Sub
