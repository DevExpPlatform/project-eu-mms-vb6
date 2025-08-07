Attribute VB_Name = "mod_Projects_DecRou"
Option Explicit

Public Function chk_ProjectNeededParams() As Boolean

    If ProjectInfo.idWorking = "Error" Or ProjectInfo.idWorking = "" Then
        MsgBox "Identificativo di lavorazione assente.", vbExclamation, "Generate Docs:"

        Exit Function
    End If
    
    chk_ProjectNeededParams = True
    
End Function
