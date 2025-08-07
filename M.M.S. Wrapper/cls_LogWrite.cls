VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_LogWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ErrMsg  As String
Dim LogFN   As Integer

Public Function LogOpen(varLogFileName) As Boolean
    
    On Error GoTo ErrHandler
    
    LogFN = FreeFile(0)

    Open varLogFileName For Append As #LogFN

    LogOpen = True

    Exit Function
    
ErrHandler:
    ErrMsg = Err.Description

End Function

Public Sub LogClose()

    Close #LogFN

End Sub

Public Sub WriteLine(varString As String)

    Print #LogFN, varString

End Sub

Public Sub WriteRow(varMessage As String, varStatus As Boolean)

    Print #LogFN, Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & TextPad(PADLEFT, varMessage, 80, " ", False) & IIf(varStatus, "[ OK ]", "[FAIL]")

End Sub
