VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_MMSConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim AppPath         As String
Dim ErrMsg          As String
Dim PrjDNS          As String
Dim PrjFolder       As String
Dim PrjFilter       As String
Dim PrjLocale       As String
Dim PrjRenderMode   As Integer
Dim PrjTNS          As String

Public Function MMSConfigOpen() As Boolean

    On Error GoTo ErrHandler

    Dim FNI             As Integer
    Dim ConfigFileName  As String
    Dim StrIn           As String
    Dim SplitData()     As String
    
    FNI = FreeFile(0)
    ConfigFileName = AppPath & "MMS.CFG"
    
    If (FDExist(ConfigFileName, False)) Then
        Open ConfigFileName For Input As #FNI
            Do Until EOF(FNI)
                Line Input #FNI, StrIn
                
                SplitData = Split(StrIn, ",")
                    
                Select Case SplitData(0)
                Case "PrjDNS"
                    PrjDNS = SplitData(1)
                
                Case "PrjFolder"
                    PrjFolder = SplitData(1)
                
                Case "PrjFilter"
                    PrjFilter = SplitData(1)
                
                Case "PrjLocale"
                    PrjLocale = SplitData(1)
                
                Case "PrjRenderMode"
                    PrjRenderMode = SplitData(1)
                
                Case "PrjTNS"
                    PrjTNS = SplitData(1)
                
                End Select
            Loop
        Close #FNI
    Else
        ErrMsg = "M.M.S. Configuration file not found."
            
        Exit Function
    End If

    MMSConfigOpen = True
    
    Exit Function

ErrHandler:
    Close #FNI
    
    ErrMsg = Err.Description
    
End Function

Public Property Get getErrMsg() As String

    getErrMsg = ErrMsg

End Property

Public Property Get getPrjDNS() As String

    getPrjDNS = PrjDNS

End Property

Public Property Get getPrjFolder() As String

    getPrjFolder = PrjFolder

End Property

Public Property Get getPrjFilter() As String

    getPrjFilter = PrjFilter

End Property

Public Property Get getPrjLocale() As String

    getPrjLocale = PrjLocale

End Property

Public Property Get getPrjRenderMode() As Integer

    getPrjRenderMode = PrjRenderMode

End Property

Public Property Get getPrjTNS() As String

    getPrjTNS = PrjTNS

End Property

Public Property Let setAppPath(ByVal varValue As String)

    AppPath = varValue

End Property
