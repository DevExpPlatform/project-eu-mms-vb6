VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type strct_PlugInIdentifier
    Id           As String
    Description  As String
    Name         As String
    Type         As String
End Type

Private PlugIn_Idntfr   As strct_PlugInIdentifier
Private SwitchError     As Boolean

Public Property Let AddFields_BarCode(ByVal XTXT_AddBarCodeFields As Boolean)
    
    DLLParams.XTXT_AddFieldsBarCode = XTXT_AddBarCodeFields

End Property

Public Property Let AddFields_Postalizzazione(ByVal XTXT_AddPostalizzazioneFields As Boolean)

    DLLParams.XTXT_AddFieldsPSTL = XTXT_AddPostalizzazioneFields

End Property

Private Sub Class_Initialize()

    With PlugIn_Idntfr
        .Description = "Importazione generica di flussi in formato TXT"
        .Id = "TXT"
        .Name = "TXT Data Importer"
        .Type = "Importer"
    End With

    DLL_Init
    
End Sub

Public Property Let DataFileName(ByVal XTXT_DataFileName As String)

    DLLParams.XTXT_FileName = XTXT_DataFileName

End Property

Public Property Let DSN(ByVal XTXT_DSN As String)

    DLLParams.XTXT_DSN = XTXT_DSN

End Property

Public Property Let ExtraParams(aParams() As String)

    On Error GoTo ErrHandler

    'DLLParams.XTXT_PstlFields = CBool(aParams(0))

    Exit Property

ErrHandler:
    SwitchError = True

    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox Err.Description, vbExclamation, "Attenzione:"
    End If

End Property

Public Property Get GetUMErrorMessage() As String
    
    GetUMErrorMessage = UMErrMsg

End Property

Public Property Let idDataCutter(ByVal XTXT_IdWorking As String)

    DLLParams.XTXT_IdDataCutter = XTXT_IdWorking

End Property

Public Property Get PlugIn_GetInfo() As strct_PlugInIdentifier

    PlugIn_GetInfo = PlugIn_Idntfr
    
End Property

Public Function StartJob() As Boolean

    If SwitchError = False Then StartJob = DB_TXTImport
        
End Function

Public Property Let TableName(ByVal XTXT_TableName As String)

    DLLParams.XTXT_TableName = XTXT_TableName

End Property

Public Property Let TNS(ByVal XTXT_TNS As String)

    DLLParams.XTXT_TNS = XTXT_TNS

End Property

Public Property Let SetUnattendedMode(ByVal varBool As Boolean)

    DLLParams.UnattendedMode = varBool

End Property

