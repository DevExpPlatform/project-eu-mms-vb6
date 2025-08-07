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

Public Property Let AddFields_BarCode(ByVal XCSV_AddBarCodeFields As Boolean)
    
    DLLParams.XCSV_AddFieldsBarCode = XCSV_AddBarCodeFields

End Property

Public Property Let AddFields_Postalizzazione(ByVal XCSV_AddPostalizzazioneFields As Boolean)

    DLLParams.XCSV_AddFieldsPSTL = XCSV_AddPostalizzazioneFields

End Property

Private Sub Class_Initialize()
    
    With PlugIn_Idntfr
        .Description = "Importazione generica di flussi in formato CSV"
        .Id = "CSV"
        .Name = "CSV Data Importer"
        .Type = "Importer"
    End With

    DLL_Init
    
End Sub

Property Let DataFileName(ByVal XCSV_DataFileName As String)

    DLLParams.XCSV_FileName = XCSV_DataFileName

End Property

Property Let DSN(ByVal XCSV_DSN As String)

    DLLParams.XCSV_DSN = XCSV_DSN

End Property

Property Let ExtraParams(aParams() As String)

    On Error GoTo ErrHandler

    DLLParams.XCSV_SubDivideChar = aParams(0)
    'DLLParams.XCSV_PstlFields = aParams(1)
    
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

Property Let idDataCutter(ByVal XTXT_IdWorking As String)

    DLLParams.XCSV_IdDataCutter = XTXT_IdWorking

End Property

Public Property Get PlugIn_GetInfo() As strct_PlugInIdentifier

    PlugIn_GetInfo = PlugIn_Idntfr
    
End Property

Public Property Let TableName(ByVal XCSV_TableName As String)

    DLLParams.XCSV_TableName = XCSV_TableName

End Property

Public Function StartJob() As Boolean

    If SwitchError = False Then StartJob = DB_CSVImport
        
End Function

Public Property Let SetUnattendedMode(ByVal varBool As Boolean)

    DLLParams.UnattendedMode = varBool

End Property

Public Property Let TNS(ByVal XCSV_TNS As String)

    DLLParams.XCSV_TNS = XCSV_TNS

End Property
