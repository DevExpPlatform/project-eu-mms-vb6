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

Public Property Let AddFields_BarCode(ByVal XSQL_AddBarCodeFields As Boolean)
    
    DLLParams.XSQL_AddFieldsBarCode = XSQL_AddBarCodeFields

End Property

Public Property Let AddFields_Postalizzazione(ByVal XSQL_AddPostalizzazioneFields As Boolean)

    DLLParams.XSQL_AddFieldsPSTL = XSQL_AddPostalizzazioneFields

End Property

Private Sub Class_Initialize()
    
    With PlugIn_Idntfr
        .Description = "Importazione generica di flussi in formato SQL"
        .Id = "SQL"
        .Name = "SQL Data Importer"
        .Type = "Importer"
    End With

    DLL_Init
    
End Sub

Property Let DSN(ByVal XSQL_DSN As String)

    DLLParams.XSQL_DSN = XSQL_DSN

End Property

Property Let ExtraParams(aParams() As String)

    On Error GoTo ErrHandler

    DLLParams.XSQL_SubDivideChar = aParams(0)
    
    Exit Property

ErrHandler:
    SwitchError = True

    UMErrMsg = Err.Description

End Property

Public Property Get GetIdWorkingLoad() As String
    
    GetIdWorkingLoad = GET_IDWORKINGLOAD

End Property

Public Property Get GetUMErrorMessage() As String
    
    GetUMErrorMessage = UMErrMsg

End Property

Property Let idDataCutter(ByVal XSQL_IdWorking As String)

    DLLParams.XSQL_IdDataCutter = XSQL_IdWorking

End Property

Public Property Get PlugIn_GetInfo() As strct_PlugInIdentifier

    PlugIn_GetInfo = PlugIn_Idntfr
    
End Property

Public Property Let setExtError(extErr As Boolean)

    SwitchError = extErr

End Property

Public Function StartJob() As Boolean

    If (SwitchError = False) Then If DB_ConnectInit Then StartJob = DB_SQLImport

End Function

Public Property Let TableName(ByVal XSQL_TableName As String)

    DLLParams.XSQL_TableName = XSQL_TableName

End Property

Public Property Let TNS(ByVal XSQL_TNS As String)

    DLLParams.XSQL_TNS = XSQL_TNS

End Property

Public Function SQLInsert(SQLItems As String)
    
    SQLInsert = DB_InsertData(SQLItems)
    
End Function

Public Sub EndJob()

    DB_OO4OClose ((UMErrMsg <> "") Or SwitchError)

End Sub
