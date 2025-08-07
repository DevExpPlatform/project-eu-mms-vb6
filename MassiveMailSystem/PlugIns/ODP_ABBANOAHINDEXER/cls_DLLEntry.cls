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
    ID           As String
    DESCRIPTION  As String
    NAME         As String
    TYPE         As String
End Type

Private PlugIn_Idntfr As strct_PlugInIdentifier

Private Sub Class_Initialize()

    With PlugIn_Idntfr
        .DESCRIPTION = "ABBANOA H2O - Billing Indexer - v. 1.03.20190709"
        .ID = "ABBANOAHINDEXER"
        .NAME = "ABBANOA H2O Indexer"
        .TYPE = "Exporter"
    End With

End Sub

Private Sub Class_Terminate()

    DB_ConnectRelease

End Sub

Public Property Let DNS(ByVal var_DNS As String)
    
    DLLParams.DSN = var_DNS

End Property

Property Let ExtraParams(aParams() As String)

    On Error GoTo ErrHandler

    With DLLParams
        .BASEFILENAME = aParams(0)
        .ZIPEXEPATH = aParams(1)
        .MODE = aParams(2)
        .TYPE = aParams(3)
    End With
    
    Exit Property

ErrHandler:
    MsgBox Err.DESCRIPTION, vbExclamation, "Attenzione:"

End Property

Public Property Let OutputFileName(ByVal OFName As String)
    
    DLLParams.OUTPUTFILEPATH = OFName

End Property

Public Property Get PlugIn_GetInfo() As strct_PlugInIdentifier

    PlugIn_GetInfo = PlugIn_Idntfr
    
End Property

Public Function StartJob() As Boolean

    If DB_ConnectInit Then StartJob = ODP_ABBANOAHINDEXER_Exporter

End Function

Public Property Let WORKINGID(ByVal var_WorkingId As String)

    DLLParams.WORKINGID = var_WorkingId

End Property

Public Property Let WORKINGTABLE(ByVal var_WorkingTable As String)

    DLLParams.WORKINGTABLE = var_WorkingTable

End Property

Public Property Let SetUnattendedMode(ByVal varBool As Boolean)

    DLLParams.UnattendedMode = varBool

End Property
