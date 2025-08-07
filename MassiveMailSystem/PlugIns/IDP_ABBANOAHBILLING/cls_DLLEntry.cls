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
    Id                  As String
    Description         As String
    Name                As String
    Type                As String
End Type

Private PlugIn_Idntfr   As strct_PlugInIdentifier

Private Sub Class_Initialize()
    
    With PlugIn_Idntfr
        .Description = "Normal. flusso ABBANOA H2O Billing - v. 1.63.20250219"
        .Id = "ABBANOAHBILLING"
        .Name = "ABBANOA H2O - Billing"
        .Type = "Normalizer"
    End With

End Sub

Property Let DSN(ByVal DSN As String)

    DLLParams.DSN = DSN

End Property

Property Let EXTRAPARAMS(aParams() As String)

    On Error GoTo ErrHandler

    DLLParams.DSN_EXT = aParams(0)
    DLLParams.CCP_BILL = aParams(1)
    DLLParams.CCP_PPA = aParams(2)
    DLLParams.CF_ENTE = aParams(3)
    DLLParams.EXTRASPATH = aParams(4)
    DLLParams.TEMPLATEORGANIZER = aParams(5)
    DLLParams.TEMPLATEVERSION = aParams(6)
    DLLParams.DOCMODE = aParams(7)
    DLLParams.PLUGMODE = aParams(8)
        
    Exit Property

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Attenzione:"

End Property

Public Property Get GetUMErrorMessage() As String
    
    GetUMErrorMessage = UMErrMsg

End Property

Property Let IDDATACUTTER(ByVal IDDATACUTTER As String)

    DLLParams.IDDATACUTTER = IDDATACUTTER

End Property

Public Property Let INPUTFILENAME(ByVal IFName As String)

    DLLParams.INPUTFILENAME = IFName

End Property

Public Property Let OUTPUTFILENAME(ByVal OFName As String)

    DLLParams.OUTPUTFILENAME = OFName

End Property

Public Property Get PlugIn_GetInfo() As strct_PlugInIdentifier

    PlugIn_GetInfo = PlugIn_Idntfr
    
End Property

Public Property Let SetUnattendedMode(ByVal varBool As Boolean)

    DLLParams.UNATTENDEDMODE = varBool

End Property

Public Function StartJob() As Boolean

    If DB_ConnectInit Then StartJob = NRM_ABBANOAHBILLING_DataProcessor
    
    DB_ConnectRelease

End Function

Public Property Let TABLENAME(ByVal TABLENAME As String)

    DLLParams.TABLENAME = TABLENAME

End Property

Public Property Let TNS(ByVal TNS As String)

    DLLParams.TNS = TNS

End Property
