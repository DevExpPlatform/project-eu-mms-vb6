VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_DataCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private dataDescription     As String
Private dataExtraParams()   As String
Private dataExtraParamsFlg  As Boolean

Public Property Get getDescription() As String

    getDescription = dataDescription
    
End Property

Public Property Get getExtraParams() As String()

    getExtraParams = dataExtraParams
    
End Property

Public Property Get getExtraParamsFlag() As Boolean

    getExtraParamsFlag = dataExtraParamsFlg
    
End Property

Public Property Let setDescrition(varParam As String)

    dataDescription = varParam

End Property

Public Property Let setExtraParams(varParams() As String)

    dataExtraParams = varParams
    dataExtraParamsFlg = True

End Property
