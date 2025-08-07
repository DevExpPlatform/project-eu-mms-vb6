VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_XMLMetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cntrExtra           As Integer
Private strExtra()          As String

Private Sub Class_Initialize()

    clearMetaData

End Sub

Private Sub clearMetaData()
    
    Erase strExtra()
    cntrExtra = -1
    
End Sub

Public Property Get getXMLMetaData() As String

    Dim I               As Integer
    Dim rValue          As String
    
    If (cntrExtra > -1) Then
        rValue = "<metadata>"
        
        For I = 0 To cntrExtra
            rValue = rValue & strExtra(I)
        Next I
        
        rValue = rValue & "</metadata>"
        
        getXMLMetaData = rValue
    End If
   
    clearMetaData

End Property

Public Property Let setMetaData(varParam As String, varValue As String)

    cntrExtra = (cntrExtra + 1)
    ReDim Preserve strExtra(cntrExtra) As String
    
    strExtra(cntrExtra) = "<" & varParam & "><![CDATA[" & varValue & "]]></" & varParam & ">"

End Property

