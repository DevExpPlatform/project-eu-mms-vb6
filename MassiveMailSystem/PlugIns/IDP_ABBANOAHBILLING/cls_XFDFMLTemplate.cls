VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_XFDFMLTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type strct_XFDFNodeField
    Id                      As String
    MultiLine               As String
End Type

Private Type strct_XFDFNodeFieldProperties
    Alignment               As String
    Coords                  As String
    FontName                As String
    FontSize                As String
    PageId                  As String
End Type

Private Type strct_XFDFNodeTemplate
    Indexes                 As String
    OutFileName             As String
    Version                 As String
End Type

Private cntrExtra           As Integer
Private cntrField           As Integer
Private nodeField           As strct_XFDFNodeField
Private nodeFieldProperties As strct_XFDFNodeFieldProperties
Private nodeTemplate        As strct_XFDFNodeTemplate
Private strExtra()          As String
Private strField()          As String

Private Sub Class_Initialize()

    clearField
    clearProperty
    clearTemplate

End Sub

Private Sub clearField()

    With nodeField
        .Id = ""
        .MultiLine = ""
    End With

End Sub

Private Sub clearProperty()

    With nodeFieldProperties
        .Alignment = ""
        .Coords = ""
        .FontName = ""
        .FontSize = ""
        .PageId = ""
    End With

End Sub

Private Sub clearTemplate()
    
    With nodeTemplate
        .Indexes = ""
        .OutFileName = ""
        .Version = ""
    End With
    
    Erase strExtra()
    Erase strField()
    
    cntrExtra = -1
    cntrField = -1
    
End Sub

Public Sub closeProperty()

    Dim fieldPropsAttrs As String

    With nodeFieldProperties
        If (.PageId <> "") Then fieldPropsAttrs = " pageid=" & Chr$(34) & .PageId & Chr$(34)
        If (.Coords <> "") Then fieldPropsAttrs = fieldPropsAttrs & " coords=" & Chr$(34) & .Coords & Chr$(34)
        If (.FontName <> "") Then fieldPropsAttrs = fieldPropsAttrs & " fontname=" & Chr$(34) & .FontName & Chr$(34)
        If (.FontSize <> "") Then fieldPropsAttrs = fieldPropsAttrs & " fontsize=" & Chr$(34) & .FontSize & Chr$(34)
        If (.Alignment <> "") Then fieldPropsAttrs = fieldPropsAttrs & " alignment=" & Chr$(34) & .Alignment & Chr$(34)
    End With
    
    strField(cntrField) = strField(cntrField) & "<properties" + fieldPropsAttrs & "/>"

End Sub

Public Sub closeField()
        
    Dim fieldAttrs As String

    With nodeField
        If (.Id <> "") Then fieldAttrs = " id=" & Chr$(34) & .Id & Chr$(34)
        If (.MultiLine <> "") Then fieldAttrs = fieldAttrs & " multiline=" & Chr$(34) & .MultiLine & Chr$(34)
    End With
    
    strField(cntrField) = "<field" & fieldAttrs & ">" & strField(cntrField) & "</field>"

    clearField
    clearProperty
    
End Sub

Public Property Get getXFDFTemplateNode() As String

    Dim I               As Integer
    Dim rValue          As String
    Dim templateAttrs   As String
    
    With nodeTemplate
        If (.Version <> "") Then templateAttrs = " version=" & Chr$(34) & .Version & Chr$(34)
        If (.OutFileName <> "") Then templateAttrs = templateAttrs & " filename=" & Chr$(34) & .OutFileName & Chr$(34)
        If (.Indexes <> "") Then templateAttrs = templateAttrs & " indexes=" & Chr$(34) & .Indexes & Chr$(34)
    End With
    
    rValue = "<template" & templateAttrs & ">"

    If (cntrField > -1) Then
        For I = 0 To cntrField
            rValue = rValue & strField(I)
        Next I
    End If
    
    If (cntrExtra > -1) Then
        rValue = rValue & "<extrainfo"
    
        For I = 0 To cntrExtra
            rValue = rValue & " " & strExtra(I)
        Next I
        
        rValue = rValue & "/>"
    End If
   
    rValue = rValue & "</template>"
    
    getXFDFTemplateNode = rValue
    
    clearTemplate

End Property

Public Property Let setExtraFields(varValue As String)
    
    cntrField = cntrField + 1

    ReDim Preserve strField(cntrField) As String

    strField(cntrField) = varValue

End Property

Public Property Let setExtraInfo(varValue As String)

    cntrExtra = cntrExtra + 1

    ReDim Preserve strExtra(cntrExtra) As String
    
    strExtra(cntrExtra) = varValue

End Property

Public Property Let setFieldId(varValue As String)

    cntrField = cntrField + 1

    ReDim Preserve strField(cntrField) As String

    nodeField.Id = varValue

End Property

Public Property Let setFieldMultiLine(varValue As String)

    nodeField.MultiLine = varValue

End Property

Public Property Let setPropertyAlignment(varValue As String)
    
    nodeFieldProperties.Alignment = varValue
    
End Property

Public Property Let setPropertyCoords(varValue As String)
    
    nodeFieldProperties.Coords = varValue
    
End Property

Public Property Let setPropertyFontName(varValue As String)
    
    nodeFieldProperties.FontName = varValue
    
End Property

Public Property Let setPropertyFontSize(varValue As String)
    
    nodeFieldProperties.FontSize = varValue
    
End Property

Public Property Let setPropertyPageId(varValue As String)
    
    nodeFieldProperties.PageId = varValue
    
End Property

Public Property Let setTemplateFileName(varValue As String)

    nodeTemplate.OutFileName = varValue

End Property

Public Property Let setTemplateIndexes(varValue As String)

    nodeTemplate.Indexes = varValue

End Property

Public Property Let setTemplateVersion(varValue As String)

    nodeTemplate.Version = varValue

End Property
