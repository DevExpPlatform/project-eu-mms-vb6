VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_XFDFMLText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type strct_XFDFNodeText
    FontLeading     As String
    FontLeadingFix  As String
    FontName        As String
    FontSize        As String
    FontStyle       As String
    RGBColor        As String
    TextAlign       As String
End Type

Private Type strct_XFDFNodeChunk
    FontName        As String
    FontSize        As String
    FontStyle       As String
    RGBColor        As String
End Type

Private cntrChunck  As Integer
Private nodeChunk   As strct_XFDFNodeChunk
Private nodeText    As strct_XFDFNodeText
Private strChunk()  As String

Public Property Let addChunk(ByVal varValue As String)
    
    Dim chunkAttributes As String
    
    ReDim Preserve strChunk(cntrChunck)
    
    With nodeChunk
        If (.FontName <> "") Then chunkAttributes = " fontname=" & Chr$(34) & .FontName & Chr$(34)
        If (.FontSize <> "") Then chunkAttributes = chunkAttributes & " fontsize=" & Chr$(34) & .FontSize & Chr$(34)
        If (.FontStyle <> "") Then chunkAttributes = chunkAttributes & " fontstyle=" & Chr$(34) & .FontStyle & Chr$(34)
        If (.RGBColor <> "") Then chunkAttributes = chunkAttributes & " rgbcolor=" & Chr$(34) & .RGBColor & Chr$(34)
    End With
    
    strChunk(cntrChunck) = "<chunk" & chunkAttributes & ">"
    strChunk(cntrChunck) = strChunk(cntrChunck) & "<![CDATA[" & varValue & "]]>"
    strChunk(cntrChunck) = strChunk(cntrChunck) & "</chunk>"
    
    cntrChunck = cntrChunck + 1

    clearChunk

End Property

Private Sub Class_Initialize()

    clearText
    clearChunk

End Sub

Private Sub clearChunk()

    With nodeChunk
        .FontName = ""
        .FontSize = ""
        .FontStyle = ""
        .RGBColor = ""
    End With

End Sub

Private Sub clearText()
    
    With nodeText
        .FontLeading = ""
        .FontLeadingFix = ""
        .FontName = ""
        .FontSize = ""
        .FontStyle = ""
        .RGBColor = ""
        .TextAlign = ""
    End With

    Erase strChunk
    
    cntrChunck = 0
    
End Sub

Public Property Get getXFDFTextNode() As String
    
    Dim rValue          As String
    Dim textAttributes  As String
    
    With nodeText
        If (.FontLeading <> "") Then textAttributes = " leading=" & Chr$(34) & .FontLeading & Chr$(34)
        If (.FontLeadingFix <> "") Then textAttributes = " leadingfixed=" & Chr$(34) & .FontLeadingFix & Chr$(34)
        If (.FontName <> "") Then textAttributes = textAttributes & " fontname=" & Chr$(34) & .FontName & Chr$(34)
        If (.FontSize <> "") Then textAttributes = textAttributes & " fontsize=" & Chr$(34) & .FontSize & Chr$(34)
        If (.FontStyle <> "") Then textAttributes = textAttributes & " fontstyle=" & Chr$(34) & .FontStyle & Chr$(34)
        If (.RGBColor <> "") Then textAttributes = textAttributes & " rgbcolor=" & Chr$(34) & .RGBColor & Chr$(34)
        If (.TextAlign <> "") Then textAttributes = textAttributes & " alignment=" & Chr$(34) & .TextAlign & Chr$(34)
    End With
    
    rValue = "<text" & textAttributes & ">"
    
    For cntrChunck = 0 To UBound(strChunk)
        rValue = rValue & strChunk(cntrChunck)
    Next cntrChunck
    
    rValue = rValue & "</text>"
    
    getXFDFTextNode = rValue
    
    clearText
    
End Property

Public Property Let setNodeChunkFontName(ByVal varValue As String)

    nodeChunk.FontName = varValue

End Property

Public Property Let setNodeChunkFontSize(ByVal varValue As String)

    nodeChunk.FontSize = varValue

End Property

Public Property Let setNodeChunkFontStyle(ByVal varValue As String)

    nodeChunk.FontStyle = varValue

End Property

Public Property Let setNodeChunkRGBColor(ByVal varValue As String)

    nodeChunk.RGBColor = varValue

End Property

Public Property Let setNodeTextFontLeading(ByVal varValue As String)

    nodeText.FontLeading = varValue

End Property

Public Property Let setNodeTextFontLeadingFixed(ByVal varValue As String)

    nodeText.FontLeadingFix = varValue

End Property

Public Property Let setNodeTextFontName(ByVal varValue As String)

    nodeText.FontName = varValue

End Property

Public Property Let setNodeTextFontSize(ByVal varValue As String)

    nodeText.FontSize = varValue

End Property

Public Property Let setNodeTextFontStyle(ByVal varValue As String)

    nodeText.FontStyle = varValue

End Property

Public Property Let setNodeTextRGBColor(ByVal varValue As String)

    nodeText.RGBColor = varValue

End Property

Public Property Let setNodeTextTextAlign(ByVal varValue As String)

    nodeText.TextAlign = varValue

End Property
