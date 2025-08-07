VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_XFDFMLTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type strct_XFDFNodeCell
    AlignH              As String
    AlignV              As String
    BackColor           As String
    BorderBottom        As String
    BorderBottomColor   As String
    BorderLeft          As String
    BorderLeftColor     As String
    BorderRight         As String
    BorderRightColor    As String
    Borders             As String
    BordersColor        As String
    BorderTop           As String
    BorderTopColor      As String
    Chunked             As Boolean
    ColSpan             As String
    FontName            As String
    FontSize            As String
    FontStyle           As String
    ForeColor           As String
    Height              As String
    Image               As String
    ImageScale          As String
    Leading             As String
    Padding             As String
    PaddingBottom       As String
    PaddingLeft         As String
    PaddingRight        As String
    PaddingTop          As String
    RoundedBorders      As String
    RoundedBordersColor As String
    RowSpan             As String
End Type

Private Type strct_XFDFNodeTable
    AlignH              As String
    AlignV              As String
    BackColor           As String
    BorderBottom        As String
    BorderBottomColor   As String
    BorderLeft          As String
    BorderLeftColor     As String
    BorderRight         As String
    BorderRightColor    As String
    Borders             As String
    BordersColor        As String
    BorderTop           As String
    BorderTopColor      As String
    Columns             As String
    FontName            As String
    FontSize            As String
    FontStyle           As String
    ForeColor           As String
    Leading             As String
    Padding             As String
    PaddingBottom       As String
    PaddingLeft         As String
    PaddingRight        As String
    PaddingTop          As String
    RowHeight           As String
    ShiftY              As String
    Widths              As String
End Type

Private cntrCell        As Integer
Private nodeCell        As strct_XFDFNodeCell
Private nodeTable       As strct_XFDFNodeTable
Private strCell()       As String

Public Property Let addCell(ByVal varValue As String)
    
    Dim cellAttributes As String
    
    cntrCell = (cntrCell + 1)
    
    ReDim Preserve strCell(cntrCell)
    
    With nodeCell
        If (.AlignH <> "") Then cellAttributes = cellAttributes & " alignmentH=" & Chr$(34) & .AlignH & Chr$(34)
        If (.AlignV <> "") Then cellAttributes = cellAttributes & " alignmentV=" & Chr$(34) & .AlignV & Chr$(34)
        If (.BackColor <> "") Then cellAttributes = cellAttributes & " rgbbackcolor=" & Chr$(34) & .BackColor & Chr$(34)
        If (.BorderBottom <> "") Then cellAttributes = cellAttributes & " borderBottom=" & Chr$(34) & .BorderBottom & Chr$(34)
        If (.BorderBottomColor <> "") Then cellAttributes = cellAttributes & " borderBottomColor=" & Chr$(34) & .BorderBottomColor & Chr$(34)
        If (.BorderLeft <> "") Then cellAttributes = cellAttributes & " borderLeft=" & Chr$(34) & .BorderLeft & Chr$(34)
        If (.BorderLeftColor <> "") Then cellAttributes = cellAttributes & " borderLeftColor=" & Chr$(34) & .BorderLeftColor & Chr$(34)
        If (.BorderRight <> "") Then cellAttributes = cellAttributes & " borderRight=" & Chr$(34) & .BorderRight & Chr$(34)
        If (.BorderRightColor <> "") Then cellAttributes = cellAttributes & " borderRightColor=" & Chr$(34) & .BorderRightColor & Chr$(34)
        If (.Borders <> "") Then cellAttributes = cellAttributes & " borders=" & Chr$(34) & .Borders & Chr$(34)
        If (.BordersColor <> "") Then cellAttributes = cellAttributes & " bordersColor=" & Chr$(34) & .BordersColor & Chr$(34)
        If (.BorderTop <> "") Then cellAttributes = cellAttributes & " borderTop=" & Chr$(34) & .BorderTop & Chr$(34)
        If (.BorderTopColor <> "") Then cellAttributes = cellAttributes & " borderTopColor=" & Chr$(34) & .BorderTopColor & Chr$(34)
        If (.Chunked = True) Then cellAttributes = cellAttributes & " chunked=" & Chr$(34) & "true" & Chr$(34)
        If (.ColSpan <> "") Then cellAttributes = cellAttributes & " colspan=" & Chr$(34) & .ColSpan & Chr$(34)
        If (.FontName <> "") Then cellAttributes = cellAttributes & " fontname=" & Chr$(34) & .FontName & Chr$(34)
        If (.FontSize <> "") Then cellAttributes = cellAttributes & " fontsize=" & Chr$(34) & .FontSize & Chr$(34)
        If (.FontStyle <> "") Then cellAttributes = cellAttributes & " fontstyle=" & Chr$(34) & .FontStyle & Chr$(34)
        If (.ForeColor <> "") Then cellAttributes = cellAttributes & " rgbforecolor=" & Chr$(34) & .ForeColor & Chr$(34)
        If (.Height <> "") Then cellAttributes = cellAttributes & " cellheight=" & Chr$(34) & .Height & Chr$(34)
        If (.Image <> "") Then cellAttributes = cellAttributes & " image=" & Chr$(34) & .Image & Chr$(34)
        If (.ImageScale <> "") Then cellAttributes = cellAttributes & " imagescale=" & Chr$(34) & .ImageScale & Chr$(34)
        If (.Leading <> "") Then cellAttributes = cellAttributes & " leading=" & Chr$(34) & .Leading & Chr$(34)
        If (.Padding <> "") Then cellAttributes = cellAttributes & " padding=" & Chr$(34) & .Padding & Chr$(34)
        If (.PaddingBottom <> "") Then cellAttributes = cellAttributes & " paddingBottom=" & Chr$(34) & .PaddingBottom & Chr$(34)
        If (.PaddingLeft <> "") Then cellAttributes = cellAttributes & " paddingLeft=" & Chr$(34) & .PaddingLeft & Chr$(34)
        If (.PaddingRight <> "") Then cellAttributes = cellAttributes & " paddingRight=" & Chr$(34) & .PaddingRight & Chr$(34)
        If (.PaddingTop <> "") Then cellAttributes = cellAttributes & " paddingTop=" & Chr$(34) & .PaddingTop & Chr$(34)
        If (.RoundedBorders <> "") Then cellAttributes = cellAttributes & " roundBorders=" & Chr$(34) & .RoundedBorders & Chr$(34)
        If (.RoundedBordersColor <> "") Then cellAttributes = cellAttributes & " roundBordersColor=" & Chr$(34) & .RoundedBordersColor & Chr$(34)
        If (.RowSpan <> "") Then cellAttributes = cellAttributes & " rowspan=" & Chr$(34) & .RowSpan & Chr$(34)
    End With
    
    strCell(cntrCell) = "<cell" & cellAttributes & ">"
    strCell(cntrCell) = strCell(cntrCell) & IIf(nodeCell.Chunked, varValue, "<![CDATA[" & varValue & "]]>")
    strCell(cntrCell) = strCell(cntrCell) & "</cell>"
    
    clearCell

End Property

Private Sub Class_Initialize()

    clearTable
    clearCell

End Sub

Private Sub clearCell()

    With nodeCell
        .AlignH = ""
        .AlignV = ""
        .BackColor = ""
        .BorderBottom = ""
        .BorderBottomColor = ""
        .BorderLeft = ""
        .BorderLeftColor = ""
        .BorderRight = ""
        .BorderRightColor = ""
        .Borders = ""
        .BordersColor = ""
        .BorderTop = ""
        .BorderTopColor = ""
        .Chunked = False
        .ColSpan = ""
        .FontName = ""
        .FontSize = ""
        .FontStyle = ""
        .ForeColor = ""
        .Height = ""
        .Image = ""
        .ImageScale = ""
        .Leading = ""
        .Padding = ""
        .PaddingBottom = ""
        .PaddingLeft = ""
        .PaddingRight = ""
        .PaddingTop = ""
        .RoundedBorders = ""
        .RoundedBordersColor = ""
        .RowSpan = ""
    End With

End Sub

Private Sub clearTable()

    With nodeTable
        .AlignH = ""
        .AlignV = ""
        .BackColor = ""
        .BorderBottom = ""
        .BorderBottomColor = ""
        .BorderLeft = ""
        .BorderLeftColor = ""
        .BorderRight = ""
        .BorderRightColor = ""
        .Borders = ""
        .BordersColor = ""
        .BorderTop = ""
        .BorderTopColor = ""
        .Columns = "1"
        .FontName = ""
        .FontSize = ""
        .FontStyle = ""
        .ForeColor = ""
        .Leading = ""
        .Padding = ""
        .PaddingBottom = ""
        .PaddingLeft = ""
        .PaddingRight = ""
        .PaddingTop = ""
        .RowHeight = ""
        .ShiftY = ""
        .Widths = ""
    End With

    Erase strCell
    
    cntrCell = -1

End Sub

Public Property Get getXFDFCellNodes() As String()

    getXFDFCellNodes = strCell
    
    clearCell

End Property

Public Property Get getXFDFTableNode() As String
    
    Dim I               As Integer
    Dim rValue          As String
    Dim tableAttributes As String
    
    If (cntrCell > -1) Then
        With nodeTable
            tableAttributes = " columns=" & Chr$(34) & .Columns & Chr$(34)
    
            If (.AlignH <> "") Then tableAttributes = tableAttributes & " alignmentH=" & Chr$(34) & .AlignH & Chr$(34)
            If (.AlignV <> "") Then tableAttributes = tableAttributes & " alignmentV=" & Chr$(34) & .AlignV & Chr$(34)
            If (.BackColor <> "") Then tableAttributes = tableAttributes & " rgbbackcolor=" & Chr$(34) & .BackColor & Chr$(34)
            If (.BorderBottom <> "") Then tableAttributes = tableAttributes & " borderBottom=" & Chr$(34) & .BorderBottom & Chr$(34)
            If (.BorderBottomColor <> "") Then tableAttributes = tableAttributes & " borderBottomColor=" & Chr$(34) & .BorderBottomColor & Chr$(34)
            If (.BorderLeft <> "") Then tableAttributes = tableAttributes & " borderLeft=" & Chr$(34) & .BorderLeft & Chr$(34)
            If (.BorderLeftColor <> "") Then tableAttributes = tableAttributes & " borderLeftColor=" & Chr$(34) & .BorderLeftColor & Chr$(34)
            If (.BorderRight <> "") Then tableAttributes = tableAttributes & " borderRight=" & Chr$(34) & .BorderRight & Chr$(34)
            If (.BorderRightColor <> "") Then tableAttributes = tableAttributes & " borderRightColor=" & Chr$(34) & .BorderRightColor & Chr$(34)
            If (.Borders <> "") Then tableAttributes = tableAttributes & " borders=" & Chr$(34) & .Borders & Chr$(34)
            If (.BordersColor <> "") Then tableAttributes = tableAttributes & " bordersColor=" & Chr$(34) & .BordersColor & Chr$(34)
            If (.BorderTop <> "") Then tableAttributes = tableAttributes & " borderTop=" & Chr$(34) & .BorderTop & Chr$(34)
            If (.BorderTopColor <> "") Then tableAttributes = tableAttributes & " borderTopColor=" & Chr$(34) & .BorderTopColor & Chr$(34)
            If (.FontName <> "") Then tableAttributes = tableAttributes & " fontname=" & Chr$(34) & .FontName & Chr$(34)
            If (.FontSize <> "") Then tableAttributes = tableAttributes & " fontsize=" & Chr$(34) & .FontSize & Chr$(34)
            If (.FontStyle <> "") Then tableAttributes = tableAttributes & " fontstyle=" & Chr$(34) & .FontStyle & Chr$(34)
            If (.ForeColor <> "") Then tableAttributes = tableAttributes & " rgbforecolor=" & Chr$(34) & .ForeColor & Chr$(34)
            If (.Leading <> "") Then tableAttributes = tableAttributes & " leading=" & Chr$(34) & .Leading & Chr$(34)
            If (.Padding <> "") Then tableAttributes = tableAttributes & " padding=" & Chr$(34) & .Padding & Chr$(34)
            If (.PaddingBottom <> "") Then tableAttributes = tableAttributes & " paddingBottom=" & Chr$(34) & .PaddingBottom & Chr$(34)
            If (.PaddingLeft <> "") Then tableAttributes = tableAttributes & " paddingLeft=" & Chr$(34) & .PaddingLeft & Chr$(34)
            If (.PaddingRight <> "") Then tableAttributes = tableAttributes & " paddingRight=" & Chr$(34) & .PaddingRight & Chr$(34)
            If (.PaddingTop <> "") Then tableAttributes = tableAttributes & " paddingTop=" & Chr$(34) & .PaddingTop & Chr$(34)
            If (.RowHeight <> "") Then tableAttributes = tableAttributes & " rowheight=" & Chr$(34) & .RowHeight & Chr$(34)
            If (.ShiftY <> "") Then tableAttributes = tableAttributes & " tableshiftY=" & Chr$(34) & .ShiftY & Chr$(34)
            If (.Widths <> "") Then tableAttributes = tableAttributes & " widths=" & Chr$(34) & .Widths & Chr$(34)
        End With
        
        rValue = "<table" & tableAttributes & ">"
        
        For I = 0 To cntrCell
            rValue = rValue & strCell(I)
        Next I
        
        rValue = rValue & "</table>"
    End If
    
    clearTable
    
    getXFDFTableNode = rValue
    
End Property

Public Property Let setCellAlignH(ByVal varValue As String)

    nodeCell.AlignH = varValue

End Property

Public Property Let setCellAlignV(ByVal varValue As String)

    nodeCell.AlignV = varValue

End Property

Public Property Let setCellBackColor(ByVal varValue As String)

    nodeCell.BackColor = varValue

End Property

Public Property Let setCellBorderBottom(ByVal varValue As String)

    nodeCell.BorderBottom = varValue

End Property

Public Property Let setCellBorderBottomColor(ByVal varValue As String)

    nodeCell.BorderBottomColor = varValue

End Property

Public Property Let setCellBorderLeft(ByVal varValue As String)

    nodeCell.BorderLeft = varValue

End Property

Public Property Let setCellBorderLeftColor(ByVal varValue As String)

    nodeCell.BorderLeftColor = varValue

End Property

Public Property Let setCellBorderRight(ByVal varValue As String)

    nodeCell.BorderRight = varValue

End Property

Public Property Let setCellBorderRightColor(ByVal varValue As String)

    nodeCell.BorderRightColor = varValue

End Property

Public Property Let setCellBorders(ByVal varValue As String)

    nodeCell.Borders = varValue

End Property

Public Property Let setCellBordersColor(ByVal varValue As String)

    nodeCell.BordersColor = varValue

End Property

Public Property Let setCellBorderTop(ByVal varValue As String)

    nodeCell.BorderTop = varValue

End Property

Public Property Let setCellBorderTopColor(ByVal varValue As String)

    nodeCell.BorderTopColor = varValue

End Property

Public Property Let setCellChuncked(ByVal varValue As Boolean)

    nodeCell.Chunked = varValue

End Property

Public Property Let setCellColSpan(ByVal varValue As String)

    nodeCell.ColSpan = varValue

End Property

Public Property Let setCellFontName(ByVal varValue As String)

    nodeCell.FontName = varValue

End Property

Public Property Let setCellFontSize(ByVal varValue As String)

    nodeCell.FontSize = varValue

End Property

Public Property Let setCellFontStyle(ByVal varValue As String)

    nodeCell.FontStyle = varValue

End Property

Public Property Let setCellForeColor(ByVal varValue As String)

    nodeCell.ForeColor = varValue

End Property

Public Property Let setCellHeight(ByVal varValue As String)

    nodeCell.Height = varValue

End Property

Public Property Let setCellImage(ByVal varValue As String)

    nodeCell.Image = varValue

End Property

Public Property Let setCellImageScale(ByVal varValue As String)

    nodeCell.ImageScale = varValue

End Property

Public Property Let setCellLeading(ByVal varValue As String)

    nodeCell.Leading = varValue

End Property

Public Property Let setCellPadding(ByVal varValue As String)

    nodeCell.Padding = varValue

End Property

Public Property Let setCellPaddingBottom(ByVal varValue As String)

    nodeCell.PaddingBottom = varValue

End Property

Public Property Let setCellPaddingLeft(ByVal varValue As String)

    nodeCell.PaddingLeft = varValue

End Property

Public Property Let setCellPaddingRight(ByVal varValue As String)

    nodeCell.PaddingRight = varValue

End Property

Public Property Let setCellPaddingTop(ByVal varValue As String)

    nodeCell.PaddingTop = varValue

End Property

Public Property Let setCellRoundedBorders(ByVal varValue As String)

    nodeCell.RoundedBorders = varValue

End Property

Public Property Let setCellRoundeBordersColor(ByVal varValue As String)

    nodeCell.RoundedBordersColor = varValue

End Property

Public Property Let setCellRowSpan(ByVal varValue As String)

    nodeCell.RowSpan = varValue

End Property

Public Property Let setTableAlignH(ByVal varValue As String)

    nodeTable.AlignH = varValue

End Property

Public Property Let setTableAlignV(ByVal varValue As String)

    nodeTable.AlignV = varValue

End Property

Public Property Let setTableBackColor(ByVal varValue As String)

    nodeTable.BackColor = varValue

End Property

Public Property Let setTableBorderBottom(ByVal varValue As String)

    nodeTable.BorderBottom = varValue

End Property

Public Property Let setTableBorderBottomColor(ByVal varValue As String)

    nodeTable.BorderBottomColor = varValue

End Property

Public Property Let setTableBorderLeft(ByVal varValue As String)

    nodeTable.BorderLeft = varValue

End Property

Public Property Let setTableBorderLeftColor(ByVal varValue As String)

    nodeTable.BorderLeftColor = varValue

End Property

Public Property Let setTableBorderRight(ByVal varValue As String)

    nodeTable.BorderRight = varValue

End Property

Public Property Let setTableBorderRightColor(ByVal varValue As String)

    nodeTable.BorderRightColor = varValue

End Property

Public Property Let setTableBorders(ByVal varValue As String)

    nodeTable.Borders = varValue

End Property

Public Property Let setTableBordersColor(ByVal varValue As String)

    nodeTable.BordersColor = varValue

End Property

Public Property Let setTableBorderTop(ByVal varValue As String)

    nodeTable.BorderTop = varValue

End Property

Public Property Let setTableBorderTopColor(ByVal varValue As String)

    nodeTable.BorderTopColor = varValue

End Property

Public Property Let setTableColumns(ByVal varValue As String)

    nodeTable.Columns = varValue

End Property

Public Property Let setTableFontName(ByVal varValue As String)

    nodeTable.FontName = varValue

End Property

Public Property Let setTableFontSize(ByVal varValue As String)

    nodeTable.FontSize = varValue

End Property

Public Property Let setTableFontStyle(ByVal varValue As String)

    nodeTable.FontStyle = varValue

End Property

Public Property Let setTableForeColor(ByVal varValue As String)

    nodeTable.ForeColor = varValue

End Property

Public Property Let setTableLeading(ByVal varValue As String)

    nodeTable.Leading = varValue

End Property

Public Property Let setTablePadding(ByVal varValue As String)

    nodeTable.Padding = varValue

End Property

Public Property Let setTablePaddingBottom(ByVal varValue As String)

    nodeTable.PaddingBottom = varValue

End Property

Public Property Let setTablePaddingLeft(ByVal varValue As String)

    nodeTable.PaddingLeft = varValue

End Property

Public Property Let setTablePaddingRight(ByVal varValue As String)

    nodeTable.PaddingRight = varValue

End Property

Public Property Let setTablePaddingTop(ByVal varValue As String)

    nodeTable.PaddingTop = varValue

End Property

Public Property Let setTableRowHeight(ByVal varValue As String)

    nodeTable.RowHeight = varValue

End Property

Public Property Let setTableShiftY(ByVal varValue As String)

    nodeTable.ShiftY = varValue

End Property

Public Property Let setTableWidths(ByVal varValue As String)

    nodeTable.Widths = varValue

End Property

Public Property Let setTableWidthsCM(WS_TABLEWIDTHS As String)

    Dim I               As Integer
    Dim WS_COLUMNS      As Integer
    Dim WS_TABLEWIDTH   As Single
    Dim WS_WIDTHS()     As String
    Dim WS_WIDTH        As String
    
    WS_WIDTHS = Split(WS_TABLEWIDTHS, ",")
    WS_COLUMNS = (UBound(WS_WIDTHS) + 1)
    
    For I = 0 To UBound(WS_WIDTHS)
        WS_TABLEWIDTH = (WS_TABLEWIDTH + Val(WS_WIDTHS(I)))
    Next I
    
    For I = 0 To UBound(WS_WIDTHS)
        WS_WIDTH = WS_WIDTH & "," & Replace$(Round(((Val(WS_WIDTHS(I)) * WS_COLUMNS) / WS_TABLEWIDTH), 2), ",", ".")
    Next I
    
    nodeTable.Widths = Mid$(WS_WIDTH, 2)

End Property

Public Property Let setXFDFCellNodes(varValue() As String)
    
    Dim I As Integer

    For I = 0 To UBound(varValue)
        cntrCell = (cntrCell + I + 1)
        
        ReDim Preserve strCell(cntrCell)
        
        strCell(cntrCell) = varValue(I)
    Next I

End Property
