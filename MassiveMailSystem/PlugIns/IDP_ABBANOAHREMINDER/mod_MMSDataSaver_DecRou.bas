Attribute VB_Name = "mod_MMSDataSaver_DecRou"
Option Explicit

Private Const WS_PAGEHEIGHTMAX = 273

Private Type strct_DOCUMENTDETAILS_DATA
    ANNO                    As String
    CODICE_SERVIZIO         As String
    DATA_COST_MORA          As String
    DATAEMISSIONE           As String
    DATASCADENZA            As String
    NUMERO_CUR              As String
    NUMERO_ORG              As String
    IMPORTO_RES             As String
    IMPORTO_TOT             As String
    SOLLECITO               As String
    NUM_SOLLECITO_BONARIO   As String
    DATA_SOLLECITO_BONARIO  As String
    DATA_RICEZIONE_SB       As String
    IMPORTO_PRESCRIVIBILE   As String
End Type

Private mySQLImporter       As SQL_Import.PlugIn
Private myXFDFMLTable       As cls_XFDFMLTable
Private WS_ANNEXED_DATA     As String
Private WS_COLSPAN          As String
Private WS_DATAEMISSIONE    As String
Private WS_ERRSCT           As String
Private WS_ERRMSG           As String
Private WS_IMPORTOTOTALE    As String
Private WS_NATIONALITY      As String
Private WS_NMR_CODANA       As String
Private WS_NMR_SOLLECITO    As String
Private WS_PAGEHEIGHT       As Single
Private WS_PAGENUM          As String
Private WS_PAGES_RD()       As String
Private WS_PAGES_RD_CNTR    As Integer
Private WS_SERVICE_HDR      As String

Private Sub ADD_DOCUMENTDETAILS_LEGEND_LXX()

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + Val(WS_CS_PXX_DOCUMENTDETAILS_LEGEND.EXTRAPARAMS(0)))

    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then WS_PAGES_RD_CNTR = (WS_PAGES_RD_CNTR + 1)
    
    ReDim Preserve WS_PAGES_RD(WS_PAGES_RD_CNTR)
    WS_PAGES_RD(WS_PAGES_RD_CNTR) = WS_PAGES_RD(WS_PAGES_RD_CNTR) & WS_CS_PXX_DOCUMENTDETAILS_LEGEND.DATADESCRIPTION
       
End Sub

Private Sub ADD_DOCUMENTDETAILS_MSG_L03()

    GET_CHK_RD_ITEMSHEIGHT 16.3

    With myXFDFMLTable
        .setCellAlignH = "justified"
        .setCellColSpan = WS_COLSPAN
        .setCellChuncked = True
        .addCell = "<text alignment='justified' fontname='helr46w.ttf' fontsize='7' rgbcolor='0,0,0'>" & _
                       "<chunk fontname='helr66w.ttf'><![CDATA[(*) N.B.:]]></chunk>" & _
                       "<chunk><![CDATA[ Comunque non prima del termine previsto per il pagamento del presente sollecito bonario.<br>]]></chunk>" & _
                       "<chunk fontname='helr66w.ttf'><![CDATA[(**)]]></chunk>" & _
                       "<chunk><![CDATA[ L’importo prescrittibile (riferito all’intera fattura e non alla sola rata) è disponibile esclusivamente per le fatture emesse dal 20/09/2021. Laddove il sollecito riguardi comunque fatture prescrittibili con scadenza successiva al 01/01/2020 o emesse a partire da tale data si invita il cliente a valutare la possibilità di eccepire la prescrizione mediante la presentazione di apposito modulo MODCLI032 disponibile nel portale (]]></chunk>" & _
                       "<chunk fontstyle='underline' rgbcolor='0,0,142'><![CDATA[]]></chunk>" & _
                       "<chunk><![CDATA[).<br>Si tenga presente che per fatture con scadenza anteriore al 01/01/2020 continua a vigere la c.d. “prescrizione quinquennale” per le quali è sempre possibile eccepire la prescrizione mediante la presentazione di apposito modulo MODCLI032 disponibile nel portale (]]></chunk>" & _
                       "<chunk fontstyle='underline' rgbcolor='0,0,142'><![CDATA[]]></chunk>" & _
                       "<chunk><![CDATA[).]]></chunk>" & _
                   "</text>"
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_MSG_L17()

    GET_CHK_RD_ITEMSHEIGHT Val(WS_CS_PXX_MSG_L17.EXTRAPARAMS(0))

    With myXFDFMLTable
        .setCellAlignH = "justified"
        .setCellColSpan = WS_COLSPAN
        .setCellChuncked = True
        .addCell = WS_CS_PXX_MSG_L17.DATADESCRIPTION
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEHEADER(addLabel As Boolean)

    Dim WS_HDR    As String
    Dim WS_STRING As String
    
    WS_HDR = "Elenco delle fatture che risultano insolute relativamente ai servizi di fornitura per un totale di € " & WS_IMPORTOTOTALE
    
    With myXFDFMLTable
        .setTableAlignH = "right"
        '.setTableBorders = "0.3"
        .setTableFontName = "helr45w.ttf"
        .setTableFontSize = "8"
        
        Select Case DLLParams.LAYOUT
        Case "L03"
            .setTableColumns = "7"
            .setTableWidths = "2,2,3.3,3.4,3.2,2,3.1"
        
            WS_COLSPAN = "7"
            WS_HDR = "Elenco fatture insolute per un totale di € " & WS_IMPORTOTOTALE
        
        Case "L07"
            .setTableColumns = "6"
            .setTableWidths = "1,1,1,1,1,1"
            
            WS_COLSPAN = "6"
        
        Case "L09"
            .setTableColumns = "5"
            .setTableWidths = "0.8,1.6,1.6,1,1"
            
            WS_COLSPAN = "5"
            WS_HDR = "Elenco fatture insolute per un totale di € " & WS_IMPORTOTOTALE
        
        Case "L10"
            .setTableColumns = "5"
            .setTableWidths = "1.6,2.7,2.7,2,1.6"
            
            WS_COLSPAN = "5"
            WS_HDR = "Elenco fatture insolute per un totale di € " & WS_IMPORTOTOTALE
        
        Case "L17"
            .setTableColumns = "9"
            .setTableFontSize = "7"
            .setTableWidths = "1.6,2.7,2.7,2,1.6,2.5,2,2,2"
        
            WS_COLSPAN = "9"
            WS_HDR = "Elenco fatture insolute per un totale di € " & WS_IMPORTOTOTALE
        
        Case Else
            .setTableColumns = "5"
            .setTableWidths = "1,1,1,1,1"
            
            WS_COLSPAN = "5"
        
        End Select

        If (addLabel) Then
            .setCellAlignH = "left"
            .setCellColSpan = WS_COLSPAN
            .setCellFontName = "helr65w.ttf"
            .setCellFontSize = "9"
            .addCell = WS_HDR

            WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.6)
            
            If ((DLLParams.BYPASS_STAMP = False) And (DLLParams.LAYOUT <> "L07")) Then
                WS_STRING = GET_STAMP_TYPE
                
                If (WS_STRING <> "") Then
                    .setCellAlignH = "left"
                    .setCellColSpan = WS_COLSPAN
                    .setCellFontName = "helr66w.ttf"
                    .setCellFontSize = "8"
                    .addCell = WS_STRING
    
                    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + 4.3)
                End If
            End If
        End If
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_L03(WS_DOCUMENTDETAILS_DATA As strct_DOCUMENTDETAILS_DATA)

    GET_CHK_RD_ITEMSHEIGHT 3.17

    With myXFDFMLTable
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO
        
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_ORG

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_CUR

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_RES

        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATASCADENZA
        
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_L07(WS_DOCUMENTDETAILS_DATA As strct_DOCUMENTDETAILS_DATA)

    GET_CHK_RD_ITEMSHEIGHT 3.17

    With myXFDFMLTable
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_ORG

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_TOT

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_RES

        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATASCADENZA
        
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.SOLLECITO
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_L09(WS_DOCUMENTDETAILS_DATA As strct_DOCUMENTDETAILS_DATA)

    GET_CHK_RD_ITEMSHEIGHT 3.17

    With myXFDFMLTable
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_ORG

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_CUR

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_RES

        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATASCADENZA
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_L10(WS_DOCUMENTDETAILS_DATA As strct_DOCUMENTDETAILS_DATA)

    GET_CHK_RD_ITEMSHEIGHT 3.17

    With myXFDFMLTable
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_ORG

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_CUR

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_RES

        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATASCADENZA
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_L17(WS_DOCUMENTDETAILS_DATA As strct_DOCUMENTDETAILS_DATA)

    GET_CHK_RD_ITEMSHEIGHT 3.17

    With myXFDFMLTable
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_ORG

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_CUR

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_RES

        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATASCADENZA
        
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUM_SOLLECITO_BONARIO
        
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATA_SOLLECITO_BONARIO
        
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATA_RICEZIONE_SB
        
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_LXX(WS_DOCUMENTDETAILS_DATA As strct_DOCUMENTDETAILS_DATA)

    GET_CHK_RD_ITEMSHEIGHT 3.17

    With myXFDFMLTable
        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.NUMERO_ORG

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_TOT

        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.IMPORTO_RES

        .setCellAlignH = "center"
        .setCellPaddingBottom = "1"
        .setCellPaddingTop = "0"
        .addCell = WS_DOCUMENTDETAILS_DATA.DATASCADENZA
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_HDR_L03()

    GET_CHK_RD_ITEMSHEIGHT 30.5

    With myXFDFMLTable
        .setCellColSpan = WS_COLSPAN
        .setCellHeight = "10"
        .addCell = ""
    
        .setCellAlignH = "left"
        .setCellColSpan = WS_COLSPAN
        .setCellFontName = "helr65w.ttf"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = WS_SERVICE_HDR
        
        .setCellAlignH = "right"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Codice<br>Servizio"
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data<br>Emissione"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num. Originaria"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num. Corrente<br>Anno/Numero/Rata"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo<br>Fatturato"

        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data<br>Scadenza"
    
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "(**) Importo<br>Prescrittibile<br>da Prescrizione<br>Biennale"
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_HDR_L07()

    GET_CHK_RD_ITEMSHEIGHT 13.41

    With myXFDFMLTable
        .setCellColSpan = WS_COLSPAN
        .setCellHeight = "10"
        .addCell = ""
    
        .setCellAlignH = "left"
        .setCellColSpan = WS_COLSPAN
        .setCellFontName = "helr65w.ttf"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = WS_SERVICE_HDR
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data Emissione"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura n."

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo Fatturato"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo Residuo"

        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Scadenza Fatture"
    
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Sollecitato"
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_HDR_L09()

    GET_CHK_RD_ITEMSHEIGHT 13.41

    With myXFDFMLTable
        .setCellColSpan = WS_COLSPAN
        .setCellHeight = "10"
        .addCell = ""
    
        .setCellAlignH = "left"
        .setCellColSpan = WS_COLSPAN
        .setCellFontName = "helr65w.ttf"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = WS_SERVICE_HDR
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data Emissione"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num. Originaria"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num. Corrente<br>Anno/Numero/Rata"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo<br>Fatturato"

        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data Scadenza"
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_HDR_L10()
    
    GET_CHK_RD_ITEMSHEIGHT 35.5

    With myXFDFMLTable
        .setCellColSpan = WS_COLSPAN
        .setCellHeight = "10"
        .addCell = ""
    
        .setCellAlignH = "left"
        .setCellColSpan = WS_COLSPAN
        .setCellFontName = "helr65w.ttf"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = WS_SERVICE_HDR
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data<br>Emissione"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num.<br>Originaria"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num. Corrente<br>Anno/Numero/Rata"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo<br>Fatturato"

        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data Scadenza"
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_HDR_L17()
    
    GET_CHK_RD_ITEMSHEIGHT 35.5

    With myXFDFMLTable
        .setCellColSpan = WS_COLSPAN
        .setCellHeight = "10"
        .addCell = ""
    
        .setCellAlignH = "left"
        .setCellColSpan = WS_COLSPAN
        .setCellFontName = "helr65w.ttf"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = WS_SERVICE_HDR
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data<br>Emissione"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num.<br>Originaria"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura<br>Num. Corrente<br>Anno/Numero/Rata"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo<br>Fatturato"

        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data Scadenza"
    
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Num.<br>Sollecito<br>Bonario"
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data<br>Sollecito<br>Bonario"
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data<br>Ricezione<br>Sollecito<br>Bonario"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo consumi<br>risalenti a più di 2 anni (*)"
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_HDR_LXX()

    GET_CHK_RD_ITEMSHEIGHT 13.41

    With myXFDFMLTable
        .setCellColSpan = WS_COLSPAN
        .setCellHeight = "10"
        .addCell = ""
    
        .setCellAlignH = "left"
        .setCellColSpan = WS_COLSPAN
        .setCellFontName = "helr65w.ttf"
        .setCellFontSize = "7"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = WS_SERVICE_HDR
        
        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Data Emissione"
        
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Fattura n."

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo Fatturato"

        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Importo Residuo"

        .setCellAlignH = "center"
        .setCellBackColor = "230,230,230"
        .setCellBorderBottom = "0.6"
        .setCellFontName = "helr65w.ttf"
        .setCellPaddingBottom = "2"
        .setCellPaddingTop = "1"
        .addCell = "Scadenza Fatture"
    End With

End Sub

Private Sub ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE()

    GET_CHK_RD_ITEMSHEIGHT 0.71

    With myXFDFMLTable
        .setCellBorderTop = "0.6"
        .setCellColSpan = WS_COLSPAN
        .setCellHeight = "2"
        .addCell = ""
    End With

End Sub

Private Function GET_CHK_RD_ITEMSHEIGHT(WS_ITEMHEIGHT As Single) As Boolean

    WS_PAGEHEIGHT = (WS_PAGEHEIGHT + WS_ITEMHEIGHT)

    If (WS_PAGEHEIGHT >= WS_PAGEHEIGHTMAX) Then
        WS_PAGEHEIGHT = 22

        ReDim Preserve WS_PAGES_RD(WS_PAGES_RD_CNTR)

        WS_PAGES_RD(WS_PAGES_RD_CNTR) = WS_PAGES_RD(WS_PAGES_RD_CNTR) & myXFDFMLTable.getXFDFTableNode
        ADD_DOCUMENTDETAILS_TABLEHEADER False

        WS_PAGES_RD_CNTR = (WS_PAGES_RD_CNTR + 1)
        GET_CHK_RD_ITEMSHEIGHT = True
    End If

End Function

Private Function GET_DOCUMENTDETAILS_L03() As String

    Dim I                           As Integer
    Dim J                           As Integer
    Dim WS_CTGR_TRFFR               As String
    Dim WS_DOCUMENTDETAILS_DATA     As strct_DOCUMENTDETAILS_DATA
    Dim WS_EXT_DATA                 As strct_DATA
    Dim WS_EXT_KEY                  As String
    Dim WS_EXT_KEY_HDR              As String
    Dim WS_STRING                   As String
    
    WS_ERRSCT = "GET_DOCUMENTDETAILS_L03"
    WS_PAGEHEIGHT = 175
    
    ADD_DOCUMENTDETAILS_TABLEHEADER True

    For I = 0 To UBound(WS_01S.ED_RXXX)
        With WS_01S
            WS_SERVICE_HDR = ""
            WS_STRING = ""
            
'            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000")
            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000") + "_" + Format$(Trim$(.ED_RXXX(I).ED_R001_S_E.CODICESERVIZIO), "0000000000")
            WS_EXT_DATA = GET_DATACACHE(EST_DATA_ANA, WS_EXT_KEY_HDR)

            If (WS_EXT_DATA.DATADESCRIPTION <> Trim$(WS_EXT_KEY_HDR)) Then
                DDS_ADD "[$TXT_CFPIVA]", WS_EXT_DATA.EXTRAPARAMS(0)

                WS_CTGR_TRFFR = WS_EXT_DATA.EXTRAPARAMS(1)
            Else
                DDS_ADD "[$TXT_CFPIVA]", "-"
            End If

            With .ED_RXXX(I)
                Select Case .TIPORECORD
                Case "AC"
                    WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_A_C.CODICEANAGRAFICO), "0000000000")
                    
                    If (WS_FLG_BO) Then
                        If ((Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "") And (Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "-")) Then
                            WS_STRING = Abs(CSng(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS))
                            WS_STRING = "<br>Quota di Bonus sociale Idrico non erogata e che potrà essere trattenuta a compensazione dell’importo insoluto oggetto di costituzione in mora: € " & NRM_IMPORT(WS_STRING, "#,##0.00", False)
                        End If
                    End If
                                                                          
                    WS_SERVICE_HDR = "Fatture" & IIf((WS_NMR_CODANA = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)), "", "<br>Categoria Tariffaria " & WS_CTGR_TRFFR & " - Num. Componenti Familiari " & Val(WS_01S.SL_R002.COMPONENTI_NUCLEO_FAMILIARE) & _
                                     "<br>Usufruisce di Bonus Idrico: " & IIf((WS_01S.SL_R002.BONUS = "S"), "SI", "NO") & _
                                     "<br>Utenza Disalimentabile: " & IIf(WS_01S.SL_R002.NON_DISALIMENTABILE = "S", "NO", "SI") & _
                                     WS_STRING & _
                                     IIf(Trim$(WS_01S.SL_R001.DATA_COSTITUZIONE_MORA) = "", "", "<br>Data Avvio Costituzione in Mora (*): " & WS_01S.SL_R001.DATA_COSTITUZIONE_MORA) & _
                                     "<br>Totale € " & Trim$(.ED_R001_S_T.IMPORTO))
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L03
                                            
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)
                    
                    For J = 0 To UBound(.ED_RXXX_A_D)
                        With .ED_RXXX_A_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO)
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = ""
                            
                            If (.FLG_IMPORTI_PRESCRIVIBILI = "S") Then
                                WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(.TOTALE_IMPORTI_PRESCRIVIBILI, "#,##0.00", False)
                            Else
                                If (WS_FLG_EST_IMP_PRESCR) Then
                                    WS_STRING = Format$(WS_01S.AF_R001.CODICELOTTO, "000000") & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000") & .ANNO & Format$(Split(WS_DOCUMENTDETAILS_DATA.NUMERO_CUR, "/")(1), "00000000")

                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = GET_DATACACHE(EST_IMP_PRESCR, WS_STRING).DATADESCRIPTION

                                    If (WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = WS_STRING) Then
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                    Else
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE, "#,##0.00", False)
                                    End If
                                Else
                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                End If
                            End If
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L03 WS_DOCUMENTDETAILS_DATA
                    Next J
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                    
                Case "SE"
                    WS_SERVICE_HDR = "Cod. Servizio " & Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    'WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_S_E.CODICESERVIZIO), "0000000000")
                
                    If (Trim$(.ED_R001_S_I.VIA & .ED_R001_S_I.NUMEROCIVICO) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & ", " & Trim$(.ED_R001_S_I.VIA) & " " & Trim$(.ED_R001_S_I.NUMEROCIVICO)
                    If (Trim$(.ED_R001_S_L.CAP & .ED_R001_S_L.LOCALITÀ & .ED_R001_S_L.PROVINCIA) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & " - " & Trim$(.ED_R001_S_L.CAP) & " " & Trim$(.ED_R001_S_L.LOCALITÀ) & " " & Trim$(.ED_R001_S_L.PROVINCIA)
                
                    If (WS_FLG_BO) Then
                        If ((Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "") And (Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "-")) Then
                            If (CSng(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) < 0) Then
                                WS_STRING = Abs(CSng(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS))
                                WS_STRING = "<br>Quota di Bonus sociale Idrico non erogata e che potrà essere trattenuta a compensazione dell’importo insoluto oggetto di costituzione in mora: € " & NRM_IMPORT(WS_STRING, "#,##0.00", False)
                            End If
                        End If
                    End If
                                                                          
                    WS_SERVICE_HDR = WS_SERVICE_HDR & "<br>Categoria Tariffaria " & WS_CTGR_TRFFR & " - Num. Componenti Familiari " & Val(WS_01S.SL_R002.COMPONENTI_NUCLEO_FAMILIARE) & _
                                                      "<br>Usufruisce di Bonus Idrico: " & IIf((WS_01S.SL_R002.BONUS = "S"), "SI", "NO") & _
                                                      "<br>Utenza Disalimentabile: " & IIf(WS_01S.SL_R002.NON_DISALIMENTABILE = "S", "NO", "SI") & _
                                                      WS_STRING & _
                                                      IIf(Trim$(WS_01S.SL_R001.DATA_COSTITUZIONE_MORA) = "", "", "<br>Data Avvio Costituzione in Mora (*): " & WS_01S.SL_R001.DATA_COSTITUZIONE_MORA) & _
                                                      "<br>Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                        
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L03
                    
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    
                    For J = 0 To UBound(.ED_RXXX_S_D)
                        With .ED_RXXX_S_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO) & "/" & Format$(.RATA, "00")
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = .DATASCADENZA
                            
                            If (.FLG_IMPORTI_PRESCRIVIBILI = "S") Then
                                WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(.TOTALE_IMPORTI_PRESCRIVIBILI, "#,##0.00", False)
                            Else
                                If ((WS_FLG_EST_IMP_PRESCR) And (Val(.RATA) < 2)) Then
                                    WS_STRING = Format$(WS_01S.AF_R001.CODICELOTTO, "000000") & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000") & .ANNO & Format$(Split(WS_DOCUMENTDETAILS_DATA.NUMERO_CUR, "/")(1), "00000000")

                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = GET_DATACACHE(EST_IMP_PRESCR, WS_STRING).DATADESCRIPTION

                                    If (WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = WS_STRING) Then
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                    Else
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE, "#,##0.00", False)
                                    End If
                                Else
                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                End If
                            End If
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L03 WS_DOCUMENTDETAILS_DATA
                    Next J
                                
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                
                End Select
            End With
        End With
    Next I
 
    ADD_DOCUMENTDETAILS_MSG_L03
 
    ' PAGES LOADER
    '
    ReDim Preserve WS_PAGES_RD(WS_PAGES_RD_CNTR)
    WS_PAGES_RD(WS_PAGES_RD_CNTR) = WS_PAGES_RD(WS_PAGES_RD_CNTR) & myXFDFMLTable.getXFDFTableNode
        
    For I = 0 To WS_PAGES_RD_CNTR
        GET_DOCUMENTDETAILS_L03 = GET_DOCUMENTDETAILS_L03 & WS_PAGES_RD(I) & IIf((I = WS_PAGES_RD_CNTR), "", "[EP]")
    Next I

    Erase WS_PAGES_RD()

End Function

Private Function GET_DOCUMENTDETAILS_L09() As String

    Dim I                           As Integer
    Dim J                           As Integer
    Dim WS_CTGR_TRFFR               As String
    Dim WS_DOCUMENTDETAILS_DATA     As strct_DOCUMENTDETAILS_DATA
    Dim WS_EXT_DATA                 As strct_DATA
    Dim WS_EXT_KEY                  As String
    Dim WS_EXT_KEY_HDR              As String
    Dim WS_STRING                   As String
    
    WS_ERRSCT = "GET_DOCUMENTDETAILS_L09"
    WS_PAGEHEIGHT = 22
    
    ADD_DOCUMENTDETAILS_TABLEHEADER True

    For I = 0 To UBound(WS_01S.ED_RXXX)
        With WS_01S
            WS_SERVICE_HDR = ""
            WS_STRING = ""
            
'            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000")
            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000") + "_" + Format$(Trim$(.ED_RXXX(I).ED_R001_S_E.CODICESERVIZIO), "0000000000")
            WS_EXT_DATA = GET_DATACACHE(EST_DATA_ANA, WS_EXT_KEY_HDR)

            If (WS_EXT_DATA.DATADESCRIPTION <> Trim$(WS_EXT_KEY_HDR)) Then
                DDS_ADD "[$TXT_CFPIVA]", WS_EXT_DATA.EXTRAPARAMS(0)

                WS_CTGR_TRFFR = WS_EXT_DATA.EXTRAPARAMS(1)
            Else
                DDS_ADD "[$TXT_CFPIVA]", "-"
            End If

            With .ED_RXXX(I)
                Select Case .TIPORECORD
                Case "AC"
                    WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_A_C.CODICEANAGRAFICO), "0000000000")
                    WS_SERVICE_HDR = "Fatture" & IIf((WS_NMR_CODANA = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)), "", "<br>Categoria Tariffaria " & WS_CTGR_TRFFR) & " - Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L09
                                            
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)
                    
                    For J = 0 To UBound(.ED_RXXX_A_D)
                        With .ED_RXXX_A_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO)
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = ""
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L09 WS_DOCUMENTDETAILS_DATA
                    Next J
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                    
                Case "SE"
                    WS_SERVICE_HDR = "Cod. Servizio " & Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    'WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_S_E.CODICESERVIZIO), "0000000000")
                
                    If (Trim$(.ED_R001_S_I.VIA & .ED_R001_S_I.NUMEROCIVICO) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & ", " & Trim$(.ED_R001_S_I.VIA) & " " & Trim$(.ED_R001_S_I.NUMEROCIVICO)
                    If (Trim$(.ED_R001_S_L.CAP & .ED_R001_S_L.LOCALITÀ & .ED_R001_S_L.PROVINCIA) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & " - " & Trim$(.ED_R001_S_L.CAP) & " " & Trim$(.ED_R001_S_L.LOCALITÀ) & " " & Trim$(.ED_R001_S_L.PROVINCIA)
                
                    WS_SERVICE_HDR = WS_SERVICE_HDR & "<br>Categoria Tariffaria " & WS_CTGR_TRFFR & " - Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                        
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L09
                    
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    
                    For J = 0 To UBound(.ED_RXXX_S_D)
                        With .ED_RXXX_S_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO) & "/" & Format$(.RATA, "00")
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = .DATASCADENZA
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L09 WS_DOCUMENTDETAILS_DATA
                    Next J
                                
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                
                End Select
            End With
        End With
    Next I
 
    ' PAGES LOADER
    '
    ReDim Preserve WS_PAGES_RD(WS_PAGES_RD_CNTR)
    WS_PAGES_RD(WS_PAGES_RD_CNTR) = WS_PAGES_RD(WS_PAGES_RD_CNTR) & myXFDFMLTable.getXFDFTableNode
        
    For I = 0 To WS_PAGES_RD_CNTR
        GET_DOCUMENTDETAILS_L09 = GET_DOCUMENTDETAILS_L09 & WS_PAGES_RD(I) & IIf((I = WS_PAGES_RD_CNTR), "", "[EP]")
    Next I

    Erase WS_PAGES_RD()

End Function

Private Function GET_DOCUMENTDETAILS_L10() As String

    Dim I                           As Integer
    Dim J                           As Integer
    Dim WS_CTGR_TRFFR               As String
    Dim WS_DOCUMENTDETAILS_DATA     As strct_DOCUMENTDETAILS_DATA
    Dim WS_EXT_DATA                 As strct_DATA
    Dim WS_EXT_KEY                  As String
    Dim WS_EXT_KEY_HDR              As String
    Dim WS_STRING                   As String
    
    WS_ERRSCT = "GET_DOCUMENTDETAILS_L10"
    WS_PAGEHEIGHT = 22
    
    ADD_DOCUMENTDETAILS_TABLEHEADER True

    For I = 0 To UBound(WS_01S.ED_RXXX)
        With WS_01S
            WS_SERVICE_HDR = ""
            WS_STRING = ""
            
'            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000")
            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000") + "_" + Format$(Trim$(.ED_RXXX(I).ED_R001_S_E.CODICESERVIZIO), "0000000000")
            WS_EXT_DATA = GET_DATACACHE(EST_DATA_ANA, WS_EXT_KEY_HDR)

            If (WS_EXT_DATA.DATADESCRIPTION <> Trim$(WS_EXT_KEY_HDR)) Then
                DDS_ADD "[$TXT_CFPIVA]", WS_EXT_DATA.EXTRAPARAMS(0)

                WS_CTGR_TRFFR = WS_EXT_DATA.EXTRAPARAMS(1)
            Else
                DDS_ADD "[$TXT_CFPIVA]", "-"
            End If

            With .ED_RXXX(I)
                Select Case .TIPORECORD
                Case "AC"
                    WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_A_C.CODICEANAGRAFICO), "0000000000")
                    WS_SERVICE_HDR = "Fatture" & IIf((WS_NMR_CODANA = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)), "", "<br>Categoria Tariffaria " & WS_CTGR_TRFFR) & " - Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L10
                                            
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)
                    
                    For J = 0 To UBound(.ED_RXXX_A_D)
                        With .ED_RXXX_A_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO)
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = ""
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L10 WS_DOCUMENTDETAILS_DATA
                    Next J
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                    
                Case "SE"
                    WS_SERVICE_HDR = "Cod. Servizio " & Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    'WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_S_E.CODICESERVIZIO), "0000000000")
                
                    If (Trim$(.ED_R001_S_I.VIA & .ED_R001_S_I.NUMEROCIVICO) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & ", " & Trim$(.ED_R001_S_I.VIA) & " " & Trim$(.ED_R001_S_I.NUMEROCIVICO)
                    If (Trim$(.ED_R001_S_L.CAP & .ED_R001_S_L.LOCALITÀ & .ED_R001_S_L.PROVINCIA) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & " - " & Trim$(.ED_R001_S_L.CAP) & " " & Trim$(.ED_R001_S_L.LOCALITÀ) & " " & Trim$(.ED_R001_S_L.PROVINCIA)
                
                    WS_SERVICE_HDR = WS_SERVICE_HDR & "<br>Categoria Tariffaria " & WS_CTGR_TRFFR & " - Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                        
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L10
                    
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    
                    For J = 0 To UBound(.ED_RXXX_S_D)
                        With .ED_RXXX_S_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO) & "/" & Format$(.RATA, "00")
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = .DATASCADENZA
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L10 WS_DOCUMENTDETAILS_DATA
                    Next J
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                
                End Select
            End With
        End With
    Next I
    
    ' PAGES LOADER
    '
    ReDim Preserve WS_PAGES_RD(WS_PAGES_RD_CNTR)
    WS_PAGES_RD(WS_PAGES_RD_CNTR) = WS_PAGES_RD(WS_PAGES_RD_CNTR) & myXFDFMLTable.getXFDFTableNode
        
    ADD_DOCUMENTDETAILS_LEGEND_LXX
    
    For I = 0 To WS_PAGES_RD_CNTR
        GET_DOCUMENTDETAILS_L10 = GET_DOCUMENTDETAILS_L10 & WS_PAGES_RD(I) & IIf((I = WS_PAGES_RD_CNTR), "", "[EP]")
    Next I

    Erase WS_PAGES_RD()
    

End Function

Private Function GET_DOCUMENTDETAILS_L17() As String

    Dim I                           As Integer
    Dim J                           As Integer
    Dim WS_CTGR_TRFFR               As String
    Dim WS_DOCUMENTDETAILS_DATA     As strct_DOCUMENTDETAILS_DATA
    Dim WS_DPL                      As String
    Dim WS_DPS                      As String
    Dim WS_EXT_DATA                 As strct_DATA
    Dim WS_EXT_KEY                  As String
    Dim WS_EXT_KEY_HDR              As String
    Dim WS_STRING                   As String
    
    WS_ERRSCT = "GET_DOCUMENTDETAILS_L17"
    WS_PAGEHEIGHT = 145
    
    If (Trim$(WS_01S.SL_R001.DATA_PREVISTA_LIMITAZIONE) <> "") Then WS_DPL = "<br>Data Avvio Limitazione: " & WS_01S.SL_R001.DATA_PREVISTA_LIMITAZIONE
    
    If (WS_01S.SL_R002.NON_DISALIMENTABILE = "S") Then
        WS_DPS = "<br>Data Avvio Sospensione Non Applicabile alla Data di Emissione della Costituzione in Mora.<br>Data Avvio Disattivazione Non Applicabile alla Data di Emissione della Costituzione in Mora."
    Else
        If (Trim$(WS_01S.SL_R001.DATA_PREVISTA_SOSPENSIONE) <> "") Then WS_DPS = "<br>Data Avvio Sospensione: " & WS_01S.SL_R001.DATA_PREVISTA_SOSPENSIONE & "<br>Data Avvio Disattivazione: " & Format$(DateAdd("d", 90, CDate(WS_01S.SL_R001.DATA_PREVISTA_SOSPENSIONE)))
    End If
    
    ADD_DOCUMENTDETAILS_TABLEHEADER True

    For I = 0 To UBound(WS_01S.ED_RXXX)
        With WS_01S
            WS_SERVICE_HDR = ""
            WS_STRING = ""
            
            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000") + "_" + Format$(Trim$(.ED_RXXX(I).ED_R001_S_E.CODICESERVIZIO), "0000000000")
            WS_EXT_DATA = GET_DATACACHE(EST_DATA_ANA, WS_EXT_KEY_HDR)

            If (WS_EXT_DATA.DATADESCRIPTION <> Trim$(WS_EXT_KEY_HDR)) Then
                DDS_ADD "[$TXT_CFPIVA]", WS_EXT_DATA.EXTRAPARAMS(0)

                WS_CTGR_TRFFR = WS_EXT_DATA.EXTRAPARAMS(1)
            Else
                DDS_ADD "[$TXT_CFPIVA]", "-"
            End If

            With .ED_RXXX(I)
                Select Case .TIPORECORD
                Case "AC"
                    WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_A_C.CODICEANAGRAFICO), "0000000000")
                    
                    If (WS_FLG_BO) Then
                        If ((Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "") And (Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "-")) Then
                            WS_STRING = Abs(CSng(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS))
                            WS_STRING = "<br>Quota di Bonus sociale Idrico non erogata e che potrà essere trattenuta a compensazione dell’importo insoluto oggetto di costituzione in mora: € " & NRM_IMPORT(WS_STRING, "#,##0.00", False)
                        End If
                    End If
                    
                    WS_SERVICE_HDR = "Fatture" & IIf((WS_NMR_CODANA = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)), "", "<br>Categoria Tariffaria " & WS_CTGR_TRFFR & " - Num. Componenti Familiari " & Val(WS_01S.SL_R002.COMPONENTI_NUCLEO_FAMILIARE) & _
                                     "<br>Usufruisce di Bonus Idrico: " & IIf((WS_01S.SL_R002.BONUS = "S"), "SI", "NO") & _
                                     "<br>Utenza Disalimentabile alla Data di Emissione della Costituzione in Mora: " & IIf(WS_01S.SL_R002.NON_DISALIMENTABILE = "S", "NO", "SI") & _
                                     WS_STRING & _
                                     WS_DPL & _
                                     WS_DPS & _
                                     "<br>Totale € " & Trim$(.ED_R001_S_T.IMPORTO))
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L17
                                            
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)
                    
                    For J = 0 To UBound(.ED_RXXX_A_D)
                        With .ED_RXXX_A_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO)
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = ""
                            WS_DOCUMENTDETAILS_DATA.NUM_SOLLECITO_BONARIO = Trim$(.NUM_SOLLECITO_BONARIO)
                            WS_DOCUMENTDETAILS_DATA.DATA_SOLLECITO_BONARIO = .DATA_SOLLECITO_BONARIO
                            WS_DOCUMENTDETAILS_DATA.DATA_RICEZIONE_SB = .DATA_RICEZIONE_SB
                            
                            If (.FLG_IMPORTI_PRESCRIVIBILI = "S") Then
                                WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(.TOTALE_IMPORTI_PRESCRIVIBILI, "#,##0.00", False)
                            Else
                                If (WS_FLG_EST_IMP_PRESCR) Then
                                    WS_STRING = Format$(WS_01S.AF_R001.CODICELOTTO, "000000") & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000") & .ANNO & Format$(Split(WS_DOCUMENTDETAILS_DATA.NUMERO_CUR, "/")(1), "00000000")

                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = GET_DATACACHE(EST_IMP_PRESCR, WS_STRING).DATADESCRIPTION

                                    If (WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = WS_STRING) Then
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                    Else
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE, "#,##0.00", False)
                                    End If
                                Else
                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                End If
                            End If
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L17 WS_DOCUMENTDETAILS_DATA
                    Next J
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                    
                Case "SE"
                    WS_SERVICE_HDR = "Cod. Servizio " & Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    'WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_S_E.CODICESERVIZIO), "0000000000")
                
                    If (Trim$(.ED_R001_S_I.VIA & .ED_R001_S_I.NUMEROCIVICO) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & ", " & Trim$(.ED_R001_S_I.VIA) & " " & Trim$(.ED_R001_S_I.NUMEROCIVICO)
                    If (Trim$(.ED_R001_S_L.CAP & .ED_R001_S_L.LOCALITÀ & .ED_R001_S_L.PROVINCIA) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & " - " & Trim$(.ED_R001_S_L.CAP) & " " & Trim$(.ED_R001_S_L.LOCALITÀ) & " " & Trim$(.ED_R001_S_L.PROVINCIA)
                
                    If (WS_FLG_BO) Then
                        If ((Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "") And (Trim$(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) <> "-")) Then
                            If (CSng(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS) < 0) Then
                                WS_STRING = Abs(CSng(WS_01S.BO_R001.IMPORTO_RESIDUO_BONUS))
                                WS_STRING = "<br>Quota di Bonus sociale Idrico non erogata e che potrà essere trattenuta a compensazione dell’importo insoluto oggetto di costituzione in mora: € " & NRM_IMPORT(WS_STRING, "#,##0.00", False)
                            End If
                        End If
                    End If
                                                                          
                    WS_SERVICE_HDR = WS_SERVICE_HDR & "<br>Categoria Tariffaria " & WS_CTGR_TRFFR & " - Num. Componenti Familiari " & Val(WS_01S.SL_R002.COMPONENTI_NUCLEO_FAMILIARE) & _
                                                      "<br>Usufruisce di Bonus Idrico: " & IIf((WS_01S.SL_R002.BONUS = "S"), "SI", "NO") & _
                                                      "<br>Utenza Disalimentabile alla Data di Emissione della Costituzione in Mora: " & IIf(WS_01S.SL_R002.NON_DISALIMENTABILE = "S", "NO", "SI") & _
                                                      WS_STRING & _
                                                      WS_DPL & _
                                                      WS_DPS & _
                                                      "<br>Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                        
                    ADD_DOCUMENTDETAILS_TABLEROW_HDR_L17
                    
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    
                    For J = 0 To UBound(.ED_RXXX_S_D)
                        With .ED_RXXX_S_D(J)
                            WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                            WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & "/" & Trim$(.NUMERO) & "/" & Format$(.RATA, "00")
                            
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY)
                                
                            If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY) Then
                                WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                
                                With WS_01S
                                    WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                                              "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                                              "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                                              "KEY NOT FOUND: " & WS_EXT_KEY
                                End With
                            Else
                                If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = .ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                                Else
                                    WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = ""
                                End If
                            End If
                            
                            WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(.IMPORTO, "#,##0.00", False)
                            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = .DATASCADENZA
                            WS_DOCUMENTDETAILS_DATA.NUM_SOLLECITO_BONARIO = Trim$(.NUM_SOLLECITO_BONARIO)
                            WS_DOCUMENTDETAILS_DATA.DATA_SOLLECITO_BONARIO = .DATA_SOLLECITO_BONARIO
                            WS_DOCUMENTDETAILS_DATA.DATA_RICEZIONE_SB = .DATA_RICEZIONE_SB
                            
                            If (.FLG_IMPORTI_PRESCRIVIBILI = "S") Then
                                WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(.TOTALE_IMPORTI_PRESCRIVIBILI, "#,##0.00", False)
                            Else
                                If ((WS_FLG_EST_IMP_PRESCR) And (Val(.RATA) < 2)) Then
                                    WS_STRING = Format$(WS_01S.AF_R001.CODICELOTTO, "000000") & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000") & .ANNO & Format$(Split(WS_DOCUMENTDETAILS_DATA.NUMERO_CUR, "/")(1), "00000000")

                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = GET_DATACACHE(EST_IMP_PRESCR, WS_STRING).DATADESCRIPTION

                                    If (WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = WS_STRING) Then
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                    Else
                                        WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = NRM_IMPORT(WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE, "#,##0.00", False)
                                    End If
                                Else
                                    WS_DOCUMENTDETAILS_DATA.IMPORTO_PRESCRIVIBILE = ""
                                End If
                            End If
                        End With
                            
                        ADD_DOCUMENTDETAILS_TABLEROW_L17 WS_DOCUMENTDETAILS_DATA
                    Next J
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                
                End Select
            End With
        End With
    Next I
 
    ADD_DOCUMENTDETAILS_MSG_L17

    ' PAGES LOADER
    '
    ReDim Preserve WS_PAGES_RD(WS_PAGES_RD_CNTR)
    WS_PAGES_RD(WS_PAGES_RD_CNTR) = WS_PAGES_RD(WS_PAGES_RD_CNTR) & myXFDFMLTable.getXFDFTableNode
        
    For I = 0 To WS_PAGES_RD_CNTR
        GET_DOCUMENTDETAILS_L17 = GET_DOCUMENTDETAILS_L17 & WS_PAGES_RD(I) & IIf((I = WS_PAGES_RD_CNTR), "", "[EP]")
    Next I

    Erase WS_PAGES_RD()

End Function

Private Function GET_DOCUMENTDETAILS_LXX() As String

    Dim I                           As Integer
    Dim J                           As Integer
    Dim WS_CALC_IMPORTO             As Single
    Dim WS_CTGR_TRFFR               As String
    Dim WS_DOCUMENTDETAILS_DATA     As strct_DOCUMENTDETAILS_DATA
    Dim WS_EXT_DATA                 As strct_DATA
    Dim WS_EXT_KEY                  As String
    Dim WS_EXT_KEY_HDR              As String
    Dim WS_EXT_KEY_TMP              As String
    Dim WS_STATUS                   As String
    Dim WS_STRING                   As String
    
    WS_ERRSCT = "GET_DOCUMENTDETAILS_LXX"
    
    WS_PAGEHEIGHT = 22
    
    ADD_DOCUMENTDETAILS_TABLEHEADER True

    For I = 0 To UBound(WS_01S.ED_RXXX)
        With WS_01S
'            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000")
            WS_EXT_KEY_HDR = Format$(WS_NMR_CODANA, "0000000000") + "_" + Format$(Trim$(.ED_RXXX(I).ED_R001_S_E.CODICESERVIZIO), "0000000000")
            WS_EXT_DATA = GET_DATACACHE(EST_DATA_ANA, WS_EXT_KEY_HDR)
            WS_STRING = ""

            If (WS_EXT_DATA.DATADESCRIPTION <> Trim$(WS_EXT_KEY_HDR)) Then
                DDS_ADD "[$TXT_CFPIVA]", WS_EXT_DATA.EXTRAPARAMS(0)

                WS_CTGR_TRFFR = WS_EXT_DATA.EXTRAPARAMS(1)
                
                If (DLLParams.LAYOUT = "L07") Then WS_STATUS = " - Stato " & WS_EXT_DATA.EXTRAPARAMS(2)
            Else
                DDS_ADD "[$TXT_CFPIVA]", "-"
            End If

            WS_CALC_IMPORTO = 0
            'WS_EXT_KEY_HDR = Format$(WS_EXT_KEY_HDR, "0000000000")
            WS_EXT_KEY_TMP = ""
            WS_SERVICE_HDR = ""

            With .ED_RXXX(I)
                Select Case .TIPORECORD
                Case "AC"
                    WS_SERVICE_HDR = "Fatture" & IIf((WS_NMR_CODANA = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)), "", "<br>Categoria Tariffaria " & WS_CTGR_TRFFR) & WS_STATUS & " - Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                    WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_A_C.CODICEANAGRAFICO), "0000000000")
                
                    Select Case DLLParams.LAYOUT
                    Case "L07"
                        ADD_DOCUMENTDETAILS_TABLEROW_HDR_L07
                    
                    Case Else
                        ADD_DOCUMENTDETAILS_TABLEROW_HDR_LXX
                    
                    End Select
                                            
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_A_C.CODICEANAGRAFICO)
                    
                    For J = 0 To UBound(.ED_RXXX_A_D)
                        With .ED_RXXX_A_D(J)
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")
                            
                            If (WS_EXT_KEY_TMP = WS_EXT_KEY) Then
                                WS_CALC_IMPORTO = (WS_CALC_IMPORTO + CSng(.IMPORTO))
                            Else
                                GoSub WRITE_ROW
                                
                                WS_EXT_KEY_TMP = WS_EXT_KEY
                                WS_CALC_IMPORTO = CSng(.IMPORTO)
                                
                                WS_DOCUMENTDETAILS_DATA.ANNO = .ANNO
                                WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                                WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & Trim$(.NUMERO)
                                WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(WS_CALC_IMPORTO, "#,##0.00", False)
                            End If
                        End With
                    Next J
                    
                    GoSub WRITE_ROW
                    
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                    
                Case "SE"
                    WS_SERVICE_HDR = "Cod. Servizio " & Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    'WS_EXT_KEY_HDR = WS_EXT_KEY_HDR & "_" & Format$(Trim$(.ED_R001_S_E.CODICESERVIZIO), "0000000000")
                
                    If (Trim$(.ED_R001_S_I.VIA & .ED_R001_S_I.NUMEROCIVICO) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & ", " & Trim$(.ED_R001_S_I.VIA) & " " & Trim$(.ED_R001_S_I.NUMEROCIVICO)
                    If (Trim$(.ED_R001_S_L.CAP & .ED_R001_S_L.LOCALITÀ & .ED_R001_S_L.PROVINCIA) <> "") Then WS_SERVICE_HDR = WS_SERVICE_HDR & " - " & Trim$(.ED_R001_S_L.CAP) & " " & Trim$(.ED_R001_S_L.LOCALITÀ) & " " & Trim$(.ED_R001_S_L.PROVINCIA)
                
                    WS_SERVICE_HDR = WS_SERVICE_HDR & "<br>Categoria Tariffaria " & WS_CTGR_TRFFR & WS_STATUS & " - Totale € " & Trim$(.ED_R001_S_T.IMPORTO)
                        
                    Select Case DLLParams.LAYOUT
                    Case "L07"
                        ADD_DOCUMENTDETAILS_TABLEROW_HDR_L07
                    
                    Case Else
                        ADD_DOCUMENTDETAILS_TABLEROW_HDR_LXX
                    
                    End Select
                    
                    WS_DOCUMENTDETAILS_DATA.CODICE_SERVIZIO = Trim$(.ED_R001_S_E.CODICESERVIZIO)
                    
                    For J = 0 To UBound(.ED_RXXX_S_D)
                        With .ED_RXXX_S_D(J)
                            WS_EXT_KEY = WS_EXT_KEY_HDR & "_" & .TIPODOCUMENTO & "_" & .ANNO & Format$(IIf((Trim$(.NUMERO) = ""), "0", Trim$(.NUMERO)), "00000000")

                            If (WS_EXT_KEY_TMP = WS_EXT_KEY) Then
                                WS_CALC_IMPORTO = (WS_CALC_IMPORTO + CSng(.IMPORTO))
                            Else
                                GoSub WRITE_ROW
                                
                                WS_EXT_KEY_TMP = WS_EXT_KEY
                                WS_CALC_IMPORTO = CSng(.IMPORTO)
                                
                                WS_DOCUMENTDETAILS_DATA.ANNO = .ANNO
                                WS_DOCUMENTDETAILS_DATA.DATAEMISSIONE = .DATAEMISSIONE
                                WS_DOCUMENTDETAILS_DATA.NUMERO_CUR = .ANNO & Trim$(.NUMERO)
                                WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(WS_CALC_IMPORTO, "#,##0.00", False)
                            End If
                        End With
                    Next J
                                
                    GoSub WRITE_ROW
                                        
                    ADD_DOCUMENTDETAILS_TABLEROW_ROWLINE
                
                End Select
            End With
        End With
    Next I
 
    ' PAGES LOADER
    '
    ReDim Preserve WS_PAGES_RD(WS_PAGES_RD_CNTR)
    WS_PAGES_RD(WS_PAGES_RD_CNTR) = WS_PAGES_RD(WS_PAGES_RD_CNTR) & myXFDFMLTable.getXFDFTableNode
        
    ADD_DOCUMENTDETAILS_LEGEND_LXX
    
    For I = 0 To WS_PAGES_RD_CNTR
        GET_DOCUMENTDETAILS_LXX = GET_DOCUMENTDETAILS_LXX & WS_PAGES_RD(I) & IIf((I = WS_PAGES_RD_CNTR), "", "[EP]")
    Next I

    Erase WS_PAGES_RD()

    Exit Function

WRITE_ROW:
    If (WS_EXT_KEY_TMP <> "") Then
        WS_EXT_DATA = GET_DATACACHE(EST_DATA, WS_EXT_KEY_TMP)
        
        If (WS_EXT_DATA.DATADESCRIPTION = WS_EXT_KEY_TMP) Then
            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = ""
            WS_DOCUMENTDETAILS_DATA.IMPORTO_TOT = ""
            WS_DOCUMENTDETAILS_DATA.SOLLECITO = ""
            
            With WS_01S
                WRITE2LOG Format$(Now, "dd/mm/yyyy hh.MM.ss") & " - " & _
                          "COD. ANAGR.: " & WS_NMR_CODANA & " - " & _
                          "NUM. DOC.: " & IIf((.AF_R001.CODICELOTTO & "/" & .IL_R001.CODICESOTTOLOTTO = "/"), "NO INFO", Format$(.AF_R001.CODICELOTTO, "000000") & "/" & Format$(.IL_R001.CODICESOTTOLOTTO, "000000")) & " - " & _
                          "KEY NOT FOUND: " & WS_EXT_KEY_TMP
            End With
        Else
            If (WS_EXT_DATA.EXTRAPARAMS(0) <> "") Then WS_DOCUMENTDETAILS_DATA.NUMERO_ORG = WS_DOCUMENTDETAILS_DATA.ANNO & WS_EXT_DATA.EXTRAPARAMS(0)
                        
            WS_DOCUMENTDETAILS_DATA.DATASCADENZA = WS_EXT_DATA.EXTRAPARAMS(2)
            WS_DOCUMENTDETAILS_DATA.IMPORTO_TOT = WS_EXT_DATA.EXTRAPARAMS(1)
            WS_DOCUMENTDETAILS_DATA.SOLLECITO = WS_EXT_DATA.EXTRAPARAMS(3)
        End If
    
        WS_DOCUMENTDETAILS_DATA.IMPORTO_RES = NRM_IMPORT(WS_CALC_IMPORTO, "#,##0.00", False)
        
        Select Case DLLParams.LAYOUT
        Case "L07"
            ADD_DOCUMENTDETAILS_TABLEROW_L07 WS_DOCUMENTDETAILS_DATA
        
        Case Else
            ADD_DOCUMENTDETAILS_TABLEROW_LXX WS_DOCUMENTDETAILS_DATA
        
        End Select
    End If
Return

End Function

Private Function GET_STAMP_TYPE() As String

    WS_ERRSCT = "GET_STAMP_TYPE"
    
    Dim WS_STAMP_TYPE   As String
    Dim WS_STRING       As String
    
    WS_STRING = Format$(WS_01S.AF_R001.CODICELOTTO, "000000") & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000")
    WS_STAMP_TYPE = GET_DATACACHE(EST_DATA_STAMP, WS_STRING).DATADESCRIPTION

    If (WS_STAMP_TYPE <> WS_STRING) Then
        GET_STAMP_TYPE = GET_DATACACHE(EST_MSG_STAMP, "MSG_STAMP_0" & WS_STAMP_TYPE).DATADESCRIPTION

        If (GET_STAMP_TYPE = WS_STAMP_TYPE) Then GET_STAMP_TYPE = ""
    End If

End Function

Private Function GET_TEMPLATEINFO() As String

    Dim I                   As Integer
    Dim J                   As Integer
    Dim myXFDFMLTemplate    As cls_XFDFMLTemplate
    Dim WS_EXTRAFIELDS      As String
    Dim WS_FIELDS           As String
    Dim WS_FILENAME         As String
    Dim WS_INT              As Integer
    Dim WS_PAGEBILL         As Integer
    Dim WS_PAGESNUM         As Integer
    Dim WS_PAGESINDXS       As String
    Dim WS_STRING           As String
    Dim WS_TEMP_UNIQUEID    As Integer

    WS_ERRSCT = "GET_TEMPLATEINFO"

    WS_ANNEXED_DATA = ""
    WS_FILENAME = "ABN_" & IIf((DLLParams.LAYOUT = ""), "", DLLParams.LAYOUT & "_")
    WS_PAGENUM = ""

    Set myXFDFMLTemplate = New cls_XFDFMLTemplate

    For I = 0 To UBound(WS_SECTIONS)
        With WS_SECTIONS(I)
            Select Case .SECTIONDESCRIPTION
            Case "BILL"
                If (DLLParams.PRINT_BILL) Then
                    With .TEMPLATES(0)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If
            
                        WS_PAGEBILL = (WS_PAGESNUM + 1)
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                    
                    If ((DLLParams.LAYOUT = "L03") Or (DLLParams.LAYOUT = "L09") Or (DLLParams.LAYOUT = "L17")) Then
                        WS_EXTRAFIELDS = WS_EXTRAFIELDS & "<field id='NMR_IMPORTOTOTALE'>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='87.7,94.2,4.5,35.5' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='87.7,248,4.5,46' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                          "</field>" & _
                                                          "<field id='NMR_SOLLECITO'>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='36.2,34.2,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='270'/>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='42.1,225,2.8,32.5' fontname='helr65w.ttf' fontsize='8' rotation='270'/>" & _
                                                          "</field>" & _
                                                          "<field id='NMR_CODANA'>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='32.4,34.2,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='270'/>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='38.7,225,2.8,32.5' fontname='helr65w.ttf' fontsize='8' rotation='270'/>" & _
                                                          "</field>" & _
                                                          "<field id='TXT_BOLLETTINOBC_BC00'><properties pageid='" & WS_PAGEBILL & "' coords='26.3,190.3,12,93' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOBC_BCT00'><properties pageid='" & WS_PAGEBILL & "' coords='23.9,190.3,2.5,93' alignment='center' fontname='helr65w.ttf' fontsize='6' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOBC_BC01'><properties pageid='" & WS_PAGEBILL & "' coords='2.35,81.06,14.97,45' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOIDCLIENTE'><properties pageid='" & WS_PAGEBILL & "' coords='56,139.5,4.8,52.2' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOID'><properties pageid='" & WS_PAGEBILL & "' coords='6.67,138.6,5.5,152.8' fontname='ocrb10n.ttf' fontsize='11' rotation='270' comb='60'/></field>"
                    Else
                        WS_EXTRAFIELDS = WS_EXTRAFIELDS & "<field id='NMR_IMPORTOTOTALE'>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='87.7,94.2,4.5,35.5' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='87.7,248,4.5,46' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/>" & _
                                                          "</field>" & _
                                                          "<field id='NMR_SOLLECITO'>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='36.2,34.2,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='270'/>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='42.1,225,2.8,32.5' fontname='helr65w.ttf' fontsize='8' rotation='270'/>" & _
                                                          "</field>" & _
                                                          "<field id='NMR_CODANA'>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='32.4,34.2,3.9,56.4' fontname='helr65w.ttf' fontsize='9' rotation='270'/>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='38.7,225,2.8,32.5' fontname='helr65w.ttf' fontsize='8' rotation='270'/>" & _
                                                          "</field>" & _
                                                          "<field id='DTA_SCADENZA'>" & _
                                                              "<properties pageid='" & WS_PAGEBILL & "' coords='10.6,22.1,4.6,26.6' fontname='helr65w.ttf' fontsize='12' alignment='center' rotation='270'/>" & _
                                                          "</field>" & _
                                                          "<field id='TXT_BOLLETTINOBC_BC00'><properties pageid='" & WS_PAGEBILL & "' coords='26.3,190.3,12,93' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOBC_BCT00'><properties pageid='" & WS_PAGEBILL & "' coords='23.9,190.3,2.5,93' alignment='center' fontname='helr65w.ttf' fontsize='6' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOBC_BC01'><properties pageid='" & WS_PAGEBILL & "' coords='2.35,81.06,14.97,45' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOIDCLIENTE'><properties pageid='" & WS_PAGEBILL & "' coords='56,139.5,4.8,52.2' fontname='ocrb10n.ttf' fontsize='11' rotation='270'/></field>" & _
                                                          "<field id='TXT_BOLLETTINOID'><properties pageid='" & WS_PAGEBILL & "' coords='6.67,138.6,5.5,152.8' fontname='ocrb10n.ttf' fontsize='11' rotation='270' comb='60'/></field>"
                    End If
                End If
                            
            Case "DETAILS"
                For J = 0 To WS_PAGES_RD_CNTR
                    With .TEMPLATES(0)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            
                            If (J = 0) Then WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If

                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                    End With
                        
                    If (J = 0) Then
                        Select Case DLLParams.LAYOUT
                        Case "L03"
                            WS_STRING = "175"
                            
                        Case "L17"
                            WS_STRING = "130"
                            
                        Case Else
                            WS_STRING = "22"
                        
                        End Select
                    Else
                        WS_STRING = "22"
                    End If
                        
                    myXFDFMLTemplate.setFieldId = "TXT_DETAILS_P" & Format$((J + 1), "000")
                    myXFDFMLTemplate.setPropertyPageId = WS_PAGESNUM
                    myXFDFMLTemplate.setPropertyCoords = "10," & WS_STRING & ",190,240"
                    myXFDFMLTemplate.closeProperty
                    myXFDFMLTemplate.closeField
                Next J

                WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(0).TEMP_BITWISE)

                If (WS_PAGESNUM And 1) Then
                    WS_PAGESINDXS = WS_PAGESINDXS & .TEMPLATES(1).TEMP_ID & ","
                    WS_PAGESNUM = (WS_PAGESNUM + .TEMPLATES(1).TEMP_PAGES)
                    WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(1).TEMP_BITWISE)
                End If
            
            Case "MODCLI032R2"
                For J = 0 To UBound(WS_SECTIONS(I).TEMPLATES)
                    With .TEMPLATES(J)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If
                        
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                Next J
            
            Case "REMINDER"
                For J = 0 To UBound(WS_SECTIONS(I).TEMPLATES)
                    With .TEMPLATES(J)
                        If (Trim$(.FIELDS) <> "") Then
                            WS_FIELDS = WS_FIELDS & GET_FIELDS(.FIELDS, WS_PAGESNUM)
                            WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        End If
                        
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                Next J
                
            Case "PAGO_PA"
                If (DLLParams.PRINT_BILL And WS_FLG_SPA) Then
                    With .TEMPLATES(0)
                        WS_FIELDS = WS_FIELDS & Replace$(GET_FIELDS(.FIELDS, WS_PAGESNUM), "XXX", "001")
                        WS_ANNEXED_DATA = WS_ANNEXED_DATA & .FIELDSDATA
                        WS_PAGESINDXS = WS_PAGESINDXS & .TEMP_ID & ","
                        WS_PAGESNUM = (WS_PAGESNUM + .TEMP_PAGES)
                        WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMP_BITWISE)
                    End With
                            
                    WS_PAGESINDXS = WS_PAGESINDXS & .TEMPLATES(1).TEMP_ID & ","
                    WS_PAGESNUM = (WS_PAGESNUM + .TEMPLATES(1).TEMP_PAGES)
                    WS_TEMP_UNIQUEID = (WS_TEMP_UNIQUEID + .TEMPLATES(1).TEMP_BITWISE)
                            
                    WS_PAGEBILL = (WS_PAGESNUM - 1)
                End If
            
            End Select
        End With
    Next I

    With myXFDFMLTemplate
        .setTemplateVersion = DLLParams.TEMPLATEVERSION
        .setTemplateFileName = WS_FILENAME & GET_TEXTPAD(PADRIGHT, Hex$(WS_TEMP_UNIQUEID), 8, "0", False) & "_P" & Format$(WS_PAGESNUM, "000")
        .setTemplateIndexes = Left$(WS_PAGESINDXS, Len(WS_PAGESINDXS) - 1)
        .setExtraFields = WS_FIELDS & WS_EXTRAFIELDS

        ' PAGE NUMBER
        '
        WS_INT = (WS_PAGEBILL - 1)
        
        For I = 1 To WS_INT
            .setFieldId = "NMR_PAGE_R" & Format$(I, "000")
            .setPropertyPageId = I
            .setPropertyCoords = "175,284,25,4"
            .setPropertyFontName = "helr45w.ttf"
            .setPropertyFontSize = "7"
            .setPropertyAlignment = "right"
            .closeProperty
            .closeField
            
            WS_PAGENUM = WS_PAGENUM & "Pag. " & I & " di " & WS_PAGESNUM & IIf((I = WS_INT), "", "|")
        Next I
        
        ' EXTRA DATA
        '
        .setExtraInfo = "pages=" & Chr$(34) & WS_PAGESNUM & Chr$(34)
        .setExtraInfo = "billpage=" & Chr$(34) & WS_PAGEBILL & Chr$(34)

        GET_TEMPLATEINFO = .getXFDFTemplateNode
    End With

End Function

Private Function GET_XMLMETADATA()

    Dim myXMLMD As cls_XMLMetaData

    Set myXMLMD = New cls_XMLMetaData

    With myXMLMD
        .setMetaData("TXT_CAP") = WS_01S.IR_R002.CAP
        .setMetaData("TXT_DESTINATARIO") = Trim$(WS_01S.AN_R001.NOMINATIVO)
        .setMetaData("TXT_INDIRIZZO_RECAPITO") = Trim$(WS_01S.IR_R001.VIA) & IIf(WS_01S.IR_R001.NUMEROCIVICO = "99999", "", " " & Trim$(WS_01S.IR_R001.NUMEROCIVICO))
        .setMetaData("TXT_CLP_RECAPITO") = Trim$(IIf(Trim$(WS_01S.IR_R002.CAP) = "0000", "", WS_01S.IR_R002.CAP) & " " & Trim$(WS_01S.IR_R002.LOCALITÀ) & " " & WS_01S.IR_R002.PROVINCIA)
        .setMetaData("TXT_NATIONALITY") = WS_NATIONALITY

        .setMetaData("IDENTIFICATIVO_DOCUMENTO") = Trim$(WS_01S.IL_R001.CODICESOTTOLOTTO)
        .setMetaData("TIPO_DOCUMENTO") = "SOL"
        .setMetaData("CODICE_CLIENTE") = Trim$(WS_01S.AN_R001.CODICEANAGRAFICO)
        .setMetaData("CODICE_SERVIZIO") = ""
        .setMetaData("PROGRESSIVO_LOTTO") = Trim$(WS_01S.AF_R001.CODICELOTTO)

        GET_XMLMETADATA = .getXMLMetaData
    End With

    Set myXMLMD = Nothing

End Function

Public Sub MMS_Close(isOk As Boolean)

    mySQLImporter.setExtError = (Not isOk)
    mySQLImporter.EndJob
    Set mySQLImporter = Nothing

    Set myXFDFMLTable = Nothing

End Sub

Public Function MMS_GetErrMsg() As String

    MMS_GetErrMsg = WS_ERRMSG

End Function

Public Function MMS_GetErrSctn() As String

    MMS_GetErrSctn = WS_ERRSCT

End Function

Public Function MMS_Insert() As Boolean

    Dim WS_ANNEXED_DATA_TMP         As String
    Dim WS_CF_PIVA                  As String
    Dim WS_COD_NAV                  As String
    Dim WS_RD_DETAIL                As String
    Dim WS_STRING                   As String
    Dim XML_TEMPLATE                As String

    Dim TXT_DESTINATARIO            As String
    Dim TXT_INDIRIZZO               As String
    Dim TXT_CLP                     As String
    Dim WS_DATASCADENZA             As String
    Dim TXT_INTESTATARIO            As String
    Dim TXT_BOLLETTINOIDCLIENTE     As String
    Dim TXT_BOLLETTINOID            As String
    Dim TXT_BOLLETTINOBC            As String
    Dim TXT_BC_PPA_QRCODE           As String
    Dim TXT_BC_PPA_DATAMATRIX       As String
    Dim TXT_DOCFILENAME             As String
    Dim XML_METADATA                As String

    ' INIT
    '
    DDS_INIT
    
    WS_PAGES_RD_CNTR = 0
    WS_SERVICE_HDR = ""
    
    WS_DATAEMISSIONE = IIf((Trim$(DLLParams.PRM_DTAEMISSIONE) = ""), WS_01S.ES_R001.DATAEMISSIONE, DLLParams.PRM_DTAEMISSIONE)
    WS_NATIONALITY = ""
    WS_NMR_SOLLECITO = Format$(WS_01S.AF_R001.CODICELOTTO, "000000") & "/" & Format$(WS_01S.IL_R001.CODICESOTTOLOTTO, "000000")
    WS_NMR_CODANA = Trim$(WS_01S.AN_R001.CODICEANAGRAFICO)
    WS_IMPORTOTOTALE = Trim$(WS_01S.IS_R001.IMPORTOCUMULATIVOSOLLECITATO)
    TXT_BOLLETTINOIDCLIENTE = Mid$(WS_01S.IS_R001.BOLLETTINOID, 2, 18)

    If (Trim$(DLLParams.PRM_GG) = "") Then
        WS_DATASCADENZA = WS_01S.IS_R002.DATAPAGAMENTO
    Else
        WS_DATASCADENZA = Format$(DateAdd("d", DLLParams.PRM_GG, WS_DATAEMISSIONE), "dd/MM/yyyy")
    End If

    ' PAGE 01
    '
    TXT_DESTINATARIO = Trim$(WS_01S.AN_R001.NOMINATIVO)

    With WS_01S.IR_R001
        TXT_INDIRIZZO = Trim$(.VIA) & IIf((.NUMEROCIVICO = "99999"), "", " " & Trim$(.NUMEROCIVICO)) & IIf((Trim$(.SUFFISSO) = ""), "", " " & Trim$(.SUFFISSO))

        WS_STRING = IIf(Trim$(.SCALA) = "", "", "SCALA " & Trim$(.SCALA)) & _
                    IIf(Trim$(.PIANO) = "", "", " - PIANO " & Trim$(.PIANO)) & _
                    IIf(Trim$(.INTERNO) = "", "", " - INT. " & Trim$(.INTERNO))

        If (Trim$(WS_STRING) <> "") Then
            If (Left$(WS_STRING, 3) = " - ") Then WS_STRING = Mid$(WS_STRING, 4)

            TXT_INDIRIZZO = TXT_INDIRIZZO & "<br>" & WS_STRING
        End If
    End With
    
    If (Trim$(WS_01S.IR_R002.CAP) <> "") Then If (Val(WS_01S.IR_R002.CAP) = 0) Then WS_01S.IR_R002.CAP = ""
    
    If ((Trim$(WS_01S.IR_R002.SIGLA_NAZIONE) = "ITA") Or (Trim$(WS_01S.IR_R002.SIGLA_NAZIONE) = "")) Then
        TXT_CLP = Trim$(WS_01S.IR_R002.CAP & " " & Trim$(WS_01S.IR_R002.LOCALITÀ) & " " & WS_01S.IR_R002.PROVINCIA)
    Else
        WS_NATIONALITY = Trim$(WS_01S.IR_R002.NAZIONALITÀ)
        TXT_CLP = Trim$(WS_01S.IR_R002.CAP & " " & Trim$(WS_01S.IR_R002.LOCALITÀ) & "<br>" & WS_NATIONALITY)
    End If

    TXT_INTESTATARIO = Trim$(WS_01S.NF_R001.NOMINATIVOFORNITURA)

    Select Case DLLParams.LAYOUT
    Case "L03"  ' SOLLECITO BONARIO
        WS_RD_DETAIL = GET_DOCUMENTDETAILS_L03
    
    Case "L09"  ' LETTERA NOTE INTERRUTTIVE
        WS_RD_DETAIL = GET_DOCUMENTDETAILS_L09
    
    Case "L10"  ' LETTERA NOTE INTERRUTTIVE - NO BOLLETTINO
        WS_RD_DETAIL = GET_DOCUMENTDETAILS_L10
    
    Case "L17"  ' COSTITUZIONE IN MORA
        WS_RD_DETAIL = GET_DOCUMENTDETAILS_L17

    Case Else
        WS_RD_DETAIL = GET_DOCUMENTDETAILS_LXX
    
    End Select

    ' TEMPLATES BUILDER
    '
    XML_TEMPLATE = GET_TEMPLATEINFO
    WS_ANNEXED_DATA_TMP = WS_ANNEXED_DATA

    ' DYNAMIC DATA SUPPORT
    '
    DDS_ADD "[$TXT_DESTINATARIO]", TXT_DESTINATARIO
    DDS_ADD "[$TXT_INTESTATARIO]", TXT_INTESTATARIO
    DDS_ADD "[$TXT_INDIRIZZO]", TXT_INDIRIZZO
    DDS_ADD "[$TXT_CLP]", TXT_CLP
    DDS_ADD "[$TXT_LOCALITÀEMISSIONE]", GET_CAPITALIZED_STRING(WS_01S.ES_R001.LOCALITÀEMISSIONE)
    DDS_ADD "[$DTA_EMISSIONE]", WS_DATAEMISSIONE
    DDS_ADD "[$DTA_SCADENZA]", WS_DATASCADENZA
    DDS_ADD "[$TXT_PROTOCOLLO]", WS_NMR_SOLLECITO
    DDS_ADD "[$TXT_CODANA]", WS_NMR_CODANA
    DDS_ADD "[$TXT_GG]", DLLParams.PRM_GG
    DDS_ADD "[$TXT_GG_DESCR]", WS_GG_DESCR
    DDS_ADD "[$IMPORTOTOTALE]", WS_IMPORTOTOTALE
    DDS_ADD "[$XML_PAGEFOOTER]", WS_PAGE_FOOTER
    DDS_ADD "[$CODELINE]", TXT_BOLLETTINOIDCLIENTE
    
    ' BILL DATA
    '
    WS_STRING = Replace$(Left$(WS_01S.IS_R001.IMPORTO, 11), "+", "")

    TXT_BOLLETTINOID = WS_01S.IS_R001.BOLLETTINOID & GET_TEXTPAD(PADRIGHT, WS_01S.IS_R001.IMPORTO, 18, " ", False) & GET_TEXTPAD(PADRIGHT, Format$(DLLParams.CCP_BILL, String$(12, "0")), 14, " ", False) & "<  896>"
    TXT_BOLLETTINOBC = "18" & TXT_BOLLETTINOIDCLIENTE & "12" & Format$(DLLParams.CCP_BILL, String$(12, "0")) & "10" & WS_STRING & "3896"
    
    ' PAGO PA
    '
    ' WS_STRING = Replace$(Left$(WS_01S.IS_R001.IMPORTO, 11), "+", "")
    
    If (WS_FLG_SPA) Then
        WS_CF_PIVA = Trim$(IIf(Trim$(WS_01S.NF_R001.PARTITA_IVA) = "", WS_01S.NF_R001.CODICE_FISCALE, WS_01S.NF_R001.PARTITA_IVA))
        WS_COD_NAV = Trim$(WS_01S.PA_R001.CODICE_NAV)
        
        DDS_ADD "[$VAR_AMOUNT]", WS_IMPORTOTOTALE
        DDS_ADD "[$CODICE_AVVISO]", GET_PAGOPACODE(WS_COD_NAV)
        DDS_ADD "[$VAR_DTAEMISSIONE]", WS_DATAEMISSIONE
        DDS_ADD "[$VAR_DOCNUM]", WS_NMR_SOLLECITO
        DDS_ADD "[$VAR_DTAEMISSIONE]", WS_DATAEMISSIONE
        DDS_ADD "[$VAR_DTASCADENZA]", WS_DATASCADENZA
        DDS_ADD "[$VAR_INTESTATARIO]", TXT_INTESTATARIO
        DDS_ADD "[$VAR_UBIC]", TXT_INDIRIZZO & "<br>" & TXT_CLP

        TXT_BC_PPA_QRCODE = "PAGOPA|002|" & WS_COD_NAV & "|" + DLLParams.CF_ENTE & "|" & WS_STRING
        TXT_BC_PPA_DATAMATRIX = "codfase=NBPA;18" & WS_COD_NAV & "12" & Format$(DLLParams.CCP_PPA, String$(12, "0")) & "10" & WS_STRING & "38961P1" & DLLParams.CF_ENTE & GET_TEXTPAD(PADLEFT, WS_CF_PIVA, 16, " ", True) & GET_TEXTPAD(PADLEFT, TXT_INTESTATARIO, 40, " ", False) + GET_TEXTPAD(PADLEFT, UCase$("Sollecito num. " & Replace$(WS_NMR_SOLLECITO, "/", "-") & " del " & Replace$(WS_DATAEMISSIONE, "/", "-")), 110, " ", False) + "            A"
    End If

    WS_ANNEXED_DATA_TMP = GET_FIELDSDATA(WS_ANNEXED_DATA_TMP)
    
    ' CFO
    '
    TXT_DOCFILENAME = "S" & WS_01S.AF_R001.CODICEAZIENDASERVER & Replace$(WS_NMR_SOLLECITO, "/", "")

    ' XML METADATA
    '
    XML_METADATA = GET_XMLMETADATA

    ' RECORD
    '
    WS_STRING = XML_TEMPLATE & "§" & WS_ANNEXED_DATA_TMP & "§" & _
                TXT_DESTINATARIO & "§" & TXT_INDIRIZZO & "§" & TXT_CLP & "§" & _
                WS_RD_DETAIL & "§" & WS_IMPORTOTOTALE & "§" & WS_DATASCADENZA & "§" & WS_NMR_SOLLECITO & "§" & WS_NMR_CODANA & "§" & _
                TXT_INTESTATARIO & "§" & _
                TXT_BOLLETTINOIDCLIENTE & "§" & TXT_BOLLETTINOID & "§" & TXT_BOLLETTINOBC & "§" & _
                TXT_BC_PPA_QRCODE & "§" & TXT_BC_PPA_DATAMATRIX & "§" & _
                TXT_DOCFILENAME & "§" & XML_METADATA & "§" & WS_PAGENUM

    ' INSERT DATA
    '
    MMS_Insert = mySQLImporter.SQLInsert(WS_STRING)
    WS_ERRMSG = mySQLImporter.GetUMErrorMessage
    WS_ERRSCT = "MMS_INSERT"

    DoEvents

End Function

Public Function MMS_Open() As Boolean

    Dim myData() As String
    ReDim myData(0) As String

    myData(0) = "§"

    Set myXFDFMLTable = New cls_XFDFMLTable
    Set mySQLImporter = New SQL_Import.PlugIn

    With mySQLImporter
        .DSN = DLLParams.DSN
        .EXTRAPARAMS = myData
        .IDDATACUTTER = DLLParams.IDDATACUTTER
        .TABLENAME = DLLParams.TABLENAME
        .TNS = DLLParams.TNS

        MMS_Open = .StartJob

        If MMS_Open Then
            DLLParams.IDWORKINGLOAD = .GetIdWorkingLoad
        Else
            MsgBox .GetUMErrorMessage, vbExclamation, "Guru Meditation:"
        End If
    End With

End Function
