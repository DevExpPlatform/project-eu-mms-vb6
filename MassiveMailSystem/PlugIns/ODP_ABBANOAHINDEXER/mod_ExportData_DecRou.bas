Attribute VB_Name = "mod_ExportData_DecRou"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)

Private Type strct_POSTEL_HDR
    DESCRIZIONE         As String * 13  ' NESSUN VALORE
    NOMELOTTO           As String * 8   ' CODICE ENG IDENTIFICATIVO DEL LOTTO DI ELABORAZIONE STAMPA. NOME UNIVOCO IDENTIFICATIVO DEL LOTTO. SOLLECITI: S1XXXXXX; - PROGETTO L01, S3XXXXXX; - PROGETTO L03, S5XXXXXX; - PROGETTO L05 | BOLLETTE: B1XXXXXX; DOVE XXXXXX È IL PROGRESSIVO ELABORAZIONE MMS ALLINEATO A DESTRA FILLER 0 AD EX: B1000001. S5000001 È IL LOTTO EFFETTIVAMENTE PRESENTE IN AMBIENTE DI STAGE DEL 28/02/2018
    ZUTENTE             As String * 8   ' NOME UTENTE POSTEL DEDICATO DA ABBANOA
    PROCEDURA           As String * 8   ' NOME DEL PROGETTO POSTEL DEDICATO DA ABBANOA
    TOTALEINDIRIZZI     As String * 6   ' NUMERO TOTALE DOCUMENTI PRESENTI NELL'ELABORAZIONE
    RESPONSABILE        As String * 45  ' CONSIGLIATO DA POSTEL
    INDIRIZZO01         As String * 45  ' CONSIGLIATO DA POSTEL
    INDIRIZZO02         As String * 45  ' CONSIGLIATO DA POSTEL
    TELEFONO            As String * 15  ' CONSIGLIATO DA POSTEL
    EMAIL               As String * 45  ' CONSIGLIATO DA POSTEL
    FAX                 As String * 15  ' CONSIGLIATO DA POSTEL
    TIPOBUSTA           As String * 8   ' STANDARD
    TIPOSTAMPA          As String * 1   ' D
    TIPOCARTA           As String * 8   ' STANDARD
    MODALITÀINVIO_OLD   As String * 1   ' *
    MITTENTELINE01      As String * 44  ' ABBANOA - SI PER SOLLECITI. CONSIGLIATO PER LE BOLLETTE
    MITTENTELINE02      As String * 44  ' ABBANOA - SI PER SOLLECITI. CONSIGLIATO PER LE BOLLETTE
    MITTENTELINE03      As String * 44  ' ABBANOA - SI PER SOLLECITI. CONSIGLIATO PER LE BOLLETTE
    MITTENTELINE04      As String * 44  ' ABBANOA - SI PER SOLLECITI. CONSIGLIATO PER LE BOLLETTE
    LAVORAZIONE         As String * 3   ' SS  SI  SOLO STAMPA 3
    ADDRDOM01           As String * 44  ' NESSUN VALORE
    ADDRDOM02           As String * 44  ' NESSUN VALORE
    ADDRDOM03           As String * 44  ' NESSUN VALORE
    ADDRDOM04           As String * 44  ' NESSUN VALORE
    COLORE              As String * 2   ' FC
    MODALITÀINVIO_NEW   As String * 3   ' PC BOLLETTE - XM SOLLECITI
End Type

Private Type strct_POSTEL_ROW
    CAP                 As String * 5   ' CAP DESTINAZIONE
    INSERTO01           As String * 8   ' NESSUN VALORE
    TIPOINSERTO01       As String * 3   ' NESSUN VALORE
    NUMFOGLIINSERTO01   As String * 2   ' NESSUN VALORE
    INSERTO02           As String * 8   ' NESSUN VALORE
    TIPOINSERTO02       As String * 3   ' NESSUN VALORE
    NUMFOGLIINSERTO02   As String * 2   ' NESSUN VALORE
    CATEGORIA           As String * 3   ' NESSUN VALORE
    RIGA01              As String * 44  ' PRIMA RIGA INDIRIZZO - DESTINATARIO
    RIGA02              As String * 44  ' SECONDA RIGA INDIRIZZO - DESTINATARIO (PRESSO)
    RIGA03              As String * 44  ' TERZA RIGA INDIRIZZO - VIA, PIAZZA, ...
    RIGA04              As String * 44  ' QUARTA RIGA INDIRIZZO - CAP LOCALITÀ PROV.
    RIGA05              As String * 44  ' QUINTA RIGA INDIRIZZO - STATO (SE ESTERO)
    NOMEPDF             As String * 20  ' NOME FILE
    PAG_DA              As String * 8   ' PAGINA ALL'INTERNO DEL PDF ALLA QUALE COMINCIA LA LETTERA, LA PRIMA PAGINA È LA "00000001"
    PAG_A               As String * 8   ' PAGINA ALL'INTERNO DEL PDF ALLA QUALE TERMINA LA LETTERA. DEVE ESSERE MAGGIORE O UGUALE AL VALORE DEL CAMPO "DA PAG", OPPURE "00000000" SE LA LETTERA COMPRENDE TUTTO IL PDF. AD ESEMPIO SE LA LETTERA È COSTITUITA DALLA SOLA PAGINA 3, "DA PAG" VALE "00000003" ED "A PAG" VALE "00000003"
    CODICEUNIVOCO       As String * 20  ' COME NOME FILE
    CENTROCOSTO         As String * 8   ' B PER LE BOLLETTE - S PER I SOLLECITI
    PAGINABOLLETTINO    As String * 40  ' DA VALORIZZARE SE PRESENTE IL BOLLETTINO POSTALE. ELENCO PAGINE DEL DOCUMENTO CONTENENTI UN BOLLETTINO POSTALE DA PERFORARE; LA PRIMA PAGINA È LA PAGINA 1; LE PAGINE SONO SEPARATE DA PUNTO E VIRGOLA. AD ES. PER PERFORARE LE PAGINE 2, 3 E 5 DEL DOCUMENTO INDICARE: "2;3;5"
End Type

Private Type strct_POSTEL
    HEADER              As strct_POSTEL_HDR
    ROW                 As strct_POSTEL_ROW
End Type

Private Function GET_POSTEL_HDR(WS_DATADST As strct_POSTEL_HDR) As String

    Dim WS_LEN      As Long
    Dim WS_STRING   As String

    WS_LEN = Len(WS_DATADST)
    WS_STRING = Space$(WS_LEN)

    CopyMemory ByVal WS_STRING, WS_DATADST, WS_LEN

    GET_POSTEL_HDR = WS_STRING

End Function

Private Function GET_POSTEL_ROW(WS_DATADST As strct_POSTEL_ROW) As String

    Dim WS_LEN      As Long
    Dim WS_STRING   As String

    WS_LEN = Len(WS_DATADST)
    WS_STRING = Space$(WS_LEN)

    CopyMemory ByVal WS_STRING, WS_DATADST, WS_LEN

    GET_POSTEL_ROW = WS_STRING

End Function

Public Function ODP_ABBANOAHINDEXER_Exporter() As Boolean

    On Error GoTo ErrHandler

    Dim Idx             As Integer
    Dim MAXPACKS        As Integer
    Dim myAPB           As cls_APB
    Dim RS              As ADODB.Recordset
    Dim xAttribute      As IXMLDOMNamedNodeMap
    Dim xNode           As MSXML2.DOMDocument60
    Dim WS_CNTR         As Long
    Dim WS_FILEDST      As String
    Dim WS_FILEDST_BOL  As String
    Dim WS_FILEDST_PDF  As String
    Dim WS_FILEDST_PDZ  As String
    Dim WS_FILEDST_XML  As String
    Dim WS_FILENAMES    As New Collection
    Dim WS_FILESRC      As String
    Dim WS_FLGCFO       As Boolean
    Dim WS_PAGEBILLS    As String
    Dim WS_PAGESCNTR    As Long
    Dim WS_PAGESNUM     As Long
    Dim WS_POSTEL       As strct_POSTEL
    Dim WS_STRING       As String
    
    ' INIT
    '
    Set myAPB = New cls_APB
    Set RS = New ADODB.Recordset
    
    Set xNode = New MSXML2.DOMDocument60
    
    WS_CNTR = 1
    WS_FLGCFO = (DLLParams.ZIPEXEPATH = "")
    
    WS_STRING = String$(Len(WS_POSTEL), " ")
    CopyMemory ByVal VarPtr(WS_POSTEL), ByVal StrPtr(WS_STRING), Len(WS_POSTEL) * 2
    
    DBConn.Open
    
    With WS_POSTEL.HEADER
        .ZUTENTE = "Z0004504"
        .TIPOBUSTA = "STANDARD"
        .TIPOSTAMPA = "D"
        .TIPOCARTA = "STANDARD"
        .MODALITÀINVIO_OLD = "*"
        .LAVORAZIONE = "SS"
        
        Select Case DLLParams.MODE
        Case "B"
            .PROCEDURA = "ABBANFAT"
            .NOMELOTTO = "B1" & Format$(DB_GetValueByID("SELECT ID_WRKCNTR FROM MMS.EST_WABBNAH2OLOG WHERE (SUBSTR(STR_MODE, 0, 1) = 'B') AND (ID_WORKINGLOAD = " & DLLParams.WORKINGID & ")"), "000000")
            .MODALITÀINVIO_NEW = "PC4"
            .COLORE = "FC"
        
        Case "L"
            .PROCEDURA = "ABBANLTR"
            .NOMELOTTO = "L" & Format$(DB_GetValueByID("SELECT ID_WRKCNTR FROM MMS.EST_WABBNAH2OLOG WHERE (SUBSTR(STR_MODE, 0, 1) = '" & Left$(DLLParams.TYPE, 1) & "') AND (ID_WORKINGLOAD = " & DLLParams.WORKINGID & ")"), "000000")
            .MITTENTELINE01 = "ABBANOA"
            .MITTENTELINE02 = "CSA MILANO"
            .MITTENTELINE03 = "PIAZZA VESUVIO, 6"
            .MITTENTELINE04 = "20144 MILANO MI"
            .MODALITÀINVIO_NEW = "XM"
            .COLORE = "BN"
        
        Case "S"
            .PROCEDURA = "ABBANSOL"
            .NOMELOTTO = "S" & Format$(DB_GetValueByID("SELECT ID_WRKCNTR FROM MMS.EST_WABBNAH2OLOG WHERE (SUBSTR(STR_MODE, 0, 1) = '" & Left$(DLLParams.TYPE, 1) & "') AND (ID_WORKINGLOAD = " & DLLParams.WORKINGID & ")"), "000000")
            .MITTENTELINE01 = "ABBANOA"
            .MITTENTELINE02 = "CSA MILANO"
            .MITTENTELINE03 = "PIAZZA VESUVIO, 6"
            .MITTENTELINE04 = "20144 MILANO MI"
            .MODALITÀINVIO_NEW = "XM"
            .COLORE = "BN"
        
        End Select
    End With
    
    MAXPACKS = DB_GetValueByID("SELECT MAX(ID_PACCO) AS MAXPACKS FROM " & DLLParams.WORKINGTABLE & " WHERE ID_WORKINGLOAD = " & DLLParams.WORKINGID)
    
    With myAPB
        .APBMode = PBSingle
        .APBCaption = "Export Data Processor:"
        .APBMaxItems = MAXPACKS
        .APBShow
    End With
    
    ' EXEC
    '
    For Idx = 1 To MAXPACKS
        myAPB.APBItemsLabel = "Package: " & Format$(Idx, "000") & "/" & Format$(MAXPACKS, "000")
        myAPB.APBItemsProgress = Idx
        
        Set RS = DBConn.Execute("SELECT XML_TEMPLATE, XML_METADATA, TXT_DOCFILENAME FROM " & DLLParams.WORKINGTABLE & _
                                " WHERE (ID_WORKINGLOAD = " & DLLParams.WORKINGID & ") AND (ID_PACCO = " & Idx & ")" & IIf((WS_FLGCFO Or (DLLParams.MODE = "L") Or (DLLParams.MODE = "S")), "", " AND (FLG_NOMERGE IS NULL)") & _
                                " ORDER BY ID_POSIZIONE")

        If RS.RecordCount > 0 Then
            If (WS_FLGCFO) Then
                If (DLLParams.MODE = "B") Then
                    WS_FILEDST_BOL = Left$(DLLParams.OUTPUTFILEPATH, Len(DLLParams.OUTPUTFILEPATH) - 4)
                Else
                    WS_FILEDST_BOL = DLLParams.OUTPUTFILEPATH
                End If
                
                WS_FILEDST_BOL = WS_FILEDST_BOL & "A" & DLLParams.MODE & DLLParams.WORKINGID & ".BOL"
                WS_POSTEL.HEADER.TOTALEINDIRIZZI = DB_GetValueByID("SELECT COUNT(*) AS NMR_DOCS FROM " & DLLParams.WORKINGTABLE & " WHERE ID_WORKINGLOAD = " & DLLParams.WORKINGID)
                
                If (Idx = 1) Then Open WS_FILEDST_BOL For Output As #1
                
                Print #1, GET_POSTEL_HDR(WS_POSTEL.HEADER)
                
                Do Until RS.EOF
                    ' GET INDEXED METADATA
                    '
                    xNode.loadXML RS("XML_METADATA")
                
                    WS_STRING = String$(Len(WS_POSTEL.ROW), " ")
                    CopyMemory ByVal VarPtr(WS_POSTEL.ROW), ByVal StrPtr(WS_STRING), Len(WS_POSTEL.ROW) * 2
                    
                    With WS_POSTEL.ROW
                        .CAP = xNode.getElementsByTagName("TXT_CAP").Item(0).Text
                        .RIGA01 = xNode.getElementsByTagName("TXT_DESTINATARIO").Item(0).Text
                        .RIGA03 = xNode.getElementsByTagName("TXT_INDIRIZZO_RECAPITO").Item(0).Text
                        .RIGA04 = xNode.getElementsByTagName("TXT_CLP_RECAPITO").Item(0).Text
                        .RIGA05 = xNode.getElementsByTagName("TXT_NATIONALITY").Item(0).Text
                        .NOMEPDF = RS("TXT_DOCFILENAME")
                        .PAG_DA = "00000000"
                        .PAG_A = "00000000"
                        .CODICEUNIVOCO = .NOMEPDF
                        .CENTROCOSTO = DLLParams.MODE
                        .PAGINABOLLETTINO = WS_PAGEBILLS
                    End With
                    
                    Print #1, GET_POSTEL_ROW(WS_POSTEL.ROW)
                
                    RS.MoveNext
                        
                    If ((RS.AbsolutePosition Mod 200) = 0) Then DoEvents
                Loop
                
                If (Idx = MAXPACKS) Then Close #1
            Else
                WS_FILESRC = DLLParams.OUTPUTFILEPATH & DLLParams.BASEFILENAME & "_P" & Format$(Idx, "000") & ".PDF"
                WS_FILEDST = DLLParams.OUTPUTFILEPATH & "A" & DLLParams.MODE & DLLParams.WORKINGID & Format$(Idx, "000")
                WS_FILEDST_BOL = WS_FILEDST & ".BOL"
                WS_FILEDST_PDF = WS_FILEDST & ".PDF"
                WS_FILEDST_PDZ = WS_FILEDST & ".PDZ"
                WS_FILEDST_XML = WS_FILEDST & ".XML"
                WS_POSTEL.HEADER.TOTALEINDIRIZZI = RS.RecordCount
                
                WS_FILENAMES.Add WS_FILEDST & ".t"
                
                Name WS_FILESRC As WS_FILEDST_PDF
            
                DoEvents
                
                Open WS_FILEDST_BOL For Output As #1
                Open WS_FILEDST_XML For Output As #2
                    Print #1, GET_POSTEL_HDR(WS_POSTEL.HEADER)
                    Print #2, "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine & _
                              "<DATASET>" & vbNewLine & _
                              vbTab & "<GLOBAL>" & vbNewLine & _
                              vbTab & vbTab & "<PROC_ARC/>" & vbNewLine & _
                              vbTab & vbTab & "<PRODOTTO/>" & vbNewLine & _
                              vbTab & "</GLOBAL>"
                    
                    WS_PAGESCNTR = 1
                    WS_CNTR = 1
    
                    Do Until RS.EOF
                        xNode.loadXML RS("XML_TEMPLATE")
                        Set xAttribute = xNode.getElementsByTagName("extrainfo").Item(0).Attributes
    
                        WS_PAGESNUM = Val(xAttribute.getNamedItem("pages").nodeValue)
                        WS_PAGEBILLS = xAttribute.getNamedItem("billpage").nodeValue
                        
                        ' GET INDEXED METADATA
                        '
                        xNode.loadXML RS("XML_METADATA")
                        
                        WS_STRING = String$(Len(WS_POSTEL.ROW), " ")
                        CopyMemory ByVal VarPtr(WS_POSTEL.ROW), ByVal StrPtr(WS_STRING), Len(WS_POSTEL.ROW) * 2
                        
                        With WS_POSTEL.ROW
                            .CAP = xNode.getElementsByTagName("TXT_CAP").Item(0).Text
                            .RIGA01 = xNode.getElementsByTagName("TXT_DESTINATARIO").Item(0).Text
                            .RIGA03 = xNode.getElementsByTagName("TXT_INDIRIZZO_RECAPITO").Item(0).Text
                            .RIGA04 = xNode.getElementsByTagName("TXT_CLP_RECAPITO").Item(0).Text
                            .RIGA05 = xNode.getElementsByTagName("TXT_NATIONALITY").Item(0).Text
                            .NOMEPDF = GET_BASENAME(WS_FILEDST, False)
                            .PAG_DA = Format$(WS_PAGESCNTR, "00000000")
                            .PAG_A = Format$((WS_PAGESCNTR + (WS_PAGESNUM - 1)), "00000000")
                            .CODICEUNIVOCO = .NOMEPDF
                            .CENTROCOSTO = DLLParams.MODE
                            .PAGINABOLLETTINO = WS_PAGEBILLS
                        End With
    
                        WS_PAGESCNTR = (WS_PAGESCNTR + WS_PAGESNUM)
    
                        Print #1, GET_POSTEL_ROW(WS_POSTEL.ROW)
                        Print #2, vbTab & "<DOCUMENT iddoc=""" & Format$(WS_CNTR, "0000000") & """>" & vbNewLine & _
                                  vbTab & vbTab & "<IDENTIFICATIVO_DOCUMENTO value=""" & xNode.getElementsByTagName("IDENTIFICATIVO_DOCUMENTO").Item(0).Text & """/>" & vbNewLine & _
                                  vbTab & vbTab & "<TIPO_DOCUMENTO value=""" & xNode.getElementsByTagName("TIPO_DOCUMENTO").Item(0).Text & """/>" & vbNewLine & _
                                  vbTab & vbTab & "<CODICE_CLIENTE value=""" & xNode.getElementsByTagName("CODICE_CLIENTE").Item(0).Text & """/>" & vbNewLine & _
                                  vbTab & vbTab & "<CODICE_SERVIZIO value=""" & xNode.getElementsByTagName("CODICE_SERVIZIO").Item(0).Text & """/>" & vbNewLine & _
                                  vbTab & vbTab & "<PROGRESSIVO_LOTTO value=""" & xNode.getElementsByTagName("PROGRESSIVO_LOTTO").Item(0).Text & """/>" & vbNewLine & _
                                  vbTab & "</DOCUMENT>"
                        
                        WS_CNTR = (WS_CNTR + 1)
                        
                        RS.MoveNext
                        
                        If ((RS.AbsolutePosition Mod 200) = 0) Then DoEvents
                    Loop
                    
                    Print #2, "</DATASET>"
                    
                    RS.Close
                Close #1
                Close #2
                
                If (FDEXIST(WS_FILEDST_PDZ, False)) Then Kill WS_FILEDST_PDZ
        
                ExecuteAndWait DLLParams.ZIPEXEPATH & " " & WS_FILEDST_PDZ & " " & WS_FILEDST_BOL & " " & WS_FILEDST_PDF & " " & WS_FILEDST_XML
                
                If (FDEXIST(WS_FILEDST_PDZ, False)) Then
                    Kill WS_FILEDST_BOL
                    Kill WS_FILEDST_PDF
                    Kill WS_FILEDST_XML
                End If
            End If
        End If

        DoEvents
    Next Idx
    
    If (WS_FILENAMES.Count > 0) Then
        For Idx = 1 To WS_FILENAMES.Count
            Open WS_FILENAMES.Item(Idx) For Output As #1
            Close #1
        Next Idx
    End If
    
    GoSub CleanUp
    
    ODP_ABBANOAHINDEXER_Exporter = True
    
    Exit Function
    
CleanUp:
    Set WS_FILENAMES = Nothing
    Set xAttribute = Nothing
    Set xNode = Nothing
    Set RS = Nothing
        
    DBConn.Close
    
    myAPB.APBClose
    Set myAPB = Nothing
Return

ErrHandler:
    Close #1
    Close #2

    GoSub CleanUp
    
    MsgBox Err.DESCRIPTION, vbExclamation, "Attenzione:"

End Function
