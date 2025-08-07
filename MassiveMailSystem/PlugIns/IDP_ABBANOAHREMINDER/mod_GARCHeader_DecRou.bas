Attribute VB_Name = "mod_GARCHeader_DecRou"
Option Explicit

' **************************************************
' *
' * LETTERA SOLLECITO
' *
' **************************************************
'
' GARC HEADER
'
Public Type strct_GARC_HEADER
    GROUP                           As String * 3
    SUBGROUP                        As String * 2
    ROWNUMBER                       As String * 3
End Type

' 01S AF
'
Private Type strct_01S_AF_R001
    CODICEAZIENDASERVER             As String * 2
    DESCRIZIONEAZIENDA              As String * 40
    CODICELOTTO                     As String * 6
    TIPODOCUMENTO                   As String * 2       ' 11 SOLLECITO SEMPLICE - 12 AVVISO DI SOSPENSIONE
End Type

' 01S AN
'
Public Type strct_01S_AN_R001
    CODICEANAGRAFICO                As String * 10
    NOMINATIVO                      As String * 40
End Type

' 01S BO
'
Public Type strct_01S_BO_R001
    IMPORTO_RESIDUO_BONUS           As String * 15
    PROTOCOLLO_RICHIESTA            As String * 21
    PROTOCOLLO_INOLTRO              As String * 21
End Type

' 01S ED
'
Private Type strct_01S_ED_R001_S_E
    CODICESERVIZIO                  As String * 10
End Type

Private Type strct_01S_ED_R001_S_I
    VIA                             As String * 30
    NUMEROCIVICO                    As String * 5       ' SE NUMERO CIVICO 0 OPPURE 99999, ASSEGNARE SPAZIO
    SUFFISSO                        As String * 3
    SCALA                           As String * 3
    PIANO                           As String * 3
    INTERNO                         As String * 3
End Type

Private Type strct_01S_ED_R001_S_L
    CAP                             As String * 5
    LOCALITÀ                        As String * 30
    PROVINCIA                       As String * 2
End Type

Private Type strct_01S_ED_R001_S_T
    CODICEMESSAGGIO                 As String * 2       ' PUÒ NON ESSERE PRESENTE IL CODICE.
    FILLER                          As String * 62
    IMPORTO                         As String * 15
End Type

Private Type strct_01S_ED_R001_A_C
    CODICEANAGRAFICO                As String * 10
End Type

Private Type strct_01S_ED_RXXX_A_D
    TIPODOCUMENTO                   As String * 1       ' B = BOLLETTA; F = FATTURA; C = CORRISPETTIVO.
    DESCRIZIONE                     As String * 15
    TIPOSERVIZIO                    As String * 2       ' AP = ACQUA; EN = ENERGIA EL; GM = GAS METANO
    DESCRIZIONETIPOSERVIZIO         As String * 10
    TIPOSERVIZIOPERNUMERAZIONE      As String * 2       ' CON GLI ZERI MENO SIGNIFICATIVI.
    ANNO                            As String * 4
    NUMERO                          As String * 8       ' ULTIMI OTTO CARATTERI
    RATA                            As String * 2       ' CON GLI ZERI MENO SIGNIFICATIVI.
    DATAEMISSIONE                   As String * 10      ' NEL FORMATO GG/MM/AAAA
    DATASCADENZA                    As String * 10      ' NEL FORMATO GG/MM/AAAA
    IMPORTO                         As String * 15
    FLG_IMPORTI_PRESCRIVIBILI       As String * 1       ' PRESENTE E VALORIZZATO SOLAMENTE SE A SISTEMA VI È OPZ.COD.TSE ATTDEL547B ATTIVA. VALORE S/N
    TOTALE_IMPORTI_PRESCRIVIBILI    As String * 15      ' PRESENTE E VALORIZZATO SOLAMENTE SE A SISTEMA VI È OPZ.COD.TSE ATTDEL547B ATTIVA.
    NUM_SOLLECITO_BONARIO           As String * 13
    DATA_SOLLECITO_BONARIO          As String * 10
    DATA_RICEZIONE_SB               As String * 10
End Type

Private Type strct_01S_ED_RXXX_S_D
    TIPODOCUMENTO                   As String * 1       ' B = BOLLETTA; F = FATTURA; C = CORRISPETTIVO.
    DESCRIZIONE                     As String * 15
    TIPOSERVIZIO                    As String * 2       ' EN = ENERGIA EL; GM = GAS METANO.
    DESCRIZIONETIPOSERVIZIO         As String * 10
    TIPOSERVIZIOPERNUMERAZIONE      As String * 2       ' CON GLI ZERI MENO SIGNIFICATIVI.
    ANNO                            As String * 4
    NUMERO                          As String * 8       ' ULTIMI OTTO CARATTERI
    RATA                            As String * 2       ' CON GLI ZERI MENO SIGNIFICATIVI.
    DATAEMISSIONE                   As String * 10      ' NEL FORMATO GG/MM/AAAA
    DATASCADENZA                    As String * 10      ' NEL FORMATO GG/MM/AAAA
    IMPORTO                         As String * 15
    FLG_IMPORTI_PRESCRIVIBILI       As String * 1       ' PRESENTE E VALORIZZATO SOLAMENTE SE A SISTEMA VI È OPZ.COD.TSE ATTDEL547B ATTIVA. VALORE S/N
    TOTALE_IMPORTI_PRESCRIVIBILI    As String * 15      ' PRESENTE E VALORIZZATO SOLAMENTE SE A SISTEMA VI È OPZ.COD.TSE ATTDEL547B ATTIVA.
    NUM_SOLLECITO_BONARIO           As String * 13
    DATA_SOLLECITO_BONARIO          As String * 10
    DATA_RICEZIONE_SB               As String * 10
End Type

Public Type strct_01S_ED_RXXX
    ED_R001_A_C                     As strct_01S_ED_R001_A_C
    ED_R001_S_E                     As strct_01S_ED_R001_S_E
    ED_R001_S_I                     As strct_01S_ED_R001_S_I
    ED_R001_S_L                     As strct_01S_ED_R001_S_L
    ED_R001_S_T                     As strct_01S_ED_R001_S_T
    ED_RXXX_A_D()                   As strct_01S_ED_RXXX_A_D
    ED_RXXX_S_D()                   As strct_01S_ED_RXXX_S_D
    TIPORECORD                      As String * 2
End Type

' 01S ES
'
Private Type strct_01S_ES_R001
    LOCALITÀEMISSIONE               As String * 30
    DATAEMISSIONE                   As String * 10      ' NEL FORMATO GG/MM/AAAA
End Type

' 01S IL
'
Private Type strct_01S_IL_R001
    INIZIOLETTERA                   As String * 2       ' VALORE FISSO ‘IL’
    CODICESOTTOLOTTO                As String * 6
End Type

' 01S IR
'
Private Type strct_01S_IR_R001
    VIA                             As String * 30
    NUMEROCIVICO                    As String * 5       ' SE NUMERO CIVICO 0 OPPURE 99999, ASSEGNARE SPAZIO
    SUFFISSO                        As String * 3
    SCALA                           As String * 3
    PIANO                           As String * 3
    INTERNO                         As String * 3
End Type

Private Type strct_01S_IR_R002
    CAP                             As String * 5
    LOCALITÀ                        As String * 30
    PROVINCIA                       As String * 2
    SIGLA_NAZIONE                   As String * 3
    NAZIONALITÀ                     As String * 30
End Type

' 01S IS
'
Private Type strct_01S_IS_R001
    IMPORTOCUMULATIVOSOLLECITATO    As String * 15
    BOLLETTINOID                    As String * 20      ' NEL FORMATO <999999999999999999> COMPOSTA DA: 2 CRT – VALORE FISSO 99; 2 CRT – CODICE AZIENDA SIU; 6 CRT – NUMERO DI LOTTO; 6 CRT – NUMERO DI SOTTOLOTTO; 2 CRT – CIN DI CONTROLLO.
    IMPORTO                         As String * 12      ' NEL FORMATO 99999999+99> COMPOSTA DA: 8 CRT – IMPORTO DA PAGARE (PARTE INTERA); 2 CRT – IMPORTO DA PAGARE (PARTE DECIMALE).
    STRINGAPOSTECAMPO03             As String * 9       ' NEL FORMATO 99999999< COMPOSTA DA: 8 CRT – CCP POSTALE
    STRINGAPOSTECAMPO04             As String * 4       ' NEL FORMATO <999 COMPOSTA DA: 3 CRT – FISSO IL VALORE 896;
End Type

Private Type strct_01S_IS_R002
    SOMMABOLLETTESERVIZIO           As String * 4       ' SE UNA STESSA BOLLETTA CONTIENE PIÙ SERVIZI QUESTA VIENE CONTATA TANTE VOLTE QUANTI SONO I SERVIZI IN ESSA CONTENUTI.
    DATAPAGAMENTO                   As String * 10      ' DD/MM/YYYY UGUALE A DATA SOLLECITO +15GG PRESA DAL RECORD 01SES
End Type

' 01S1 NF
'
Private Type strct_01S_NF_R001
    NOMINATIVOFORNITURA             As String * 40
    PARTITA_IVA                     As String * 12
    CODICE_FISCALE                  As String * 16
End Type

' 01S PA
'
Private Type strct_01S_PA_R001
    CODICE_NAV                      As String * 35
End Type

' 01S1 SL
'
Private Type strct_01S_SL_R001
    TIPO_RECORD                     As String * 2       ' DS = DETTAGLIO SOTTO-LOTTO
    TIPO_SOLLECITO                  As String * 2       ' SB = SOLLECITO BONARIO, CM = COST.MORA, LL = LETTERE LIMIT.
    DATA_ELABORAZIONE_LOTTO         As String * 10      ' FORMATO GG/MM/YYYY
    TIPO_DATA_ELABORAZIONE          As String * 1       ' I = INVIO, E = EMISSIONE
    DATA_SCADENZA                   As String * 10      ' DEL SOTTO-LOTTO. FORMATO GG/MM/YYYY
    DATA_COSTITUZIONE_MORA          As String * 10      ' FORMATO GG/MM/YYYY
    TERMINE_ULTIMO_PAGAMENTO        As String * 10      ' FORMATO GG/MM/YYYY
    DATA_PREVISTA_LIMITAZIONE       As String * 10      ' FORMATO GG/MM/YYYY
    DATA_PREVISTA_SOSPENSIONE       As String * 10      ' FORMATO GG/MM/YYYY
    COST_MORA_ULTIMI_18_MESI        As String * 1       ' FLAG S/N
    DURATA_LIMITAZIONE              As String * 3       ' NUMERO GIORNI; PAD 0 A SINISTRA
    SPESE_LIMITAZIONE               As String * 1       ' FLAG S/N
    GIORNI_TERMINE_PREV_REGOLATORIE As String * 3       ' DEL 221/20: SOLO PER SOTTOLOTTO SB E CM. GIORNI PER TERMINE PREVISIONI REGOLATORIE  (CAMPO SSS_ GG_TERM_PREV_REG EL SOTTOLOTTO). TALE CAMPO ASSUMERÀ IL VALORE DEI GIORNI PARAMETRICI (DI DEFAULT 40) PREVISTI DALLA NORMATIVA COME TERMINE MINIMO DAL RICEVIMENTO DEL SOLLECITO PER L’ADEMPIMENTO DEGLI OBBLIGHI DI PAGAMENTO.
    DATA_RICEZIONE_ESITO_SB         As String * 10      ' DEL 221/20: SOLO PER SOTTOLOTTO CM. SPONDA PER CALCOLO DEL TULP (CAMPO SSS_DATA_SPONDA). FORMATO GG/MM/YYYY
    TIPO_SOLLECITO_BS_ORIGINE       As String * 10      ' DEL 221/20: SOLO PER SOTTOLOTTO CM. INDICARE SE LA DATA SPONDA È PEC O RACC (CAMPO SSS_TIPO_SPONDA)
End Type

Private Type strct_01S_SL_R002
    TIPO_RECORD                     As String * 2       ' TU = TIPOLOGIA UTENZA
    TIPOLOGIA_UTENZA_665            As String * 40      ' DESCRIZIONE TIPOLOGIA UTENZA (TIPUSO_NORM.TPN_DESCRIZIONE)
    RESIDENTE                       As String * 1       ' FLAG S/N
    CONDOMINIO                      As String * 1       ' FLAG S/N
    PADRE_FIGLIO                    As String * 1       ' FLAG P/F
    BONUS                           As String * 1       ' FLAG S/N
    NON_DISALIMENTABILE             As String * 1       ' FLAG S/N
    CONCESSIONI_RESIDENTI           As String * 6       ' PAD 0 A SINISTRA
    COMPONENTI_NUCLEO_FAMILIARE     As String * 6       ' PAD 0 A SINISTRA
    COEFFICIENTE_FASCIA_AGEVOLATA   As String * 15      ' NUMERO NEL FORMATO 999999,99999999
End Type

' 01S
'
Private Type strct_01S
    AF_R001                         As strct_01S_AF_R001
    AN_R001                         As strct_01S_AN_R001
    BO_R001                         As strct_01S_BO_R001
    ED_RXXX()                       As strct_01S_ED_RXXX
    ES_R001                         As strct_01S_ES_R001
    IL_R001                         As strct_01S_IL_R001
    IR_R001                         As strct_01S_IR_R001
    IR_R002                         As strct_01S_IR_R002
    IS_R001                         As strct_01S_IS_R001
    IS_R002                         As strct_01S_IS_R002
    NF_R001                         As strct_01S_NF_R001
    PA_R001                         As strct_01S_PA_R001
    SL_R001                         As strct_01S_SL_R001
    SL_R002                         As strct_01S_SL_R002
End Type

Public WS_01S                       As strct_01S
Public WS_FLG_BO                    As Boolean
Public WS_FLG_SPA                   As Boolean

Public Sub GARC_DATA_CLEAR()
    
    Dim WS_STRING As String
    
    WS_FLG_BO = False
    WS_FLG_SPA = False
    
    ' 01S
    '
    With WS_01S
        WS_STRING = String$(Len(.AN_R001), " ")
        CopyMemory ByVal VarPtr(.AN_R001), ByVal StrPtr(WS_STRING), Len(.AN_R001) * 2
        
        WS_STRING = String$(Len(.BO_R001), " ")
        CopyMemory ByVal VarPtr(.BO_R001), ByVal StrPtr(WS_STRING), Len(.BO_R001) * 2
        
        Erase .ED_RXXX
        
        WS_STRING = String$(Len(.ES_R001), " ")
        CopyMemory ByVal VarPtr(.ES_R001), ByVal StrPtr(WS_STRING), Len(.ES_R001) * 2
        
        WS_STRING = String$(Len(.IL_R001), " ")
        CopyMemory ByVal VarPtr(.IL_R001), ByVal StrPtr(WS_STRING), Len(.IL_R001) * 2
        
        WS_STRING = String$(Len(.IR_R001), " ")
        CopyMemory ByVal VarPtr(.IR_R001), ByVal StrPtr(WS_STRING), Len(.IR_R001) * 2
    
        WS_STRING = String$(Len(.IR_R002), " ")
        CopyMemory ByVal VarPtr(.IR_R002), ByVal StrPtr(WS_STRING), Len(.IR_R002) * 2
    
        WS_STRING = String$(Len(.IS_R001), " ")
        CopyMemory ByVal VarPtr(.IS_R001), ByVal StrPtr(WS_STRING), Len(.IS_R001) * 2
    
        WS_STRING = String$(Len(.IS_R002), " ")
        CopyMemory ByVal VarPtr(.IS_R002), ByVal StrPtr(WS_STRING), Len(.IS_R002) * 2
    
        WS_STRING = String$(Len(.NF_R001), " ")
        CopyMemory ByVal VarPtr(.NF_R001), ByVal StrPtr(WS_STRING), Len(.NF_R001) * 2
        
        WS_STRING = String$(Len(.PA_R001), " ")
        CopyMemory ByVal VarPtr(.PA_R001), ByVal StrPtr(WS_STRING), Len(.PA_R001) * 2
    
        WS_STRING = String$(Len(.SL_R001), " ")
        CopyMemory ByVal VarPtr(.SL_R001), ByVal StrPtr(WS_STRING), Len(.SL_R001) * 2
        
        WS_STRING = String$(Len(.SL_R002), " ")
        CopyMemory ByVal VarPtr(.SL_R002), ByVal StrPtr(WS_STRING), Len(.SL_R002) * 2
    End With

End Sub

Public Sub GARC_DATA_INIT()

    Dim WS_STRING As String
    
    With WS_01S
        WS_STRING = String$(Len(.AF_R001), " ")
        CopyMemory ByVal VarPtr(.AF_R001), ByVal StrPtr(WS_STRING), Len(.AF_R001) * 2
    End With

End Sub
