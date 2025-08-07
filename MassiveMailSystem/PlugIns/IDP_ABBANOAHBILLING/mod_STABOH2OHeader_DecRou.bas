Attribute VB_Name = "mod_STABOH2OHeader_DecRou"
Option Explicit

' ************************************************************
' *                                                          *
' * TRACCIATO STABOL STANDARD ACQUA (STDH2O)                 *
' *                                                          *
' ************************************************************
'
' ********************************************************************************************************************************************
' * STABO H2O ROWHEADER
' ********************************************************************************************************************************************
'
Public Type strct_STABODF_Header
    GROUP                               As String * 2
    ROWNUMBER                           As String * 3
    TIPOLOGIASEZIONE                    As String * 1
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
End Type

' G00 - R000
'
Private Type strct_G00
    TIPOESTRAZIONE                      As String * 1       ' PUÒ ASSUMERE I SEGUENTI VALORI: • 1 – DATI ESTRATTI DA ORDSTA1 • 2 – DATI ESTRATTI DA ORDSTA
    AZIENDASIU                          As String * 2
    PROGRESSIVOINIZIO                   As String * 6
    PROGRESSIVOFINE                     As String * 6
    TIPOBOLLETTE                        As String * 1       ' PUÒ ASSUMERE I SEGUENTI VALORI: • 'T' – TUTTE • 'P' – SOLO POSITIVE • 'N' – SOLO NEGATIVE • '0' – SOLO ZERO • 'M' – MAGGIORI UGUALI A ZERO
    NUMEROPOSIZIONAMENTOCARTA           As String * 2       ' VALORE FISSO A '0'
    STAMPASOLLECITO                     As String * 1       ' PUÒ ASSUMERE I VALORI 'S' O 'N'
    DATALIMITESOLLECITO                 As String * 10      ' DD/MM/YYYY. SE NON VALORIZZATA SARANNO MOSTRATI "----------"
    TIPIDOCUMENTO                       As String * 4       ' LE TIPOLOGIE DOCUMENTO POSSONO ESSERE LE SEGUENTI: • B = BOLLETTA • F = FATTURA • N=NOTA CREDITO • C = CORRISPETTIVO SE NON VALORIZZATA SARANNO MOSTRATI "----"
    TIPOSOLLECITO                       As String * 1       ' PUÒ ASSUMERE I VALORI: • 'T' – TUTTI • 'S' – SOLO SALDO POSITIVO • 'P' – ALMENO UNA BOLLETTA POSITIVA SE NON VALORIZZATO SARÀ MOSTRATO "-"
    STATIMORAINCLUDERE                  As String * 20      ' SE NON VALORIZZATO SARANNO MOSTRATI 20 CARATTERI "-"
    STATIMORAECLUDERE                   As String * 20      ' SE NON VALORIZZATO SARANNO MOSTRATI 20 CARATTERI "-"
    TIPOSELEZIONE                       As String * 1       ' PUÒ ASSUMERE I VALORI: • 'N' – NORMALE • 'R' – RISTAMPA
    TIPOSTAMPA                          As String * 1       ' VALORE FISSO 'N'
    SEPARAZIONEBOLLETTE                 As String * 1       ' PUÒ ASSUMERE I VALORI: • 'T' – TUTTE • 'S' – SOLO QUELLE CON BOLLO • 'N' – SOLO QUELLE SENZA BOLLO
    SEQUENZASTAMPA                      As String * 10      ' NUMERO DI SEQUENZA UNIVOCO
    BASEDATI                            As String * 1       ' PUÒ ASSUMERE I VALORI: • 'B' – BOLLETTAZIONE • 'G' – REALE
    PROGRESSIVOFATTURAZIONE             As String * 7       ' SE BASE DATI DEFINITIVA ASSUME IL VALORE FISSO 0.
    MONETABASE                          As String * 3       ' VALORE FISSO 'EUR'
    MONETAEURO                          As String * 3       ' VALORE FISSO 'EUR'
    STAMPAWEB                           As String * 1       ' VALORE FISSO '-'
    TIPOCCP                             As String * 1       ' VALORE FISSO 'N'
    STAMPABOLLETTEERRATE                As String * 1       ' PUÒ ASSUMERE I VALORI 'S' O 'N'
    LIVELLOLOG                          As String * 1       ' PUÒ ASSUMERE I SEGUENTI VALORI: • 0 – SPENTO • 1 – PERFORMANCE • 2 – MINIMO • 3 – INTERMEDIO • 4 – MASSIMO
    AUTOBASKET                          As String * 1       ' PUÒ ASSUMERE I VALORI: • 'S' – SI • 'N' – NO
    FLAGCOMUNECCP                       As String * 1       ' PUÒ ASSUMERE I SEGUENTI VALORI: • 0 – NON ABILITATO • 1 – NON ATTIVO • 2 – ATTIVO
End Type

' G01 - R001
'
Private Type strct_G01_R001
    TIPOSERVIZIO                        As String * 2       ' NB: CAMPO VALORIZZATO SOLO PER BOLLETTE MONOSITO (RECORD 001 CAMPO "TIPO BOLLETTAZIONE" DIVERSO DA "S")
    DESCRIZIONE                         As String * 20      ' NB: CAMPO VALORIZZATO SOLO PER BOLLETTE MONOSITO (RECORD 001 CAMPO "TIPO BOLLETTAZIONE" DIVERSO DA "S")
    SEPARATOR_00                        As String * 1       ' "/"
    SEZIONALE                           As String * 2
    SEPARATOR_01                        As String * 1       ' "/"
    TIPONUMERAZIONE                     As String * 1
    SEPARATOR_02                        As String * 1       ' "/"
    CODSERVIZIONUMERAZIONE              As String * 2
    SEPARATOR_03                        As String * 1       ' "/"
    ANNOBOLLETTA                        As String * 4
    SEPARATOR_04                        As String * 1       ' "/"
    NUMEROBOLLETTA                      As String * 8
    SEPARATORE_05                       As String * 1       ' "/"
    RATABOLLETTA                        As String * 1
    SEPARATOR_06                        As String * 1       ' "/"
    PERIODO                             As String * 20      ' NB: NEL CASO DI FATTURAZIONE MULTI SITO SI METTE IL PERIODO RELATIVO AL PRIMO SERVIZIO IN BOLLETTA
    DATAEMISSIONE                       As String * 10      ' NELLA FORMA GG/MM/AAAA
    TIPOBOLLETTAZIONE                   As String * 1       ' P = SOLE PARTITE VARIE; A = ACCONTO NORMALE; L = LETTURA; D = CONGUAGLIO + ACCONTO; S = MULTISITO. NB: DAL VALORE DI QUESTO CAMPO, LO STAMPATORE È IN GRADO DI CAPIRE SE IL TRACCIATO CHE VA A STAMPARE È DI UNA BOLLETTA MONOSITO O MULTISITO.
    CCP                                 As String * 1       ' IL FLAG DI PRESENZA C.C.P VALUTA L'OPZIONE H2_CCP1 SOLO SE IL TIPO PAGAMENTO È "RISCONTRO AUTOMATICO DA SPORTELLO" (BOT_TIPOPAG = 1); IN QUESTO CASO SE L'OPZIONE È ATTIVA STAMPA S, ALTRIMENTI N. IL CASO PER CUI IL FLAG CCP È POSTO A "S" ANCHE PRIMA DELL'ATTIVAZIONE DELL'OPZIONE, È QUELLO IN CUI LA BOLLETTA È "NON DOMICIALIATA".
    DATALETTURAPRECEDENTE               As String * 10      ' NELLA FORMA GG/MM/AAAA. NB: CAMPO VALORIZZATO SOLO PER BOLLETTE MONOSITO (RECORD 001 CAMPO "TIPO BOLLETTAZIONE" DIVERSO DA "S")
    DATALETTURAATTUALE                  As String * 10      ' NELLA FORMA GG/MM/AAAA. NB: CAMPO VALORIZZATO SOLO PER BOLLETTE MONOSITO (RECORD 001 CAMPO "TIPO BOLLETTAZIONE" DIVERSO DA "S")
    DATALETTURAACCONTO                  As String * 10      ' IN BOLLETTA L+M. NELLA FORMA GG/MM/AAAA. NB: CAMPO VALORIZZATO SOLO PER BOLLETTE MONOSITO (RECORD 001 CAMPO "TIPO PRECEDENTE BOLLETTAZIONE" DIVERSO DA 'S')
    DATALETTURAATTUALEACCONTO           As String * 10      ' IN BOLLETTA L+M. NELLA FORMA GG/MM/AAAA. NB: CAMPO VALORIZZATO SOLO PER BOLLETTE MONOSITO (RECORD 001 CAMPO "TIPO BOLLETTAZIONE" DIVERSO DA 'S')
    CODICELINGUAANAGRAFICACLIENTE       As String * 3
End Type

' G01 R002 - ATTIVABILE LA SEPARAZIONE DEI FLUSSI STABO IN BASE AL CANALE DI INOLTRO CONTATTANDO IL PERSONALE ENG IN QUANTO FUNZIONALITÀ DISPONIBILE SOLO SU SPECIFICA RICHIESTA DEL CLIENTE.
'
Private Type strct_G01_R002
    CANALEINOLTRO                       As String * 2       ' PUÒ ASSUMERE I SEGUENTI VALORI: 1. '01' STAMPA, 2. '02' E-MAIL, 3. '03' STAMPA + E-MAIL, 4. '04' FAX, 5. '05' STAMPA + FAX, 6. '06' E-MAIL + FAX, 7. '07' STAMPA + E-MAIL + FAX
    FILLER_00                           As String * 6       ' CONTIENE DEGLI SPAZI
    INDIRIZZOEMAIL                      As String * 60      ' EVENTUALE INDIRIZZO E-MAIL PER L'INOLTRO DELLA FATTURA (NOA_NOTE OPPURE ANA_EMAIL)
    NUMEROFAX                           As String * 30      ' EVENTUALE NUMERO DI FAX PER L'INOLTRO DELLA FATTURA (NOA_NOTE1)
End Type

Private Type strct_G01_R004
    DESCRIZIONE                         As String * 50      ' VALORIZZATO CON: 'IL PRESENTE DOCUMENTO ANNULLA LA BOLLETTA'.
    ANNO                                As String * 4       ' ANNO BOLLETTA ANNULLATA
    SEPARATORE                          As String * 1       ' '/'
    NUMERO                              As String * 8       ' NUMERO BOLLETTA ANNULLATA
End Type

' G01 R005 - ATTIVABILE SOTTO OPZIONE CODIFICATA. SISTEMA ALIMENTANTE_NUMERO DOCUMENTO. ATTIVABILE TRAMITE OPZIONE CODIFICATA H2_FEALPDF. RIGA NON PRESENTE SE OPZIONE CODIFICATA NON ABILITATA
'
Private Type strct_G01_R005
    NOME_FE                             As String * 30
End Type

' G01 R006 - ATTIVABILE TRAMITE OPZIONE CODIFICATA H2_FEALPDF. RIGA NON PRESENTE SE OPZIONE CODIFICATA NON ABILITATA.
'
Private Type strct_G01_R006
    AZIENDASIU                          As String * 2
    SEZIONALE                           As String * 2
    TIPONUMERAZIONE                     As String * 1
    CODICESERVIZIONUMERAZIONE           As String * 2
    RATABOLLETTA                        As String * 1
    PROGRESSIVOFATTURAZIONE             As String * 6
End Type

' G01 R007 - INFORMAZIONI PREVISTE PER LA FATTURAZIONE ELETTRONICA
'
Private Type strct_G01_R007
    TIPOLOGIA_FEL                       As String * 5       ' TIPOLOGIA FEL
    MODALITÀ_INVIO                      As String * 5       ' MODALITÀ DI INVIO FEL
    CODICE_DESTINATARIO                 As String * 7       ' CODICE DESTINATARIO FEL
    CODICE_CUP                          As String * 15      ' CODICE CUP FEL
    CODICE_CIG                          As String * 15      ' CODICE CIG FEL
    RINUNCIA_COPIA_ANALOGICA            As String * 1       ' FLAG RINUNCIA ALLA COPIA ANALOGICA
    INVIO_PEC                           As String * 1       ' INVIO TRAMITE PEC
    INDIRIZZO_PEC                       As String * 255     ' INDIRIZZO PEC
    SOTTOTIPOLOGIA_FEL                  As String * 6       ' SOTTOTIPOLOGIA FEL
    ESITO_STATO_FEL                     As String * 10      ' ESITO STATO FEL
End Type

Private Type strct_G01
    R001                                As strct_G01_R001
    R002                                As strct_G01_R002
    R004                                As strct_G01_R004
    R005                                As strct_G01_R005
    R006                                As strct_G01_R006
    R007                                As strct_G01_R007
End Type

' G02 - INTESTATARIO
'
Private Type strct_G02_R001
    CODICEANAGRAFICO                    As String * 10      ' CODICE 001 CLIENTE/UTENTE.
    CODICESERVIZIO                      As String * 10      ' CODICE SERVIZIO/PUNTOPRESA. NB: CAMPO VALORIZZATO SOLO PER BOLLETTE MONOSITO (RECORD 001 CAMPO "TIPO BOLLETTAZIONE" DIVERSO DA "S")
End Type

Private Type strct_G02_R002
    RAGIONESOCIALEINTESTATARIO          As String * 40      ' INTESTATARIO
End Type

Private Type strct_G02_R003
    INDIRIZZORESIDENZA                  As String * 40      ' VIA, CIVICO/SUFFISSO
End Type

Private Type strct_G02_R004
    LOCALITÀRESIDENZA                   As String * 40      ' PRECEDUTA DAL CAP ED EVENTUALMENTE SEGUITA DALLA PROVINCIA.
End Type

Private Type strct_G02_R005
    CODICEFISCALE                       As String * 16
    FILLER                              As String * 1
    PARTITAIVA                          As String * 11
End Type

Private Type strct_G02
    R001                                As strct_G02_R001
    R002                                As strct_G02_R002
    R003                                As strct_G02_R003
    R004                                As strct_G02_R004
    R005                                As strct_G02_R005
End Type

' G03 - RECAPITO
'
Private Type strct_G03_R001
    NOME_RAGIONESOCIALE                 As String * 40      ' RECAPITO
End Type

Private Type strct_G03_R002
    DESCR_INTERNO                       As String * 8       ' DESCR. FISSA "INTERNO" (SOSTITUITA DA SPAZI SE INTERNO NON VALORIZZATO)
    INTERNO                             As String * 3       ' INTERNO INDIRIZZO DI RECAPITO
    FILLER_00                           As String * 1       ' SPAZIO
    DESCR_SCALA                         As String * 6       ' DESCR. FISSA "SCALA" (SOSTITUITA DA SPAZI SE SCALA NON VALORIZZATA)
    SCALA                               As String * 3       ' INDIRIZZO DI RECAPITO
    FILLER_01                           As String * 1       ' SPAZIO
    DESCR_PIANO                         As String * 6       ' DESCR. FISSA "PIANO" (SOSTITUITA DA SPAZI SE PIANO NON VALORIZZATO)
    PIANO                               As String * 3       ' PIANO INDIRIZZO DI RECAPITO
    FILLER_02                           As String * 9       ' SPAZI
End Type
    
Private Type strct_G03_R003
    INDIRIZZO                           As String * 40      ' VIA, CIVICO/SUFFISSO
    DESCRIZIONERECAPITO                 As String * 50
    DESCRIZIONEESTESARECAPITO           As String * 150
End Type

Private Type strct_G03_R004
    LOCALITÀ                            As String * 40      ' PRECEDUTA DAL CAP ED EVENTUALMENTE SEGUITA DALLA PROVINCIA.
    SIGLA_NAZIONE                       As String * 2
End Type

Private Type strct_G03
    R001()                              As strct_G03_R001
    R002                                As strct_G03_R002
    R003                                As strct_G03_R003
    R004                                As strct_G03_R004
End Type

' G05
'
Private Type strct_G05_R001
    INDIRIZZOFORNITURA                  As String * 40      ' VIA, CIVICO/SUFFISSO
End Type

Private Type strct_G05_R002
    LOCALITÀFORNITURA                   As String * 30      ' SENZA CAP E PROVINCIA.
End Type

Private Type strct_G05
    R001                                As strct_G05_R001
    R002                                As strct_G05_R002
End Type

' G06
'
Private Type strct_G06_R001
    TOTALELIRE                          As String * 15      ' TOTALE A PAGARE NELLA FORMA 99.999.999.999-
    TOTALEEURO                          As String * 14      ' TOTALE A PAGARE NELLA FORMA 99.999.999,99-
    DATASCADENZA                        As String * 10      ' DATA SCADENZA NELLA FORMA 99/99/9999.
    CODICEMESSAGGIO                     As String * 4       ' VEDERE "USO DEI MESSAGGI IN BOLLETTA"
    TOTALEADDEBITOACCREDITO             As String * 14      ' NELLA FORMA 99.999.999,99-
    IMPORTORESIDUO                      As String * 14      ' NELLA FORMA 99.999.999,99-
    TOTALEADDEBITOACCREDITOSUCCESSIVA   As String * 14      ' NELLA FORMA 99.999.999,99-. ATTIVABILE TRAMITE OPPORTUNA OPZIONE CODIFICATA.
End Type

Private Type strct_G06_R002
    DESCRIZIONEBANCA                    As String * 40      ' SOLO SE PRESENTE DOMICILIAZIONE BANCARIA.
    DESCRIZIONEBANCASECONDALINGUA       As String * 40      ' SOLO SE PRESENTE DOMICILIAZIONE BANCARIA.
End Type

Private Type strct_G06_R003
    IBAN                                As String * 100     ' COORDINATE BANCARIE DELL"AZIENDA IBAN O ABI + CAB + CCBAN
End Type

Private Type strct_G06
    R001                                As strct_G06_R001
    R002                                As strct_G06_R002
    R003                                As strct_G06_R003
End Type

' G07
'
Private Type strct_G07_R003
    CODICE_ESECUTORE                    As String * 3
    DESCRIZIONE_ESECUTORE               As String * 40      ' SE ESISTE, VIENE ESTRATTA LA DESCRIZIONE ESTERNA
    DATA_LETTURA                        As String * 10      ' NELLA FORMA 99/99/9999
    LETTURA                             As String * 15      ' ULTIMA LETTURA NON VALIDATA. NELLA FORMA 99999999,999999
    CONSUMO                             As String * 16      ' NELLA FORMA 99999999,999999-
    CAUSALE_NON_VALIDAZIONE             As String * 2       ' LSC_CAUS_LETT1
    DESCRIZIONE_NON_VALIDAZIONE         As String * 40      ' SE ESISTE, VIENE ESTRATTA LA DESCRIZIONE ESTERNA
    FLAG_AUTOLETTURA                    As String * 1       ' VALE 'S' O 'N'
    MOTIVAZIONE_SCARTO                  As String * 256     ' SE VALORIZZATA
    FLAG_NON_LETTURA                    As String * 1
    ORARIO                              As String * 8
End Type

Private Type strct_G07
    R003                                As strct_G07_R003
End Type

' G09
'
Private Type strct_G09_R001
    MATRICOLACONTATORE                  As String * 25      ' SE CONTATORE PRESENTE (SER_TIPOPRESA != 0) E NON MULTI CONTATORE (ESPOSTA MATRICOLA DEL PRIMO CO). ALTRIMENTI SPAZI. POPOLATO CON ACO_MATRICON OPPURE, SE NON VALORIZZATO, ACO_MATCON.
    TIPOLOGIAMISURATORE                 As String * 40      ' SE CONTATORE PRESENTE (SER_TIPOPRESA != 0) E NON MULTI CONTATORE (ESPOSTA MATRICOLA DEL PRIMO CO). ALTRIMENTI SPAZI. DESCRIZIONE (POPOCON) ASSOCIATA AI CAMPI ACO_POPOMIN E ACO_POPOMAX (ANACON) DEL CONTATORE N(2) NUMERO DI CONTATORI SUL SERVIZIO SER_NUMCONTINS (NUMERO CONTATORI INSTALLATI)
    NUMEROCONTATORISERVIZIO             As String * 2       ' SER_NUMCONTINS (NUMERO CONTATORI INSTALLATI)
    SIGLA                               As String * 3       ' SE CONTATORE PRESENTE (SER_TIPOPRESA != 0). ALTRIMENTI SPAZI
    PORTATA_POTENZACONTRATTUALE         As String * 11      ' NELLA FORMA -999999,999
    NUMEROIDRANTITIPO1                  As String * 4       ' NELLA FORMA 9999
    NUMEROIDRANTITIPO2                  As String * 4       ' NELLA FORMA 9999
    NUMEROIDRANTITIPO3                  As String * 4       ' NELLA FORMA 9999
    MINIMOANNUOCONTRATTUALE             As String * 8       ' NELLA FORMA -9999999. SE NON PRESENTI LE CONCESSIONI, VENGONO MOSTRATI 8 SPAZI.
    MINIMOFATTURATO                     As String * 8       ' NELLA FORMA -9999999
    TIPO                                As String * 1       ' SE "D" = DEPOSITO CAUZIONALE, SE "A" = ANTICIPO FORNITURA.
    IMPORTO_DEPOSITO_ANTICIPO           As String * 13      ' NELLA FORMA -999999999,99
    FLAGACCESSIBILITÀCONTATORE          As String * 1       ' ATTIVABILE CONTATTANDO IL PERSONALE ENG IN QUANTO FUNZIONALITÀ DISPONIBILE SOLO SU SPECIFICA RICHIESTA DEL CLIENTE. CAMPO UBI_ACCESSIBILE DELLA TABELLA TBUBIC. OPZ. COD. = H2_ACCCONT. SE DISATTIVA VERRÀ MOSTRATO UNO SPAZIO.
    CODICECLASSECONTATORE               As String * 4       ' CAMPO DI ANACON.ACO_CLATENTICO
    DESCRIZIONECLASSECONTATORE          As String * 40      ' CAMPO TECLCON.TCC_DES_40
    DESCRIZIONEMODELLOCONTATORE         As String * 40      ' CONTAT.MCO_DES_40
    CODICEMESSAGGIOACCESSOCONTATORE     As String * 2       ' ATTIVABILE TRAMITE OPPORTUNA OPZIONE CODIFICATA. CAMPO SE OPZIONE CODIFICATA NON ABILITATA.ù
    INDICAZIONEACCESSIBILITÀ_218_16     As String * 1       ' CAMPO ESPOSTO SE ATTIVA A SISTEMA LA DEL.218/16
    NUMERO_LETTURE_ANNUE_218_16         As String * 2       ' CAMPO ESPOSTO SE ATTIVA A SISTEMA LA DEL.218/16
    IMPORTO_PAGATO_DEPOSITO_ANTICIPO    As String * 13      ' NELLA FORMA -999999999,99. ATTIVABILE TRAMITE OPPORTUNA OPZIONE CODIFICATA. CAMPO NON VALORIZZATO SE OPZIONE CODIFICATA NON ABILITATA.
End Type

Private Type strct_G09_R002
    CODICEPUNTORICONSEGNA               As String * 23      ' CONCATENAZIONE SER_COD_IMPIAN E SER_PUNTO_EROGA
    CODICEREMIVIRTUALE                  As String * 25
    CODICERETEUTENZA                    As String * 18
    CODICEPLAYER                        As String * 35      ' CAMPO PREVISTO MA ATTUALMENTE NON VALORIZZATO.
    RAGIONESOCIALEPLAYER                As String * 40      ' CAMPO PREVISTO MA ATTUALMENTE NON VALORIZZATO.
    NUMEROPRONTOINTERVENTO              As String * 15      ' CAMPO PREVISTO MA ATTUALMENTE NON VALORIZZATO.
    NUMEROTELEFONICO                    As String * 15      ' CAMPO PREVISTO MA ATTUALMENTE NON VALORIZZATO.
    NUMEROFAX                           As String * 15      ' CAMPO PREVISTO MA ATTUALMENTE NON VALORIZZATO.
    DESCRIZIONEREMIVIRTUALE             As String * 40
End Type

Private Type strct_G09_R003
    CODICECATEGORIAUTENZA               As String * 4
    CODICEUSOMERCEOLOGICO               As String * 4
    CODICETARIFFA                       As String * 4
    DESCRIZIONETARIFFA                  As String * 40      ' DESCRIZIONE DEL CODICE TARIFFA CON DECORRENZA E VERSIONE NULL
    CODICETARIFFA2                      As String * 4       ' SER_COD_TARIF2
    DESCRIZIONETARIFFA2                 As String * 40      ' DESCRIZIONE DEL CODICE TARIFFA2 CON DECORRENZA E VERSIONE NULL
    PERCENTUALEPROMISCUO                As String * 7       ' SER_PERC_TARIFF2
    CODICEUSOTARIFFARIO                 As String * 4
    CODICEZONA                          As String * 3
    CODICESTATISTICO                    As String * 6
    CODICELIMITATORE                    As String * 3
    TIPOPRESA                           As String * 1       ' SER_TIPOPRESA
    TIPOLOGIAPRESA                      As String * 1       ' PUÒ VALERE: SE TIPOPRESA = "0": A (BOCCA ANTINCENDIO) SE SER_IDRANTIPX VALORIZZATO, OPPURE SER_POPOCON1 > 0 E DIVERSO DA 999999; B (BOCCA TASSATA) ALTRIMENTI C
    TIPOFORNITURA                       As String * 4
    DESCRIZIONETIPOFORNITURA            As String * 40
    DATAINIZIOFORNITURA                 As String * 10      ' NELLA FORMA 99/99/9999.
    DESCRIZIONETARIFFAPERIODO           As String * 40      ' SARÀ RIPORTATA LA DESCRIZIONE DEL PRIMO CODICE TARIFFA APPLICATO NEL PERIODO. LA DECORRENZA E VERSIONE SARANNO QUELLE NULL. VALE SPAZI IN CASO DI FATTURAZIONE DI SOLE PARTITE O SE NON SONO FATTURATE TARIFFE
    DESCRIZIONEUSOTARIFFARIO            As String * 40      ' CAMPO CODUSO.CDU_DES_40. VALE SPAZI IN CASO DI FATTURAZIONE DI SOLE PARTITE
    NUMEROTOTALECONCESSIONI             As String * 4       ' POTREBBE NON ESSERE PRESENTE
End Type

Private Type strct_G09_R005
    FLAGRATEIZZABILITÀ                  As String * 1       ' VALE: 'S' SE BOLLETTA RATEIZZABILE IN BASE DATI DEFINITIVA. 'N' (DEFAULT) ALTRIMENTI, IN BASE DATI DEFINITIVA. ' ' SE STAMPA IN BASE DATI PROVVISORIA.
    FILLER                              As String * 3       '
    FLAGINDENNIZZI                      As String * 1       '
    CODICEMESSAGGIO                     As String * 4       ' SOLO SE ESISTE IL BOLLO
    FLAG_AZZERAMENTO_BOLLO              As String * 1       ' VALE S/N
    AZZERAMENTO_BOLLO                   As String * 18      ' NEL FORMATO -99.999.999.999,99
    RATEIZZATO_NORMATIVA                As String * 2       ' VALE S/N
    PERCENTUALE_INTERESSI_DILATORI      As String * 6       ' FORMATO 999,99
    TERMINE_ULTIMO_RICHIESTA            As String * 10      ' FORMATO: DD/MM/YYYY
End Type

Private Type strct_G09_R008
    BASECOMPUTO                         As String * 17      ' FORMATO: 999999999,999999-
    FLAGTIPOCALCOLOCONSUMO              As String * 2       ' VALE: "AP" PER CONSUMO ANNO PRECEDENTE; "12" PER CONSUMO 12 MESI PRECEDENTI; "IF" PER CONSUMO DALLA DATA INIZIO FORNITURA;
    CONSUMOANNUO                        As String * 17      ' NEL FORMATO 999999999,999999-. IN CASO DI FLAG = "AP" SARÀ ESPOSTO IL VALORE DEL CAMPO SER_CONSTO_PREC (IN CASO DI RISTAMPA IL VALORE POTREBBE NON ESSERE CORRETTO)
    DATAINIZIOPERIODOCONSIDERATO        As String * 10      ' NELLA FORMA 99/99/9999.
    DATAFINEPERIODOCONSIDERATO          As String * 10      ' NELLA FORMA 99/99/9999.
    GIORNIPERIODO                       As String * 3       ' NELLA FORMA 999
End Type

Private Type strct_G09_R012                                 ' GRUPPO ESPOSTO SE ATTIVA A SISTEMA LA DEL.218/16
    CONSUMO_MEDIO_ANNUO_AC              As String * 17      ' NEL FORMATO 999999999,999999-. VALORE DETERMINATO IN BASE AL VALORE DELL"OPZIONE CODIFICATA DEL_218_DATARIF
    CONSUMO_MEDIO_ANNUO_AS              As String * 17      ' NEL FORMATO 999999999,999999-. VALORE DETERMINATO IN BASE AL VALORE DELL"OPZIONE CODIFICATA DEL_218_DATARIF
End Type

Private Type strct_G09
    R001                                As strct_G09_R001
    R002                                As strct_G09_R002
    R003                                As strct_G09_R003
    R005                                As strct_G09_R005
    R008                                As strct_G09_R008
    R012                                As strct_G09_R012
End Type

' G11
' R001 (SEMPRE PRESENTE)
' R002 (PRESENTE PER LA PARTE IN ACCONTO DI UNA FATTURAZIONE L + M)
'
Private Type strct_G11_R001R002
    DATALETTURAPRECEDENTE               As String * 10      ' NELLA FORMA 99/99/9999. NON PRESENTE PER SER_TIPOPRESA = 0.
    DATALETTURAATTUALE                  As String * 10      ' NELLA FORMA 99/99/9999. NON PRESENTE PER SER_TIPOPRESA = 0.
    GIORNIFATTURATI_LETT_PREC_ATT       As String * 4       ' NELLA FORMA 9999. NON PRESENTE PER SER_TIPOPRESA = 0.
    FLAGCONTATOREFORZATOPERIODO         As String * 1       ' SOLO SE BOLLETTA NON IN ACCONTO. VALORIZZATO IN CASO DI CONSUMO FORZATO TRA LE LETTURE ATTUALI E LE PRECEDENTI NEL PERIODO. NON PRESENTE PER SER_TIPOPRESA = 0.
    CONSUMO_ATT_PREC                    As String * 16      ' SE LA BOLLETTA È ANNULLATA IL SEGNO È INVERTITO. NELLA FORMA 99999999,999999-. NON PRESENTE PER SER_TIPOPRESA = 0.
    CONSUMOCONTATORISOSTITUITI          As String * 16      ' SE LA BOLLETTA È ANNULLATA IL SEGNO È INVERTITO. NELLA FORMA 99999999,999999-. NON PRESENTE PER SER_TIPOPRESA = 0.
    CONSUMO_SOMMARE_DETRARRE            As String * 16      ' NELLA FORMA 99999999,999999-. NON PRESENTE PER SER_TIPOPRESA = 0.
    ADEGUAMENTO_MINIMO                  As String * 16      ' NELLA FORMA 99999999,999999-. NON PRESENTE PER SER_TIPOPRESA = 0.
    CONSUMOFORFAITANNUO                 As String * 16      ' NELLA FORMA 99999999,999999-
    DATAINIZIO_CONSUMISTIMATI_BOLPREC   As String * 10      ' NELLA FORMA 99/99/9999. VALORIZZATO CON BAT_DA_LEPREC SOLO SE BOLLETTA PRECEDENTE È UN ACCONTO. NON PRESENTE PER SER_TIPOPRESA = 0.
    DATAFINE_CONSUMISTIMATI_BOLPREC     As String * 10      ' NELLA FORMA 99/99/9999. VALORIZZATO SOLO SE BOLLETTA PRECEDENTE È UN ACCONTO. NON PRESENTE PER SER_TIPOPRESA = 0.
    GIORNIFATT_INIFINE_CONSSTIM_BOLPREC As String * 4       ' NELLA FORMA 9999. NON PRESENTE PER SER_TIPOPRESA = 0.
    MC_ACCONTIFATTURATI                 As String * 16      ' SE LA BOLLETTA È ANNULLATA, COMPAIONO IN NEGATIVO. NELLA FORMA 99999999,999999-. NON PRESENTE PER SER_TIPOPRESA = 0.
    TOTALECONSUMOFATTURATO              As String * 16      ' NELLA FORMA 99999999,999999-
    TOTALECONSUMORILEVATO               As String * 16      ' NELLA FORMA 99999999,999999-. NON PRESENTE PER SER_TIPOPRESA = 0.
End Type

Private Type strct_G11_R003
    TOTALECONSUMOFATTURATO              As String * 17      ' NELLA FORMA 999999999,999999-
End Type

Private Type strct_G11
    R001                                As strct_G11_R001R002
    R002                                As strct_G11_R001R002
    R003                                As strct_G11_R003
End Type

' G12
'
Private Type strct_G12_RXXX
    TIPOPER                             As String * 1       ' IL TIPOPER E CODICE TARIFFA/IMPOSTA SONO RIFERITI AL MAX BOLCON
    CODICE_TARIFFA_IMPOSTA              As String * 4       ' IL TIPOPER E CODICE TARIFFA/IMPOSTA SONO RIFERITI AL MAX BOLCON
    DESCRIZIONE_TARIFFA_IMPOSTA         As String * 40      ' LA DESCRIZIONE TARIFFA/IMPOSTA È RIFERITA AL MAX BOLCON
    INFO_SERVIZIO_DEPURAZIONE           As String * 1       ' E' INDICATO AL SINGOLO UTENTE FINALE SE È FORNITO O MENO DA UN IMPIANTO DI DEPURAZIONE: 'A' - ATTIVO, 'B' - NON ATTIVO MA IN CORSO, 'C' - TEMPORANEAMENTE INATTIVO (ATTIVATO SU FATTURAZIONE TRAMITE OPZIONE "ESEIMPOPUNTUALI"), 'D' - NON PRESENTE.
End Type

' G13
'
' IN QUESTO RIQUADRO VENGONO RIPORTATE, IN ORDINE DI LET_NPROG_LETT DECRESCENTE, TUTTE LE LETTURE FATTURATE NELLA BOLLETTA CORRENTE A PARTIRE DALLA LETTURA PRECEDENTE.
' LA LETTURA PRECEDENTE (CHE VA OVVIAMENTE RIPORTATA NELLA RIGA 001) CORRISPONDE:
' - IN CASO DI CONGUAGLIO, ALL'ULTIMA LETTURA EFFETTIVA FATTURATA PRIMA DELLA BOLLETTA CORRENTE
' - IN CASO DI ACCONTO, ALL'ULTIMA LETTURA FATTURATA (NON IMPORTA SE IN ACCONTO O EFFETTIVA) PRIMA DELLA BOLLETTA CORRENTE
' SULLA RIGA DELLA LETTURA PRECEDENTE (RIGA 001) IL CONSUMO MISURATO E IL CONSUMO FATTURATO DEVONO ESSERE VUOTI.
'
Private Type strct_G13_RXXX
    NUMEROSEQUENZACONTATORE             As String * 2       ' RIPORTA IL NUMERO DI SEQUENZA DEL CONTATORE
    TIPOLETTURA                         As String * 1       ' RIPORTARE IL VALORE DI LET_TIP_LETTUR)
    AUTOLETTURA                         As String * 1       ' VALE "S" SE È UN'AUTOLETTURA, ALTRIMENTI "N". RIPORTARE IL VALORE DI TBESEC.TBE_FLG_AUTOLETTURE
    TIPOCONSUMO                         As String * 1       ' RIPORTARE IL VALORE DI LET_TIP_CONSUM
    DATALETTURA                         As String * 10      ' NELLA FORMA 99/99/9999
    LETTURA                             As String * 15      ' NELLA FORMA 99999999,999999
    CONSUMOMISURATO                     As String * 16      ' NELLA FORMA 99999999,999999-. RIPORTARE IL VALORE DI LET_CONSUMO
    CONSUMOFATTURATO                    As String * 16      ' NELLA FORMA 99999999,999999-. E' LA SOMMA DEI CONSUMI IN BOLCON RELATIVI ALLO STESSO PERIODO DEL CONSUMO MISURATO.
    GIORNI                              As String * 4       ' GIORNI RILEVATI FRA LA DATA LETTURA ANTECEDENTE E DATA LETTURA CONSIDERATA (DATA LETTURA CONSIDERATA - DATA LETTURA PRECEDENTE). NON PRESENTI PER LA LETTURA PRECEDENTE
    TIPOLETTURAAEEG                     As String * 20      ' DESCRIZIONE TIPOLOGIA LETTURA DEFINITA DALL'AEEG. I VALORI POSSIBILI SONO VALORIZZATI SECONDO LA SEGUENTE LOGICA(NELL'ORDINE): - "AUTOLETTURA", SE IL CODICE ESECUTORE È PARAMETRIZZATO COME ESECUTORE DI AUTOLETTURA; - "STIMATA", SE IL TIPO LETTURA SIU È LETTURA STIMATA; - "RILEVATA", SE NON SI RIENTRA NEI 2 CASI PRECEDENTI.
    MATRICOLACONTATORE                  As String * 25      ' ACO_MATRICON OPPURE, SE NON VALORIZZATO, ACO_MATCON
    TIPOLOGIAMISURATORE                 As String * 40      ' DESCRIZIONE (POPOCON) ASSOCIATA AI CAMPI ACO_POPOMIN E ACO_POPOMAX (ANACON) DEL CONTATORE.
    CODICEESECUTORE                     As String * 3
    DESCRIZIONEESECUTORE                As String * 40
End Type

' G14
'
' QUESTO RIQUADRO È PRESENTE SOLO IN CASO DI BOLLETTE CONGUAGLIO + ACCONTO E CONTIENE LE LETTURE RELATIVE ALLA PARTE IN ACCONTO (COMPRESA LA
' LETTURA PRECEDENTE CHE DEVE VENIRE RIPORTATA NELLA RIGA 001 E CHE È IN PRATICA LA RIPETIZIONE DELL'ULTIMA LETTURA DELLA PARTE A CONGUAGLIO)
'
Private Type strct_G14_RXXX
    NUMEROSEQUENZACONTATORE             As String * 2       ' RIPORTA IL NUMERO DI SEQUENZA DEL CONTATORE
    TIPOLETTURA                         As String * 1       ' RIPORTARE IL VALORE DI LET_TIP_LETTUR)
    AUTOLETTURA                         As String * 1       ' VALE "S" SE È UN'AUTOLETTURA, ALTRIMENTI "N". RIPORTARE IL VALORE DI TBESEC.TBE_FLG_AUTOLETTURE
    TIPOCONSUMO                         As String * 1       ' RIPORTARE IL VALORE DI LET_TIP_CONSUM
    DATALETTURA                         As String * 10      ' NELLA FORMA 99/99/9999
    LETTURA                             As String * 15      ' NELLA FORMA 99999999,999999
    CONSUMOMISURATO                     As String * 16      ' NELLA FORMA 99999999,999999-. RIPORTARE IL VALORE DI LET_CONSUMO
    CONSUMOFATTURATO                    As String * 16      ' NELLA FORMA 99999999,999999-. E' LA SOMMA DEI CONSUMI IN BOLCON RELATIVI ALLO STESSO PERIODO DEL CONSUMO MISURATO.
    GIORNI                              As String * 4       ' GIORNI RILEVATI FRA LA DATA LETTURA ANTECEDENTE E DATA LETTURA CONSIDERATA (DATA LETTURA CONSIDERATA - DATA LETTURA PRECEDENTE). NON PRESENTI PER LA LETTURA PRECEDENTE
    TIPOLETTURAAEEG                     As String * 20      ' DESCRIZIONE TIPOLOGIA LETTURA DEFINITA DALL'AEEG.
    MATRICOLACONTATORE                  As String * 25      ' ACO_MATRICON OPPURE, SE NON VALORIZZATO, ACO_MATCON
    TIPOLOGIAMISURATORE                 As String * 40      ' DESCRIZIONE (POPOCON) ASSOCIATA AI CAMPI ACO_POPOMIN E ACO_POPOMAX (ANACON) DEL CONTATORE.
    CODICEESECUTORE                     As String * 3
    DESCRIZIONEESECUTORE                As String * 40
End Type

' G16
'
Private Type strct_G16_RXXX
    DATA_DECORRENZA_CONCESSIONE         As String * 10      ' NELLA FORMA 99/99/9999.
    PROGRESSIVO_RIGA                    As String * 1       ' DA 1, ..., 8
    DESCRIZIONE_TIPOLOGIA_RIGA          As String * 100     ' SE OPZIONE 665/17 SPENTA NUOVO CAMPO SU DESBASE ALTRIMENTI NUOVO CAMPO SU ANAGRAFICA CONCESSIONI CATEGORIE NORMATE
    TIPOLOGIA_RIGA                      As String * 3       ' SE OPZIONE 665/17 SPENTA NUOVO CAMPO SU DESBASE ALTRIMENTI NUOVO CAMPO SU ANAGRAFICA CONCESSIONI CATEGORIE NORMATE
    TIPOLOGIA_USO_665_17                As String * 25      ' ESPOSTO SOLAMENTE SE ATTIVA OPZIONE 665/17
    RESIDENTE                           As String * 1       ' ESPOSTO SOLAMENTE SE ATTIVA OPZIONE 665/17
    NUCLEO_FAMILIARE                    As String * 6       ' ESPOSTO SOLAMENTE SE ATTIVA OPZIONE 665/17
    TIPOLOGIA_RIGA_CONCESSIONE          As String * 1       ' CON_TIPO1
    CODICE_TARIFFA_APPLICATA            As String * 4       ' NELLA FORMA 9999. CODICE TARIFFA APPLICATA
    DESCRIZIONE_TARIFFA_APPLICATA       As String * 40      ' DESCRIZIONE TARIFFA
    CODICE_QUOTA_FISSA_APPLICATA        As String * 4       ' NELLA FORMA 9999. CODICE QUOTA FISSA APPLICATA
    DESCRIZIONE_QUOTA_FISSA_APPLICATA   As String * 40      ' DESCRIZIONE QUOTA FISSA APPLICATA
    NUMERO_CONCESSIONI                  As String * 4       ' NELLA FORMA 9999 (SPAZI SE NON PRESENTE)
    NUMERO_TOTALE_CONCESSIONI           As String * 4       ' NELLA FORMA 9999 (SPAZI SE NON PRESENTE)
    MINIMO_GARANTITO_TOTALE_ANNUO       As String * 8       ' IL VALORE PIÙ RECENTE. NELLA FORMA -9999999 (SPAZI SE NON PRESENTE)
    QUANTITÀ_RIGA_CONCESSIONE           As String * 8       ' NELLA FORMA 99999999 (SPAZI SE NON PRESENTE)
    QUANTITÀ_2_RIGA_CONCESSIONE         As String * 8       ' NELLA FORMA 99999999 (SPAZI SE NON PRESENTE)
    MESE_INIZIO                         As String * 2       ' NELLA FORMA 99
    MESE_FINE                           As String * 2       ' NELLA FORMA 99
    ZZ                                  As String * 5
End Type

' G17
'
Private Type strct_G17_RXXX
    DATA_DECORRENZA                     As String * 10      ' NELLA FORMA 99/99/9999 (DATA INIZIO FORNITURA)
    TIPOLOGIA_UTENZA                    As String * 100     ' VALORE ESPOSTO IN BASE AD OPZIONE CODIFICATA
    DESCRIZIONE_TARIFFA_APPLICATA       As String * 100     ' VALORE ESPOSTO IN BASE AD OPZIONE CODIFICATA
    NUMERO_CONCESSIONI                  As String * 4       ' NELLA FORMA 9999
    PERCENTUALE                         As String * 3       ' VALORIZZATO SOLAMENTE PER PROMISCUI
    PROGRESSIVO                         As String * 1       ' RIGA DA 1,-,8
    DESCRIZIONE_TIPOLOGIA_RIGA          As String * 100     ' NUOVO CAMPO SU DESBASE
    TIPOLOGIA_RIGA_STABOL               As String * 3       ' NUOVO CAMPO SU DESBASE
    TIPOLOGIA_USO_665_17                As String * 25      ' ESPOSTO SOLAMENTE SE ATTIVA OPZIONE 665/17
    RESIDENTE                           As String * 1       ' ESPOSTO SOLAMENTE SE ATTIVA OPZIONE 665/17
    NUCLEO_FAMILIARE                    As String * 6       ' ESPOSTO SOLAMENTE SE ATTIVA OPZIONE 665/17
End Type

' G21 - PAGAMENTI PRECEDENTI
'
Private Type strct_G21_R001R004
    CODICEMESSAGGIOMOROSITÀ             As String * 4       ' VALE SOLO PER IL RECORD "21001". VEDERE "USO DEI MESSAGGI DA 001 A NNN IN BOLLETTA". SE SONO PIÙ DI 'NNN' COMPARE "ALTRE BOLLETTE" CON L'IMPORTO RESIDUO
    FILLER_00                           As String * 2
    FILLER_01                           As String * 1
    ANNOFATTURA                         As String * 4
    FILLER_02                           As String * 1
    NUMEROFATTURA                       As String * 8
    FILLER_03                           As String * 1
    RATA                                As String * 2
    FILLER_04                           As String * 7
    DATASCADENZA                        As String * 10
    IMPORTO                             As String * 18
End Type

Private Type strct_G21_R005
    CODICEMESSAGGIOMOROSITÀ             As String * 4
    DESCRIZIONE                         As String * 36
    IMPORTO                             As String * 18
End Type

Private Type strct_G21_R999
    IMPORTOTOTALE                       As String * 18      ' NELLA FORMA 99.999.999.999,99-
End Type

Private Type strct_G21
    RXXX()                              As strct_G21_R001R004
    R005                                As strct_G21_R005
    R999                                As strct_G21_R999
End Type

' G22
'
Private Type strct_G22_RXXX
    DATAINIZIALEPERIODO                 As String * 10      ' NELLA FORMA GG/MM/AAAA
    DATAFINALEPERIODO                   As String * 10      ' NELLA FORMA GG/MM/ AAAA
    TOTALEGIORNIPERIODO                 As String * 5       ' NELLA FORMA 99999
    CONSUMOPERIODO                      As String * 17      ' NELLA FORMA 999999999,999999-
    CONSUMOMEDIOGIORNALIEROPERIODO      As String * 17      ' NELLA FORMA 999999,999999-
End Type

' G23
'
Private Type strct_G23_RXXX
    HEADER                              As String * 27      '
    CL_FILLER_00                        As String * 1       ' "<"
    CL_BOLLETTINO_ID                    As String * 18      '
    CL_FILLER_01                        As String * 1       ' ">"
    CL_IMPORTO                          As String * 11      '
    CL_FILLER_02                        As String * 1       ' ">"
    CL_CC                               As String * 8       '
    CL_BOLLETTINO_TYPE                  As String * 5       ' <896>
    CL_FILLER_03                        As String * 8       '
    NUMERORATA                          As String * 2       ' VALE 00 - BOLLETTINO GLOBALE
    IMPORTO                             As String * 12      ' NELLA FORMA 999.999,99-
    SCADENZA                            As String * 10      ' NELLA FORMA GG/MM/AAAA
    FLAGBOLLETTINOPAGATO                As String * 1       ' VALE 'P' IN CASO DI BOLLETTINO PAGATO ALTRIMENTI ' '.
End Type

' G26
'
Private Type strct_G26_RXXX
    RIGAMESSAGGIO                       As String * 80      ' MESSAGGIO A LIVELLO DI BOLLETTA. VEDERE "USO DEI MESSAGGI IN BOLLETTA"
End Type

' G34
'
Private Type strct_G34_RXXX
    CODICE_NAV                          As String * 35
    NUMERO_RATA                         As String * 2       ' VALE 00 - PER DOCUMENTO NON RATEIZZATO
End Type

' GAA - DA 001 A NNN (RECORD RELATIVO AL DETTAGLIO DELL'AREA ASSOGGETTATI DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GAA_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' X = DETTAGLIO AREA ASSOGGETTATI;
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 24
    PARCAU                              As String * 4
    FILLER_01                           As String * 1
    DESCRIZIONE                         As String * 141
    FILLER_02                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_03                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GAR - DA 001 A NNN (RECORD RELATIVO ALL'ARROTONDAMENTO ATTUALE/PRECEDENTE BOLLETTA DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GAR_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' A = AREA; V = VOCE; R = RAGGRUPPAMENTO; " " = CHE NON CONCORRE AL TOTALE SEZIONE
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 168
    FILLER_01                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_02                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GAZ - DA 001 A NNN (RECORD RELATIVO AL DETTAGLIO DELL'AREA AZZERAMENTI DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GAZ_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' X = DETTAGLIO AREA ASSOGGETTATI;
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 24
    PARCAU                              As String * 4
    FILLER_01                           As String * 1
    DESCRIZIONE                         As String * 141
    FILLER_02                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_03                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GBO - DA 001 A NNN (Record relativo al bollo di quietanza dei servizi H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GBO_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' R = RAGGRUPPAMENTO
    IDENTIFICATIVO                      As String * 3       '
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 168     ' SE ATTIVA OPZIONE CODIFICATA RBXEDESBO LA DESCRIZIONE È NEL FORMATO: “CAUSALE FISSA BOLLO QUIETANZA” (40CHR) + "- " (1 CHR) + “DESCRIZIONE CAUSALE FISSA BOLLO QUIETANZA” (40 CHR)
    FILLER_01                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_02                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GBS - BONUS SOCIALE
'
Private Type strct_GBS_RXXX
    DATA_INIZIO_PERIODO               As String * 10
    DATA_FINE_PERIODO                 As String * 10
End Type

' GDF
'
Public Type strct_GDF_Data
    ROW                                 As String
    SORT_KEY                            As String
End Type

Public Type strct_GDF_RXXX
    RXXX()                              As strct_GDF_Data
End Type

' GDO - DA 001 A NNN (RECORD RELATIVO ALLE VARIE CONFLUITE AUTOMATICAMENTE NELL'AREA DI DEFAULT DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GDO_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' A = AREA; V = VOCE; R = RAGGRUPPAMENTO; " " = CHE NON CONCORRE AL TOTALE SEZIONE
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    PERIODO                             As String * 21
    FILLER_01                           As String * 1
    PARCAU                              As String * 4
    FILLER_02                           As String * 1
    DESCRIZIONE                         As String * 141
    FILLER_03                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_04                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GDM DA 001 A NNN (RECORD RELATIVO AGLI INTERESSI DI MORA DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GDM_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' R = RAGGRUPPAMENTO;
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    PARCAU                              As String * 4
    FILLER_01                           As String * 1
    DESCRIZIONE                         As String * 100
    FILLER_02                           As String * 1
    TIPOLOGIASOTTOTIPO                  As String * 1       ' M: INTERESSI MORATORI - D: INTERESSI DILATORI
    FILLER_03                           As String * 3
    DESCRIZIONESOTTOTIPO                As String * 50
    FILLER_04                           As String * 10
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_05                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GIV - DA 001 A NNN (RECORD RELATIVO AL RIEPILOGO IVA DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GIV_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' A = AREA; V = VOCE; R = RAGGRUPPAMENTO; " " = CHE NON CONCORRE AL TOTALE SEZIONE
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 168
    FILLER_01                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_02                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GNV - DA 001 A NNN (RECORD RELATIVO ALLE NO VARIE DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GNV_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' R = RAGGRUPPAMENTO; " " = CHE NON CONCORRE AL TOTALE SEZIONE
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    PERIODO                             As String * 21
    FILLER_01                           As String * 1
    DESCRIZIONE                         As String * 45
    FILLER_02                           As String * 4
    CONCESSIONI                         As String * 4
    FILLER_03                           As String * 1
    UNITÀMISURACONCESSIONI              As String * 20
    FILLER_04                           As String * 1
    TEMPO                               As String * 4
    FILLER_05                           As String * 1
    UNITÀMISURATEMPO                    As String * 4
    FILLER_06                           As String * 1
    QUANTITÀ                            As String * 19      ' NEL FORMATO SENZA PUNTO MIGLIAIA
    FILLER_07                           As String * 1
    UNITÀMISURAQUANTITÀ                 As String * 8
    FILLER_08                           As String * 1
    PREZZO                              As String * 14      ' NEL FORMATO SENZA PUNTO MIGLIAIA
    FILLER_09                           As String * 1
    UNITÀMISURAPREZZO                   As String * 17
    FILLER_10                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_11                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GOR - NSO
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GOR_RXXX
    ORDINE_ACQUISTO_NSO                 As String * 20
    DATA_ORDINE_ACQUISTO_NSO            As String * 10
    EMITTENTE_ORDINE_NSO                As String * 40
End Type

' GPB 001 (RECORD RELATIVO ALLA PRESCRITTIBILITÀ)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GPB_R001
    TIPOLOGIA_PRESCRIZIONE              As String * 1       ' SPAZIO SE VUOTO
    DATA_PRESCRIZIONE                   As String * 10      ' SPAZIO SE VUOTO
    FILLER_00                           As String * 1
    IMPONIBILE_PRESCRIZIONE             As String * 20      ' SPAZIO SE VUOTO
    FILLER_01                           As String * 1
    IVA_PRESCRIZIONE                    As String * 20      ' SPAZIO SE VUOTO
    FILLER_02                           As String * 1
    SOMMA_TOTALE_PRESCRIZIONE           As String * 20      ' SPAZIO SE VUOTO
    FILLER_03                           As String * 1
    POTENZIALE_PRESCRIZIONE             As String * 20      ' 0 SE VUOTO
    CLASSE_UTENZA_PRESCRIZIONE          As String * 255     ' SPAZIO SE VUOTO
End Type

Private Type strct_GPB_R002
    FLAG_NON_PRESCRITTIBILITÀ           As String * 1
    DATA_INIZIO_NON_PRESCR              As String * 10
    DATA_FINE_NON_PRESCR                As String * 10
    MOTIVAZIONE                         As String * 500
End Type

Public Type strct_GPB_RXXX
    R001                                As strct_GPB_R001
    R002                                As strct_GPB_R002
End Type

' GRE DA 001 A NNN (RECORD RELATIVO ALLA RESTITUZIONE DEGLI ACCONTI DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GRE_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' Fisso a "F"
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 170
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_01                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GRM DA 001 A NNN (RECORD RELATIVO AGLI INTERESSI DI MORA DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GRM_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' R = RAGGRUPPAMENTO;
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    ANNODOCUMENTOMOROSO                 As String * 4
    FILLER_01                           As String * 1
    NUMERODOCUMENTOMOROSO               As String * 8
End Type

' GSE DA 001 A NNN (CONSUMI MEDI GIORNALIERI PER TIPOLOGIA DI UTENZA FATTURATI)
' ATTIVABILE CONTATTANDO IL PERSONALE ENG IN QUANTO FUNZIONALITÀ DISPONIBILE SOLO SPECIFICA RICHIESTA DEL CLIENTE.
'
Private Type strct_GSE_RXXX
    DATA_INIZIO_PERIODO_CONSUMI         As String * 10      ' NELLA FORMA 99/99/9999
    DATA_FINE_PERIODO_CONSUMI           As String * 10      ' NELLA FORMA 99/99/9999
    ISTANZA                             As String * 1       ' L/M
    TIPOLOGIA_UTENZA                    As String * 25
    DESCRIZIONE                         As String * 100
    GIORNI                              As String * 5       ' NELLA FORMA 99999
    CONSUMO_FATTURATO                   As String * 10      ' NELLA FORMA 999999999-
    CONSUMO_MEDIO_LITRI                 As String * 10      ' NELLA FORMA 999999999-
    CONSUMO_MEDIO_LITRI_UI              As String * 10      ' NELLA FORMA 999999999-
End Type

' GSF (ESPOSIZIONE DEGLI IMPORTI FATTURATI NEGLI ULTIMI 12 MESI)
' ATTIVABILE CONTATTANDO IL PERSONALE ENG IN QUANTO FUNZIONALITÀ DISPONIBILE SOLO SPECIFICA RICHIESTA DEL CLIENTE.
'
Private Type strct_GSF
    DATA_EMISSIONE_AP                   As String * 10      ' NELLA FORMA 99/99/9999
    DATA_EMISSIONE_AC                   As String * 10      ' NELLA FORMA 99/99/9999
    TOTALE                              As String * 14      ' NELLA FORMA 99.999.999,99-
End Type

' GSI - SINTETICO
'
Private Type strct_GSI_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' A = AREA; V = VOCE; R = RAGGRUPPAMENTO; F = RESTITUZIONE ACCONTI; I = RIEPILOGO IVA; T = TOTALE SERVIZIO; B = TOTALE BOLLETTA; X = TOTALE GIÀ ASSOGGETTATI; " " = SEZIONE SCONOSCIUTA
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONESINTETICO                As String * 100
    UNITAMISURA                         As String * 1       ' "€"
    IMPORTO_IMPOSTA                     As String * 18      ' NEL FORMATO -99.999.999.999,99
    'FILLER_01                           As String * 1
    'ALIQUOTAIVA                         As String * 4       ' NON RAPPRESENTA IL RIEPILOGO IMPONIBILE E IMPORTO IVA
End Type

' GSP
'
Private Type strct_GSP_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' FISSO: "S"
    IDENTIFICATIVO                      As String * 3       ' FISSO: "000" (RECORD IDENTIFICATIVO DELLA FATTURAZIONE ESPOSTA NEL DETTAGLIO BOLLETTA)
    INDENTIFICATIVOFATTURAZIONE         As String * 30      ' DETTAGLIO BOLLETTA POSSONO COMPARIRE LE SEGUENTI STRINGHE: DETTAGLIO PERIODO IN ACCONTO; DETTAGLIO PERIODO A CONGUAGLIO; DETTAGLIO PERIODO A PARTITE
End Type

' GTB - DA 001 A NNN (RECORD RELATIVO AL TOTALE BOLLETTA DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
'Private Type strct_GTB_RXXX
'    TIPOLOGIASEZIONE                    As String * 1       ' A = AREA; V = VOCE; R = RAGGRUPPAMENTO; " " = CHE NON CONCORRE AL TOTALE SEZIONE; X = AREA ASSOGGETTATI
'    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
'    FILLER_00                           As String * 2
'    DESCRIZIONE                         As String * 168
'    FILLER_01                           As String * 2
'    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
'End Type

' GTI - DA 001 A NNN (RECORD RELATIVO AI TITOLI AREA / VOCE / RAGGRUPPAMENTI / RESTITUZIONI ACCONTI / CONGUAGLI DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GTI_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' A = AREA; V = VOCE; R = RAGGRUPPAMENTO; " " = CHE NON CONCORRE AL TOTALE SEZIONE; X = AREA ASSOGGETTATI
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 168
    FILLER_01                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_02                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GTM - DA 001 A NNN (RECORD RELATIVO AGLI INTERESSI DI MORA DEI SERVIZI H2O). ESPOSTE AL MAX N RIGHE DI MORA PER DOCUMENTO (VEDERE OPZIONE CODIFICATA H2_MAXINTE).
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GTM_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' R = RAGGRUPPAMENTO;
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    PERIODO                             As String * 21      ' PERIODO NELLA FORMA "DD/MM/YYYY-DD/MM/YYYY" SE VALORIZZATO, ALTRIMENTI SOSTITUITO DA SPAZI
    FILLER_01                           As String * 1
    TASSOINTERESSEAPPLICATO             As String * 5       ' (ANNUO) VALORIZZATO SOLO PER INTERESSI DI MORA (TIPO 'I').
    FILLER_02                           As String * 1
    GIORNIPERIODO                       As String * 4       ' VALORIZZATO SOLO PER PENALI FISSE DI MORA (TIPO 'F').
    FILLER_03                           As String * 1
    PENALE                              As String * 5       ' (ANNUA) VALORIZZATO SOLO PER PENALI DI MORA AD IMPORTO PROPORZIONALE AL TOTALE DEL DOCUMENTO MOROSO (TIPO 'P').
    FILLER_04                           As String * 1
    IMPONIBILE                          As String * 13
    FILLER_05                           As String * 1
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_06                           As String * 5
End Type

' GTO - DA 001 A NNN (RECORD RELATIVO AI TOTALI AREA / VOCE / RAGGRUPPAMENTI DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GTO_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' A = AREA; V = VOCE; R = RAGGRUPPAMENTO; " " = CHE NON CONCORRE AL TOTALE SEZIONE; X = AREA ASSOGGETTATI
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 168
    FILLER_01                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_02                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GTP - DA 001 A NNN (RECORD RELATIVO AL TOTALE DA PAGARE DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GTP_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' VALORE FISSO "P"
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 168
    FILLER_01                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
End Type

' GTS - DA 001 A NNN (RECORD RELATIVO AL TOTALE SERVIZIO DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GTS_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' T = TOTALE SERVIZIO
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    DESCRIZIONE                         As String * 168
    FILLER_01                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_02                           As String * 1
    ALIQUOTA                            As String * 4
End Type

' GVA DA 001 A NNN (RECORD RELATIVO ALLE VARIE DEI SERVIZI H2O)
' IL RECORD POTREBBE NON ESSERE PRESENTE
'
Private Type strct_GVA_RXXX
    TIPOLOGIASEZIONE                    As String * 1       ' R = RAGGRUPPAMENTO
    IDENTIFICATIVO                      As String * 3       ' IDENTIFICATIVO DELLA SEZIONE
    FILLER_00                           As String * 2
    PERIODO                             As String * 21      ' NELLA FORMA "GG/MM/AAAA-GG/MM/AAAA" OPPURE "GG.MM.AAAAGG. MM.AAAA" NEL CASO DI SECONDA LINGUA
    FILLER_01                           As String * 1
    PARCAU                              As String * 4
    FILLER_02                           As String * 1
    DESCRIZIONE                         As String * 40
    FILLER_03                           As String * 4
    CONCESSIONI                         As String * 4
    FILLER_04                           As String * 1
    UNITÀMISURACONCESSIONI              As String * 20
    FILLER_05                           As String * 1
    TEMPO                               As String * 4
    FILLER_06                           As String * 1
    UNITÀMISURATEMPO                    As String * 4
    FILLER_07                           As String * 1
    QUANTITÀ                            As String * 19      ' NEL FORMATO SENZA PUNTO MIGLIAIA
    FILLER_08                           As String * 1
    UNITÀMISURAQUANTITÀ                 As String * 8
    FILLER_09                           As String * 1
    PREZZO                              As String * 14      ' NEL FORMATO SENZA PUNTO MIGLIAIA
    FILLER_10                           As String * 1
    UNITÀMISURAPREZZO                   As String * 17
    FILLER_11                           As String * 2
    IMPORTO                             As String * 18      ' NEL FORMATO -99.999.999.999,99
    FILLER_12                           As String * 1
    ALIQUOTA                            As String * 4
    DESCRIZIONEAGGIUNTIVA               As String * 40
End Type

' ********************************************************************************************************************************************
' * STABO H2O DATA
' ********************************************************************************************************************************************
'
Public WS_G00                           As strct_G00
Public WS_G01                           As strct_G01
Public WS_G02                           As strct_G02
Public WS_G03                           As strct_G03
Public WS_G05                           As strct_G05
Public WS_G06                           As strct_G06
Public WS_G07                           As strct_G07
Public WS_G09                           As strct_G09
Public WS_G11                           As strct_G11
Public WS_G12()                         As strct_G12_RXXX
Public WS_G13()                         As strct_G13_RXXX
Public WS_G14()                         As strct_G14_RXXX
Public WS_G16()                         As strct_G16_RXXX
Public WS_G17()                         As strct_G17_RXXX
Public WS_G21                           As strct_G21
Public WS_G22()                         As strct_G22_RXXX
Public WS_G23()                         As strct_G23_RXXX
Public WS_G26()                         As strct_G26_RXXX
Public WS_G34()                         As strct_G34_RXXX
Public WS_GAA                           As strct_GAA_RXXX
Public WS_GAR                           As strct_GAR_RXXX
Public WS_GAZ                           As strct_GAZ_RXXX
Public WS_GBO                           As strct_GBO_RXXX
Public WS_GBS                           As strct_GBS_RXXX
Public WS_GDF()                         As strct_GDF_RXXX
Public WS_GDF_QF()                      As strct_GDF_RXXX
Public WS_GDF_NV018()                   As strct_GDF_RXXX
Public WS_GDF_NV019()                   As strct_GDF_RXXX
Public WS_GDF_NV020()                   As strct_GDF_RXXX
Public WS_GDF_NV021()                   As strct_GDF_RXXX
Public WS_GDF_NV022()                   As strct_GDF_RXXX
Public WS_GDF_NV023()                   As strct_GDF_RXXX
Public WS_GDF_NV024()                   As strct_GDF_RXXX
Public WS_GDF_NV025()                   As strct_GDF_RXXX
Public WS_GDM                           As strct_GDM_RXXX
Public WS_GDO                           As strct_GDO_RXXX
Public WS_GIV                           As strct_GIV_RXXX
Public WS_GNV                           As strct_GNV_RXXX
Public WS_GOR                           As strct_GOR_RXXX
Public WS_GPB                           As strct_GPB_RXXX
Public WS_GRE                           As strct_GRE_RXXX
Public WS_GRM                           As strct_GRM_RXXX
Public WS_GSE()                         As strct_GSE_RXXX
Public WS_GSF                           As strct_GSF
Public WS_GSI()                         As strct_GSI_RXXX
Public WS_GSP                           As strct_GSP_RXXX
'Public WS_GTB                           As strct_GTB_RXXX
Public WS_GTI                           As strct_GTI_RXXX
Public WS_GTM                           As strct_GTM_RXXX
Public WS_GTO                           As strct_GTO_RXXX
Public WS_GTP                           As strct_GTP_RXXX
Public WS_GTS                           As strct_GTS_RXXX
Public WS_GVA                           As strct_GVA_RXXX

Public WS_CHK_G07                       As Boolean
Public WS_CHK_G13                       As Boolean
Public WS_CHK_G14                       As Boolean
Public WS_CHK_G16                       As Boolean
Public WS_CHK_G22                       As Boolean
Public WS_CHK_G23                       As Boolean
Public WS_CHK_G34                       As Boolean
Public WS_CHK_GOR                       As Boolean
Public WS_CHK_GPB                       As Boolean
Public WS_CHK_GSE                       As Boolean

Public Function CHK_GDF_NVXXX(WS_DATA() As strct_GDF_Data) As Boolean
    
    On Error GoTo ErrHandler

    Dim I As Integer

    I = UBound(WS_DATA)

    CHK_GDF_NVXXX = True

ErrHandler:

End Function

Public Sub H2O_DATA_CLR()

    Dim WS_STRING As String
    
    WS_CHK_G07 = False
    WS_CHK_G13 = False
    WS_CHK_G14 = False
    WS_CHK_G16 = False
    WS_CHK_G22 = False
    WS_CHK_G23 = False
    WS_CHK_G34 = False
    WS_CHK_GOR = False
    WS_CHK_GPB = False
    WS_CHK_GSE = False
    
    ' 01
    '
    WS_STRING = String$(Len(WS_G01), " ")
    CopyMemory ByVal VarPtr(WS_G01), ByVal StrPtr(WS_STRING), Len(WS_G01) * 2
    
    ' 02
    '
    WS_STRING = String$(Len(WS_G02), " ")
    CopyMemory ByVal VarPtr(WS_G02), ByVal StrPtr(WS_STRING), Len(WS_G02) * 2
    
    ' 03
    '
    Erase WS_G03.R001()
        
    WS_STRING = String$(Len(WS_G03.R002), " ")
    CopyMemory ByVal VarPtr(WS_G03.R002), ByVal StrPtr(WS_STRING), Len(WS_G03.R002) * 2

    WS_STRING = String$(Len(WS_G03.R003), " ")
    CopyMemory ByVal VarPtr(WS_G03.R003), ByVal StrPtr(WS_STRING), Len(WS_G03.R003) * 2

    WS_STRING = String$(Len(WS_G03.R004), " ")
    CopyMemory ByVal VarPtr(WS_G03.R004), ByVal StrPtr(WS_STRING), Len(WS_G03.R004) * 2

    ' 05
    '
    WS_STRING = String$(Len(WS_G05), " ")
    CopyMemory ByVal VarPtr(WS_G05), ByVal StrPtr(WS_STRING), Len(WS_G05) * 2

    ' 06
    '
    WS_STRING = String$(Len(WS_G06), " ")
    CopyMemory ByVal VarPtr(WS_G06), ByVal StrPtr(WS_STRING), Len(WS_G06) * 2

    ' 06
    '
    WS_STRING = String$(Len(WS_G09), " ")
    CopyMemory ByVal VarPtr(WS_G09), ByVal StrPtr(WS_STRING), Len(WS_G09) * 2

    ' 11
    '
    WS_STRING = String$(Len(WS_G11), " ")
    CopyMemory ByVal VarPtr(WS_G11), ByVal StrPtr(WS_STRING), Len(WS_G11) * 2

    ' 12
    '
    Erase WS_G12()
    
    ' 13
    '
    Erase WS_G13()
    
    ' 14
    '
    Erase WS_G14()
    
    ' 16
    '
    Erase WS_G16()
    
    ' 17
    '
    Erase WS_G17()
    
    ' 21
    '
    Erase WS_G21.RXXX()
    
    WS_STRING = String$(Len(WS_G21.R005), " ")
    CopyMemory ByVal VarPtr(WS_G21.R005), ByVal StrPtr(WS_STRING), Len(WS_G21.R005) * 2
    
    WS_STRING = String$(Len(WS_G21.R999), " ")
    CopyMemory ByVal VarPtr(WS_G21.R999), ByVal StrPtr(WS_STRING), Len(WS_G21.R999) * 2
    
    ' 22
    '
    Erase WS_G22()
        
    ' 23
    '
    Erase WS_G23()
    
    ' 26
    '
    Erase WS_G26()
    
    ' 34
    '
    Erase WS_G34()
    
    ' AA
    '
    WS_STRING = String$(Len(WS_GAA), " ")
    CopyMemory ByVal VarPtr(WS_GAA), ByVal StrPtr(WS_STRING), Len(WS_GAA) * 2
    
    ' AR
    '
    WS_STRING = String$(Len(WS_GAR), " ")
    CopyMemory ByVal VarPtr(WS_GAR), ByVal StrPtr(WS_STRING), Len(WS_GAR) * 2
    
    ' AZ
    '
    WS_STRING = String$(Len(WS_GAZ), " ")
    CopyMemory ByVal VarPtr(WS_GAZ), ByVal StrPtr(WS_STRING), Len(WS_GAZ) * 2
    
    ' BO
    '
    WS_STRING = String$(Len(WS_GBO), " ")
    CopyMemory ByVal VarPtr(WS_GBO), ByVal StrPtr(WS_STRING), Len(WS_GBO) * 2
    
    ' BS
    '
    WS_STRING = String$(Len(WS_GBS), " ")
    CopyMemory ByVal VarPtr(WS_GBS), ByVal StrPtr(WS_STRING), Len(WS_GBS) * 2
    
    ' DF
    '
    Erase WS_GDF()
    Erase WS_GDF_QF()
    Erase WS_GDF_NV018()
    Erase WS_GDF_NV019()
    Erase WS_GDF_NV020()
    Erase WS_GDF_NV021()
    Erase WS_GDF_NV022()
    Erase WS_GDF_NV023()
    Erase WS_GDF_NV024()
    Erase WS_GDF_NV025()
    
    ' DM
    '
    WS_STRING = String$(Len(WS_GDM), " ")
    CopyMemory ByVal VarPtr(WS_GDM), ByVal StrPtr(WS_STRING), Len(WS_GDM) * 2
    
    ' DO
    '
    WS_STRING = String$(Len(WS_GDO), " ")
    CopyMemory ByVal VarPtr(WS_GDO), ByVal StrPtr(WS_STRING), Len(WS_GDO) * 2
    
    ' IV
    '
    WS_STRING = String$(Len(WS_GIV), " ")
    CopyMemory ByVal VarPtr(WS_GIV), ByVal StrPtr(WS_STRING), Len(WS_GIV) * 2
    
    ' NV
    '
    WS_STRING = String$(Len(WS_GNV), " ")
    CopyMemory ByVal VarPtr(WS_GNV), ByVal StrPtr(WS_STRING), Len(WS_GNV) * 2

    ' PB
    '
    WS_STRING = String$(Len(WS_GPB), " ")
    CopyMemory ByVal VarPtr(WS_GPB), ByVal StrPtr(WS_STRING), Len(WS_GPB) * 2

    ' OR
    '
    WS_STRING = String$(Len(WS_GOR), " ")
    CopyMemory ByVal VarPtr(WS_GOR), ByVal StrPtr(WS_STRING), Len(WS_GOR) * 2

    ' RE
    '
    WS_STRING = String$(Len(WS_GRE), " ")
    CopyMemory ByVal VarPtr(WS_GRE), ByVal StrPtr(WS_STRING), Len(WS_GRE) * 2

    ' RM
    '
    WS_STRING = String$(Len(WS_GRM), " ")
    CopyMemory ByVal VarPtr(WS_GRM), ByVal StrPtr(WS_STRING), Len(WS_GRM) * 2

    ' SE
    '
    Erase WS_GSE()

    ' SF
    '
    WS_STRING = String$(Len(WS_GSF), " ")
    CopyMemory ByVal VarPtr(WS_GSF), ByVal StrPtr(WS_STRING), Len(WS_GSF) * 2

    ' SI
    '
    Erase WS_GSI()

    ' SP
    '
    WS_STRING = String$(Len(WS_GSP), " ")
    CopyMemory ByVal VarPtr(WS_GSP), ByVal StrPtr(WS_STRING), Len(WS_GSP) * 2

    ' TB
    '
    'WS_STRING = String$(Len(WS_GTB), " ")
    'CopyMemory ByVal VarPtr(WS_GTB), ByVal StrPtr(WS_STRING), Len(WS_GTB) * 2
    
    ' TI
    '
    WS_STRING = String$(Len(WS_GTI), " ")
    CopyMemory ByVal VarPtr(WS_GTI), ByVal StrPtr(WS_STRING), Len(WS_GTI) * 2

    ' TM
    '
    WS_STRING = String$(Len(WS_GTM), " ")
    CopyMemory ByVal VarPtr(WS_GTM), ByVal StrPtr(WS_STRING), Len(WS_GTM) * 2
    
    ' TO
    '
    WS_STRING = String$(Len(WS_GTO), " ")
    CopyMemory ByVal VarPtr(WS_GTO), ByVal StrPtr(WS_STRING), Len(WS_GTO) * 2
    
    ' TP
    '
    WS_STRING = String$(Len(WS_GTP), " ")
    CopyMemory ByVal VarPtr(WS_GTP), ByVal StrPtr(WS_STRING), Len(WS_GTP) * 2
    
    ' TS
    '
    WS_STRING = String$(Len(WS_GTS), " ")
    CopyMemory ByVal VarPtr(WS_GTS), ByVal StrPtr(WS_STRING), Len(WS_GTS) * 2
    
    ' VA
    '
    WS_STRING = String$(Len(WS_GVA), " ")
    CopyMemory ByVal VarPtr(WS_GVA), ByVal StrPtr(WS_STRING), Len(WS_GVA) * 2

End Sub

Public Sub H2O_DATA_INIT()

    Dim WS_STRING As String
    
    ' 00
    '
    WS_STRING = String$(Len(WS_G00), " ")
    CopyMemory ByVal VarPtr(WS_G00), ByVal StrPtr(WS_STRING), Len(WS_G00) * 2
    
End Sub
