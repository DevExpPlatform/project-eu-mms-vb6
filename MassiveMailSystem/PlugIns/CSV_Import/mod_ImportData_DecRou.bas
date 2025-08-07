Attribute VB_Name = "mod_ImportData_DecRou"
Option Explicit

Private Type strct_XCSVDC
    FieldLenghtDB           As String
    FieldName               As String
    FieldStart              As Long
    FieldType               As String
End Type

Private XCSV_DataCutter()   As strct_XCSVDC

Private Function DB_CreateTable() As Boolean
    
    On Error GoTo ErrHandler
    
    Dim FieldName   As String
    Dim I           As Integer
    Dim MaxFields   As Integer
    Dim SQLLOBData  As String
    Dim SQLFields   As String
    Dim SQLTable    As String
    
    If DB_CheckTable(DLLParams.XCSV_TableName) = False Then
        MaxFields = UBound(XCSV_DataCutter)
        
        For I = 0 To MaxFields
            Select Case Left$(XCSV_DataCutter(I).FieldName, 6)
                Case "$BLANK"
                    FieldName = ""
                
                Case "$FIELD"
                    FieldName = Right$(XCSV_DataCutter(I).FieldName, Len(XCSV_DataCutter(I).FieldName) - InStr(1, XCSV_DataCutter(I).FieldName, "|"))
                
                    If (FieldName = "#SERIALBARCODE") Then FieldName = ""
                
                Case Else
                    FieldName = XCSV_DataCutter(I).FieldName
                
            End Select
            
            If FieldName <> "" Then
                SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & FieldName & " " & _
                            XCSV_DataCutter(I).FieldType & _
                            IIf(XCSV_DataCutter(I).FieldLenghtDB = "", "", "(" & XCSV_DataCutter(I).FieldLenghtDB & ")")
            
                If (XCSV_DataCutter(I).FieldType = "CLOB") Then
                    SQLLOBData = SQLLOBData & _
                                 " LOB (" & XCSV_DataCutter(I).FieldName & ") STORE AS (" & _
                                 "DISABLE STORAGE IN ROW " & _
                                 "PCTVERSION 0 " & _
                                 "CHUNK 512 " & _
                                 "CACHE " & _
                                 "TABLESPACE MMS_LOBDATA " & _
                                 "STORAGE (" & _
                                 "INITIAL 1048576 " & _
                                 "NEXT 5242880)" & _
                                 ")"
                End If
            End If
        Next I
        
        SQLTable = Replace$(DLLParams.XCSV_TableName, "_", "")
        
        DBConn.Execute "CREATE TABLE " & DLLParams.XCSV_TableName & " (" & _
                       "ID_WORKCNTR NUMBER(14) NOT NULL, " & _
                       "ID_WORKINGLOAD NUMBER(14) NOT NULL, " & _
                       "ID_PACCO NUMBER(5), " & _
                       "ID_POSIZIONE NUMBER(5), " & _
                       IIf(DLLParams.XCSV_AddFieldsPSTL, "ID_PROVINCIA NUMBER(5), FLG_MIX NUMBER(3), ", "") & _
                       IIf(DLLParams.XCSV_AddFieldsBarCode, "NMR_ITFCODE NUMBER(11), ", "") & _
                       SQLFields & _
                       ", CONSTRAINT PK_" & SQLTable & " PRIMARY KEY (ID_WORKCNTR, ID_WORKINGLOAD) USING INDEX TABLESPACE MMS_IDXDATA) " & _
                       "CACHE " & _
                       "NOLOGGING " & _
                       "TABLESPACE MMS_WRKDATA " & _
                       "STORAGE (" & _
                       "INITIAL 1048576 " & _
                       "NEXT 5242880 " & _
                       "MAXEXTENTS UNLIMITED " & _
                       "BUFFER_POOL DEFAULT) " & _
                       SQLLOBData
                       
        On Error Resume Next
        DBConn.Execute "DROP SEQUENCE SQN_" & SQLTable
        
        On Error GoTo ErrHandler
        DBConn.Execute "CREATE SEQUENCE SQN_" & SQLTable & " MINVALUE 1 MAXVALUE 99999999999999 INCREMENT BY 1 START WITH 1 NOCACHE ORDER NOCYCLE"
        
        DBConn.Execute "CREATE INDEX IDX_" & SQLTable & " ON " & DLLParams.XCSV_TableName & " (ID_WORKINGLOAD ASC, ID_PACCO ASC, ID_POSIZIONE ASC) " & _
                       "NOLOGGING " & _
                       "TABLESPACE MMS_IDXDATA " & _
                       "COMPRESS 2"
        
        DBConn.Execute "CREATE TRIGGER TRG_" & SQLTable & " BEFORE INSERT ON " & DLLParams.XCSV_TableName & " " & _
                       "FOR EACH ROW " & _
                       "BEGIN " & _
                       "SELECT SQN_" & SQLTable & ".NEXTVAL INTO :NEW.ID_WORKCNTR FROM DUAL; " & _
                       "END;"
    End If
    
    DB_CreateTable = True
    
    Exit Function

ErrHandler:
    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox Err.Description, vbCritical, "TXT PlugIn Error:"
    End If
    
End Function

Private Function DB_DataCutter_SELECT() As Boolean

    On Error GoTo ErrHandler

    Dim RS  As ADODB.Recordset

    Erase XCSV_DataCutter

    Set RS = DBConn.Execute("SELECT * FROM EDT_DATACUTTER WHERE ID_DATACUTTER = " & DLLParams.XCSV_IdDataCutter & " ORDER BY NMR_FIELDORDER ASC")

    If RS.RecordCount > 0 Then
        ReDim XCSV_DataCutter(RS.RecordCount - 1)
                
        Do Until RS.EOF
            With XCSV_DataCutter(RS.AbsolutePosition - 1)
                .FieldName = RS("descr_FieldName")
                .FieldType = RS("str_DBFieldType")
                
                If IsNull(RS("nmr_DBFieldLenght")) Then
                    Select Case RS("str_DBFieldType")
                        Case "DATE", "CLOB"
                            .FieldLenghtDB = ""
                        
                        Case Else
                            .FieldLenghtDB = RS("nmr_FileFieldLenght")
                    
                    End Select
                Else
                    .FieldLenghtDB = RS("nmr_DBFieldLenght")
                End If
            End With
            
            RS.MoveNext
        Loop
    
        DB_DataCutter_SELECT = True
    End If
    
    RS.Close

    Set RS = Nothing

    Exit Function

ErrHandler:
    Set RS = Nothing
    
    Erase XCSV_DataCutter
    
    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox "Error loading data import capabilities.", vbCritical, "Data Caps:"
    End If

End Function

Private Function DB_InsertData() As Boolean

    On Error GoTo ErrHandler
    
    Dim CntrCLOB        As Integer
    'Dim CntrMaxValue    As Long
    Dim DataCLOB()      As String
    'Dim DataField       As String
    Dim dta_WLoad       As String
    Dim ErrMsg          As String
    Dim FN              As Integer
    Dim FileLenInfo     As Long
    Dim FileLocInfo     As Long
    Dim I               As Integer
    Dim InptStr         As String
    Dim MaxFields       As Integer
    Dim myAPB           As cls_APB
    Dim myOracleCLOB    As OraClob
    Dim myOracleDB      As OraDatabase
    Dim myOracleSession As OraSessionClass
    Dim RetVal          As Integer
    Dim SplitData()     As String
    Dim SQLFields       As String
    Dim SQLString       As String
    'Dim SQLTable        As String
    
    If (DLLParams.UnattendedMode = False) Then
        Set myAPB = New cls_APB
        
        With myAPB
            .APBMode = PBSingle
            .APBCaption = "Importing Data:"
            .APBMaxItems = 1
            .APBShow
        End With
    End If

    SplitData = Split(DLLParams.XCSV_TNS, "|")

    Set myOracleSession = CreateObject("OracleInProcServer.XOraSession")
    Set myOracleDB = myOracleSession.DbOpenDatabase(SplitData(0), SplitData(1), 0&)
    Set myOracleCLOB = myOracleDB.CreateTempClob

    myOracleSession.BeginTrans

    dta_WLoad = Format$(Now, "yyyyMMddhhmmss")
    FN = FreeFile
    MaxFields = UBound(XCSV_DataCutter)
    SQLFields = ""

    ' Insert Data
    '
    SQLFields = ""
    
    For I = 0 To MaxFields
        Select Case Left$(XCSV_DataCutter(I).FieldName, 6)
            Case "$BLANK", "$FIELD"
            
            Case Else
                Select Case XCSV_DataCutter(I).FieldType
                    Case "CLOB"
                        ReDim Preserve DataCLOB(CntrCLOB)

                        DataCLOB(CntrCLOB) = "CLOB" & Replace$(XCSV_DataCutter(I).FieldName, "_", "")
                        myOracleDB.Parameters.Add DataCLOB(CntrCLOB), Null, ORAPARM_INPUT, ORATYPE_CLOB

                        CntrCLOB = CntrCLOB + 1

                End Select
            
                SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & XCSV_DataCutter(I).FieldName
        
        End Select
    Next I
    
    'SQLFields = "INSERT INTO " & DLLParams.XCSV_TableName & " (ID_WORKCNTR, ID_WORKINGLOAD, " & SQLFields & ") VALUES ("
    SQLFields = "INSERT INTO " & DLLParams.XCSV_TableName & " (ID_WORKINGLOAD, " & SQLFields & ") VALUES ("
    
    Open DLLParams.XCSV_FileName For Input As #FN
        'CntrMaxValue = Val(DB_GetValueByID("SELECT MAX(ID_WORKCNTR) FROM " & DLLParams.XCSV_TableName))
        If (DLLParams.UnattendedMode = False) Then
            FileLenInfo = (LOF(FN) \ 1024)
            myAPB.APBMaxItems = FileLenInfo
        End If
        
        Do Until EOF(FN)
            Line Input #FN, InptStr
            
            CntrCLOB = 0
            
            If (DLLParams.UnattendedMode = False) Then
                FileLocInfo = (Loc(FN) \ 8)
                'CntrMaxValue = CntrMaxValue + 1
    
                myAPB.APBItemsLabel = "Read " & Format$(FileLocInfo, "##,##") & " KB of " & Format$(FileLenInfo, "##,##") & " KB"
                If FileLocInfo > 0 Then myAPB.APBItemsProgress = FileLocInfo
            End If
            
            SQLString = ""
                
            If Trim$(InptStr) <> "" Then
                'If Left$(InptStr, 1) = Chr$(34) Then InptStr = Replace$(InptStr, Chr$(34), "")

                SplitData = Split(InptStr, DLLParams.XCSV_SubDivideChar)

                If chk_Array(SplitData) Then
                    If UBound(SplitData) = MaxFields Then
                        For I = 0 To MaxFields
                            Select Case Left$(XCSV_DataCutter(I).FieldName, 6)
                            Case "$BLANK", "$FIELD"
                              
                            Case Else
                                Select Case XCSV_DataCutter(I).FieldType
                                Case "CLOB"
                                    If InStr(1, SplitData(I), "<br>") > 0 Then SplitData(I) = Replace$(SplitData(I), "<br>", vbNewLine)
                                    If InStr(1, SplitData(I), "<bs>") > 0 Then SplitData(I) = Replace$(SplitData(I), "<bs>", "")
                                    If InStr(1, SplitData(I), "<es>") > 0 Then SplitData(I) = Replace$(SplitData(I), "<es>", "")

                                    If SplitData(I) = "" Then
                                         SplitData(I) = "NULL"
                                    Else
                                        myOracleCLOB.Erase myOracleCLOB.Size
                                        myOracleCLOB.Write SplitData(I)
                                        myOracleCLOB.Trim Len(SplitData(I))
                                        
                                        myOracleDB.Parameters(DataCLOB(CntrCLOB)).Value = myOracleCLOB
                                            
                                        SplitData(I) = ":" & DataCLOB(CntrCLOB)
    
                                        CntrCLOB = CntrCLOB + 1
                                    End If
    
                                Case "DATE"
                                    SplitData(I) = Conv_Time2SQLServerTime(SplitData(I))
                                
                                Case "NUMBER"
                                    SplitData(I) = Conv_Str2Num(SplitData(I))
                                
                                Case Else
                                    SplitData(I) = Conv_String2SQLString(SplitData(I))

                                    If InStr(1, SplitData(I), "<br>") > 0 Then SplitData(I) = Replace$(SplitData(I), "<br>", vbNewLine)
                                    If InStr(1, SplitData(I), "<bs>") > 0 Then SplitData(I) = Replace$(SplitData(I), "<bs>", "")
                                    If InStr(1, SplitData(I), "<es>") > 0 Then SplitData(I) = Replace$(SplitData(I), "<es>", "")
                            
                            End Select
                            
                            SQLString = SQLString & ", " & SplitData(I)
                          
                          End Select
                        Next I
                    Else
                        'XCSV_ErrMsg = "Colonne incongruenti"
                    End If
                Else
                    'XCSV_ErrMsg = "Colonne mancanti"
                End If
            End If
            
            'SQLString = SQLFields & CntrMaxValue & ", " & dta_WLoad & SQLString & ")"
            SQLString = SQLFields & dta_WLoad & SQLString & ")"
            
            ' Debug.Print SQLString

            RetVal = myOracleDB.ExecuteSQL(SQLString)
            
            If RetVal = 0 Then
                ErrMsg = "Errore durante la scrittura dei records."

                GoTo ErrHandler
            End If
        Loop
    Close #FN
        
    For I = 0 To (CntrCLOB - 1)
        myOracleDB.Parameters.Remove DataCLOB(I)
    Next I

    Erase DataCLOB

    myOracleCLOB.Erase myOracleCLOB.Size
    Set myOracleCLOB = Nothing

    myOracleSession.CommitTrans
    myOracleDB.Close

    Set myOracleSession = Nothing
    Set myOracleDB = Nothing

    If (DLLParams.UnattendedMode = False) Then
        myAPB.APBClose
        Set myAPB = Nothing
    End If

    DB_InsertData = True
    
    Exit Function

ErrHandler:
    myOracleSession.Rollback
    myOracleDB.Close

    Set myOracleSession = Nothing
    Set myOracleDB = Nothing

    If (DLLParams.UnattendedMode = False) Then
        myAPB.APBClose
        Set myAPB = Nothing
    End If
    
    Close #FN

    If (DLLParams.UnattendedMode) Then
        If ErrMsg = "" Then
            UMErrMsg = Err.Description
        Else
            UMErrMsg = ErrMsg
        End If
    Else
        If ErrMsg = "" Then
            MsgBox Purge_ErrDescr(Err.Description), vbCritical, "CSV PlugIn Error:"
        Else
            MsgBox ErrMsg, vbCritical, "CSV PlugIn Error:"
        End If
    End If

End Function

Public Function DB_CSVImport() As Boolean
    
    Dim CanDo As Boolean
    
    If DB_ConnectInit Then
        CanDo = DB_DataCutter_SELECT
        
        If CanDo Then CanDo = DB_CreateTable
        
        DB_ConnectRelease
    
        If CanDo Then CanDo = DB_InsertData
        
        DB_CSVImport = CanDo
    End If
 
End Function
