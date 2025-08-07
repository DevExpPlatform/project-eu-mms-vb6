Attribute VB_Name = "mod_ImportData_DBRtns"
Option Explicit

Private Type strct_XTXTDC
    FieldLenghtDB           As String
    FieldLenghtFile         As Long
    FieldName               As String
    FieldStart              As Long
    FieldType               As String
End Type

Dim XTXT_DataCutter()       As strct_XTXTDC

Private Function DB_CreateTable() As Boolean
    
    On Error GoTo ErrHandler
    
    Dim FieldName   As String
    Dim I           As Integer
    Dim MaxFields   As Integer
    Dim SQLLOBData  As String
    Dim SQLFields   As String
    Dim SQLTable    As String
        
    If DB_CheckTable(DLLParams.XTXT_TableName) = False Then
        MaxFields = UBound(XTXT_DataCutter)
        
        For I = 0 To MaxFields
            Select Case Left$(XTXT_DataCutter(I).FieldName, 6)
                Case "$BLANK"
                    FieldName = ""
                
                Case "$FIELD"
                    FieldName = Right$(XTXT_DataCutter(I).FieldName, Len(XTXT_DataCutter(I).FieldName) - InStr(1, XTXT_DataCutter(I).FieldName, "|"))
                
                    If (FieldName = "#SERIALBARCODE") Then FieldName = ""
                
                Case Else
                    FieldName = XTXT_DataCutter(I).FieldName
                
            End Select
            
            If FieldName <> "" Then
                SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & FieldName & " " & _
                            XTXT_DataCutter(I).FieldType & _
                            IIf(XTXT_DataCutter(I).FieldLenghtDB = "", "", "(" & XTXT_DataCutter(I).FieldLenghtDB & ")")
            
                If (XTXT_DataCutter(I).FieldType = "CLOB") Then
                    SQLLOBData = SQLLOBData & _
                                 " LOB (" & XTXT_DataCutter(I).FieldName & ") STORE AS (" & _
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
        
        SQLTable = Replace$(DLLParams.XTXT_TableName, "_", "")
        
        DBConn.Execute "CREATE TABLE " & DLLParams.XTXT_TableName & " (" & _
                       "ID_WORKCNTR NUMBER(14) NOT NULL, " & _
                       "ID_WORKINGLOAD NUMBER(14) NOT NULL, " & _
                       "ID_PACCO NUMBER(5), " & _
                       "ID_POSIZIONE NUMBER(5), " & _
                       IIf(DLLParams.XTXT_AddFieldsPSTL, "ID_PROVINCIA NUMBER(5), FLG_MIX NUMBER(3), ", "") & _
                       IIf(DLLParams.XTXT_AddFieldsBarCode, "NMR_ITFCODE NUMBER(11), ", "") & _
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
        
        DBConn.Execute "CREATE INDEX IDX_" & SQLTable & " ON " & DLLParams.XTXT_TableName & " (ID_WORKINGLOAD ASC, ID_PACCO ASC, ID_POSIZIONE ASC) " & _
                       "NOLOGGING " & _
                       "TABLESPACE MMS_IDXDATA " & _
                       "PARALLEL 2 " & _
                       "COMPRESS 2"
        
        DBConn.Execute "CREATE TRIGGER TRG_" & SQLTable & " BEFORE INSERT ON " & DLLParams.XTXT_TableName & " " & _
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

    Dim RS              As ADODB.Recordset
    Dim tmpStringSize   As Long

    Erase XTXT_DataCutter

    Set RS = DBConn.Execute("SELECT * FROM EDT_DATACUTTER WHERE ID_DATACUTTER = " & DLLParams.XTXT_IdDataCutter & " ORDER BY NMR_FIELDORDER ASC")

    If RS.RecordCount > 0 Then
        ReDim XTXT_DataCutter(RS.RecordCount - 1)
                
        Do Until RS.EOF
            With XTXT_DataCutter(RS.AbsolutePosition - 1)
                .FieldLenghtFile = RS("nmr_FileFieldLenght")
                .FieldName = RS("descr_FieldName")
                .FieldStart = tmpStringSize + 1
                .FieldType = RS("str_DBFieldType")
            
                If IsNull(RS("nmr_DBFieldLenght")) Then
                    Select Case RS("str_DBFieldType")
                        Case "DATE", "CLOB"
                            .FieldLenghtDB = ""
                        
                        Case Else
                            .FieldLenghtDB = .FieldLenghtFile
                    
                    End Select
                Else
                    .FieldLenghtDB = RS("nmr_DBFieldLenght")
                End If
            End With
        
            If Left$(RS("descr_FieldName"), 6) <> "$FIELD" Then tmpStringSize = tmpStringSize + RS("nmr_FileFieldLenght")
        
            RS.MoveNext
        Loop
    
        DB_DataCutter_SELECT = True
    End If
    
    RS.Close

    Set RS = Nothing

    Exit Function

ErrHandler:
    Set RS = Nothing
    
    Erase XTXT_DataCutter
    
    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox "Error loading data import capabilities.", vbCritical, "Data Caps:"
    End If

End Function

Public Function DB_InsertData() As Boolean
                
    On Error GoTo ErrHandler
    
    Dim CntrCLOB        As Integer
'    Dim CntrMaxValue    As Long
    Dim DataCLOB()      As String
    Dim DataField       As String
    Dim dta_WLoad       As String
    Dim ErrMsg          As String
    Dim FileLenInfo     As Long
    Dim FileLocInfo     As Long
    Dim FN              As Integer
    Dim I               As Integer
    Dim InptStr         As String
    Dim MaxFields       As Integer
    Dim myAPB           As cls_APB
    Dim myOracleCLOB    As OraClob
    Dim myOracleDB      As OraDatabase
    Dim myOracleSession As OraSessionClass
    Dim RetVal          As Long
    Dim SplitData()     As String
    Dim SQLFields       As String
    Dim SQLString       As String
    
    If (DLLParams.UnattendedMode = False) Then
        Set myAPB = New cls_APB
        
        With myAPB
            .APBMode = PBSingle
            .APBCaption = "Importing Data:"
            .APBMaxItems = 1
            .APBShow
        End With
    End If
    
    SplitData = Split(DLLParams.XTXT_TNS, "|")
    
    Set myOracleSession = CreateObject("OracleInProcServer.XOraSession")
    Set myOracleDB = myOracleSession.DbOpenDatabase(SplitData(0), SplitData(1), 0&)
    Set myOracleCLOB = myOracleDB.CreateTempClob
                
    myOracleSession.BeginTrans

    dta_WLoad = Format$(Now, "yyyyMMddhhmmss")
    FN = FreeFile
    MaxFields = UBound(XTXT_DataCutter)
    SQLFields = ""
            
    ' Get Field Names
    '
    For I = 0 To MaxFields
        Select Case Left$(XTXT_DataCutter(I).FieldName, 6)
            Case "$BLANK", "$FIELD"
                        
            Case Else
                Select Case XTXT_DataCutter(I).FieldType
                    Case "CLOB"
                        ReDim Preserve DataCLOB(CntrCLOB)
                        
                        DataCLOB(CntrCLOB) = "CLOB" & Replace$(XTXT_DataCutter(I).FieldName, "_", "")
                        myOracleDB.Parameters.Add DataCLOB(CntrCLOB), Null, ORAPARM_INPUT, ORATYPE_CLOB
                        
                        CntrCLOB = CntrCLOB + 1
                        
                End Select
                
                SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & XTXT_DataCutter(I).FieldName
                            
        End Select
    Next I
    
    ' Insert Data
    '
    'SQLFields = "INSERT INTO " & DLLParams.XTXT_TableName & " (ID_WORKCNTR, ID_WORKINGLOAD, " & SQLFields & ") VALUES ("
    SQLFields = "INSERT INTO " & DLLParams.XTXT_TableName & " (ID_WORKINGLOAD, " & SQLFields & ") VALUES ("
            
    Open DLLParams.XTXT_FileName For Input As #FN
'        CntrMaxValue = Val(DB_GetValueByID("SELECT MAX(ID_WORKCNTR) FROM " & DLLParams.XTXT_TableName))
        If (DLLParams.UnattendedMode = False) Then
            FileLenInfo = (LOF(FN) \ 1024)
            
            myAPB.APBMaxItems = FileLenInfo
        End If
        
        Do Until EOF(FN)
            Line Input #FN, InptStr
                    
            CntrCLOB = 0
            
            If (DLLParams.UnattendedMode = False) Then
                FileLocInfo = (Loc(FN) \ 8)
    '            CntrMaxValue = CntrMaxValue + 1
                
                myAPB.APBItemsLabel = "Read " & Format$(FileLocInfo, "##,##") & " KB of " & Format$(FileLenInfo, "##,##") & " KB"
                If FileLocInfo > 0 Then myAPB.APBItemsProgress = FileLocInfo
            End If
            
            SQLString = ""
            
            For I = 0 To MaxFields
                Select Case Left$(XTXT_DataCutter(I).FieldName, 6)
                    Case "$BLANK", "$FIELD"
                    
                    Case Else
                        DataField = RTrim$(Mid$(InptStr, XTXT_DataCutter(I).FieldStart, XTXT_DataCutter(I).FieldLenghtFile))
                        
                        ' Debug.Print DataField
                        
                        Select Case XTXT_DataCutter(I).FieldType
                            Case "CLOB"
                                If InStr(1, DataField, "<br>") > 0 Then DataField = Replace$(DataField, "<br>", vbNewLine)
                                If InStr(1, DataField, "<bs>") > 0 Then DataField = Replace$(DataField, "<bs>", "")
                                If InStr(1, DataField, "<es>") > 0 Then DataField = Replace$(DataField, "<es>", "")
                            
                                If DataField = "" Then
                                    DataField = "NULL"
                                Else
                                    myOracleCLOB.Erase myOracleCLOB.Size
                                    myOracleCLOB.Write DataField
                                    myOracleCLOB.Trim Len(DataField)
                                    
                                    myOracleDB.Parameters(DataCLOB(CntrCLOB)).Value = myOracleCLOB
    
                                    DataField = ":" & DataCLOB(CntrCLOB)
                                    
                                    CntrCLOB = CntrCLOB + 1
                                End If
                                                                
                            Case "DATE"
                                DataField = Conv_Time2SQLServerTime(DataField)
                            
                            Case "NUMBER"
                                DataField = Conv_Str2Num(DataField)
                            
                            Case Else
                                DataField = Conv_String2SQLString(DataField)
                                
                                If InStr(1, DataField, "<br>") > 0 Then DataField = Replace$(DataField, "<br>", vbNewLine)
                                If InStr(1, DataField, "<bs>") > 0 Then DataField = Replace$(DataField, "<bs>", "")
                                If InStr(1, DataField, "<es>") > 0 Then DataField = Replace$(DataField, "<es>", "")
                        
                        End Select
                        
                        SQLString = SQLString & ", " & DataField
                
                End Select
            Next I

'            SQLString = SQLFields & CntrMaxValue & ", " & dta_WLoad & SQLString & ")"
            SQLString = SQLFields & dta_WLoad & SQLString & ")"
            
            'Debug.Print SQLString
            
            RetVal = myOracleDB.ExecuteSQL(SQLString)
            
            If RetVal = 0 Then
                ErrMsg = "Errore durante la scrittura dei records."
                
                GoTo ErrHandler
            End If
        Loop
    Close #FN
            
    If (DLLParams.UnattendedMode = False) Then
        myAPB.APBClose
        Set myAPB = Nothing
    End If
    
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
            MsgBox Purge_ErrDescr(Err.Description), vbCritical, "TXT PlugIn Error:"
        Else
            MsgBox ErrMsg, vbCritical, "TXT PlugIn Error:"
        End If
    End If

End Function

Public Function DB_TXTImport() As Boolean
    
    Dim CanDo As Boolean
    
    If DB_ConnectInit Then
        CanDo = DB_DataCutter_SELECT
        
        If CanDo Then CanDo = DB_CreateTable
        
        DB_ConnectRelease
        
        If CanDo Then CanDo = DB_InsertData
        
        DB_TXTImport = CanDo
    End If
 
End Function
