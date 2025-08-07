Attribute VB_Name = "mod_ImportData_DecRou"
Option Explicit

Private Type strct_XSQLDC
    FieldLenghtDB           As String
    FieldName               As String
    FieldStart              As Long
    FieldType               As String
End Type

Private DataCLOB()          As String
Private DTA_WLOAD           As String
Private MaxFields           As Integer
Private myOracleCLOB        As OraClob
Private myOracleDB          As OraDatabase
Private myOracleSession     As OraSessionClass
Private SQLFields           As String
Private XSQL_DataCutter()   As strct_XSQLDC

Private Function DB_CreateTable() As Boolean
    
    On Error GoTo ErrHandler
    
    Dim FieldName   As String
    Dim I           As Integer
    Dim MaxFields   As Integer
    Dim SQLLOBData  As String
    Dim SQLFields   As String
    Dim SQLTable    As String
    
    If DB_CheckTable(DLLParams.XSQL_TableName) = False Then
        MaxFields = UBound(XSQL_DataCutter)
        
        For I = 0 To MaxFields
            Select Case Left$(XSQL_DataCutter(I).FieldName, 6)
                Case "$BLANK"
                    FieldName = ""
                
                Case "$FIELD"
                    FieldName = Right$(XSQL_DataCutter(I).FieldName, Len(XSQL_DataCutter(I).FieldName) - InStr(1, XSQL_DataCutter(I).FieldName, "|"))
                
                    If (FieldName = "#SERIALBARCODE") Then FieldName = ""
                
                Case Else
                    FieldName = XSQL_DataCutter(I).FieldName
                
            End Select
            
            If FieldName <> "" Then
                SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & FieldName & " " & _
                            XSQL_DataCutter(I).FieldType & _
                            IIf(XSQL_DataCutter(I).FieldLenghtDB = "", "", "(" & XSQL_DataCutter(I).FieldLenghtDB & ")")
            
                If ((XSQL_DataCutter(I).FieldType = "CLOB") Or (XSQL_DataCutter(I).FieldType = "NCLOB")) Then
                    SQLLOBData = SQLLOBData & _
                                 " LOB (" & XSQL_DataCutter(I).FieldName & ") STORE AS (" & _
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
        
        SQLTable = Replace$(DLLParams.XSQL_TableName, "_", "")
        
        DBConn.Execute "CREATE TABLE " & DLLParams.XSQL_TableName & " (" & _
                       "ID_WORKCNTR NUMBER(14) NOT NULL, " & _
                       "ID_WORKINGLOAD NUMBER(14) NOT NULL, " & _
                       "ID_PACCO NUMBER(5), " & _
                       "ID_POSIZIONE NUMBER(5), " & _
                       IIf(DLLParams.XSQL_AddFieldsPSTL, "ID_PROVINCIA NUMBER(5), FLG_MIX NUMBER(3), ", "") & _
                       IIf(DLLParams.XSQL_AddFieldsBarCode, "NMR_ITFCODE NUMBER(11), ", "") & _
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
        DBConn.Execute "CREATE SEQUENCE SQN_" & SQLTable & " MINVALUE 1 MAXVALUE 99999999999999 INCREMENT BY 1 START WITH 1 CACHE 20"
        
        DBConn.Execute "CREATE INDEX IDX_" & SQLTable & " ON " & DLLParams.XSQL_TableName & " (ID_WORKINGLOAD ASC, ID_PACCO ASC, ID_POSIZIONE ASC) " & _
                       "NOLOGGING " & _
                       "TABLESPACE MMS_IDXDATA " & _
                       "PARALLEL 2 " & _
                       "COMPRESS 2"
        
        DBConn.Execute "CREATE TRIGGER TRG_" & SQLTable & " BEFORE INSERT ON " & DLLParams.XSQL_TableName & " " & _
                       "FOR EACH ROW " & _
                       "BEGIN " & _
                       "SELECT SQN_" & SQLTable & ".NEXTVAL INTO :NEW.ID_WORKCNTR FROM DUAL; " & _
                       "END;"
    End If
    
    DB_CreateTable = True
    
    Exit Function

ErrHandler:
    UMErrMsg = Err.Description

End Function

Private Function DB_DataCutter_SELECT() As Boolean

    On Error GoTo ErrHandler

    Dim RS  As ADODB.Recordset

    Erase XSQL_DataCutter

    Set RS = DBConn.Execute("SELECT * FROM EDT_DATACUTTER WHERE ID_DATACUTTER = " & DLLParams.XSQL_IdDataCutter & " ORDER BY NMR_FIELDORDER ASC")

    If RS.RecordCount > 0 Then
        ReDim XSQL_DataCutter(RS.RecordCount - 1)
                
        Do Until RS.EOF
            With XSQL_DataCutter(RS.AbsolutePosition - 1)
                .FieldName = RS("descr_FieldName")
                .FieldType = RS("str_DBFieldType")
                
                If IsNull(RS("nmr_DBFieldLenght")) Then
                    Select Case RS("str_DBFieldType")
                        Case "DATE", "CLOB", "NCLOB"
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
    
    Erase XSQL_DataCutter
    
    UMErrMsg = Err.Description

End Function

Public Function DB_InsertData(InptStr As String) As Boolean

    On Error GoTo ErrHandler
    
    Dim I               As Integer
    Dim CntrCLOB        As Integer
    Dim SplitData()     As String
    Dim SQLString       As String
    Dim RetVal          As Integer
    
    CntrCLOB = 0
    SQLString = ""
    UMErrMsg = ""
    
    If Trim$(InptStr) <> "" Then
        SplitData = Split(InptStr, DLLParams.XSQL_SubDivideChar)

        If chk_Array(SplitData) Then
            If UBound(SplitData) = MaxFields Then
                For I = 0 To MaxFields
                    Select Case Left$(XSQL_DataCutter(I).FieldName, 6)
                    Case "$BLANK", "$FIELD"
                    Case Else
                        Select Case XSQL_DataCutter(I).FieldType
                        Case "CLOB", "NCLOB"
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
            End If
        End If
    End If
            
    SQLString = SQLFields & DTA_WLOAD & SQLString & ")"
            
    ' Debug.Print SQLString

    RetVal = myOracleDB.ExecuteSQL(SQLString)
            
    If RetVal = 0 Then
        UMErrMsg = "Errore durante la scrittura dei records."

        GoTo ErrHandler
    End If
    
    DB_InsertData = True
    
    Exit Function

ErrHandler:
    If (UMErrMsg = "") Then UMErrMsg = Err.Description

End Function

Public Sub DB_OO4OClose(IsOnError As Boolean)

    Dim I As Integer
    
    If (chk_Array(DataCLOB)) Then
        For I = 0 To (UBound(DataCLOB))
            myOracleDB.Parameters.Remove DataCLOB(I)
        Next I
    
        Erase DataCLOB
        
        myOracleCLOB.Erase myOracleCLOB.Size
    End If

    Set myOracleCLOB = Nothing

    If (IsOnError) Then
        myOracleSession.Rollback
    Else
        myOracleSession.CommitTrans
    End If
    
    myOracleDB.Close

    Set myOracleSession = Nothing
    Set myOracleDB = Nothing
    
End Sub

Private Function DB_OO4OInit()

    On Error GoTo ErrHandler

    Dim CntrCLOB        As Integer
    Dim I               As Integer
    Dim IsConnected     As Boolean
    Dim SplitData()     As String

    SplitData = Split(DLLParams.XSQL_TNS, "|")

    Set myOracleSession = CreateObject("OracleInProcServer.XOraSession")
    Set myOracleDB = myOracleSession.DbOpenDatabase(SplitData(0), SplitData(1), 0&)
    Set myOracleCLOB = myOracleDB.CreateTempClob(True)
    
    IsConnected = True

    myOracleSession.BeginTrans

    DTA_WLOAD = Format$(Now, "yyyyMMddhhmmss")
    MaxFields = UBound(XSQL_DataCutter)
    SQLFields = ""
    
    For I = 0 To MaxFields
        Select Case Left$(XSQL_DataCutter(I).FieldName, 6)
        Case "$BLANK", "$FIELD"
        
        Case Else
            Select Case XSQL_DataCutter(I).FieldType
            Case "CLOB", "NCLOB"
                ReDim Preserve DataCLOB(CntrCLOB)

                DataCLOB(CntrCLOB) = "CLOB" & Replace$(XSQL_DataCutter(I).FieldName, "_", "")
                myOracleDB.Parameters.Add DataCLOB(CntrCLOB), Null, ORAPARM_INPUT, ORATYPE_CLOB

                CntrCLOB = (CntrCLOB + 1)
        
            End Select
        
            SQLFields = SQLFields & IIf(SQLFields = "", "", ", ") & XSQL_DataCutter(I).FieldName
        
        End Select
    Next I

    SQLFields = "INSERT INTO " & DLLParams.XSQL_TableName & " (ID_WORKINGLOAD, " & SQLFields & ") VALUES ("

    DB_OO4OInit = True

    Exit Function

ErrHandler:
    If (IsConnected) Then myOracleDB.Close

    Set myOracleCLOB = Nothing
    Set myOracleDB = Nothing
    Set myOracleSession = Nothing
    
    UMErrMsg = Err.Description

End Function

Public Function DB_SQLImport() As Boolean
    
    DB_SQLImport = DB_DataCutter_SELECT
    
    If (DB_SQLImport = False) Then Exit Function
    
    DB_SQLImport = DB_CreateTable
        
    If (DB_SQLImport = False) Then Exit Function
    
    DB_ConnectRelease
    
    DB_SQLImport = DB_OO4OInit
    
End Function

Public Function GET_IDWORKINGLOAD() As String

    GET_IDWORKINGLOAD = DTA_WLOAD

End Function
