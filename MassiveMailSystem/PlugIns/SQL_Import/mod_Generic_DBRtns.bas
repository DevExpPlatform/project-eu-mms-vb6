Attribute VB_Name = "mod_Generic_DBRtns"
Option Explicit

Public DBConn As ADODB.Connection

Public Function DB_Check(ByVal SQLStr As String, Optional ByVal ChkField As String) As Boolean

    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset
    
    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If

    Set RS = DBConn.Execute(SQLStr)

    With RS
        If .RecordCount > 0 Then
            If ChkField <> "" Then
                DB_Check = Not IsNull(RS(ChkField))
            Else
                DB_Check = True
            End If
        End If
        
        .Close
    End With

    Set RS = Nothing
    
    If CanDisconnect Then DBConn.Close

    Exit Function

ErrHandler:

End Function

Public Function DB_CheckTable(ByVal TableName As String)
    
    On Error GoTo ErrHandler
    
    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset
    
    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If

    Set RS = DBConn.Execute("SELECT * FROM " & TableName & " WHERE ROWNUM = 1")
    Set RS = Nothing
    
    If CanDisconnect Then DBConn.Close

    DB_CheckTable = True
    
    Exit Function

ErrHandler:
    If CanDisconnect Then DBConn.Close

End Function

Public Function DB_ConnectInit() As Boolean
    
    On Error GoTo ErrHandler
    
    Set DBConn = New ADODB.Connection
    
    With DBConn
        .ConnectionString = "DSN=" & DLLParams.XSQL_DSN
        .CursorLocation = adUseClient
        .CommandTimeout = 0
        .Open
    End With

    DB_ConnectInit = True

    Exit Function

ErrHandler:
    Set DBConn = Nothing
    
    DB_ConnectInit = False
    
    UMErrMsg = Err.Description
    
End Function

Public Sub DB_ConnectRelease()
    
    If DBConn.State Then DBConn.Close
    
    Set DBConn = Nothing

End Sub

Public Function DB_ExecuteQuery(ByVal SQLString As String, ByVal ShowQuestion As Boolean, ByVal ShowMsgOk As Boolean, ByVal ShowMsgErr As Boolean, ByVal QueryDescr As String) As Boolean
    
    On Error GoTo ErrHandler
    
    If ShowQuestion Then
        If MsgBox("Sicuro di voler proseguire?", vbQuestion + vbYesNo, QueryDescr) = vbNo Then
            Exit Function
        End If
    End If
    
    ' Start
    '
    Dim CanDisconnect   As Boolean
    Dim RValue          As Integer
    
    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
            
    DBConn.Execute SQLString, RValue
            
    If CanDisconnect Then DBConn.Close
    
    DB_ExecuteQuery = (RValue > 0)
    
    If DB_ExecuteQuery Then
        If ShowMsgOk Then MsgBox "Operazione eseguita correttamente.", vbExclamation, QueryDescr
    Else
        If ShowMsgErr Then MsgBox "Errore durante l'aggiornamento del case selezionato.", vbExclamation, QueryDescr
    End If
    
    Exit Function

ErrHandler:
    'MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Modify Case Description:"
    
    If CanDisconnect Then DBConn.Close

End Function

Public Function DB_GetLastIdentity(ByVal myTable As String) As Long

    Dim RS As ADODB.Recordset
    
    Set RS = DBConn.Execute("SELECT @@identity AS NewID FROM " & myTable)

    DB_GetLastIdentity = RS("NewID")

    RS.Close
    Set RS = Nothing
    
End Function

Public Function DB_GetValueByID(ByVal SQLString As String, Optional ByVal MultiRow As Boolean) As String

    On Error GoTo ErrHandler

    Dim I                   As Integer
    Dim RFnd                As Boolean
    Dim RS                  As ADODB.Recordset
    Dim Tmp_DescrField()    As String
    Dim Tmp_Str             As String
    Dim CanDisconnect       As Boolean
    
    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If

    Tmp_Str = Mid$(SQLString, 8, InStr(8, SQLString, "FROM") - 8)
    Tmp_DescrField = Split(Tmp_Str, ", ")

    Set RS = DBConn.Execute(SQLString)

    With RS
        If InStrRev(SQLString, "%") > 0 Or MultiRow Then
            RFnd = (.RecordCount > 0)
        Else
            RFnd = (.RecordCount = 1)
        End If
        
        If RFnd Then
            Dim IValue As Integer
            
            If Trim$(Tmp_DescrField(0)) = "*" Then ReDim Tmp_DescrField(RS.Fields.Count - 1)
            
            For I = 0 To UBound(Tmp_DescrField)
                If Tmp_DescrField(I) <> "" Then
                    IValue = InStr(1, Tmp_DescrField(I), "AS")
                    
                    If IValue > 0 Then
                        IValue = IValue + 3
                        
                        Tmp_DescrField(I) = Mid$(Tmp_DescrField(I), IValue, InStr(IValue, Tmp_DescrField(I), " ") - IValue)
                    End If
                    
                    Tmp_DescrField(I) = Trim$(Tmp_DescrField(I))
                Else
                    Tmp_DescrField(I) = RS(I).Name
                End If

                DB_GetValueByID = DB_GetValueByID & IIf(Not IsNull(RS(Tmp_DescrField(I))), RS(Tmp_DescrField(I)), "") & "|"
            Next I
'        Else
'            GoTo ErrHandler
        End If
        
        .Close
    End With

    GoSub CleanUp

    If DB_GetValueByID <> "" Then DB_GetValueByID = Left$(DB_GetValueByID, Len(DB_GetValueByID) - 1)

    Exit Function

CleanUp:
    Set RS = Nothing

    If CanDisconnect Then DBConn.Close
Return

ErrHandler:
    GoSub CleanUp

    If Err.Description <> "" Then
        DB_GetValueByID = "Error"
        
        MsgBox Err.Description, vbExclamation, "GetValueById:"
    End If

End Function
