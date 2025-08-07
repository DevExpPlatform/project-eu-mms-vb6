Attribute VB_Name = "mod_Projects_DBRtns"
Option Explicit

Public Function DB_SubProject_SELECT(ByVal idProject As Long) As Integer

    On Error GoTo ErrHandler

    Dim RS          As ADODB.Recordset
    
    DB_SubProject_SELECT = -1
    
    DBConn.Open

    Set RS = DBConn.Execute("SELECT id_SubProject FROM edt_SubProjects WHERE id_Project = " & idProject)
    
    If RS.RecordCount > 0 Then
        DB_SubProject_SELECT = RS("id_SubProject")
    End If
    
    DBConn.Close
        
    Exit Function
        
ErrHandler:
    DB_SubProject_SELECT = -1

End Function

Public Sub DB_Projects_SELECT()

    On Error GoTo ErrHandler
    
    frm_Main.cmb_Projects.Clear
    frm_Main.cmb_Workings.Clear
    
    '
    '
    If AppSettings.PrjFilter = "" Then
        MsgBox "Nessun filtro di progetto definito.", vbExclamation, "Attenzione:"
        
        Exit Sub
    End If
    
    '
    '
    Dim I           As Byte
    Dim RS          As ADODB.Recordset
    Dim SplitData() As String
    Dim SQLString   As String
    Dim tmpString   As String
    
    If AppSettings.PrjFilter <> "*" Then
        SplitData = Split(AppSettings.PrjFilter, "|")
    
        If chk_Array(SplitData) Then
            For I = 0 To UBound(SplitData)
                tmpString = tmpString & IIf(tmpString <> "", " OR ", "") & "id_Project = " & SplitData(I)
            Next I
        End If
        
        tmpString = " WHERE " & tmpString
    
        Erase SplitData
    End If
        
    SQLString = "SELECT id_Project, descr_Project FROM view_Projects_FE" & tmpString
    
    DBConn.Open
    
    Set RS = DBConn.Execute(SQLString)
    
    If RS.RecordCount > 0 Then
        With frm_Main.cmb_Projects
            .Clear
            
            Do Until RS.EOF
                .AddItem RS("descr_Project")
                .Tag = .Tag & RS("id_Project") & "|"
                
                RS.MoveNext
            Loop
        
            .ListIndex = 0
        End With
    End If
            
    RS.Close
    
    GoSub CleanUp
    
    Exit Sub

CleanUp:
    Set RS = Nothing
    
    DBConn.Close
Return

ErrHandler:
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Load Gestori:"

    GoSub CleanUp

End Sub
