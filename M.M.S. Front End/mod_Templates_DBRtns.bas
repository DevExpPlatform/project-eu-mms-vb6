Attribute VB_Name = "mod_Templates_DBRtns"
Option Explicit

Public Sub DB_TemplatesDetails_SELECListView()

    On Error GoTo ErrHandler

    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset

    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
    
    Set RS = DBConn.Execute("SELECT id_TemplateDef, str_TemplateFileName, nmr_Sheets FROM edt_Templates WHERE id_Template = " & frm_Main.lvw_TemplatesRef.SelectedItem.Tag & " ORDER BY nmr_TemplateOrder")
    
    With frm_Main.lvw_TemplatesDetails
        .ListItems.Clear
        
        If RS.RecordCount > 0 Then
            Dim myItem As ListItem
            
            Do Until RS.EOF
                Set myItem = .ListItems.Add(, , RS("str_TemplateFileName"))
                
                With myItem
                    .Tag = RS("id_TemplateDef")
                    .SubItems(1) = RS("nmr_Sheets")
                End With
                
                RS.MoveNext
            Loop
            
            Set myItem = Nothing
        
            lvw_Autosize frm_Main.lvw_TemplatesDetails, lvwITEMS
        Else
            lvw_Autosize frm_Main.lvw_TemplatesDetails, lvwCONTROL
        End If
    End With
            
    GoSub CleanUp
    
    Exit Sub

CleanUp:
    Set RS = Nothing

    If CanDisconnect Then DBConn.Close
Return

ErrHandler:
    GoSub CleanUp
    
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Templates References:"

End Sub

Public Sub DB_TemplatesReferences_SELECListView()

    On Error GoTo ErrHandler

    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset

    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
    
    Set RS = DBConn.Execute("SELECT id_Template, descr_Template, str_QField, str_QValue FROM ref_Templates WHERE id_SubProject = " & frm_Main.SelectedNode & " ORDER BY str_QField, str_QValue")
    
    With frm_Main.lvw_TemplatesRef
        .ListItems.Clear
        
        If RS.RecordCount > 0 Then
            Dim myItem As ListItem
            
            Do Until RS.EOF
                Set myItem = .ListItems.Add(, , RS("descr_Template"))
                
                With myItem
                    .Tag = RS("id_Template")
                    
                    If IsNull(RS("str_QField")) Then
                        .SubItems(1) = " "
                        .SubItems(2) = " "
                    Else
                        .SubItems(1) = RS("str_QField")
                        .SubItems(2) = RS("str_QValue")
                    End If
                End With
                
                RS.MoveNext
            Loop
            
            Set myItem = Nothing
        
            lvw_Autosize frm_Main.lvw_TemplatesRef, lvwITEMS
        
            DB_TemplatesDetails_SELECListView
        Else
            lvw_Autosize frm_Main.lvw_TemplatesRef, lvwCONTROL
        End If
    End With
            
    GoSub CleanUp
    
    Exit Sub

CleanUp:
    Set RS = Nothing

    If CanDisconnect Then DBConn.Close
Return

ErrHandler:
    GoSub CleanUp
    
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Templates References:"

End Sub
