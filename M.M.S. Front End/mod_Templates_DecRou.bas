Attribute VB_Name = "mod_Templates_DecRou"
Option Explicit

Public Sub GUI_Templates_CLR()

    With frm_Main
        .lvw_TemplatesRef.ListItems.Clear
        .txt_TemplatesRefDescr.Text = ""
        .cmb_TemplatesRefField.ListIndex = 0
        .txt_TemplatesRefFValue.Text = ""
        
        .lvw_TemplatesDetails.ListItems.Clear
    End With

End Sub

Public Sub GUI_TemplRefCmdsEnabler(ByVal BValue As Boolean)
    
    With frm_Main
        If .lvw_TemplatesRef.ListItems.Count = 0 Then Exit Sub
        
        .fme_ProjectsList.Enabled = BValue
        
        .lvw_TemplatesRef.Enabled = BValue
        .cmd_TemplatesRefDEL.Enabled = BValue
        
        If BValue Then
            .txt_TemplatesRefDescr.Text = ""
            .cmb_TemplatesRefField.ListIndex = 0
        Else
            .txt_TemplatesRefDescr.Text = .lvw_TemplatesRef.SelectedItem.Text
            If Trim$(.lvw_TemplatesRef.SelectedItem.SubItems(1)) <> "" Then .cmb_TemplatesRefField.Text = .lvw_TemplatesRef.SelectedItem.SubItems(1)
            If Trim$(.lvw_TemplatesRef.SelectedItem.SubItems(2)) <> "" Then .txt_TemplatesRefFValue.Text = .lvw_TemplatesRef.SelectedItem.SubItems(2)
        End If
        
        .cmd_TemplatesRefAM.Caption = IIf(BValue, "A", "M")
    End With

End Sub

Public Sub lvw_TemplatesDetails_INIT()
    
    With frm_Main
        With .lvw_TemplatesDetails
            .FullRowSelect = True
            .HideSelection = False
            .LabelEdit = lvwManual
            .View = lvwReport
            
            With .ColumnHeaders
                .Clear
    
                .Add , , "File Name"
                .Add , , "Sheets", , vbCenter
            End With
        End With
    
        lvw_Autosize .lvw_TemplatesDetails, lvwMAX
    End With

End Sub

Public Sub lvw_TemplatesReferences_INIT()
    
    With frm_Main
        With .lvw_TemplatesRef
            .FullRowSelect = True
            .HideSelection = False
            .LabelEdit = lvwManual
            .View = lvwReport
            
            With .ColumnHeaders
                .Clear
    
                .Add , , "Description"
                .Add , , "Field"
                .Add , , "Value", , vbCenter
            End With
        End With
    
        lvw_Autosize .lvw_TemplatesRef, lvwMAX
    End With

End Sub


