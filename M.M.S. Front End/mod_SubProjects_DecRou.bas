Attribute VB_Name = "mod_SubProjects_DecRou"
Option Explicit

Public Sub GUI_SubPrjCFOCmdsEnabler(ByVal BValue As Boolean)
    
    With frm_Main
        If .lvw_PrjSortFields.ListItems.Count = 0 Then Exit Sub
        
        .fme_ProjectsList.Enabled = BValue
        
        .lvw_SubPrjCFOrganizer.Enabled = BValue
        .cmd_SubPrjCFOFieldUD(0).Enabled = BValue
        .cmd_SubPrjCFOFieldUD(1).Enabled = BValue
        .cmd_SubPrjCFOFieldDEL.Enabled = BValue
        
        If BValue Then
            .cmb_SubPrjCFOField.ListIndex = 0
            .cmb_SubPrjCFOSortMode.ListIndex = 0
            .txt_SubPrjCFOFieldAlias.Text = ""
            .cmb_SubPrjCFOAliasType.ListIndex = 0
        Else
            .cmb_SubPrjCFOField.Text = .lvw_SubPrjCFOrganizer.SelectedItem.Text
            .cmb_SubPrjCFOSortMode.Text = .lvw_SubPrjCFOrganizer.SelectedItem.SubItems(1)
            .txt_SubPrjCFOFieldAlias.Text = .lvw_SubPrjCFOrganizer.SelectedItem.SubItems(2)
            .cmb_SubPrjCFOAliasType.Text = .lvw_SubPrjCFOrganizer.SelectedItem.SubItems(3)
        End If
        
        .cmd_SubPrjCFOFieldAM.Caption = IIf(BValue, "A", "M")
    End With

End Sub

Public Sub GUI_SubPrjFrontEnd_CLR()

    With frm_Main
        .txt_SubPrjDescr.Text = ""
        .txt_SubPrjWorkDir.Text = ""
        .txt_SubPrjBaseFName.Text = ""
                    
        .lvw_SubPrjCFOrganizer.ListItems.Clear
        If .cmb_SubPrjCFOField.ListCount > 0 Then .cmb_SubPrjCFOField.ListIndex = 0
        .cmb_SubPrjCFOSortMode.ListIndex = 0
        .txt_SubPrjCFOFieldAlias.Text = ""
        .cmb_SubPrjCFOAliasType.ListIndex = 0
                    
        .chk_SubPrjGenOMR.Value = 0
        .chk_SubPrjMakePackages.Value = 0
    End With

End Sub

Public Sub GUI_SubPrjInsertCmds(ByVal BValue As Boolean)
    
    With frm_Main
        If BValue Then
            GUI_SubPrjFrontEnd_CLR
            GUI_Templates_CLR
            
            .cmd_SubPrjSettingsNew.Caption = "Esci"
            .cmd_SubPrjSettingsAM.Caption = "Inserisci"
        Else
            DB_SubPrjINFO_SELECT
            DB_TemplatesReferences_SELECListView
            
            .cmd_SubPrjSettingsNew.Caption = "Nuovo"
            .cmd_SubPrjSettingsAM.Caption = "Modifica"
        End If
        
        .cmd_SubPrjSettingsNew.Enabled = (.SelectedNode <> "")
        .sst_GeneralSettings.TabEnabled(2) = (Not BValue)
        
        .InsertModeSubPrj = BValue
    End With

End Sub

Public Sub SubPrj_GetCFOData(ByVal CustomerFileOrganizer As String)
    
    Dim SplitData() As String
                
    frm_Main.lvw_SubPrjCFOrganizer.ListItems.Clear
    
    SplitData = Split(CustomerFileOrganizer, "|")
    
    If chk_Array(SplitData) Then
        Dim I               As Integer
        Dim myItem          As ListItem
        Dim SubSplitData()  As String
        
        For I = 0 To UBound(SplitData)
            SubSplitData = Split(SplitData(I), ";")
                                    
            Set myItem = frm_Main.lvw_SubPrjCFOrganizer.ListItems.Add(, , SubSplitData(0))
            
            myItem.SubItems(1) = IIf(SubSplitData(1) = "", "ASC", SubSplitData(1))
            myItem.SubItems(2) = SubSplitData(2)
            myItem.SubItems(3) = IIf(SubSplitData(3) = 0, "F", "D")
        Next I
                
        Set myItem = Nothing
        
        Erase SubSplitData
        
        lvw_Autosize frm_Main.lvw_SubPrjCFOrganizer, lvwITEMS
    Else
        lvw_Autosize frm_Main.lvw_SubPrjCFOrganizer, lvwCONTROL
    End If

    Erase SplitData

End Sub

'Public Sub SubPrj_GetQueryFilterData(ByVal QueryFilter As String)
'
'    Dim SplitData() As String
'
'    SplitData = Split(QueryFilter, ";")
'
'    If chk_Array(SplitData) Then
'        frm_Main.cmb_SubPrjFilterField.Text = SplitData(0)
'        frm_Main.txt_SubPrjFilterFieldValue.Text = SplitData(1)
'    End If
'
'    Erase SplitData
'
'End Sub

Public Sub lvw_SubPrjCFO_INIT()
    
    With frm_Main
        With .lvw_SubPrjCFOrganizer
            .FullRowSelect = True
            .HideSelection = False
            .LabelEdit = lvwManual
            .View = lvwReport
            
            With .ColumnHeaders
                .Clear
    
                .Add , , "Field"
                .Add , , "Mode"
                .Add , , "Alias"
                .Add , , "Type", , vbCenter
            End With
        End With
    
        lvw_Autosize .lvw_SubPrjCFOrganizer, lvwCONTROL
    
        With .cmb_SubPrjCFOSortMode
            .AddItem "ASC"
            .AddItem "DESC"
        
            .ListIndex = 0
        End With
        
        With .cmb_SubPrjCFOAliasType
            .AddItem "D"
            .AddItem "F"
                
            .ListIndex = 0
        End With
    End With

End Sub
