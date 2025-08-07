Attribute VB_Name = "mod_Projects_DecRou"
Option Explicit

Public Sub cmb_PrjNormalizer_INIT(ByRef Normalizers() As String)
        
    If chk_Array(Normalizers) Then
        Dim I           As Byte
        Dim SplitData() As String
            
        With frm_Main.cmb_PrjNormalizer
            .AddItem "Nessuno"
            .Tag = "NULL|"
    
            For I = 0 To UBound(Normalizers)
                SplitData = Split(Normalizers(I), "|")
                
                .AddItem SplitData(1)
                .Tag = .Tag & SplitData(0) & "|"
            Next I
        
            .ListIndex = 0
        End With
        
        Erase SplitData
    End If

End Sub

Private Sub GUI_PrjFrontEnd_CLR()

    With frm_Main
        .txt_PrjDescr.Text = ""
        .txt_PrjWrkDir.Text = ""
        .txt_PrjJobId.Text = ""
        .txt_PrjRefTable.Text = ""
        .txt_PrjWeight.Text = "5"
        
        .cmb_PrjDataCutter.ListIndex = 0
        .cmb_PrjSerializeMode.ListIndex = 1
        .chk_PrjPstlExtraSort.Value = 0
        
        .cmb_PrjBarCodeType.ListIndex = 0
        .chk_PrjShowBarCodeTxt.Value = 0
        
        .cmb_PrjProduct.ListIndex = 0
        .chk_PrjWeightTolerance.Value = 0
    End With

End Sub

Public Sub GUI_PrjInsertCmds(ByVal BValue As Boolean)
    
    With frm_Main
        .fme_ProjectsList.Enabled = (Not BValue)

        If BValue Then
            GUI_PrjFrontEnd_CLR
            GUI_SubPrjFrontEnd_CLR
 
            .cmd_PrjSettingsNew.Caption = "Esci"
            .cmd_PrjSettingsAM.Caption = "Inserisci"
            
            .cmd_SubPrjSettingsNew.Enabled = (.SelectedNode <> "")
        Else
            .cmd_PrjSettingsNew.Caption = "Nuovo"
            .cmd_PrjSettingsAM.Caption = "Modifica"
 
            DB_PrjINFO_SELECT
            
            If .SelectedNode <> "" Then
                DB_SubPrjINFO_SELECT
                ' DB_Templates_SELECListView
            End If
        End If
     
        .cmd_SubPrjSettingsNew.Enabled = (Not BValue) And (.SelectedNode <> "")
        .cmd_SubPrjSettingsAM.Enabled = (Not BValue)
     
     '   .fme_SubPrjTemplatesConsolle.Enabled = (Not BValue)

        .InsertModePrj = BValue
    End With

End Sub

Public Sub GUI_PrjSortCmdsEnabler(ByVal BValue As Boolean)
    
    With frm_Main
        If .lvw_PrjSortFields.ListItems.Count = 0 Then Exit Sub
        
        .fme_ProjectsList.Enabled = BValue
        
        .lvw_PrjSortFields.Enabled = BValue
        .cmd_PrjSortFieldUD(0).Enabled = BValue
        .cmd_PrjSortFieldUD(1).Enabled = BValue
        .cmd_PrjSortFieldDEL.Enabled = BValue
        
        If BValue Then
            .cmb_PrjSortField.ListIndex = 0
            .cmb_PrjSortFieldMode.ListIndex = 0
        Else
            .cmb_PrjSortField.Text = .lvw_PrjSortFields.SelectedItem.Text
            .cmb_PrjSortFieldMode.Text = .lvw_PrjSortFields.SelectedItem.SubItems(1)
        End If
        
        .cmd_PrjSortFieldAM.Caption = IIf(BValue, "A", "M")
    End With

End Sub

Public Sub lvw_PrjSortFields_INIT()
    
    With frm_Main
        With .lvw_PrjSortFields
            .FullRowSelect = True
            .HideSelection = False
            .LabelEdit = lvwManual
            .View = lvwReport
            
            With .ColumnHeaders
                .Clear
    
                .Add , , "Field"
                .Add , , "Mode"
            End With
        End With
        
        lvw_Autosize .lvw_PrjSortFields, lvwCONTROL
        
        With .cmb_PrjSortFieldMode
            .AddItem "ASC"
            .AddItem "DESC"
            
            .ListIndex = 0
        End With
    End With

End Sub

Public Sub Prj_GetSortData(ByVal OrderFields As String)

    Dim SplitData() As String

    frm_Main.lvw_PrjSortFields.ListItems.Clear

    SplitData = Split(OrderFields, "|")
            
    If chk_Array(SplitData) Then
        Dim I               As Integer
        Dim myItem          As ListItem
        Dim SubSplitData()  As String
        
        For I = 0 To UBound(SplitData)
            SubSplitData = Split(SplitData(I), ";")
            
            Set myItem = frm_Main.lvw_PrjSortFields.ListItems.Add(, , SubSplitData(0))
            
            myItem.SubItems(1) = IIf(SubSplitData(1) = "", "ASC", SubSplitData(1))
        Next I
    
        Set myItem = Nothing
        
        Erase SubSplitData
        
        lvw_Autosize frm_Main.lvw_PrjSortFields, lvwITEMS
    Else
        lvw_Autosize frm_Main.lvw_PrjSortFields, lvwCONTROL
    End If

    Erase SplitData

End Sub

Public Sub tvw_Projects_INIT()

    With frm_Main.tvw_Projects
        .HideSelection = False
        .ImageList = frm_Main.iml_PrjTree
        .LabelEdit = lvwManual
        .Style = tvwTreelinesPlusMinusPictureText
    End With

End Sub
