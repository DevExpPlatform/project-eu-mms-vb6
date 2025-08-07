Attribute VB_Name = "mod_Projects_DBRtns"
Option Explicit

Public Sub DB_PrjDataCutter_SELECTCombo()
    
    Dim RS As ADODB.Recordset
    
    DBConn.Open
    
    Set RS = DBConn.Execute("SELECT * FROM ref_DataCutter ORDER BY descr_DataCutter")
    
    If RS.RecordCount > 0 Then
        With frm_Main.cmb_PrjDataCutter
            .AddItem "Sconosciuto"
            .Tag = "NULL|"
            
            Do Until RS.EOF
                .AddItem RS("descr_DataCutter")
                .Tag = .Tag & RS("id_DataCutter") & "|"
                
                RS.MoveNext
            Loop
            
            .ListIndex = 0
        End With
    End If

    RS.Close
    
    Set RS = Nothing
    
    DBConn.Close
    
End Sub

Public Sub DB_PrjDataCutterFields_SELECTCombo(ByVal id_DC As Integer)

    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset

    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
    
    Set RS = DBConn.Execute("SELECT descr_FieldName FROM edt_DataCutter WHERE id_DataCutter = " & id_DC & " ORDER BY nmr_FieldOrder")
    
    If RS.RecordCount > 0 Then
        With frm_Main
            .cmb_PrjPstlField.Clear
            .cmb_PrjSortField.Clear
            .cmb_SubPrjCFOField.Clear
            .cmb_TemplatesRefField.Clear
                
            .cmb_TemplatesRefField.AddItem ""
            
            Do Until RS.EOF
                .cmb_PrjPstlField.AddItem RS("descr_FieldName")
                .cmb_PrjSortField.AddItem RS("descr_FieldName")
                .cmb_TemplatesRefField.AddItem RS("descr_FieldName")
                .cmb_SubPrjCFOField.AddItem RS("descr_FieldName")
                
                RS.MoveNext
            Loop
            
            .cmb_PrjPstlField.ListIndex = 0
            .cmb_PrjSortField.ListIndex = 0
            .cmb_PrjSortFieldMode.ListIndex = 0
            .cmb_TemplatesRefField.ListIndex = 0
            .cmb_SubPrjCFOField.ListIndex = 0
        End With
    End If

    RS.Close
    
    Set RS = Nothing
    
    If CanDisconnect Then DBConn.Close

End Sub

Public Sub DB_PrjINFO_SELECT()

    On Error GoTo ErrHandler

    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset

    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
    
    Set RS = DBConn.Execute("SELECT * FROM edt_Projects WHERE id_Project = " & frm_Main.SelectedPrj)
    
    If RS.RecordCount = 1 Then
        With frm_Main
            ' Generic Parameters
            '
            .txt_PrjDescr.Text = RS("descr_Project")
            .txt_PrjWrkDir.Text = RS("str_PrjWorkDir")
            
            If IsNull(RS("id_JobId")) Then
                .txt_PrjJobId.Text = ""
            Else
                .txt_PrjJobId.Text = RS("id_JobId")
            End If
            
            .txt_PrjRefTable.Text = RS("str_RefTableName")
            .txt_PrjWeight.Text = RS("nmr_PrjWeight")
                        
            ' Data Cutter Section
            '
            If IsNull(RS("id_IDPPlugIn")) Then
                .cmb_PrjNormalizer.ListIndex = 0
            Else
                .cmb_PrjNormalizer.ListIndex = cmb_GetListIndex(.cmb_PrjNormalizer, RS("id_IDPPlugIn"))
            End If
            
            .cmb_PrjDataCutter.ListIndex = cmb_GetListIndex(.cmb_PrjDataCutter, RS("id_DataCutter"))
            .cmb_PrjSerializeMode.ListIndex = IIf(RS("id_SerializeMode") = 1, 0, 1)
            
            If Not IsNull(RS("str_PstlField")) Then .cmb_PrjPstlField.Text = RS("str_PstlField")
            
            .chk_PrjPstlExtraSort.Value = Abs(RS("flg_PstlExtraSort"))
        
            ' BarCode Section
            '
            If IsNull(RS("FLG_BARCODETYPE")) Then
                .cmb_PrjBarCodeType.ListIndex = 0
            Else
                .cmb_PrjBarCodeType.ListIndex = cmb_GetListIndex(.cmb_PrjBarCodeType, RS("FLG_BARCODETYPE"))
                .chk_PrjShowBarCodeTxt.Value = Abs(RS("FLG_BARCODETXT"))
            End If
            
            ' Product Section
            '
            .cmb_PrjProduct.ListIndex = cmb_GetListIndex(.cmb_PrjProduct, RS("id_Product"))
            .chk_PrjWeightTolerance.Value = Abs(RS("flg_ApplyTolerance"))
            
            ' Sort Section
            '
            If IsNull(RS("str_OrderFields")) Then
                .lvw_PrjSortFields.ListItems.Clear
                lvw_Autosize .lvw_PrjSortFields, lvwCONTROL
            Else
                Prj_GetSortData RS("str_OrderFields")
            End If
        End With
    End If
    
    GoSub CleanUp
    
    Exit Sub

CleanUp:
    Set RS = Nothing

    If CanDisconnect Then DBConn.Close
Return

ErrHandler:
    GoSub CleanUp
    
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Project Info:"


End Sub

Public Sub DB_PrjProducts_SELECTCombo()

    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset

    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
    
    Set RS = DBConn.Execute("SELECT id_Product, id_Omologazione FROM edt_Products ORDER BY id_Omologazione")
    
    If RS.RecordCount > 0 Then
        With frm_Main.cmb_PrjProduct
            .Clear
            
            Do Until RS.EOF
                .AddItem RS("id_Omologazione")
                .Tag = .Tag & RS("id_Product") & "|"
                
                RS.MoveNext
            Loop
            
            .ListIndex = 0
        End With
    End If

    RS.Close
    
    Set RS = Nothing
    
    If CanDisconnect Then DBConn.Close

End Sub

Public Function DB_PrjSettings_AM() As Boolean

    On Error GoTo ErrHandler
    
    Dim flds_OrderFields    As String
    Dim I                   As Integer
    Dim RValue              As Boolean
    Dim SplitData           As String
    Dim SQLString           As String
            
    DBConn.Open
    DBConn.BeginTrans
        
    With frm_Main
        For I = 1 To .lvw_PrjSortFields.ListItems.Count
            With .lvw_PrjSortFields.ListItems(I)
                flds_OrderFields = flds_OrderFields & IIf(flds_OrderFields = "", "", "|") & .Text & ";" & IIf(.SubItems(1) = "DESC", "DESC", "")
            End With
        Next I
        
        If .InsertModePrj Then
            SQLString = "INSERT INTO edt_Projects (descr_Project, id_JobId, str_PrjWorkDir, str_RefTableName, nmr_PrjWeight, id_DataCutter, id_IDPPlugIn, id_SerializeMode, str_PstlField, flg_PstlExtraSort, str_OrderFields, flg_BarCodeType, flg_BarCodeTxt, id_Product, flg_ApplyTolerance) VALUES(" & _
                        Conv_String2SQLString(.txt_PrjDescr.Text) & ", " & _
                        Conv_String2SQLString(.txt_PrjJobId) & ", " & _
                        Conv_String2SQLString(.txt_PrjWrkDir.Text) & ", " & _
                        Conv_String2SQLString(.txt_PrjRefTable.Text) & ", " & _
                        Conv_Str2Num(.txt_PrjWeight) & ", " & _
                        cmb_GetTagValue(.cmb_PrjDataCutter, True, True) & ", " & _
                        cmb_GetTagValue(.cmb_PrjNormalizer) & ", " & _
                        cmb_GetTagValue(.cmb_PrjSerializeMode, True) & ", " & _
                        Conv_String2SQLString(IIf(.cmb_PrjPstlField.Enabled, .cmb_PrjPstlField.Text, "")) & ", " & _
                        .chk_PrjPstlExtraSort.Value & ", " & _
                        Conv_String2SQLString(flds_OrderFields) & ", " & _
                        cmb_GetTagValue(.cmb_PrjBarCodeType) & ", " & _
                        .chk_PrjShowBarCodeTxt.Value & ", " & _
                        cmb_GetTagValue(.cmb_PrjProduct, True) & ", " & _
                        .chk_PrjWeightTolerance.Value & ")"
        
            DBConn.Execute SQLString, RValue
        
            If RValue Then
                Dim flds_CFO            As String
                Dim nmr_NewPrjId        As Long
                
                RValue = False
                nmr_NewPrjId = DB_GetLastIdentity("edt_Projects")
                
                For I = 1 To .lvw_SubPrjCFOrganizer.ListItems.Count
                    With .lvw_SubPrjCFOrganizer.ListItems(I)
                        flds_CFO = flds_CFO & IIf(flds_CFO = "", "", "|") & .Text & ";" & IIf(.SubItems(1) = "DESC", "DESC", "") & ";" & .SubItems(2) & ";" & IIf(.SubItems(3) = "D", 1, 0)
                    End With
                Next I
                
                'If .cmb_SubPrjFilterField.ListIndex > 0 And Trim$(.txt_SubPrjFilterFieldValue.Text) <> "" Then
                '    flds_QueryFilter = .cmb_SubPrjFilterField.Text & ";" & .txt_SubPrjFilterFieldValue.Text
                'End If
                
                SQLString = "INSERT INTO edt_SubProjects (id_Project, descr_SubProject, str_SubPrjWorkDir, str_BaseFileName, str_CustomerFileOrganizer, flg_OMRGen, flg_MakePackages) VALUES(" & _
                            nmr_NewPrjId & ", " & _
                            Conv_String2SQLString(.txt_SubPrjDescr.Text) & ", " & _
                            Conv_String2SQLString(.txt_SubPrjWorkDir.Text) & ", " & _
                            Conv_String2SQLString(.txt_SubPrjBaseFName.Text) & ", " & _
                            Conv_String2SQLString(flds_CFO) & ", " & _
                            .chk_SubPrjGenOMR.Value & ", " & _
                            .chk_SubPrjMakePackages.Value & ")"
                
                DBConn.Execute SQLString, RValue
            End If
        Else
            SplitData = DB_GetValueByID("SELECT id_DataCutter FROM edt_Projects WHERE id_Project = " & .SelectedPrj)
            
            If (cmb_GetTagValue(.cmb_PrjDataCutter, True) <> SplitData) Then
                'SQLString = "UPDATE edt_SubProjects SET str_QueryFilter = NULL, str_CustomerFileOrganizer = NULL WHERE id_Project = " & .SelectedPrj
                
                'DBConn.Execute SQLString, RValue
            Else
                RValue = True
            End If
            
            If RValue Then
                RValue = False
                
                SQLString = "UPDATE edt_Projects SET " & _
                            "descr_Project = " & Conv_String2SQLString(.txt_PrjDescr.Text) & ", " & _
                            "id_JobId = " & Conv_String2SQLString(.txt_PrjJobId) & ", " & _
                            "str_PrjWorkDir = " & Conv_String2SQLString(.txt_PrjWrkDir.Text) & ", " & _
                            "str_RefTableName = " & Conv_String2SQLString(.txt_PrjRefTable.Text) & ", " & _
                            "nmr_PrjWeight = " & Conv_Str2Num(.txt_PrjWeight) & ", " & _
                            "id_DataCutter = " & cmb_GetTagValue(.cmb_PrjDataCutter, True, True) & ", " & _
                            "id_IDPPlugIn = " & cmb_GetTagValue(.cmb_PrjNormalizer) & ", " & _
                            "id_SerializeMode = " & cmb_GetTagValue(.cmb_PrjSerializeMode, True) & ", " & _
                            "str_PstlField = " & Conv_String2SQLString(IIf(.cmb_PrjPstlField.Enabled, .cmb_PrjPstlField.Text, "")) & ", " & _
                            "flg_PstlExtraSort = " & IIf(.cmb_PrjSerializeMode.ListIndex = 0, .chk_PrjPstlExtraSort.Value, 0) & ", " & _
                            "str_OrderFields = " & Conv_String2SQLString(flds_OrderFields) & ", " & _
                            "flg_BarCodeType = " & cmb_GetTagValue(.cmb_PrjBarCodeType) & ", " & _
                            "flg_BarCodeTxt = " & .chk_PrjShowBarCodeTxt.Value & ", " & _
                            "id_Product = " & cmb_GetTagValue(.cmb_PrjProduct, True) & ", " & _
                            "flg_ApplyTolerance = " & .chk_PrjWeightTolerance.Value & _
                            " WHERE id_Project = " & .SelectedPrj
                
                DBConn.Execute SQLString, RValue
            End If
        End If
        
        If RValue Then
            DBConn.CommitTrans
        Else
            DBConn.RollbackTrans
        End If
                
        DBConn.Close
                    
        DB_PrjSettings_AM = RValue
    End With

    Exit Function

ErrHandler:
    If DBConn.State = 1 Then
        DBConn.RollbackTrans
        DBConn.Close
    End If
    
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, IIf(frm_Main.InsertModePrj, "Insert", "Update") & " Progetto:"

End Function

Public Sub DB_Projects_SELECTTreeView()
    
    On Error GoTo ErrHandler

    Dim RS  As ADODB.Recordset
    
    frm_Main.tvw_Projects.Nodes.Clear
    
    DBConn.Open
        
    Set RS = DBConn.Execute("SELECT * FROM view_Projects_FE")
    
    If RS.RecordCount > 0 Then
        Dim myNode      As Node
        Dim tmp_Project As String
        Dim tmp_Status  As String
        
        Do Until RS.EOF
            If tmp_Project <> RS("descr_Project") Then
                tmp_Project = RS("descr_Project")
            
                Set myNode = frm_Main.tvw_Projects.Nodes.Add(, , "P" & RS("id_Project"), RS("descr_Project"), 1)
                'myNode.Expanded = True
            End If
                
            Set myNode = frm_Main.tvw_Projects.Nodes.Add("P" & RS("id_Project"), tvwChild, "S" & RS("id_SubProject"), RS("descr_SubProject") & " " & tmp_Status, 2)
            
            RS.MoveNext
        Loop
    
        myNode.Selected = True
        myNode.EnsureVisible
    
        frm_Main.SelectedNode = Right$(myNode.Key, Len(myNode.Key) - 1)
        frm_Main.SelectedPrj = Right$(myNode.Parent.Key, Len(myNode.Parent.Key) - 1)
        
        DB_PrjINFO_SELECT
        DB_SubPrjINFO_SELECT
        DB_TemplatesReferences_SELECListView
    End If
            
    RS.Close
    
    GoSub CleanUp
    
    Exit Sub

CleanUp:
    Set myNode = Nothing
    Set RS = Nothing
    
    DBConn.Close
Return

ErrHandler:
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, "Load Gestori:"

    GoSub CleanUp

End Sub
