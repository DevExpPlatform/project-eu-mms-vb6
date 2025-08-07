Attribute VB_Name = "mod_SubProjects_DBRtns"
Option Explicit

Public Sub DB_SubPrjINFO_SELECT()

    On Error GoTo ErrHandler

    Dim CanDisconnect   As Boolean
    Dim RS              As ADODB.Recordset

    If DBConn.State = 0 Then
        DBConn.Open
        
        CanDisconnect = True
    End If
    
    Set RS = DBConn.Execute("SELECT * FROM edt_SubProjects WHERE id_SubProject = " & frm_Main.SelectedNode)
    
    If RS.RecordCount > 0 Then
        With frm_Main
            .txt_SubPrjDescr.Text = RS("descr_SubProject")
            .txt_SubPrjWorkDir.Text = RS("str_SubPrjWorkDir")
            .txt_SubPrjBaseFName.Text = RS("str_BaseFileName")
            
            ' Query Filters
            '
            'If IsNull(RS("str_QueryFilter")) Then
            '    .cmb_SubPrjFilterField.ListIndex = 0
            '    .txt_SubPrjFilterFieldValue.Text = ""
            'Else
            '    SubPrj_GetQueryFilterData RS("str_QueryFilter")
            'End If
        
            ' Customer File Organization
            '
            If IsNull(RS("str_CustomerFileOrganizer")) Then
                .lvw_SubPrjCFOrganizer.ListItems.Clear
                lvw_Autosize .lvw_SubPrjCFOrganizer, lvwCONTROL
            Else
                SubPrj_GetCFOData RS("str_CustomerFileOrganizer")
            End If
            
            .chk_SubPrjGenOMR.Value = Abs(RS("flg_OMRGen"))
            .chk_SubPrjMakePackages.Value = Abs(RS("flg_MakePackages"))
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

Public Function DB_SubPrjSettings_AM() As Boolean

    On Error GoTo ErrHandler
    
    Dim flds_CFO            As String
    'Dim flds_QueryFilter    As String
    Dim I                   As Integer
    Dim RValue              As Boolean
    Dim SQLString           As String
    
    With frm_Main
        For I = 1 To .lvw_SubPrjCFOrganizer.ListItems.Count
            With .lvw_SubPrjCFOrganizer.ListItems(I)
                flds_CFO = flds_CFO & IIf(flds_CFO = "", "", "|") & .Text & ";" & IIf(.SubItems(1) = "DESC", "DESC", "") & ";" & .SubItems(2) & ";" & IIf(.SubItems(3) = "D", 1, 0)
            End With
        Next I
        
        'If .cmb_SubPrjFilterField.ListIndex > 0 And Trim$(.txt_SubPrjFilterFieldValue.Text) <> "" Then
        '    flds_QueryFilter = .cmb_SubPrjFilterField.Text & ";" & .txt_SubPrjFilterFieldValue.Text
        'End If
        
        If .InsertModeSubPrj Then
            Dim SubProjectOrder As Byte
        
            SubProjectOrder = DB_GetValueByID("SELECT MAX(nmr_SubProjectOrder) AS SubProjectOrder FROM edt_SubProjects WHERE id_SubProject = " & .SelectedNode) + 1

            SQLString = "INSERT INTO edt_SubProjects (id_Project, nmr_SubProjectOrder, descr_SubProject, str_SubPrjWorkDir, str_BaseFileName, str_CustomerFileOrganizer, flg_OMRGen, flg_MakePackages) VALUES(" & _
                        .SelectedPrj & ", " & _
                        SubProjectOrder & ", " & _
                        Conv_String2SQLString(.txt_SubPrjDescr.Text) & ", " & _
                        Conv_String2SQLString(.txt_SubPrjWorkDir.Text) & ", " & _
                        Conv_String2SQLString(.txt_SubPrjBaseFName.Text) & ", " & _
                        Conv_String2SQLString(flds_CFO) & ", " & _
                        .chk_SubPrjGenOMR.Value & ", " & _
                        .chk_SubPrjMakePackages.Value & ")"
        Else
            SQLString = "UPDATE edt_SubProjects SET " & _
                        "descr_SubProject = " & Conv_String2SQLString(.txt_SubPrjDescr.Text) & ", " & _
                        "str_SubPrjWorkDir = " & Conv_String2SQLString(.txt_SubPrjWorkDir.Text) & ", " & _
                        "str_BaseFileName = " & Conv_String2SQLString(.txt_SubPrjBaseFName.Text) & ", " & _
                        "str_CustomerFileOrganizer = " & Conv_String2SQLString(flds_CFO) & ", " & _
                        "flg_OMRGen = " & .chk_SubPrjGenOMR.Value & ", " & _
                        "flg_MakePackages = " & .chk_SubPrjMakePackages.Value & _
                        " WHERE id_SubProject = " & .SelectedNode
        End If
        
        ' Write to DB
        '
        DBConn.Open
        DBConn.BeginTrans
            
        DBConn.Execute SQLString, RValue
        
        If RValue Then
            DBConn.CommitTrans
        Else
            DBConn.RollbackTrans
        End If
        
        DBConn.Close
        
        DB_SubPrjSettings_AM = RValue
    End With

    Exit Function

ErrHandler:
    If DBConn.State = 1 Then
        DBConn.RollbackTrans
        DBConn.Close
    End If
    
    MsgBox Purge_ErrDescr(Err.Description), vbExclamation, IIf(frm_Main.InsertModeSubPrj, "Insert", "Update") & " Progetto:"

End Function

