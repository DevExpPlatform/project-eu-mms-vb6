Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Private Type strct_AppSettings
    DLLParams       As String
    PrjDNS          As String
    PrjFilter       As String
    PrjFolder       As String
    PrjLocale       As String
    PrjRenderMode   As Integer
    PrjTNS          As String
End Type

Public AppPath      As String
Public AppSettings  As strct_AppSettings
Public IsIDE        As Boolean
Public myMMS        As cls_MMS

Public Function AddPreSuff2String(ByVal Mode As Byte, ByVal myString As String, ByVal NumChar As Byte, ByVal String2Repeat As String) As String

    Dim tmpString As String

    myString = Trim$(myString)
    tmpString = String$(NumChar - Len(myString), String2Repeat)

    If Mode = 0 Then
        AddPreSuff2String = tmpString & myString
    Else
        AddPreSuff2String = myString & tmpString
    End If

End Function

Public Function chk_Array(ByVal myArray As Variant) As Boolean
    
    On Error GoTo ErrHandler
    
    If UBound(myArray) > -1 Then chk_Array = True
    
    Exit Function

ErrHandler:

End Function

Public Function chk_IDE() As Boolean
    
    IsIDE = True
    chk_IDE = True

End Function

Public Function cmb_GetTagValue(ByVal myCombo As ComboBox, Optional ByVal RetNumeric As Boolean, Optional ByVal RetNumericNULL As Boolean = False) As String

    Dim SplitData() As String
    
    SplitData = Split(myCombo.Tag, "|")
    
    cmb_GetTagValue = SplitData(myCombo.ListIndex)
        
    If RetNumeric Then
        If cmb_GetTagValue = "NULL" And RetNumericNULL = False Then cmb_GetTagValue = 0
    Else
        If cmb_GetTagValue <> "NULL" Then cmb_GetTagValue = cmb_GetTagValue
    End If
    
    Erase SplitData

End Function

Public Function FDExist(ByVal NomeFD As String, ByVal ChkDir As Boolean) As Boolean

    On Error GoTo ErrHandler

    Dim Dummy As String
    
    If NomeFD = "" Then Exit Function

    If ChkDir Then
        Dummy = Dir$(NomeFD, vbDirectory)
    Else
        Dummy = Dir$(NomeFD)
    End If

    If (Len(Dummy) > 0) And (Err = 0) Then FDExist = True
   
ErrHandler:

End Function

Public Function Fix_Paths(ByVal myPath As String, Optional ByVal myStrComp As String = "\") As String

    Fix_Paths = IIf(Right$(myPath, 1) = myStrComp, myPath, myPath & myStrComp)
    
End Function

Public Sub GetPrjWorkings()
        
    Dim PrjWorkings() As String

    frm_Main.cmb_Workings.Clear
    frm_Main.cmb_Workings.Tag = ""

    PrjWorkings = myMMS.GetWorkings
    
    If chk_Array(PrjWorkings) Then
        If PrjWorkings(0) <> "Error" Then
            Dim I       As Integer
            Dim tmpStr  As String
            
            With frm_Main.cmb_Workings
                For I = 0 To UBound(PrjWorkings)
                    tmpStr = Mid$(PrjWorkings(I), 7, 2) & "/" & Mid$(PrjWorkings(I), 5, 2) & "/" & Mid$(PrjWorkings(I), 1, 4) & " - " & Mid$(PrjWorkings(I), 9, 2) & "." & Mid$(PrjWorkings(I), 11, 2) & "." & Mid$(PrjWorkings(I), 13, 2)
                    
                    .AddItem (I + 1) & ". " & tmpStr
                    .Tag = .Tag & PrjWorkings(I) & "|"
                Next I
            
                .ListIndex = .ListCount - 1
            End With
        End If
    End If
    
    Erase PrjWorkings

End Sub

Public Sub GUI_MenuAdminEnabler(ByVal BValue As Boolean)
        
    With frm_Main
        .mnu_BaseWorkDir.Visible = BValue
        .mnu_PrjFilter.Visible = BValue
        .mnu_Space00.Visible = BValue
        .mnu_SetDNS.Visible = BValue
        .mnu_SetTNS.Visible = BValue
        .mnu_Space01.Visible = BValue
    End With

End Sub

Public Function Purge_ErrDescr(ErrDescr As String) As String
    
    Purge_ErrDescr = Mid$(ErrDescr, InStrRev(ErrDescr, "]") + 1, Len(ErrDescr))
    
End Function

