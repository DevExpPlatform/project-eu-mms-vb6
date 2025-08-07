Attribute VB_Name = "mod_PlugInsManager_DecRou"
Option Explicit

Public Function PlugIn_Load(ByVal myPlugIn As String, myPlugVar As Object) As Boolean

    On Error GoTo ErrHandler
    
    Set myPlugVar = CreateObject(myPlugIn & ".PlugIn")
            
    PlugIn_Load = True

    Exit Function

ErrHandler:
    If (DLLParams.UnattendedMode) Then
        UMErrMsg = Err.Description
    Else
        MsgBox Err.Description, vbExclamation, "Attenzione:"
    End If

End Function

Public Function PlugIns_Load(ByVal PlugFilter As String) As String()

    Dim I               As Byte
    Dim PlugCntr        As Byte
    Dim PlugIn          As Object
    Dim PlugLoad()      As String
    Dim ProgId          As String
    Dim extPlugIns()    As String
    
    extPlugIns = Get_FolderFiles(AppPath & "PlugIns\", "*dll")

    If chk_Array(extPlugIns) Then
        For I = 0 To UBound(extPlugIns)
            ProgId = Get_BaseName(extPlugIns(I), 4) & ".PlugIn"

            On Error Resume Next

'RtnToLoad:
            Set PlugIn = CreateObject(ProgId)
            
            If Err = 0 Then
                Dim PlugInfo As Variant
                
                PlugInfo = PlugIn.PlugIn_GetInfo
            
                If PlugInfo.Type = PlugFilter Then
                    ReDim Preserve PlugLoad(PlugCntr)
                        
                    PlugLoad(PlugCntr) = PlugInfo.id & "|" & PlugInfo.Description
                            
                    PlugCntr = PlugCntr + 1
                End If
            Else
                MsgBox "Errore durante il loading del PlugIn " & Get_BaseName(extPlugIns(I), 4) & "." & vbNewLine & vbNewLine & _
                        Err.Description, vbExclamation, "Attenzione:"

                'Select Case Err.Number
                '    Case 429
                '        If RegServer(extPlugIns(I)) Then GoTo RtnToLoad
                '
                'End Select
            End If
        
            Set PlugIn = Nothing
        Next I
    End If
    
    If chk_Array(PlugLoad) Then PlugIns_Load = PlugLoad
    
    Erase extPlugIns

End Function
