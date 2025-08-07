Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Public Enum TextPadJustify
    PADLEFT
    PADRIGHT
    PADCENTER
End Enum

' Public LogFile As String

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

Public Function Get_BaseName(ByVal myPath As String, Optional ByVal CutExt As Byte = 0, Optional ByVal myStrComp As String = "\") As String
    
    Get_BaseName = Right$(myPath, Len(myPath) - InStrRev(myPath, myStrComp))
    
    If CutExt > 0 Then
        Get_BaseName = Left$(Get_BaseName, Len(Get_BaseName) - CutExt)
    End If
    
End Function

Public Function Get_LogFileName(ByVal varFile As String) As String

    If (Mid$(varFile, Len(varFile) - 3, 1) = ".") Then varFile = Left$(varFile, Len(varFile) - 4)
    
    Get_LogFileName = varFile & "_MMS.LOG"

End Function

Public Function Get_PathName(ByVal myPath As String, Optional ByVal myStrComp As String = "\") As String
    
    Get_PathName = Left$(myPath, InStrRev(myPath, myStrComp))
    
End Function

Public Function TextPad(ByVal Mode As TextPadJustify, ByVal myString As String, ByVal NumChar As Integer, ByVal String2Repeat As String, ByVal AutoTrim As Boolean) As String

    Dim TmpString As String

    If AutoTrim Then myString = Trim$(myString)

    If Len(myString) < NumChar Then
        TmpString = String$(((NumChar - Len(myString)) \ IIf(Mode = 2, 2, 1)), String2Repeat)
        
        Select Case Mode
            Case PADLEFT
                TextPad = myString & TmpString
            
            Case PADRIGHT
                TextPad = TmpString & myString
            
            Case PADCENTER
                TextPad = TmpString & IIf((Len(myString) And 1), String2Repeat, "") & myString & TmpString
        
        End Select
    Else
        TextPad = Left$(myString, NumChar)
    End If

End Function

Public Sub Wait(ByVal PauseTime As Integer)

    Dim Start As Long
   
    Start = Timer
    
    Do While Timer < (Start + PauseTime)
       DoEvents
    Loop

End Sub

'Public Sub Write2Log(ByVal Descr_Err As String)
'
'    Dim LogFN As Integer
'
'    LogFN = FreeFile(0)
'
'    Open LogFile For Append As #LogFN
'        Print #LogFN, Descr_Err
'    Close #LogFN
'
'End Sub

