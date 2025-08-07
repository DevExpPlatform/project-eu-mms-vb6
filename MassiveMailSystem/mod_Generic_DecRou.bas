Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Private Type strct_DLLParams
    AutoMergePacks      As Boolean
    BaseWorkDir         As String
    DSN                 As String
    TSN                 As String
    UnattendedMode      As Boolean
End Type

Public AppPath          As String
Public DLLParams        As strct_DLLParams
Public myAPB            As cls_APB
Public UMErrMsg         As String

Public Function chk_Array(ByVal myArray As Variant) As Boolean
    
    On Error GoTo ErrHandler
    
    If UBound(myArray) > -1 Then chk_Array = True
    
    Exit Function

ErrHandler:

End Function

Public Sub chk_Directory(ByVal DirName As String, ByVal mode As Byte)
    
    If FDExist(DirName, True) Then
        If (mode And 1) Then EmptyDir DirName
    Else
        If (mode And 2) Then MkDir DirName
    End If
    
End Sub

Public Function Conv_Name2ConventionalName(ByVal InputString As String) As String

    Dim I           As Byte
    Dim SrchStr(7)  As String
    
    SrchStr(0) = "/"
    SrchStr(1) = ":"
    SrchStr(2) = "*"
    SrchStr(3) = "?"
    SrchStr(4) = "<"
    SrchStr(5) = ">"
    SrchStr(6) = "|"
    SrchStr(7) = "."
    
    'SrchStr(0) = "\"
    'SrchStr(8) = "-"
    'SrchStr(9) = "+"
    'SrchStr(11) = " "
    
    For I = 0 To 7
        InputString = Replace$(InputString, SrchStr(I), " ")
    Next I
    
    For I = 0 To 2
        InputString = Replace$(InputString, "  ", " ")
    Next I
        
    Conv_Name2ConventionalName = Trim$(InputString)
    
End Function

Public Function Conv_String2SQLString(ByVal InputTXT As String) As String

    InputTXT = Trim$(InputTXT)

    If InputTXT <> "" Then
        InputTXT = Replace$(InputTXT, "'", "''")
        
        Conv_String2SQLString = "'" & InputTXT & "'"
    Else
        Conv_String2SQLString = "NULL"
    End If

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

Public Function Get_BaseName(ByVal myPath As String, Optional ByVal CutExt As Byte = 0, Optional ByVal myStrComp As String = "\") As String
    
    Get_BaseName = Right$(myPath, Len(myPath) - InStrRev(myPath, myStrComp))
    
    If CutExt > 0 Then
        Get_BaseName = Left$(Get_BaseName, Len(Get_BaseName) - CutExt)
    End If
    
End Function

Public Function Get_FolderFiles(ByVal Folder As String, Optional Filter = ".*", Optional retFullPath As Boolean = True) As String()
   
    Dim FileNames() As String
    
    If FDExist(Folder, True) Then
        Dim Extension   As String
        Dim FoundFile   As String
        
        If Left$(Filter, 1) = "*" Then Extension = Mid$(Filter, 2, Len(Filter))
        If Left$(Filter, 1) <> "." Then Filter = "." & Filter
        
        FoundFile = Dir$(Folder & "\*" & Filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
        
        While FoundFile <> ""
            Push_Array FileNames(), IIf(retFullPath = True, Folder & FoundFile, FoundFile)
            
            FoundFile = Dir$()
        Wend
    End If
    
    Get_FolderFiles = FileNames

End Function

Public Function Get_PathName(ByVal myPath As String, Optional ByVal myStrComp As String = "\") As String
    
    Get_PathName = Left$(myPath, InStrRev(myPath, myStrComp))

End Function

Public Function Purge_ErrDescr(ErrDescr As String) As String
    
    Purge_ErrDescr = Mid$(ErrDescr, InStrRev(ErrDescr, "]") + 1, Len(ErrDescr))
    
End Function

Public Sub Push_Array(myArray As Variant, Value As Variant)
    
    On Error GoTo InitArray
    
    Dim ArrayCount As Integer
    
    ArrayCount = UBound(myArray) + 1
    
    ReDim Preserve myArray(ArrayCount)
    
    myArray(ArrayCount) = Value
    
    Exit Sub

InitArray:
    ReDim myArray(0)
    
    myArray(0) = Value

End Sub

Public Sub Sort_Quick(ByRef Arr As Variant, Optional ByVal numEls As Variant, Optional ByVal Descending As Boolean)

    Dim I               As Long
    Dim J               As Long
    Dim LeftNdx         As Long
    Dim LeftStk(32)     As Long
    Dim RightNdx        As Long
    Dim RightStk(32)    As Long
    Dim SP              As Integer
    Dim Temp            As Variant
    Dim Value           As Variant
    
    If IsMissing(numEls) Then numEls = UBound(Arr)
    
    ' Init pointers
    '
    LeftNdx = LBound(Arr)
    RightNdx = numEls
    
    ' Init stack
    '
    SP = 1
    LeftStk(SP) = LeftNdx
    RightStk(SP) = RightNdx

    Do
        If RightNdx > LeftNdx Then
            Value = Arr(RightNdx)
            I = LeftNdx - 1
            J = RightNdx
            
            ' Find the pivot item
            '
            If Descending Then
                Do
                    Do: I = I + 1: Loop Until Arr(I) <= Value
                    Do: J = J - 1: Loop Until J = LeftNdx Or Arr(J) >= Value
                    
                    Temp = Arr(I)
                    Arr(I) = Arr(J)
                    Arr(J) = Temp
                Loop Until J <= I
            Else
                Do
                    Do: I = I + 1: Loop Until Arr(I) >= Value
                    Do: J = J - 1: Loop Until J = LeftNdx Or Arr(J) <= Value
                    
                    Temp = Arr(I)
                    Arr(I) = Arr(J)
                    Arr(J) = Temp
                Loop Until J <= I
            End If
            
            ' Swap found items
            '
            Temp = Arr(J)
            Arr(J) = Arr(I)
            Arr(I) = Arr(RightNdx)
            Arr(RightNdx) = Temp
            
            ' Push on the stack the pair of pointers that differ most
            '
            SP = SP + 1
            
            If (I - LeftNdx) > (RightNdx - I) Then
                LeftStk(SP) = LeftNdx
                RightStk(SP) = I - 1
                LeftNdx = I + 1
            Else
                LeftStk(SP) = I + 1
                RightStk(SP) = RightNdx
                RightNdx = I - 1
            End If
        Else
            ' Pop a new pair of pointers off the stacks
            '
            LeftNdx = LeftStk(SP)
            RightNdx = RightStk(SP)
            SP = SP - 1
            
            If SP = 0 Then Exit Do
        End If
    Loop

End Sub

Public Sub Wait(ByVal PauseTime As Integer)

    Dim start As Long
   
    start = Timer
    
    Do While Timer < (start + PauseTime)
       DoEvents
    Loop

End Sub
