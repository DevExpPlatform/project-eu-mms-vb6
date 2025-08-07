Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Private Type strctAppSettings
    PrjRenderMode   As Integer
End Type

Public AppSettings  As strctAppSettings

Public Function chk_Array(ByVal myArray As Variant) As Boolean
    
    On Error GoTo ErrHandler
    
    If UBound(myArray) > -1 Then chk_Array = True
    
    Exit Function

ErrHandler:

End Function

Public Function cmb_GetListIndex(ByVal myCombo As ComboBox, ByVal SrchText As String)

    Dim I           As Integer
    Dim SplitData() As String
    
    SplitData = Split(myCombo.Tag, "|")

    For I = 0 To UBound(SplitData)
        If SplitData(I) = SrchText Then
            cmb_GetListIndex = I
        
            Exit For
        End If
    Next I
    
    Erase SplitData

End Function

Public Function cmb_GetTagValue(ByVal myCombo As ComboBox, Optional ByVal RetNumeric As Boolean, Optional ByVal RetNumericNULL As Boolean = False) As String

    Dim SplitData() As String
    
    SplitData = Split(myCombo.Tag, "|")
    
    cmb_GetTagValue = SplitData(myCombo.ListIndex)
        
    If RetNumeric Then
        If cmb_GetTagValue = "NULL" And RetNumericNULL = False Then cmb_GetTagValue = 0
    Else
        If cmb_GetTagValue <> "NULL" Then cmb_GetTagValue = Conv_String2SQLString(cmb_GetTagValue)
    End If
    
    Erase SplitData

End Function

Public Function Conv_Str2Num(InputNum As String, Optional ByVal ZeroReturn As Boolean = False) As String

    If InputNum <> "" Then
        Dim FndStr As Integer
        
        InputNum = CDbl(InputNum)
        
        FndStr = InStr(1, InputNum, ",")
        
        If FndStr > 0 Then
            Mid$(InputNum, FndStr, 1) = "."
        End If
        
        If (InputNum = 0 And ZeroReturn) Then
            Conv_Str2Num = 0
        Else
            Conv_Str2Num = IIf(InputNum <> 0, InputNum, "NULL")
        End If
    Else
        Conv_Str2Num = "NULL"
    End If

End Function

Public Function Conv_String2SQLString(ByVal InputTXT As String) As String

    If Trim$(InputTXT) <> "" Then
        Dim ChrCntr As Integer
        
        Do
            ChrCntr = InStr(ChrCntr + 1, InputTXT, "'")
        
            If ChrCntr > 0 Then
                InputTXT = Left$(InputTXT, ChrCntr) & "'" & Right$(InputTXT, Len(InputTXT) - ChrCntr)
                
                ChrCntr = ChrCntr + 1
            End If
        Loop Until ChrCntr = 0
    
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

Public Sub NumOnlyFilter(ASCIIKey As Integer, NText As String)
    
    Select Case ASCIIKey
        Case 46
            ASCIIKey = 44
        
        Case Else
            If (Not IsNumeric(Chr$(ASCIIKey))) And _
                ASCIIKey <> 8 And _
                ASCIIKey <> 44 And _
                ASCIIKey <> 45 And _
                ASCIIKey <> 47 Then
                
                ASCIIKey = 0
            End If
    
    End Select
    
    If ASCIIKey = 44 Or ASCIIKey = 45 Or ASCIIKey = 47 Then
        If InStr(1, NText, Chr$(ASCIIKey)) > 0 Or Len(NText) = 0 Then
            ASCIIKey = 0
            
            Exit Sub
        End If
    End If

End Sub

Public Function Purge_ErrDescr(ErrDescr As String) As String
    
    Purge_ErrDescr = Mid$(ErrDescr, InStrRev(ErrDescr, "]") + 1, Len(ErrDescr))
    
End Function

