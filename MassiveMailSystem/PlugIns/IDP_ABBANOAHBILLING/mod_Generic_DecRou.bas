Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Private Type strct_DLLParams
    CCP_BILL            As String
    CCP_PPA             As String
    CF_ENTE             As String
    DOCMODE             As String
    DSN                 As String
    DSN_EXT             As String
    IDDATACUTTER        As String
    IDWORKINGLOAD       As String
    INPUTFILENAME       As String
    EXTRASPATH          As String
    OUTPUTFILENAME      As String
    PLUGMODE            As String
    TABLENAME           As String
    TNS                 As String
    TEMPLATEORGANIZER   As String
    UNATTENDEDMODE      As Boolean
    TEMPLATEVERSION     As String
End Type

Public Enum TextPadJustify
    PADLEFT
    PADRIGHT
    PADCENTER
End Enum

Public DLLParams        As strct_DLLParams
Public UMErrMsg         As String

Public Function CHK_ARRAY(ByVal myArray As Variant) As Boolean
    
    On Error GoTo ErrHandler
    
    If UBound(myArray) > -1 Then CHK_ARRAY = True
    
    Exit Function

ErrHandler:

End Function

Public Function FDEXIST(NomeFD As String, ChkDir As Boolean) As Boolean

    On Error GoTo ErrHandler

    Dim Dummy As String
    
    If NomeFD = "" Then Exit Function

    If ChkDir Then
        Dummy = Dir$(NomeFD, vbDirectory)
    Else
        Dummy = Dir$(NomeFD)
    End If

    If (Len(Dummy) > 0) And (Err = 0) Then FDEXIST = True
   
ErrHandler:

End Function

Public Function GET_BASENAME(ByVal myPath As String, Optional ByVal CutExt As Boolean, Optional ByVal myStrComp As String = "\") As String
    
    Dim ExtPos As Integer
    
    GET_BASENAME = Right$(myPath, Len(myPath) - InStrRev(myPath, myStrComp))
    
    If (CutExt) Then
        ExtPos = (InStrRev(GET_BASENAME, ".") - 1)
        If (ExtPos > -1) Then GET_BASENAME = Left$(GET_BASENAME, ExtPos)
    End If
    
End Function

Public Function GET_PAGOPACODE(pagoPACode As String)

    Dim I       As Integer
    Dim sShift  As Integer
    Dim lString As Integer
        
    lString = Len(pagoPACode)
    sShift = 4

    For I = 1 To lString Step 4
        GET_PAGOPACODE = GET_PAGOPACODE & Mid$(pagoPACode, I, 4) & " "
    Next I
    
    GET_PAGOPACODE = Trim$(GET_PAGOPACODE)
    
End Function

Public Function GET_PATHNAME(ByVal myPath As String, Optional ByVal myStrComp As String = "\") As String
    
    GET_PATHNAME = Left$(myPath, InStrRev(myPath, myStrComp))

End Function

Public Function GET_TEXTPAD(ByVal Mode As TextPadJustify, ByVal myString As String, ByVal NumChar As Integer, ByVal String2Repeat As String, ByVal AutoTrim As Boolean) As String

    Dim TmpString As String

    If AutoTrim Then myString = Trim$(myString)

    If Len(myString) < NumChar Then
        TmpString = String$(((NumChar - Len(myString)) \ IIf(Mode = 2, 2, 1)), String2Repeat)
        
        Select Case Mode
            Case PADLEFT
                GET_TEXTPAD = myString & TmpString
            
            Case PADRIGHT
                GET_TEXTPAD = TmpString & myString
            
            Case PADCENTER
                GET_TEXTPAD = TmpString & IIf((Len(myString) And 1), String2Repeat, "") & myString & TmpString
        
        End Select
    Else
        GET_TEXTPAD = Left$(myString, NumChar)
    End If

End Function

Function GET_VAL2HEX(ByVal VALUE As Currency) As String
    
    Dim division  As Currency
    Dim remainder As Currency
    
    While VALUE > 0
        division = Int(VALUE / 16)
        remainder = (VALUE - (division * 16))
        VALUE = division
        
        GET_VAL2HEX = Hex$(remainder) & GET_VAL2HEX
    Wend

End Function

Public Function NRM_IMPORT(ByVal varImport As String, varFormat As String, forceNull As Boolean) As String
    
    varImport = Trim$(varImport)
    
    If (varImport = "") Then
        NRM_IMPORT = IIf(forceNull, "", "0")
    Else
        If (CSng(varImport) = 0) Then
            NRM_IMPORT = IIf(forceNull, "", Format$(varImport, varFormat))
        Else
            NRM_IMPORT = Format$(varImport, varFormat)
        End If
    End If

End Function

Public Function NRM_REMOVEZEROES(ByVal varData As String, varTrail As Boolean) As String
    
    Dim minusSign   As String
    
    NRM_REMOVEZEROES = Trim$(varData)
    
    If (varTrail) Then
        If (Right$(NRM_REMOVEZEROES, 1) = "-") Then
            NRM_REMOVEZEROES = Left$(NRM_REMOVEZEROES, Len(NRM_REMOVEZEROES) - 1)
            minusSign = "-"
        End If
        
        Do While Right$(NRM_REMOVEZEROES, 1) = "0"
            NRM_REMOVEZEROES = Left$(NRM_REMOVEZEROES, Len(NRM_REMOVEZEROES) - 1)
        Loop
        
        NRM_REMOVEZEROES = minusSign & NRM_REMOVEZEROES
    
        If (Right$(NRM_REMOVEZEROES, 1) = ",") Then NRM_REMOVEZEROES = Left$(NRM_REMOVEZEROES, Len(NRM_REMOVEZEROES) - 1)
    Else
        Do While Left$(NRM_REMOVEZEROES, 1) = "0"
            NRM_REMOVEZEROES = Right$(NRM_REMOVEZEROES, Len(NRM_REMOVEZEROES) - 1)
        Loop
    End If

End Function

Public Sub WRITE2LOG(descrError As String)
    
    Open WS_LOGFILEPATH For Append As #2
        Print #2, descrError
    Close #2

End Sub
