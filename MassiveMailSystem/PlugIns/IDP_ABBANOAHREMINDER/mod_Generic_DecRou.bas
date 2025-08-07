Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)

Private Type strct_DLLParams
    BYPASS_STAMP        As Boolean
    CCP_BILL            As String
    CCP_PPA             As String
    CF_ENTE             As String
    DSN                 As String
    DSN_EXT             As String
    EXTRASPATH          As String
    IDDATACUTTER        As String
    IDWORKINGLOAD       As String
    INPUTFILENAME       As String
    LAYOUT              As String
    OUTPUTFILENAME      As String
    PRINT_BILL          As Boolean
    PRM_DTAEMISSIONE    As String
    PRM_GG              As String
    PRM_DCM             As String
    TABLENAME           As String
    TEMPLATEVERSION     As String
    TNS                 As String
    UNATTENDEDMODE      As Boolean
End Type

Public Enum TextPadJustify
    PADLEFT
    PADRIGHT
    PADCENTER
End Enum

Public DLLParams        As strct_DLLParams
Public UMErrMsg         As String

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

Public Function GET_CAPITALIZED_STRING(varSTR As String) As String
    
    Dim I           As Integer
    Dim WS_DATA()   As String

    WS_DATA = Split(Trim$(varSTR), " ")
    
    For I = 0 To UBound(WS_DATA)
        GET_CAPITALIZED_STRING = GET_CAPITALIZED_STRING & UCase$(Left$(WS_DATA(I), 1)) & LCase(Mid$(WS_DATA(I), 2)) & " "
    Next I

    GET_CAPITALIZED_STRING = Trim$(GET_CAPITALIZED_STRING)

End Function

Public Function GET_DUPLICATE_STRING(varSTR As String, numRepeat As Integer) As String
    
    Dim I As Integer
        
    For I = 1 To numRepeat
        GET_DUPLICATE_STRING = GET_DUPLICATE_STRING & varSTR
    Next I

End Function

Public Function GET_NUM2STRING(inputNum As String) As String
    
    If (Trim$(inputNum) = "") Then Exit Function
    
    Dim CalcUnità       As Boolean
    Dim Cntr            As Byte
    Dim Centinaia       As Variant
    Dim Decine          As Variant
    Dim I               As Integer
    Dim Unità           As Variant
    Dim SplitValue()    As String
    Dim Terzina         As Byte
    Dim tmpCentinaia    As String
    Dim tmpDecine       As String
    Dim TmpStr          As String
    Dim tmpUnità        As String
    Dim tmpTerzina      As String
    
    Centinaia = Array("Cento", "Duecento", "Trecento", "Quattrocento", "Cinquecento", "Seicento", "Settecento", "Ottocento", "Novecento")
    Decine = Array("Venti", "Trenta", "Quaranta", "Cinquanta", "Sessanta", "Settanta", "Ottanta", "Novanta")
    Unità = Array("Uno", "Due", "Tre", "Quattro", "Cinque", "Sei", "Sette", "Otto", "Nove", "Dieci", "Undici", "Dodici", "Tredici", "Quattordici", "Quindici", "Sedici", "Diciassette", "Diciotto", "Diciannove")
    
    SplitValue = Split(inputNum, ",")
    
    For I = Len(SplitValue(0)) To 1 Step -1
        TmpStr = Mid$(SplitValue(0), I, 1) & TmpStr
        
        Cntr = Cntr + 1
        
        If Cntr > 2 Or (I = 1) Then
            CalcUnità = True
            Terzina = Terzina + 1
            Cntr = Len(TmpStr)
            
            If Cntr = 3 And Left$(TmpStr, 1) > 0 Then
                tmpCentinaia = Centinaia(Left$(TmpStr, 1) - 1)
            End If
            
            If Cntr >= 2 Then
                If Mid$(TmpStr, Cntr - 2 + 1, 2) <> "00" Then
                    If Mid$(TmpStr, Cntr - 2 + 1, 2) > 19 Then
                        tmpDecine = Decine(Mid$(TmpStr, Cntr - 1, 1) - 2)
                    Else
                        tmpDecine = Unità(Right$(TmpStr, 2) - 1)
                        
                        CalcUnità = False
                    End If
                End If
                
                If tmpCentinaia <> "" And Mid$(TmpStr, 2, 1) = 8 Then
                    tmpCentinaia = Left$(tmpCentinaia, Len(tmpCentinaia) - 1)
                End If
            End If
            
            If CalcUnità And Right$(TmpStr, 1) > 0 Then
                If Cntr >= 1 Then
                    tmpUnità = Unità(Right$(TmpStr, 1) - 1)
                        
                    If tmpDecine <> "" And (Right$(TmpStr, 1) = 1 Or Right$(TmpStr, 1) = 8) Then
                        tmpDecine = Left$(tmpDecine, Len(tmpDecine) - 1)
                    End If
                End If
            End If
                
            Select Case Terzina
                Case 2
                    If Val(TmpStr) > 0 Then
                        If Val(TmpStr) > 1 Then
                            tmpTerzina = "Mila"
                        Else
                            tmpTerzina = "Mille"
                            tmpUnità = ""
                            tmpDecine = ""
                        End If
                    End If
                    
                Case 3
                    If Right$(TmpStr, 1) > 1 Then
                        tmpTerzina = "Milioni"
                    Else
                        tmpTerzina = "Milione"
                        
                        tmpUnità = Left$(tmpUnità, Len(tmpUnità) - 1)
                    End If
                
            End Select
            
            GET_NUM2STRING = tmpCentinaia & tmpDecine & tmpUnità & tmpTerzina & GET_NUM2STRING
            
            Cntr = 0
            TmpStr = ""
            tmpCentinaia = ""
            tmpDecine = ""
            tmpUnità = ""
        End If
    Next I
    
    If UBound(SplitValue) > 0 Then GET_NUM2STRING = GET_NUM2STRING & "/" & SplitValue(1)
    
    Erase SplitValue

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

Public Function GET_TEXTPAD(Mode As TextPadJustify, myString As String, NumChar As Integer, String2Repeat As String, AutoTrim As Boolean) As String

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

Public Sub WRITE2LOG(descrError As String)
    
    Open WS_LOGFILEPATH For Append As #2
        Print #2, descrError
    Close #2

End Sub
