Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Private Type strct_DLLParams
    BASEFILENAME    As String
    DSN             As String
    MODE            As String
    OUTPUTFILEPATH  As String
    OUTFILENAME     As String
    TYPE            As String
    UnattendedMode  As Boolean
    WORKINGID       As String
    WORKINGTABLE    As String
    ZIPEXEPATH      As String
End Type

Public DLLParams    As strct_DLLParams

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
