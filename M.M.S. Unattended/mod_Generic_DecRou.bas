Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Public Function chk_Array(ByVal myArray As Variant) As Boolean
    
    On Error GoTo ErrHandler
    
    If UBound(myArray) > -1 Then chk_Array = True
    
    Exit Function

ErrHandler:

End Function

Public Function PDF_Open(ByVal PDFFileName As String) As Boolean

    On Error GoTo ErrHandler
  
    Dim AcrobatPath As String
    Dim rValue      As Long
  
    AcrobatPath = String(128, 32)
    rValue = FindExecutable(PDFFileName, vbNullString, AcrobatPath)
  
    If rValue <= 32 Then
        MsgBox "Acrobat could not be found on this computer.", vbExclamation, "PDF Open:"
      
        Exit Function
    End If
    
    AcrobatPath = Left$(AcrobatPath, Len(Trim$(AcrobatPath)) - 1)
    rValue = Shell(Chr$(34) & AcrobatPath & Chr$(34) & " " & PDFFileName, vbNormalFocus)
       
    If (rValue >= 0) And (rValue <= 32) Then
        MsgBox "An error occured launching Acrobat.", vbExclamation, "PDF Open:"
    
        Exit Function
    End If
    
    PDF_Open = True
    
    Exit Function

ErrHandler:
    MsgBox Err.Description, vbExclamation, "PDF Open:"

End Function

