VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_GenXFDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FN As Integer

Private Function GET_TXT2UCTXT(ByVal InputTXT As String) As String

    If (InputTXT <> "") Then
        GET_TXT2UCTXT = StrConv(UniCode2ANSI(InputTXT, CP_UTF8, 0), vbUnicode)
        GET_TXT2UCTXT = Left$(GET_TXT2UCTXT, InStr(GET_TXT2UCTXT, vbNullChar) - 1)
    End If

End Function

Public Sub XFDF_Close()
    
    Print #FN, vbTab & "</fields>"
    Print #FN, "</xfdf>"
    
    Close #FN

End Sub

Public Sub XFDF_FieldBarCode(ByVal FieldName As String, ByVal FieldSymbology As String, ByVal FieldValue As String)
    
    FieldValue = Replace$(FieldValue, "<", "&lt;")
    FieldValue = Replace$(FieldValue, ">", "&gt;")
    FieldValue = Replace$(FieldValue, "Ê", GET_TXT2UCTXT("Ê"))
    
    Print #FN, vbTab & vbTab & "<field name=" & Chr$(34) & FieldName & Chr$(34) & " type=""barcode"" symbology=" & Chr$(34) & FieldSymbology & Chr$(34) & ">"
    Print #FN, vbTab & vbTab & vbTab & "<value>" & FieldValue & "</value>"
    Print #FN, vbTab & vbTab & "</field>"

End Sub

Public Sub XFDF_FieldImage(ByVal FieldName As String, ByVal FieldImage As String, ByVal FieldValue As String)
        
    Print #FN, vbTab & vbTab & "<field name=" & Chr$(34) & FieldName & Chr$(34) & " type=""img"" href=" & Chr$(34) & FieldImage & Chr$(34) & ">"
    Print #FN, vbTab & vbTab & vbTab & "<value>" & FieldValue & "</value>"
    Print #FN, vbTab & vbTab & "</field>"

End Sub

Public Sub XFDF_FieldText(ByVal FieldName As String, ByVal FieldValue As String, ByVal isML As Boolean)
    
    Dim strEnd      As Long
    Dim strLen      As Long
    Dim strRow      As String
    Dim strStart    As Long
    
    strStart = 1
    strEnd = 1
    
    If (isML) Then
        FieldValue = Trim$(FieldValue)
        
        Do
            strStart = (InStr(strEnd, FieldValue, "[CDATA["))
            
            If (strStart > 0) Then
                strEnd = (InStr(strStart, FieldValue, "]]"))
                strStart = strStart + 7
                strLen = (strEnd - strStart)
                strRow = Mid$(FieldValue, strStart, strLen)
                
                FieldValue = Left$(FieldValue, strStart - 1) & GET_TXT2UCTXT(strRow) & Mid$(FieldValue, strEnd)
                strEnd = (InStr(strStart, FieldValue, "]]"))
            End If
        Loop Until strStart = 0
    Else
        FieldValue = Replace$(FieldValue, "&", "&amp;")
        FieldValue = Replace$(FieldValue, "<", "&lt;")
        FieldValue = Replace$(FieldValue, ">", "&gt;")
        
        FieldValue = GET_TXT2UCTXT(FieldValue)
    End If
    
    Print #FN, vbTab & vbTab & "<field name=" & Chr$(34) & FieldName & Chr$(34) & IIf(isML, " type=" & Chr$(34) & "ml" & Chr$(34), "") & ">"
    Print #FN, vbTab & vbTab & vbTab & "<value>" & FieldValue & "</value>"
    Print #FN, vbTab & vbTab & "</field>"

End Sub

Public Sub XFDF_Open(ByVal XFDFFileName As String, ByVal TemplateFileName As String, ByVal prjbasepath As String)
    
    FN = FreeFile
    
    Open XFDFFileName For Output As #FN
    
    Print #FN, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #FN, "<xfdf xmlns=""http://ns.adobe.com/xfdf/"">"
    Print #FN, vbTab & "<f href=" & Chr$(34) & TemplateFileName & Chr$(34) & "/>"
    Print #FN, vbTab & "<extraparams prjbasepath=" & Chr$(34) & prjbasepath & Chr$(34) & "/>"
    Print #FN, vbTab & "<fields>"

End Sub
