Attribute VB_Name = "mod_Generic_DecRou"
Option Explicit

Private Type strct_DLLParams
    UnattendedMode          As Boolean
    XTXT_AddFieldsBarCode   As Boolean
    XTXT_AddFieldsPSTL      As Boolean
    XTXT_DSN                As String
    XTXT_FileName           As String
    XTXT_IdDataCutter       As Integer
    XTXT_TableName          As String
    XTXT_TNS                As String
End Type

Public DLLParams            As strct_DLLParams
Public UMErrMsg             As String

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

    InputTXT = Trim$(InputTXT)

    If InputTXT <> "" Then
        InputTXT = Replace$(InputTXT, "'", "''")
        
        Conv_String2SQLString = "'" & InputTXT & "'"
    Else
        Conv_String2SQLString = "NULL"
    End If

End Function

Public Function Conv_Time2SQLServerTime(ByVal InputDate As String, Optional ByVal SQLFormat As Boolean = True) As String
    
    Dim Conv_Str2Date As String
    Dim HaveTimeStamp As Boolean
    
    Select Case Trim$(InputDate)
        Case "00000000", "", String$(8, Chr$(0))
            If SQLFormat Then
                Conv_Time2SQLServerTime = "NULL"
            Else
                Conv_Time2SQLServerTime = Space$(10)
            End If
        
        Case Else
            If Not IsDate(InputDate) Then
                Dim X_Left As Byte
                Dim X_Middle As Byte
                Dim X_Right As Byte
                
                If CInt(Right$(InputDate, 4)) > 1231 Then
                    X_Left = 2
                    X_Middle = 3
                    X_Right = 4
                Else
                    X_Left = 4
                    X_Middle = 5
                    X_Right = 2
                End If

                Conv_Str2Date = Format$(Left$(InputDate, X_Left) & "/" & Mid$(InputDate, X_Middle, 2) & "/" & Right$(InputDate, X_Right), "dd/mm/yyyy")
            Else
                Conv_Str2Date = InputDate
            End If
                        
            If SQLFormat Then
                HaveTimeStamp = (Len(CStr(Conv_Str2Date)) > 10)
                Conv_Time2SQLServerTime = "TO_DATE('" & Format$(Conv_Str2Date, "dd/mm/yyyy") & IIf(HaveTimeStamp, " " & Format$(Hour(Conv_Str2Date), "00") & ":" & Format$(Minute(Conv_Str2Date), "00") & ":" & Format$(Second(Conv_Str2Date), "00"), "") & "', " & _
                                          "'DD/MM/YYYY" & IIf(HaveTimeStamp, " HH24:MI:SS", "") & "')"
            Else
                Conv_Time2SQLServerTime = Conv_Str2Date
            End If
        
    End Select
    
End Function

Public Sub DLL_Init()

    Dim DLLParamsEmpty As strct_DLLParams

    DLLParams = DLLParamsEmpty

End Sub

Public Function Purge_ErrDescr(ErrDescr As String) As String
    
    Purge_ErrDescr = Mid$(ErrDescr, InStrRev(ErrDescr, "]") + 1, Len(ErrDescr))
    
End Function
