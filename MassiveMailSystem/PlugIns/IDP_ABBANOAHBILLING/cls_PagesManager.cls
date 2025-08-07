VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_PagesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pageHeight      As Single
Private pageHeightMax   As Single
Private pageHeightMin   As Single
Private pagesCounter    As Integer
Private vctrPagesData() As String

Private Sub Class_Initialize()
    
    pageHeight = 10
    pageHeightMin = 10
    pageHeightMax = 280
    pagesCounter = -1

End Sub

Private Sub addData(item As String)
    
    pagesCounter = (pagesCounter + 1)
    ReDim Preserve vctrPagesData(pagesCounter) As String
    
    vctrPagesData(pagesCounter) = item

End Sub

Public Sub clean()

    Erase vctrPagesData()

    pageHeight = pageHeightMin
    pagesCounter = -1
    
End Sub

Public Property Get getData() As String
    
    Dim I As Integer
    
    If (isEmpty) Then
        getData = ""
    Else
        For I = 0 To pagesCounter
            getData = getData & vctrPagesData(I) & "[EP]"
        Next I
    
        getData = Left$(getData, (Len(getData) - 4))
    End If

End Property

Public Property Get getHeight() As String

    getHeight = pageHeight

End Property

Public Property Get getPages() As String

    getPages = (pagesCounter + 1)
    
End Property

Public Property Get isEmpty() As Boolean
    
    isEmpty = (pagesCounter = -1)
    
End Property

Public Sub setData(itemCloser As String, item As String, itemHeight As Single, itemOpen As String, itemOpenHeight As Single, Optional newMin As Single)

    Dim tmpPage As String

    If (isEmpty) Then
        pagesCounter = 0
        
        ReDim vctrPagesData(pagesCounter) As String
    End If

    pageHeight = (pageHeight + itemHeight)

    If (pageHeight > pageHeightMax) Then
        If (newMin <> 0) Then setMin = newMin
    
        pageHeight = (pageHeightMin + itemHeight)
    
        If (itemCloser <> "") Then
            tmpPage = vctrPagesData(pagesCounter) & itemCloser
            
            If (Right$(tmpPage, 16) <> "</table></table>") Then vctrPagesData(pagesCounter) = tmpPage
        End If
    
        If (itemOpen = "") Then
            addData item
        Else
            pageHeight = (pageHeight + itemOpenHeight)
            
            If (Left$(item, 6) = "<table") Then
                addData item
            Else
                addData itemOpen & item
            End If
        End If
    Else
        vctrPagesData(pagesCounter) = vctrPagesData(pagesCounter) & item
    End If

End Sub

Public Property Let setMax(varValue As Single)

    pageHeightMax = varValue

End Property

Public Property Let setMin(varValue As Single)

    pageHeightMin = varValue
    pageHeight = varValue

End Property
