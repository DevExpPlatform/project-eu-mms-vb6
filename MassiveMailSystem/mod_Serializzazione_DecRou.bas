Attribute VB_Name = "mod_Serializzazione_DecRou"
Option Explicit

Public Function Calc_ScaglionePeso() As Byte
            
    Dim I                   As Byte
    Dim ScaglioniPeso(6, 1) As Integer
    
    ScaglioniPeso(0, 0) = 0
    ScaglioniPeso(0, 1) = 20
    ScaglioniPeso(1, 0) = 21
    ScaglioniPeso(1, 1) = 50
    ScaglioniPeso(2, 0) = 51
    ScaglioniPeso(2, 1) = 100
    ScaglioniPeso(3, 0) = 101
    ScaglioniPeso(3, 1) = 250
    ScaglioniPeso(4, 0) = 251
    ScaglioniPeso(4, 1) = 350
    ScaglioniPeso(5, 0) = 351
    ScaglioniPeso(5, 1) = 1000
    ScaglioniPeso(6, 0) = 1001
    ScaglioniPeso(6, 1) = 2000
            
    For I = 0 To 6
        If ProjectInfo.PesoProdotto >= ScaglioniPeso(I, 0) And ProjectInfo.PesoProdotto <= ScaglioniPeso(I, 1) Then
            Calc_ScaglionePeso = (I + 1)
            
            Exit For
        End If
    Next I
            
    Erase ScaglioniPeso

End Function
