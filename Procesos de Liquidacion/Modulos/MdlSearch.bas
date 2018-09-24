Attribute VB_Name = "MdlSearch"
Option Explicit

Public Function BinarySearch(strArray() As String, strSearch As String) As Long
Dim lngIndex As Long
Dim lngFirst As Long
Dim lngLast As Long
Dim lngMiddle As Long
Dim bolInverseOrder As Boolean
     
    lngFirst = LBound(strArray)
    lngLast = UBound(strArray)
     
    bolInverseOrder = (strArray(lngFirst) > strArray(lngLast))
    
    BinarySearch = lngFirst - 1
    Do
        lngMiddle = (lngFirst + lngLast) \ 2
        If strArray(lngMiddle) = strSearch Then
            BinarySearch = lngMiddle
            Exit Do
        ElseIf ((strArray(lngMiddle) < strSearch) Xor bolInverseOrder) Then
             lngFirst = lngMiddle + 1
        Else
             lngLast = lngMiddle - 1
        End If
    Loop Until lngFirst > lngLast
End Function



Public Function IndiceContrato(TerceroBuscado As Long) As Long
Dim lngIndex As Long
Dim lngFirst As Long
Dim lngLast As Long
Dim lngMiddle As Long
Dim bolInverseOrder As Boolean
     
    lngFirst = LBound(Arr_Contrato)
    lngLast = UBound(Arr_Contrato)
     
     If lngFirst > 0 Then
        bolInverseOrder = (Arr_Contrato(lngFirst).ternro > Arr_Contrato(lngLast).ternro)
        
        IndiceContrato = lngFirst - 1
        Do
            lngMiddle = (lngFirst + lngLast) \ 2
            If Arr_Contrato(lngMiddle).ternro = TerceroBuscado Then
                IndiceContrato = lngMiddle
                Exit Do
            ElseIf ((Arr_Contrato(lngMiddle).ternro < TerceroBuscado) Xor bolInverseOrder) Then
                 lngFirst = lngMiddle + 1
            Else
                 lngLast = lngMiddle - 1
            End If
        Loop Until lngFirst > lngLast
    Else
        IndiceContrato = 0
    End If
End Function

