Attribute VB_Name = "Module1"
Option Explicit

Public Function INDTAX(AMT As LongLong)

Dim firstRes As LongLong
Dim secRes As LongLong

firstRes = (500000 - 250000) * 0.05
secRes = (1000000 - 500000) * 0.2

If (AMT >= 1000000) Then
INDTAX = firstRes + secRes + ((AMT - 1000000) * 0.3)
ElseIf (AMT >= 500000) Then
INDTAX = firstRes + ((AMT - 500000) * 0.2)
Else
INDTAX = (AMT - 250000) * 0.05
End If

End Function

