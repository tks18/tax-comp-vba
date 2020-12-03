Attribute VB_Name = "Module1"
Option Explicit

Public Function INDTAX(AMT As LongLong, ANS_TYPE As String)

Dim firstRes As LongLong
Dim secRes As LongLong
Dim interRes As LongLong
Dim thirdRes As LongLong
Dim surch As LongLong
Dim cess As LongLong
Dim finalAns As LongLong

firstRes = (500000 - 250000) * 0.05
secRes = (1000000 - 500000) * 0.2

If (AMT > 50000000) Then
    interRes = ((AMT - 1000000) * 0.3)
    thirdRes = firstRes + secRes + interRes
    surch = thirdRes * 0.37
ElseIf (AMT >= 20000000) Then
    interRes = ((AMT - 1000000) * 0.3)
    thirdRes = firstRes + secRes + interRes
    surch = thirdRes * 0.25
ElseIf (AMT >= 10000000) Then
    interRes = ((AMT - 1000000) * 0.3)
    thirdRes = firstRes + secRes + interRes
    surch = thirdRes * 0.15
ElseIf (AMT >= 5000000) Then
    interRes = ((AMT - 1000000) * 0.3)
    thirdRes = firstRes + secRes + interRes
    surch = thirdRes * 0.1
Else
    If (AMT >= 1000000) Then
        interRes = ((AMT - 1000000) * 0.3)
        thirdRes = firstRes + secRes + interRes
        surch = 0
    ElseIf (AMT >= 500000) Then
        interRes = ((AMT - 500000) * 0.2)
        thirdRes = firstRes + ((AMT - 500000) * 0.2)
        surch = 0
    ElseIf (AMT <= 250000) Then
        interRes = 0
        thirdRes = 0
        surch = 0
    Else
        interRes = (AMT - 250000) * 0.05
        thirdRes = interRes
        surch = 0
    End If
End If

cess = (thirdRes + surch) * 0.04
finalAns = Application.WorksheetFunction.MRound(thirdRes + surch + cess, 10)

If (IsNull(ANS_TYPE)) Then
    INDTAX = finalAns
ElseIf LCase(ANS_TYPE) = "s1" Then
    If (AMT >= 500000) Then
        INDTAX = firstRes
    ElseIf (AMT <= 250000) Then
        INDTAX = 0
    Else
        INDTAX = interRes
    End If
ElseIf LCase(ANS_TYPE) = "s2" Then
    If (AMT >= 1000000) Then
        INDTAX = secRes
    ElseIf (AMT <= 500000) Then
        INDTAX = 0
    Else
        INDTAX = interRes
    End If
ElseIf LCase(ANS_TYPE) = "s3" Then
    If (AMT >= 1000000) Then
        INDTAX = interRes
    Else
        INDTAX = 0
    End If
ElseIf LCase(ANS_TYPE) = "surch" Then
    INDTAX = surch
ElseIf LCase(ANS_TYPE) = "cess" Then
    INDTAX = cess
ElseIf LCase(ANS_TYPE) = "noround" Then
    INDTAX = thirdRes + surch + cess
Else
    INDTAX = finalAns
End If


End Function


