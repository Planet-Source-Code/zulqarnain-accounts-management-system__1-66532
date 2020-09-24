Attribute VB_Name = "mod_Validation"
'***************************************************************'
'For Validating Users's Input in TextBoxes
'***************************************************************'

Public Sub OnlyAlphabets(KeyAscii As Integer)

If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii >= 0 And KeyAscii <= 7 Or KeyAscii >= 9 And KeyAscii <= 31 Or KeyAscii >= 33 And KeyAscii <= 44 Or KeyAscii > 46 And KeyAscii <= 47 Or KeyAscii >= 58 And KeyAscii <= 64 Or KeyAscii >= 91 And KeyAscii <= 96 Or KeyAscii >= 123 And KeyAscii <= 500 Then
    KeyAscii = 0
End If

End Sub

Public Sub OnlyNumbers(KeyAscii As Integer)

If KeyAscii = 45 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= 0 And KeyAscii <= 7 Or KeyAscii >= 9 And KeyAscii <= 31 Or KeyAscii >= 33 And KeyAscii <= 44 Or KeyAscii >= 47 And KeyAscii <= 47 Or KeyAscii >= 58 And KeyAscii <= 96 Or KeyAscii >= 123 And KeyAscii <= 500 Then
    KeyAscii = 0
End If

End Sub

Public Sub NoApostrophie(KeyAscii As Integer)

If KeyAscii = 39 Then
    KeyAscii = 0
End If

End Sub

Public Sub NoSpace(KeyAscii As Integer)

If KeyAscii = 32 Then
    KeyAscii = 0
End If

End Sub

Public Sub ForNumericFields(KeyAscii As Integer, txt As TextBox)

Call mod_Validation.OnlyNumbers(KeyAscii)
Call mod_Validation.NoApostrophie(KeyAscii)
If txt.SelStart = 0 Then
    Call mod_Validation.NoSpace(KeyAscii)
End If

End Sub

Public Sub ForStringFields(KeyAscii As Integer, txt As TextBox)

Call mod_Validation.NoApostrophie(KeyAscii)
If txt.SelStart = 0 Then
    Call mod_Validation.NoSpace(KeyAscii)
End If

End Sub

Public Sub ForGeneralFields(KeyAscii As Integer, txt As TextBox)

If txt.SelStart = 0 Then
    Call mod_Validation.NoSpace(KeyAscii)
End If
Call mod_Validation.NoApostrophie(KeyAscii)

End Sub


