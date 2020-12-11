Option Explicit

Dim moji

moji = "a"

If Len(moji) = LenByte(moji) Then
   MsgBox "‚º‚ñ‚Ô”¼Šp"
Else
    MsgBox "‘SŠp‚ª‚Ü‚¶‚Á‚Ä‚Ü‚·‚È"
End If

moji = "‚ "

If Len(moji) = LenByte(moji) Then
    MsgBox "‚º‚ñ‚Ô”¼Šp"
Else
    MsgBox "‘SŠp‚ª‚Ü‚¶‚Á‚Ä‚Ü‚·‚È"
End If


Function LenByte(ByVal s)

    Dim c, i, k

    c = 0

    For i = 0 To Len(s) - 1
        k = Mid(s, i + 1, 1)

        If (Asc(k) And &HFF00) = 0 Then
            c = c + 1
        Else
            c = c + 2
        End If
    Next

    LenByte = c

End Function