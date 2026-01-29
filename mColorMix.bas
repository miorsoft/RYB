Attribute VB_Name = "mColorMix"
Option Explicit
'https://stackoverflow.com/questions/1351442/is-there-an-algorithm-for-color-mixing-that-works-like-mixing-real-colors
' Came Cameron


Private Function Min(ByVal A!, ByVal B!) As Single
    If A < B Then Min = A Else: Min = B
End Function

Public Sub PigmentMix(ByVal R1!, ByVal G1!, ByVal B1!, _
                      ByVal R2!, ByVal G2!, ByVal B2!, _
                      ByRef RR!, ByRef GG!, ByRef BB!)

    Dim W1!
    Dim W2!
    Dim Wcolor!
    Dim nR1!, nG1!, nB1!
    Dim nR2!, nG2!, nB2!

    Dim Wavg!

    '[1] Remove white from all colors, keeping the white parts and color parts
    W1 = Min(R1, Min(G1, B1))
    'W1 = (R1 + G1 + B1) * 0.33333

    W2 = Min(R2, Min(G2, B2))
    'W2 = (R2 + G2 + B2) * 0.33333

    nR1 = R1 - W1: nG1 = G1 - W1: nB1 = B1 - W1
    nR2 = R2 - W2: nG2 = G2 - W2: nB2 = B2 - W2

    '[2] Average the RGB values of the white parts removed from the colors
    Wavg = (W1 + W2) * 0.5

    '[3] Average the RGB values of the color parts
    RR = (nR1 + nR2) * 0.5
    GG = (nG1 + nG2) * 0.5
    BB = (nB1 + nB2) * 0.5

    '[4] Take out the white from the averaged color parts
    Wcolor = Min(RR, Min(GG, BB))
    RR = RR - Wcolor
    GG = GG - Wcolor
    BB = BB - Wcolor

    '[5] Half the white value removed and add that value to the Green of the averaged color parts
    ' GG = GG + Wcolor * 0.5
    ' ( Changing the green portion added back in step 5 should help. I thought 0.75 or 0.8 looked better than 0.5 )
    GG = GG + Wcolor * 0.75             ' 0.75 '0.75


    '[6] Add the averaged white parts back in and make whole number
    RR = RR + Wavg
    GG = GG + Wavg
    BB = BB + Wavg


Debug.Print R1 & "  " & G1 & "  " & B1
Debug.Print R2 & "  " & G2 & "  " & B2
Debug.Print RR & "  " & GG & "  " & BB
Debug.Print "  "

End Sub


