Attribute VB_Name = "mColorMix"
Option Explicit
'https://stackoverflow.com/questions/1351442/is-there-an-algorithm-for-color-mixing-that-works-like-mixing-real-colors
' Came Cameron


Private Function min(ByVal A!, ByVal B!) As Single
    If A < B Then min = A Else: min = B
End Function

Public Sub PigmentMix(ByVal R1!, ByVal G1!, ByVal B1!, _
                      ByVal R2!, ByVal G2!, ByVal B2!, ByVal Perc!, _
                      ByRef rr!, ByRef GG!, ByRef bB!)

    Dim W1!
    Dim W2!
    Dim Wcolor!
    Dim nR1!, nG1!, nB1!
    Dim nR2!, nG2!, nB2!
    Dim Wavg!


    '---- Percentace by reexre
    Dim Q!
    Q = 1 - Perc


    '[1] Remove white from all colors, keeping the white parts and color parts
    W1 = min(R1, min(G1, B1))
    'W1 = (R1 + G1 + B1) * 0.33333

    W2 = min(R2, min(G2, B2))
    'W2 = (R2 + G2 + B2) * 0.33333

    nR1 = R1 - W1: nG1 = G1 - W1: nB1 = B1 - W1
    nR2 = R2 - W2: nG2 = G2 - W2: nB2 = B2 - W2

    '[2] Average the RGB values of the white parts removed from the colors
    'Wavg = (W1 + W2) * 0.5
    'Using Perc
    Wavg = (W1 * Q + Perc * W2)


    '[3] Average the RGB values of the color parts
    '    rr = (nR1 + nR2) * 0.5
    '    GG = (nG1 + nG2) * 0.5
    '    bB = (nB1 + nB2) * 0.5

    'Using Perc
    rr = nR1 * Q + Perc * nR2
    GG = nG1 * Q + Perc * nG2
    bB = nB1 * Q + Perc * nB2


    '[4] Take out the white from the averaged color parts
    Wcolor = min(rr, min(GG, bB))
    rr = rr - Wcolor
    GG = GG - Wcolor
    bB = bB - Wcolor

    '[5] Half the white value removed and add that value to the Green of the averaged color parts
    ' GG = GG + Wcolor * 0.5
    ' ( Changing the green portion added back in step 5 should help. I thought 0.75 or 0.8 looked better than 0.5 )
    GG = GG + Wcolor * 0.75             ' 0.75 '0.75


    '[6] Add the averaged white parts back in and make whole number
    rr = rr + Wavg
    GG = GG + Wavg
    bB = bB + Wavg


    ''Debug.Print R1 & "  " & G1 & "  " & B1
    ''Debug.Print R2 & "  " & G2 & "  " & B2
    ''Debug.Print RR & "  " & GG & "  " & BB
    ''Debug.Print "  "

End Sub


