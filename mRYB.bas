Attribute VB_Name = "mRYB"
Option Explicit

'     REEXRE - MiorSoft Implementation of RGB->RYB and RYB->RGB colorspace conversion and (RGB2HSL HSL2RGB conversion)
'
'     RGB2ryb0
'     ryb2RGB0

'     RGB2HSLmy
'     HSL2RGBmy

'     Copyright (c) 2026 Roberto Mior (miorsoft - reexre)
'
'     Permission is hereby granted, free of charge, to any person obtaining a copy
'     of this software and associated documentation files (the "Software"), to deal
'     in the Software without restriction, including without limitation the rights
'     to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'     copies of the Software, and to permit persons to whom the Software is
'     furnished to do so, subject to the following conditions:
'
'     The above copyright notice and this permission notice shall be included in all
'     copies or substantial portions of the Software.
'
'     THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'     IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'     FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'     AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'     LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'     OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'     SOFTWARE.

Private Type tVec2
    X             As Single
    Y             As Single
End Type

Private Const PI  As Single = 3.14159265358979
Private Const PIh As Single = 1.5707963267949
Private Const PI2 As Single = 6.28318530717959
Private Const InvPI2 As Single = 1 / 6.28318530717959
Private Const PI2_13 As Single = 6.28318530717959 / 3
Private Const PI2_23 As Single = 6.28318530717959 * 2 / 3
Private Const inv255 As Single = 1 / 255

Public Sub RGB2ryb0(R As Single, G As Single, B As Single)

    '     R--|--G--|--B--|--R
    '     0  1  2  3  4  5  6
    '     0           1
    '     0    0.5    1

    Dim H!, S!, L!
    RGB2HSLmy R, G, B, H, S, L
    If H < 0.6666666! Then              '4/6             '< Blue  (Range R=0 Y=.1666 G=0.33333)
        H = H * 1.5!                    '6/4
        '        H = H ^ 0.5   'Prima versione
        H = ForwardHUEtransform(H)
        H = H * 0.6666666!              '4/6
    End If
    L = 1 - L
    HSL2RGBmy H, S, L, R, G, B

End Sub
Public Sub ryb2RGB0(R As Single, Y As Single, B As Single)

    '     R--|--G--|--B--|--R
    '     0  1  2  3  4  5  6
    '     0           1
    '     0    0.5    1

    Dim H!, S!, L!
    RGB2HSLmy R, Y, B, H, S, L
    If H < 0.6666666! Then              '4/6
        H = H * 1.5!                    '6/4              ' / 0.6666667
        'H = H * H  'Prima versione
        H = InverseHUEtransform(H)
        H = H * 0.6666666!              '4/6
    End If
    L = 1 - L
    HSL2RGBmy H, S, L, R, Y, B

End Sub

Private Function Vec2(X!, Y!) As tVec2
    Vec2.X = X
    Vec2.Y = Y
End Function

Private Function Clamp01(V!) As Single
    If V < 0! Then
        Clamp01 = 0!
    ElseIf V > 1! Then
        Clamp01 = 1!
    Else
        Clamp01 = V
    End If
End Function

Public Function ForwardHUEtransform(ByVal X!) As Single
    ' 1st Version
    '  X ^ 0.5
    ' 2nd Version
    'ForwardHUEtransform = (2.72 - Exp(1 - 2.5 * X)) * 0.4
    'ForwardHUEtransform = ForwardHUEtransform * 0.5 + X ^ 0.5 * 0.5
    ' 3rd Version
    ForwardHUEtransform = 1.206639 * (1 - Exp(-1.764618 * (X ^ 0.860773)))
End Function
Public Function InverseHUEtransform(ByVal X!) As Single
    ' 1st Version
    '  X * X
    ' 2nd Version
    'InverseHUEtransform = -(Log(-2.5 * X + 2.72) - 1) * 0.4
    'InverseHUEtransform = InverseHUEtransform * 0.5 + X * X * 0.5
    ' 3rd Version
    '    InverseHUEtransform = (-Log(1 - (X / 1.206639)) / 1.764618) ^ (1 / 0.860773)
    ' 3rd Verdion (avoid Division)
    InverseHUEtransform = (-Log(1 - (X * 0.828748)) * 0.566694) ^ (1.161746)
End Function
Private Function ForwardHUEtransformEX(X!) As Single
    If X <= 0.25 Then
        ForwardHUEtransformEX = -27.17042496 * X * X * X + 10.79259688 * X * X + X
    ElseIf X <= 0.5 Then
        ForwardHUEtransformEX = 1.206639 * (1 - Exp(-1.764618 * (X ^ 0.860773)))
    Else
        ForwardHUEtransformEX = 3.055516 * X * X * X - 6.63879 * X * X + 5.111032 * X - 0.527758
    End If
End Function

Private Function InverseHUEtransformEX(X!) As Single
    Dim u!, R!
    If X <= 0.5 Then
        InverseHUEtransformEX = 0.132405115 + 0.345248 * Cos(1 / 3 * Acos(0.925157 - 3.578117 * X) - PI2_13)
    ElseIf X <= 0.75 Then
        InverseHUEtransformEX = (-Log(1 - (X * 0.828748)) * 0.566694) ^ (1.161746)
    Else
        u = 0.1392415 - 0.163593 * X
        R = (u * u + 0.000036) ^ 0.5
        If (-u - R) > 0 Then
Stop
            InverseHUEtransformEX = (-u + R) ^ (1 / 3) + (-u - R) ^ (1 / 3) + 0.724461
        Else
            InverseHUEtransformEX = (-u + R) ^ (1 / 3) + 0 + 0.724461
        End If
    End If
End Function


Private Sub RGB2HSLmy(R As Single, G As Single, B As Single, _
                      ByRef H As Single, ByRef S As Single, ByRef L As Single)
    Dim V1x!, V1y!
    Dim V2x!, V2y!
    Dim V3x!, V3y!
    Dim sumX!
    Dim sumY!

    V1x = R
    V2x = -0.5 * G
    V2y = 0.86602537 * G
    V3x = -0.5 * B
    V3y = -0.86602537 * B

    sumX = V1x + V2x + V3x
    sumY = V1y + V2y + V3y

    S = Sqr(sumX * sumX + sumY * sumY)

    L = (R + G + B) * 0.333333333333333

    H = Atan2(sumX, sumY)

    While H < 0: H = H + PI2: Wend
    While H > PI2: H = H - PI2: Wend
    H = (H) * InvPI2

End Sub

Private Sub HSL2RGBmy(ByVal H As Single, ByVal S As Single, ByVal L As Single, _
                      R As Single, G As Single, B As Single)
    Dim V1        As tVec2
    Dim V2        As tVec2
    Dim V3        As tVec2
    Dim HUEvec    As tVec2

    V1 = Vec2(1, 0)                     '  0 Degree
    V2 = Vec2(-0.5, 0.86602537)         '120 Degree
    V3 = Vec2(-0.5, -0.86602537)        '240 Degree

    H = H * PI2

    HUEvec.X = Cos(H) * S
    HUEvec.Y = Sin(H) * S

    If H < PI2_13 Then                  ' PI2 / 3
        DecomposeVector HUEvec, V1, V2, R, G: B = 0
    ElseIf H < PI2_23 Then              'PI2 * 2 / 3
        DecomposeVector HUEvec, V2, V3, G, B: R = 0
    Else
        DecomposeVector HUEvec, V3, V1, B, R: G = 0
    End If

    L = L - (R + G + B) * 0.333333333333333

    R = Clamp01(R + L)
    G = Clamp01(G + L)
    B = Clamp01(B + L)
End Sub
Private Sub DecomposeVector(Vdecompose As tVec2, alongVA As tVec2, alongVB As tVec2, retMagA As Single, retMagb As Single)
    Dim det       As Single
    ' Calcola il determinante
    det = alongVA.X * alongVB.Y - alongVA.Y * alongVB.X
    ' Calcola i coefficienti usando la regola di Cramer
    If det Then
        retMagA = (Vdecompose.X * alongVB.Y - Vdecompose.Y * alongVB.X) / det
        retMagb = (alongVA.X * Vdecompose.Y - alongVA.Y * Vdecompose.X) / det
    End If
End Sub

Public Function Atan2(DX As Single, DY As Single) As Single
    If DX Then Atan2 = Atn(DY / DX) + PI * (DX < 0!) _
       Else _
       Atan2 = -PIh - (DY > 0!) * PI
End Function

Public Sub PigmentMixREEXRE(ByVal R1!, ByVal G1!, ByVal B1!, _
                            ByVal R2!, ByVal G2!, ByVal B2!, _
                            ByRef RR!, ByRef GG!, ByRef BB!)

    R1 = R1 * inv255
    G1 = G1 * inv255
    B1 = B1 * inv255
    R2 = R2 * inv255
    G2 = G2 * inv255
    B2 = B2 * inv255

    RGB2ryb0 R1, G1, B1
    RGB2ryb0 R2, G2, B2

    RR = (R1 + R2) * 0.5
    GG = (G1 + G2) * 0.5
    BB = (B1 + B2) * 0.5

    ryb2RGB0 RR, GG, BB

    RR = Clamp01(RR)
    GG = Clamp01(GG)
    BB = Clamp01(BB)

    RR = RR * 255
    GG = GG * 255
    BB = BB * 255

End Sub


Public Sub CreateWheels()
    ' VERY ROW  Wheel PNG creator

    Dim SRF       As cCairoSurface
    Dim CC        As cCairoContext
    Dim B()       As Byte

    Dim pW        As Long
    Dim pH        As Long
    Dim RR!, GG!, BB!, YY!
    Dim NA!
    Dim H!, L!, S!
    Dim NewA!

    pW = 1024
    pH = pW * 0.525

    Set SRF = Cairo.CreateSurface(pW, pH, ImageSurface)
    Set CC = SRF.CreateContext
    CC.SetSourceRGB 0.8, 0.8, 0.8: CC.Paint

    SRF.BindToArray B

    Dim cX!, cY!, X!, Y!, A!, R!, maxR!

    maxR = pW * 0.215
    cX = pW * 0.25
    cY = pH * 0.55


    For R = 1 To maxR Step 0.7
        For A = 0 To PI2 Step 1 / R * 0.75
            L = R / maxR
            S = 0.85

            HSL2RGBmy A * InvPI2, S, L, RR, GG, BB
            X = cX + Cos(-A) * R
            Y = cY + Sin(-A) * R
            X = Int(X)
            Y = Int(Y)

            B(X * 4 + 0, Y) = BB * 255
            B(X * 4 + 1, Y) = GG * 255
            B(X * 4 + 2, Y) = RR * 255
            B(X * 4 + 3, Y) = 255

            ' H = A * InvPI2
            '            If H < 0.6666667! Then      '4/6
            '                H = H * 1.5!            '6/4              ' / 0.6666667
            '                H = ForwardHUEtransform(H)
            '                H = H * 0.6666667!      '4/6
            '            End If
            '
            ''            L = 1 - L
            '
            '            HSL2RGBmy H, 0.9, L, RR, YY, BB
            '            ryb2RGB0 RR, YY, BB
            '         NewA = H * PI2
            ryb2RGB0 RR, GG, BB


            NewA = A
            X = cX + Cos(-NewA) * R + pW * 0.5
            Y = cY + Sin(-NewA) * R
            X = Int(X)
            Y = Int(Y)

            B(X * 4 + 0, Y) = BB * 255
            B(X * 4 + 1, Y) = GG * 255
            B(X * 4 + 2, Y) = RR * 255
            B(X * 4 + 3, Y) = 255


        Next
    Next


    CC.SetSourceRGBA 0, 0, 0, 0.25
    For A = 0 To PI2 - 0.0001 Step PI2 / 3
        CC.MoveTo cX, cY
        CC.LineTo cX + maxR * Cos(-A), cY + maxR * Sin(-A)
        CC.Stroke
        CC.MoveTo cX + pW * 0.5, cY
        CC.LineTo cX + pW * 0.5 + maxR * Cos(-A), cY + maxR * Sin(-A)
        CC.Stroke
    Next

    CC.SetSourceRGBA 0, 0, 0, 0.25
    For A = PI2 / 6 To PI2 - 0.0001 Step PI2 / 3
        CC.MoveTo cX + maxR * Cos(-A) * 0.75, cY + maxR * Sin(-A) * 0.75
        CC.LineTo cX + maxR * Cos(-A), cY + maxR * Sin(-A)
        CC.Stroke
        CC.MoveTo cX + pW * 0.5 + maxR * Cos(-A), cY + maxR * Sin(-A)
        CC.LineTo cX + pW * 0.5 + maxR * Cos(-A) * 0.75, cY + maxR * Sin(-A) * 0.75
        CC.Stroke
    Next
    CC.Arc cX, cY, maxR * 0.55: CC.Stroke
    CC.Arc cX, cY, maxR * 0.45: CC.Stroke
    CC.Arc cX + pW * 0.5, cY, maxR * 0.55: CC.Stroke
    CC.Arc cX + pW * 0.5, cY, maxR * 0.45: CC.Stroke




    CC.SelectFont "Segoe UI", 14, , True

    CC.TextOut pW * 0.25 - 10, 8, "RGB"
    CC.TextOut pW * 0.75 - 10, 8, "RYB"

    CC.SelectFont "Segoe UI", 10

    R = maxR + 15
    A = 0
    X = cX + Cos(-A) * R - 45
    Y = cY + Sin(-A) * R - 45
    CC.DrawText X * 1, Y * 1, 90, 90, "0" & vbCrLf & "(0.0)", False, vbCenter, , True: X = X + pW * 0.5
    CC.DrawText X * 1, Y * 1, 90, 90, "0" & vbCrLf & "(0.0)", False, vbCenter, , True

    A = PI2 * 1 / 6
    X = cX + Cos(-A) * R - 45
    Y = cY + Sin(-A) * R - 45
    CC.DrawText X * 1, Y * 1, 90, 90, "60" & vbCrLf & "(0.25)", False, vbCenter, , True: X = X + pW * 0.5
    CC.DrawText X * 1, Y * 1, 90, 90, "60" & vbCrLf & "(0.25)", False, vbCenter, , True:


    A = PI2 * 2 / 6
    X = cX + Cos(-A) * R - 45
    Y = cY + Sin(-A) * R - 45
    CC.DrawText X * 1, Y * 1, 90, 90, "120" & vbCrLf & "(0.5)", False, vbCenter, , True: X = X + pW * 0.5
    CC.DrawText X * 1, Y * 1, 90, 90, "120" & vbCrLf & "(0.5)", False, vbCenter, , True:

    A = PI2 * 3 / 6
    X = cX + Cos(-A) * R - 45
    Y = cY + Sin(-A) * R - 45
    CC.DrawText X * 1, Y * 1, 90, 90, "180" & vbCrLf & "(0.75)", False, vbCenter, , True: X = X + pW * 0.5
    CC.DrawText X * 1, Y * 1, 90, 90, "180" & vbCrLf & "(0.75)", False, vbCenter, , True:


    A = PI2 * 4 / 6
    X = cX + Cos(-A) * R - 45
    Y = cY + Sin(-A) * R - 45
    CC.DrawText X * 1, Y * 1, 90, 90, "240" & vbCrLf & "(1.0)", False, vbCenter, , True: X = X + pW * 0.5
    CC.DrawText X * 1, Y * 1, 90, 90, "240" & vbCrLf & "(1.0)", False, vbCenter, , True:


    A = PI2 * 5 / 6
    X = cX + Cos(-A) * R - 45
    Y = cY + Sin(-A) * R - 45
    CC.DrawText X * 1, Y * 1, 90, 90, "300", False, vbCenter, , True: X = X + pW * 0.5
    CC.DrawText X * 1, Y * 1, 90, 90, "300", False, vbCenter, , True:


    SRF.WriteContentToPngFile App.Path & "\Images\HUEwheels.PNG"


End Sub
Private Function mix(A!, B!, p!) As Single
    mix = A + (B - A) * p
End Function
Public Sub CreateGradients(R1!, G1!, B1!, R2!, G2!, B2!, I As Long)
    Dim SRF       As cCairoSurface
    Dim CC        As cCairoContext
    Dim p!
    Dim pW&, pH&
    Dim X!

    pW = 720
    pH = 200

    Set SRF = Cairo.CreateSurface(pW, pH, ImageSurface)
    Set CC = SRF.CreateContext
    CC.SetSourceRGB 0.8, 0.8, 0.8: CC.Paint
    CC.SetLineWidth 1

    '    Dim R1!, G1!, B1!
    '    Dim R2!, G2!, B2!
    Dim RR1!, YY1!, BB1!
    Dim RR2!, YY2!, BB2!
    Dim R!, G!, B!
    Dim RR!, YY!, BB!

    'R1 = 1: G1 = 1: B1 = 0
    'R2 = 0: G2 = 0: B2 = 1

    RR1 = R1: YY1 = G1: BB1 = B1
    RR2 = R2: YY2 = G2: BB2 = B2

    RGB2ryb0 RR1, YY1, BB1
    RGB2ryb0 RR2, YY2, BB2

    CC.TextOut 10, pH * 0.1 - 17, "RGB"
    CC.TextOut 35, pH * 0.1 + 0, Round(R1, 2)
    CC.TextOut 35, pH * 0.1 + 17, Round(G1, 2)
    CC.TextOut 35, pH * 0.1 + 34, Round(B1, 2)
    CC.TextOut pW - 40, pH * 0.1 + 0, Round(R2, 2)
    CC.TextOut pW - 40, pH * 0.1 + 17, Round(G2, 2)
    CC.TextOut pW - 40, pH * 0.1 + 34, Round(B2, 2)

    CC.TextOut 10, pH * 0.6 - 17, "RYB"
    CC.TextOut 35, pH * 0.6 + 0, Round(RR1, 2)
    CC.TextOut 35, pH * 0.6 + 17, Round(YY1, 2)
    CC.TextOut 35, pH * 0.6 + 34, Round(BB1, 2)
    CC.TextOut pW - 40, pH * 0.6 + 0, Round(RR2, 2)
    CC.TextOut pW - 40, pH * 0.6 + 17, Round(YY2, 2)
    CC.TextOut pW - 40, pH * 0.6 + 34, Round(BB2, 2)


    For X = 60 To 660
        p = (X - 60) * 1.66666666666667E-03
        R = mix(R1, R2, p)
        G = mix(G1, G2, p)
        B = mix(B1, B2, p)
        CC.SetSourceRGB R, G, B
        CC.MoveTo X + 0.5, pH * 0.1
        CC.LineTo X + 0.5, pH * 0.4
        CC.Stroke

        RR = mix(RR1, RR2, p)
        YY = mix(YY1, YY2, p)
        BB = mix(BB1, BB2, p)
        ryb2RGB0 RR, YY, BB

        CC.SetSourceRGB RR, YY, BB
        CC.MoveTo X + 0.5, pH * 0.6
        CC.LineTo X + 0.5, pH * 0.9
        CC.Stroke

    Next


    SRF.WriteContentToPngFile App.Path & "\Images\Gradients\Gradient" & I & ".PNG"




End Sub
Private Function Acos(Value As Single) As Single
    '    If CSng(Value) = -1# Then Acos = PI: Exit Function
    '    If CSng(Value) = 1# Then Acos = 0: Exit Function
    '    Acos = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
    '


    ' Radians.
    ' value is forced between 1 and -1.  This prevents IEEE rounding from causing any problems.
    Dim dRoot     As Single
    Dim D2        As Single
    '
    Select Case True
        Case Value > 1!: D2 = 1!
        Case Value < -1!: D2 = -1!
        Case Else: D2 = Value
    End Select
    '
    dRoot = Sqr(1! - D2 * D2)
    If Abs(dRoot) <> 0! Then
        Acos = PIh - Atn(D2 / dRoot)
    Else
        Acos = PIh - PIh * Sgn(D2)
    End If


End Function
