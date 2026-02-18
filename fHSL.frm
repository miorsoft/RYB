VERSION 5.00
Begin VB.Form fHSL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HSL Test"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   930
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar sRGB 
      Height          =   735
      Index           =   2
      Left            =   840
      Max             =   255
      TabIndex        =   17
      Top             =   6000
      Width           =   3615
   End
   Begin VB.HScrollBar sRGB 
      Height          =   735
      Index           =   1
      Left            =   840
      Max             =   255
      TabIndex        =   15
      Top             =   5160
      Width           =   3615
   End
   Begin VB.HScrollBar sRGB 
      Height          =   735
      Index           =   0
      Left            =   840
      Max             =   255
      TabIndex        =   13
      Top             =   4320
      Width           =   3615
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   6120
      ScaleHeight     =   2385
      ScaleWidth      =   4665
      TabIndex        =   9
      Top             =   600
      Width           =   4695
   End
   Begin VB.HScrollBar HSL 
      Height          =   735
      Index           =   2
      Left            =   840
      Max             =   100
      TabIndex        =   6
      Top             =   2280
      Width           =   3615
   End
   Begin VB.HScrollBar HSL 
      Height          =   735
      Index           =   1
      Left            =   840
      Max             =   100
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.HScrollBar HSL 
      Height          =   735
      Index           =   0
      Left            =   840
      Max             =   100
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "HSL 3D Cilinder"
      Height          =   375
      Left            =   6120
      TabIndex        =   28
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Dot
      X1              =   888
      X2              =   792
      Y1              =   352
      Y2              =   344
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   880
      X2              =   784
      Y1              =   400
      Y2              =   392
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   864
      X2              =   768
      Y1              =   424
      Y2              =   416
   End
   Begin VB.Shape ShapeOval 
      BorderWidth     =   2
      Height          =   375
      Left            =   6120
      Shape           =   2  'Oval
      Top             =   6000
      Width           =   975
   End
   Begin VB.Shape ShapeHSL2 
      BackColor       =   &H00001111&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   6360
      Shape           =   2  'Oval
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label lcube 
      Alignment       =   2  'Center
      Caption         =   "L"
      Height          =   255
      Index           =   2
      Left            =   10560
      TabIndex        =   27
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lcube 
      Alignment       =   2  'Center
      Caption         =   "S"
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   26
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lcube 
      Alignment       =   2  'Center
      Caption         =   "H"
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   25
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "HSL 3D cube"
      Height          =   375
      Left            =   9960
      TabIndex        =   24
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Line LineL2 
      BorderStyle     =   3  'Dot
      X1              =   656
      X2              =   560
      Y1              =   456
      Y2              =   448
   End
   Begin VB.Line LineS2 
      BorderStyle     =   3  'Dot
      X1              =   664
      X2              =   568
      Y1              =   440
      Y2              =   432
   End
   Begin VB.Line LineH2 
      BorderStyle     =   3  'Dot
      X1              =   664
      X2              =   568
      Y1              =   424
      Y2              =   416
   End
   Begin VB.Shape ShapeHSL 
      BackColor       =   &H00001111&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   8400
      Shape           =   2  'Oval
      Top             =   5040
      Width           =   135
   End
   Begin VB.Line LineL 
      BorderWidth     =   2
      X1              =   592
      X2              =   504
      Y1              =   368
      Y2              =   440
   End
   Begin VB.Line LineH 
      BorderWidth     =   2
      X1              =   704
      X2              =   592
      Y1              =   424
      Y2              =   368
   End
   Begin VB.Line LineS 
      BorderWidth     =   2
      X1              =   592
      X2              =   592
      Y1              =   264
      Y2              =   368
   End
   Begin VB.Label Label2 
      Caption         =   "Change RGB value to see corrisponding HSL"
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   3720
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Change HSL value to see corrisponding RGB"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label llRGB 
      Caption         =   "Blue"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label llRGB 
      Caption         =   "Green"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label llRGB 
      Caption         =   "Red"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lRGB 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   18
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lRGB 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lRGB 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   14
      Top             =   4560
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   24
      X2              =   760
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Label lHSL2 
      Alignment       =   1  'Right Justify
      Caption         =   "H"
      Height          =   255
      Index           =   5
      Left            =   10920
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lHSL2 
      Alignment       =   1  'Right Justify
      Caption         =   "H"
      Height          =   255
      Index           =   4
      Left            =   10920
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lHSL2 
      Alignment       =   1  'Right Justify
      Caption         =   "H"
      Height          =   255
      Index           =   3
      Left            =   10920
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lHSL2 
      Alignment       =   2  'Center
      Caption         =   "H"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lSHL 
      Alignment       =   2  'Center
      Caption         =   "L"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lHSL2 
      Alignment       =   2  'Center
      Caption         =   "H"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lSHL 
      Alignment       =   2  'Center
      Caption         =   "S"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lHSL2 
      Alignment       =   2  'Center
      Caption         =   "H"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lSHL 
      Alignment       =   2  'Center
      Caption         =   "H"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "fHSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim H!
Dim S!
Dim L!
Dim R!
Dim G!
Dim B!
Dim LL!
Dim cX!, cY!
Attribute cY.VB_VarUserMemId = 1073938439

Private MOUSEY    As Boolean
Attribute MOUSEY.VB_VarUserMemId = 1073938441


Private Sub Form_Activate()
    HSL(0) = 0
    HSL(1) = 100
    HSL(2) = 50

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MOUSEY = Y > Line1.Y1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub HSL_Change(Index As Integer)
    HSL_Scroll Index

End Sub

Private Sub HSL_Scroll(Index As Integer)
    If Not MOUSEY Then
        If Index = 0 Then H = HSL(Index) * 0.01: lHSL2(Index) = H
        If Index = 1 Then S = HSL(Index) * 0.01: lHSL2(Index) = S
        If Index = 2 Then L = HSL(Index) * 0.01: lHSL2(Index) = L

        HSL2RGBmy H, S, L, R, G, B

        lHSL2(3) = Round(R * 255)
        lHSL2(4) = Round(G * 255)
        lHSL2(5) = Round(B * 255)

        sRGB(0) = R * 255
        sRGB(1) = G * 255
        sRGB(2) = B * 255
        lRGB(0) = Round(R * 255)
        lRGB(1) = Round(G * 255)
        lRGB(2) = Round(B * 255)

        PIC.BackColor = RGB(R * 255, G * 255, B * 255)

        MoveHSLdot

    End If

End Sub

Private Sub sRGB_Change(Index As Integer)
    sRGB_Scroll Index
End Sub

Private Sub sRGB_Scroll(Index As Integer)

    If MOUSEY Then

        If Index = 0 Then R = sRGB(Index) * 1 / 255: lRGB(Index) = Round(R * 255)
        If Index = 1 Then G = sRGB(Index) * 1 / 255: lRGB(Index) = Round(G * 255)
        If Index = 2 Then B = sRGB(Index) * 1 / 255: lRGB(Index) = Round(B * 255)

        RGB2HSLmy R, G, B, H, S, L

        lHSL2(3) = Round(R * 255)
        lHSL2(4) = Round(G * 255)
        lHSL2(5) = Round(B * 255)

        HSL(0) = H * 100
        HSL(1) = S * 100
        HSL(2) = L * 100
        lHSL2(0) = Round(H, 2)
        lHSL2(1) = Round(S, 2)
        lHSL2(2) = Round(L, 2)

        PIC.BackColor = RGB(R * 255, G * 255, B * 255)

        MoveHSLdot
    End If



End Sub

Private Sub Form_Load()


    LL = 120

    cX = Me.ScaleWidth * 0.85
    cY = Me.ScaleHeight * 0.75 - 4


    LineL.X1 = cX: LineL.Y1 = cY
    LineL.X2 = cX: LineL.Y2 = cY - LL

    LineH.X1 = cX: LineH.Y1 = cY
    LineH.X2 = cX + LL * 0.866025403784438: LineH.Y2 = cY + LL * 0.5

    LineS.X1 = cX: LineS.Y1 = cY
    LineS.X2 = cX - LL * 0.866025403784438: LineS.Y2 = cY + LL * 0.5

    lcube(0).Left = LineH.X2
    lcube(0).Top = LineH.Y2
    lcube(1).Left = LineS.X2
    lcube(1).Top = LineS.Y2
    lcube(2).Left = LineL.X2 + 2
    lcube(2).Top = LineL.Y2




End Sub
Private Sub MoveHSLdot()
    Dim X!, Y!
    Dim PX!, PY!
    Const PI2     As Single = 6.28318530717959


    X = H * 0.866025403784438 - S * 0.866025403784438
    Y = H * 0.5 + S * 0.5 - L
    PX = cX + X * LL
    PY = cY + Y * LL

    ShapeHSL.Left = -4.5 + PX
    ShapeHSL.Top = -4.5 + PY


    LineH2.X2 = PX
    LineH2.Y2 = PY
    LineH2.X1 = PX - LL * H * 0.866025403784438
    LineH2.Y1 = PY - LL * H * 0.5

    LineS2.X2 = PX
    LineS2.Y2 = PY
    LineS2.X1 = PX + LL * S * 0.866025403784438
    LineS2.Y1 = PY - LL * S * 0.5

    LineL2.X2 = PX
    LineL2.Y2 = PY
    LineL2.X1 = PX
    LineL2.Y1 = PY + LL * L


    ShapeOval.Left = Me.ScaleWidth * 0.5


    ShapeOval.Width = LL
    ShapeOval.Height = LL * 0.5

    PX = ShapeOval.Left + LL * 0.5
    PY = ShapeOval.Top + LL * 0.2

    ShapeHSL2.Left = -4.5 + PX + Cos(-H * PI2) * LL * 0.5 * S
    ShapeHSL2.Top = -4.5 + PY + Sin(-H * PI2) * LL * 0.25 * S - L * LL

    Line2.X1 = PX
    Line2.Y1 = PY
    Line2.X2 = PX
    Line2.Y2 = PY - LL

    PX = ShapeHSL2.Left + 4.5
    PY = ShapeHSL2.Top + 9

    Line3.X1 = PX
    Line3.Y1 = PY
    Line3.X2 = PX
    Line3.Y2 = PY + L * LL

    Line4.X1 = PX
    Line4.Y1 = PY - 4.5
    Line4.X2 = PX - Cos(-H * PI2) * LL * 0.5 * S
    Line4.Y2 = PY - 4.5 - Sin(-H * PI2) * LL * 0.25 * S


End Sub

