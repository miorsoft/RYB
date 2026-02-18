VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Mixing Colors like Pigments"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   769
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar HPerc 
      Height          =   375
      Left            =   3240
      Max             =   100
      TabIndex        =   22
      Top             =   3960
      Width           =   5175
   End
   Begin VB.PictureBox PIC5 
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
      Height          =   975
      Left            =   3120
      ScaleHeight     =   945
      ScaleWidth      =   3225
      TabIndex        =   20
      Top             =   6600
      Width           =   3255
   End
   Begin VB.PictureBox PIC4 
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
      Height          =   975
      Left            =   3120
      ScaleHeight     =   945
      ScaleWidth      =   3225
      TabIndex        =   9
      Top             =   7920
      Width           =   3255
   End
   Begin VB.PictureBox PIC1 
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
      Height          =   975
      Left            =   360
      ScaleHeight     =   945
      ScaleWidth      =   3225
      TabIndex        =   8
      Top             =   2760
      Width           =   3255
   End
   Begin VB.PictureBox PIC3 
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
      Height          =   975
      Left            =   3120
      ScaleHeight     =   945
      ScaleWidth      =   3225
      TabIndex        =   7
      Top             =   5160
      Width           =   3255
   End
   Begin VB.PictureBox PIC2 
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
      Height          =   975
      Left            =   5760
      ScaleHeight     =   945
      ScaleWidth      =   3225
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.HScrollBar HB 
      Height          =   495
      Index           =   1
      Left            =   4920
      Max             =   255
      TabIndex        =   5
      Top             =   1680
      Width           =   3615
   End
   Begin VB.HScrollBar HG 
      Height          =   495
      Index           =   1
      Left            =   4920
      Max             =   255
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.HScrollBar HR 
      Height          =   495
      Index           =   1
      Left            =   4920
      Max             =   255
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
   Begin VB.HScrollBar HB 
      Height          =   495
      Index           =   0
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.HScrollBar HG 
      Height          =   495
      Index           =   0
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.HScrollBar HR 
      Height          =   495
      Index           =   0
      Left            =   120
      Max             =   255
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lPerc 
      Alignment       =   2  'Center
      Caption         =   "RGB 1"
      Height          =   375
      Left            =   8520
      TabIndex        =   23
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lResult3 
      Alignment       =   2  'Center
      Caption         =   "R3"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   6240
      Width           =   9015
   End
   Begin VB.Label lResult1 
      Alignment       =   2  'Center
      Caption         =   "R1"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   9015
   End
   Begin VB.Label lResult2 
      Alignment       =   2  'Center
      Caption         =   "R2"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   7560
      Width           =   9015
   End
   Begin VB.Label LR 
      Caption         =   "Label2"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   17
      Top             =   360
      Width           =   615
   End
   Begin VB.Label LB 
      Caption         =   "Label2"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   16
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LG 
      Caption         =   "Label2"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   15
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label LR 
      Caption         =   "Label2"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label LB 
      Caption         =   "Label2"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LG 
      Caption         =   "Label2"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RGB 2"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lRGB1 
      Alignment       =   2  'Center
      Caption         =   "RGB 1"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2400
      Width           =   3255
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://stackoverflow.com/questions/1351442/is-there-an-algorithm-for-color-mixing-that-works-like-mixing-real-colors

Dim R1!, G1!, B1!
Attribute G1.VB_VarUserMemId = 1073938432
Attribute B1.VB_VarUserMemId = 1073938432
Dim R2!, G2!, B2!
Attribute R2.VB_VarUserMemId = 1073938435
Attribute G2.VB_VarUserMemId = 1073938435
Attribute B2.VB_VarUserMemId = 1073938435

Dim R3!, G3!, B3!
Attribute R3.VB_VarUserMemId = 1073938436
Attribute G3.VB_VarUserMemId = 1073938436
Attribute B3.VB_VarUserMemId = 1073938436

Private Sub Form_Activate()
    HR.Item(0).Value = 255
    HG.Item(0).Value = 255
    HB.Item(0).Value = 1

    HR.Item(1).Value = 1
    HG.Item(1).Value = 1
    HB.Item(1).Value = 255
    HPerc = 50

End Sub

Private Sub Form_Load()
    Dim I         As Long

    Dim rr!, GG!, bB!

    LoadTables


    PigmentMix 1, 0.95, 0, _
               0.1, 0.1, 1, 0.5, rr, GG, bB

    PigmentMix 0.5, 0.5, 0.5, _
               0.1, 0.8, 1, 0.5, rr, GG, bB

    Exit Sub

    CreateWheels
    ''        End

    Dim R!, G!, B!
    Dim AA&, Re&
    Dim J&, K&
    For I = 0 To 7
        For J = I + 1 To 7
            R1 = -((I And &H1) = 1)
            G1 = -((I And &H2) = 2)
            B1 = -((I And &H4) = 4)

            R2 = -((J And &H1) = 1)
            G2 = -((J And &H2) = 2)
            B2 = -((J And &H4) = 4)
            CreateGradients R1, G1, B1, R2, G2, B2, K
            K = K + 1
        Next
    Next

    'Dim St!
    'St = 1
    '
    'For R1 = 0 To 1 Step St
    'For G1 = 0 To 1 Step St
    'For B1 = 0 To 1 Step St
    '
    'For R2 = 0 To 1 Step St
    'For G2 = 0 To 1 Step St
    'For B2 = 0 To 1 Step St
    'If R1 <> R2 Or G1 <> G2 Or B1 <> B2 Then
    '            CreateGradients R1, G1, B1, R2, G2, B2, K
    '            K = K + 1
    'End If
    'Next
    'Next
    'Next
    'Next
    'Next
    'Next

    End


End Sub

Private Sub SetMIXEDColor()
    Dim P         As Double
    Dim Q         As Double

    P = HPerc.Value * 0.01
    Q = 1 - P

    PIC1.BackColor = RGB(R1, G1, B1)
    PIC2.BackColor = RGB(R2, G2, B2)


    PigmentMixREEXRE R1, G1, B1, R2, G2, B2, HPerc.Value * 0.01, R3, G3, B3
    PIC3.BackColor = RGB(R3, G3, B3)
    lResult1.Caption = "MY Pigment Mix (Subtractive): " & Round(R3) & "-" & Round(G3) & "-" & Round(B3)

    PIC4.BackColor = RGB(R1 * Q + P * R2, G1 * Q + P * G2, B1 * Q + P * B2)
    lResult2.Caption = "Light Mix (Additive): " & Round(R1 * Q + P * R2) & "-" & Round(G1 * Q + P * G2) & "-" & Round(B1 * Q + P * B2)


    PigmentMix R1, G1, B1, R2, G2, B2, P, R3, G3, B3
    PIC5.BackColor = RGB(R3, G3, B3)
    lResult3.Caption = "Pigment Mix (Subtractive): " & Round(R3) & "-" & Round(G3) & "-" & Round(B3)


End Sub

Private Sub Form_Resize()
    Dim W         As Long
    Dim I         As Long
    W = Me.ScaleWidth

    If W < 10 Then Exit Sub

    HR(0).Width = W * 0.5 - LR(0).Width - 8
    HR(0).Left = 8
    LR(0).Left = HR(0).Left + HR(0).Width
    HR(1).Width = W * 0.5 - LR(1).Width - 8
    HR(1).Left = W * 0.5 + 8
    LR(1).Left = HR(1).Left + HR(1).Width

    For I = 0 To 1
        HG(I).Width = HR(I).Width
        HG(I).Left = HR(I).Left
        LG(I).Left = LR(I).Left
        HB(I).Width = HR(I).Width
        HB(I).Left = HR(I).Left
        LB(I).Left = LR(I).Left
    Next

    PIC1.Left = W * 0.25 - PIC1.Width * 0.5
    PIC2.Left = W * 0.75 - PIC2.Width * 0.5
    lRGB1.Left = W * 0.25 - lRGB1.Width * 0.5
    Label1.Left = W * 0.75 - Label1.Width * 0.5

    lResult1.Left = W * 0.5 - lResult1.Width * 0.5
    lResult2.Left = W * 0.5 - lResult2.Width * 0.5
    lResult3.Left = W * 0.5 - lResult3.Width * 0.5

    PIC3.Left = W * 0.5 - PIC3.Width * 0.5
    PIC4.Left = W * 0.5 - PIC4.Width * 0.5
    PIC5.Left = W * 0.5 - PIC5.Width * 0.5

    HPerc.Left = W * 0.5 - HPerc.Width * 0.5
    lPerc.Left = HPerc.Left + HPerc.Width + 2
End Sub

Private Sub HPerc_Change()
    lPerc = HPerc & "%"
    SetMIXEDColor
End Sub

Private Sub HPerc_Scroll()
    lPerc = HPerc & "%"
    SetMIXEDColor
End Sub

Private Sub HR_Change(Index As Integer)
    HR_Scroll Index
End Sub
Private Sub HG_Change(Index As Integer)
    HG_Scroll Index
End Sub

Private Sub HB_Change(Index As Integer)
    HB_Scroll Index
End Sub

Private Sub HB_Scroll(Index As Integer)
    If Index = 0 Then
        B1 = HB.Item(Index).Value
    Else
        B2 = HB.Item(Index).Value
    End If
    LB(Index) = HB.Item(Index).Value
    SetMIXEDColor
End Sub


Private Sub HG_Scroll(Index As Integer)
    If Index = 0 Then
        G1 = HG.Item(Index).Value
    Else
        G2 = HG.Item(Index).Value
    End If
    LG(Index) = HG.Item(Index).Value

    SetMIXEDColor
End Sub

Private Sub HR_Scroll(Index As Integer)
    If Index = 0 Then
        R1 = HR.Item(Index).Value
    Else
        R2 = HR.Item(Index).Value
    End If
    LR(Index) = HR.Item(Index).Value

    SetMIXEDColor
End Sub

