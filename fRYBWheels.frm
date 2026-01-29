VERSION 5.00
Begin VB.Form fRYBWheels 
   AutoRedraw      =   -1  'True
   Caption         =   "RGB - RYB Wheels"
   ClientHeight    =   9840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16185
   LinkTopic       =   "Form1"
   ScaleHeight     =   656
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1079
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   600
      ScaleHeight     =   519
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   895
      TabIndex        =   0
      Top             =   960
      Width           =   13455
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00AAAAAA&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00AAAAAA&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   960
         Shape           =   2  'Oval
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click and Drag Mouse to see the corrisponding point in the opposit Wheel."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   180
      Width           =   10095
   End
End
Attribute VB_Name = "fRYBWheels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pW!
Dim pH!

Private Sub Form_Load()
Label1.Left = (Me.ScaleWidth - Label1.Width) * 0.5
Shape1.Move -20, -20
Shape2.Move -20, -20

End Sub

Private Sub Form_Activate()
Cairo.ImageList.AddImage "WHEELS", App.Path & "\Images\HUEwheels.PNG"

PIC.Width = Cairo.ImageList.Item("WHEELS").Width
PIC.Height = Cairo.ImageList.Item("WHEELS").Height
pW = PIC.Width
pH = PIC.Height

PIC.Left = (Me.ScaleWidth - PIC.Width) * 0.5
PIC.Top = (Me.ScaleHeight - PIC.Height) * 0.5



Cairo.ImageList.Item("WHEELS").DrawToDC PIC.hDC


End Sub



Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cX!, cY!, maxR!
    Dim A!, R!
    Dim DX!, DY!
    Dim H!, S!, L!
    If Button = 1 Then
        maxR = pW * 0.215
        cY = pH * 0.55

        If X < pW * 0.5 Then
        
            cX = pW * 0.25
            DX = X - cX
            DY = Y - cY

            A = Atan2(DX, -DY)
            R = Sqr(DX * DX + DY * DY)
            If R > maxR Then R = maxR

            H = A / (Atn(1) * 8)
            If H < 0 Then H = H + 1
            L = R / maxR
            S = 0.85

            If H < 0.6666667! Then      '4/6
                H = H * 1.5!            '6/4              ' / 0.6666667
                H = ForwardHUEtransform(H)
                H = H * 0.6666667!      '4/6
            End If
            L = 1 - L
            Shape1.Move X - 8, Y - 8


            cX = cX + pW * 0.5
            cY = cY
            A = H * (Atn(1) * 8)
            cX = cX + Cos(A) * L * maxR
            cY = cY - Sin(A) * L * maxR
            Shape2.Move cX - 8, cY - 8

        Else
            cX = pW * 0.75
            DX = X - cX
            DY = Y - cY

            A = Atan2(DX, -DY)
            R = Sqr(DX * DX + DY * DY)
            If R > maxR Then R = maxR

            H = A / (Atn(1) * 8)
            If H < 0 Then H = H + 1
            L = R / maxR
            S = 0.85

            If H < 0.6666667! Then      '4/6
                H = H * 1.5!            '6/4              ' / 0.6666667
                H = InverseHUEtransform(H)
                H = H * 0.6666667!      '4/6
            End If
            L = 1 - L
            Shape2.Move X - 8, Y - 8

            cX = cX - pW * 0.5
            cY = cY

            A = H * (Atn(1) * 8)

            cX = cX + Cos(A) * L * maxR
            cY = cY - Sin(A) * L * maxR

            Shape1.Move cX - 8, cY - 8
        End If

    End If

End Sub
