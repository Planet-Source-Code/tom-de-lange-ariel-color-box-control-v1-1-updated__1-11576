VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmColorSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Colour"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   915
      Index           =   1
      Left            =   5400
      ScaleHeight     =   855
      ScaleWidth      =   1515
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame fr 
      Caption         =   "RGB"
      Height          =   1575
      Index           =   1
      Left            =   5460
      TabIndex        =   14
      Top             =   1140
      Width           =   1575
      Begin VB.TextBox txtRed 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   17
         Text            =   "0"
         Top             =   360
         Width           =   675
      End
      Begin VB.TextBox txtGreen 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   16
         Text            =   "0"
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtBlue 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   15
         Text            =   "0"
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   20
         Top             =   405
         Width           =   300
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   19
         Top             =   765
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   18
         Top             =   1125
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   435
      Left            =   5580
      TabIndex        =   3
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Frame fr 
      Caption         =   "HSL to RGB Converter"
      Height          =   3195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5295
      Begin VB.PictureBox pSat 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   1260
         ScaleHeight     =   15
         ScaleMode       =   0  'User
         ScaleWidth      =   236.25
         TabIndex        =   11
         Top             =   1980
         Width           =   3840
      End
      Begin VB.TextBox txtSat 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "0"
         Top             =   2100
         Width           =   675
      End
      Begin VB.PictureBox pLum 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   1260
         ScaleHeight     =   15
         ScaleMode       =   0  'User
         ScaleWidth      =   236.25
         TabIndex        =   7
         Top             =   1140
         Width           =   3840
      End
      Begin VB.TextBox txtLum 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "0"
         Top             =   1380
         Width           =   675
      End
      Begin VB.PictureBox pHue 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   1260
         ScaleHeight     =   15
         ScaleMode       =   0  'User
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3840
      End
      Begin VB.TextBox txtHue 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "0"
         Top             =   600
         Width           =   675
      End
      Begin ComctlLib.Slider sldHue 
         Height          =   510
         Left            =   1095
         TabIndex        =   4
         Top             =   660
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   900
         _Version        =   327682
         LargeChange     =   48
         Max             =   240
         TickStyle       =   1
         TickFrequency   =   20
      End
      Begin ComctlLib.Slider sldLum 
         Height          =   510
         Left            =   1095
         TabIndex        =   8
         Top             =   1440
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   900
         _Version        =   327682
         LargeChange     =   15
         Max             =   240
         SelStart        =   120
         TickStyle       =   1
         TickFrequency   =   30
         Value           =   120
      End
      Begin ComctlLib.Slider sldSat 
         Height          =   510
         Left            =   1095
         TabIndex        =   12
         Top             =   2280
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   900
         _Version        =   327682
         LargeChange     =   15
         Max             =   240
         TickStyle       =   1
         TickFrequency   =   30
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Sat"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1860
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Lum"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1140
         Width           =   300
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Hue"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmColorSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Dim Hue, Sat, Lum
Dim hsl As HSLColor
Dim cRed, cGreen, cBlue
Dim dRed As Double, dGreen As Double, dBlue As Double

Sub PaintHue()
Dim Hue, Color
Dim cRed, cGreen, cBlue
For Hue = 0 To MaxHSL - 1
  Call Hue2RGB(Hue, cRed, cGreen, cBlue)
  Color = RGB(cRed, cGreen, cBlue)
  pHue.Line (Hue, 0)-(Hue + 1, 16), Color, BF
Next
End Sub
Sub PaintLum()
Dim Lum, Color
Dim r, g, b

For Lum = 0 To MaxHSL / 2
  r = (2 * dRed * Lum) / MaxHSL * MaxRGB
  g = (2 * dGreen * Lum) / MaxHSL * MaxRGB
  b = (2 * dBlue * Lum) / MaxHSL * MaxRGB
  Color = RGB(r, g, b)
  pLum.Line (Lum, 0)-(Lum + 1, 16), Color, BF
Next
For Lum = MaxHSL / 2 + 1 To MaxHSL - 1
  r = (2 * Lum * (1 - dRed) + 2 * MaxHSL * dRed - MaxHSL) / MaxHSL * MaxRGB
  g = (2 * Lum * (1 - dGreen) + 2 * MaxHSL * dGreen - MaxHSL) / MaxHSL * MaxRGB
  b = (2 * Lum * (1 - dBlue) + 2 * MaxHSL * dBlue - MaxHSL) / MaxHSL * MaxRGB

  'r = 2 * Lum - 2 * Lum * dRed + 512 * dRed - 256
  'g = 2 * Lum - 2 * Lum * dGreen + 512 * dGreen - 256
  'b = 2 * Lum - 2 * Lum * dBlue + 512 * dBlue - 256
  Color = RGB(r, g, b)
  pLum.Line (Lum, 0)-(Lum + 1, 16), Color, BF
Next

End Sub

Sub PaintPic()
Dim r, g, b
Dim Color

If Lum <= MaxHSL / 2 Then
  r = (Lum + (2 * dRed - 1) * Lum * Sat / MaxHSL) / MaxHSL * MaxRGB
  g = (Lum + (2 * dGreen - 1) * Lum * Sat / MaxHSL) / MaxHSL * MaxRGB
  b = (Lum + (2 * dBlue - 1) * Lum * Sat / MaxHSL) / MaxHSL * MaxRGB
Else
  r = (Lum + Sat * (2 * dRed - 1) + Lum * Sat * (1 - 2 * dRed) / MaxHSL) / MaxHSL * MaxRGB
  g = (Lum + Sat * (2 * dGreen - 1) + Lum * Sat * (1 - 2 * dGreen) / MaxHSL) / MaxHSL * MaxRGB
  b = (Lum + Sat * (2 * dBlue - 1) + Lum * Sat * (1 - 2 * dBlue) / MaxHSL) / MaxHSL * MaxRGB
End If
If r > MaxRGB Then r = MaxRGB
If g > MaxRGB Then g = MaxRGB
If b > MaxRGB Then b = MaxRGB
Color = RGB(r, g, b)

End Sub
Sub PaintPic2()
Dim Color

Color = HSLtoRGB(hsl)
pic(1).BackColor = Color
txtRed(1) = RGBRed(Color)
txtGreen(1) = RGBGreen(Color)
txtBlue(1) = RGBBlue(Color)

End Sub


Sub PaintSat()
Dim Sat, Color
Dim r, g, b

For Sat = 0 To MaxHSL
  r = MaxRGB / 2 + (Sat * dRed - Sat / 2) / MaxHSL * MaxRGB
  g = MaxRGB / 2 + (Sat * dGreen - Sat / 2) / MaxHSL * MaxRGB
  b = MaxRGB / 2 + (Sat * dBlue - Sat / 2) / MaxHSL * MaxRGB
  Color = RGB(r, g, b)
  pSat.Line (Sat, 0)-(Sat + 1, 16), Color, BF
Next

End Sub

Private Sub cmdOk_Click()
frmColorBoxTest.ArielColorBox1.SelectedColor = pic(1).BackColor
Unload Me

End Sub

Private Sub Form_Load()

Hue = 0
Lum = MaxHSL / 2
Sat = MaxHSL
hsl.Hue = Hue
hsl.Lum = Lum
hsl.Sat = Sat

txtHue = Hue
txtLum = Lum
txtSat = Sat
sldLum.Value = Lum
sldSat.Value = Sat
sldHue.Value = Hue

Call Hue2RGB(Hue, cRed, cGreen, cBlue)
dRed = cRed / MaxRGB
dGreen = cGreen / MaxRGB
dBlue = cBlue / MaxRGB

PaintHue
PaintSat
PaintLum
PaintPic
PaintPic2

End Sub

Private Sub sldHue_Change()
Hue = sldHue.Value Mod MaxHSL
hsl.Hue = Hue
txtHue = Hue
Call Hue2RGB(Hue, cRed, cGreen, cBlue)
dRed = cRed / MaxRGB
dGreen = cGreen / MaxRGB
dBlue = cBlue / MaxRGB

PaintSat
PaintLum
PaintPic2

End Sub

Private Sub sldHue_Scroll()
Hue = sldHue.Value Mod MaxHSL
hsl.Hue = Hue
txtHue = Hue

End Sub

Private Sub sldLum_Change()
Lum = sldLum.Value
If Lum > MaxHSL Then Lum = MaxHSL
hsl.Lum = Lum
txtLum = Lum

PaintPic
PaintPic2

End Sub

Private Sub sldLum_Scroll()
Lum = sldLum.Value
If Lum > MaxHSL Then Lum = MaxHSL
hsl.Lum = Lum
txtLum = Lum

End Sub


Private Sub sldSat_Change()
Sat = sldSat.Value
If Sat > MaxHSL Then Sat = MaxHSL
hsl.Sat = Sat
txtSat = Sat

PaintPic
PaintPic2

End Sub

Private Sub sldSat_Scroll()
If Sat > MaxHSL Then Sat = MaxHSL
hsl.Sat = Sat
txtSat = Sat

End Sub
