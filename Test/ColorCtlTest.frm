VERSION 5.00
Object = "{F07B2960-8F3C-11D4-9744-004F490561B3}#2.1#0"; "Ariel Color Ctrl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColorBoxTest 
   Caption         =   "Ariel Color Box Test"
   ClientHeight    =   2220
   ClientLeft      =   5085
   ClientTop       =   2835
   ClientWidth     =   5025
   FillColor       =   &H80000010&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   5025
   Begin ArielColorCtrl.ArielColorBox ArielColorBox1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Click on the dropdown button to select color"
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectedColor   =   12582912
      Palette         =   10
   End
   Begin VB.CheckBox chkPopup 
      Alignment       =   1  'Right Justify
      Caption         =   "Popup Enabled"
      Height          =   195
      Left            =   2700
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2700
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select color"
   End
   Begin VB.ComboBox cmbPalette 
      Height          =   315
      ItemData        =   "ColorCtlTest.frx":0000
      Left            =   3300
      List            =   "ColorCtlTest.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select Color"
      Height          =   375
      Left            =   3660
      TabIndex        =   0
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sep 2000 Rev 1.1"
      Height          =   195
      Index           =   3
      Left            =   3690
      TabIndex        =   2
      Top             =   1740
      Width           =   1290
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "by Tom de Lange"
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   6
      Top             =   1500
      Width           =   1230
   End
   Begin VB.Label lbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ColorCtlTest.frx":007C
      Height          =   1515
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   2490
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Palette"
      Height          =   195
      Index           =   0
      Left            =   2700
      TabIndex        =   3
      Top             =   180
      Width           =   510
   End
End
Attribute VB_Name = "frmColorBoxTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
DefLng A-N, P-Z
DefBool O

Dim HueSel          'Selected Hue
Dim MouseMove As Boolean
Dim MouseX As Single, MouseY As Single

Function HitColor(x As Single, Y As Single) As Long
'-------------------------------------
'Determine which hue was selected
'3 rows of 8 boxes
'Each box is 12 pixels, spaced 16
'Return the array number, -1 for none
'-------------------------------------
Dim row, col, i

HitColor = -1
If x >= 0 And x <= (7 * 16 + 12) Then
  If Y >= 0 And Y <= (2 * 16 + 12) Then
    col = x \ 16
    row = Y \ 16
    i = row * 8 + col
    HitColor = i
  End If
End If

End Function

Private Sub ArielColorBox1_Popup()
'-----------------------------------
'Substitute your popup code here
'Popup() event is fired only when
'PopupEnabled property is false
'-----------------------------------
On Error GoTo Err1
Const cdlCCRGBInit = &H1

dlg.Flags = cdlCCRGBInit
dlg.Color = ArielColorBox1.SelectedColor
dlg.ShowColor
ArielColorBox1.SelectedColor = dlg.Color
Exit Sub

Err1:
Exit Sub

End Sub


Private Sub chkPopup_Click()
ArielColorBox1.PopupEnabled = (chkPopup = 1)

End Sub

Private Sub cmbPalette_Click()
ArielColorBox1.Palette = cmbPalette.ListIndex
End Sub


Private Sub cmdDrawSat()
'-------------------------------------
'Draw sat/lum variaionts of
'selected hue in picture box
'-------------------------------------
End Sub

Private Sub cmdSelect_Click()
frmColorSel.Show

End Sub






Private Sub Form_Load()
cmbPalette.ListIndex = 10

End Sub


