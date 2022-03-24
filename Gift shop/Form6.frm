VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   250
      Left            =   14400
      Top             =   6840
   End
   Begin VB.Timer Timer2 
      Interval        =   900
      Left            =   14640
      Top             =   7560
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   13800
      Top             =   7440
   End
   Begin VB.Image Image2 
      Height          =   4455
      Left            =   8640
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form6.frx":44CA
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   10320
      TabIndex        =   0
      Top             =   600
      Width           =   11295
   End
   Begin VB.Image Image1 
      Height          =   10935
      Left            =   0
      Picture         =   "Form6.frx":459E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Timer1_Timer()
 Main.Show
    Unload Me
    Main.SuplInfo.Enabled = True
    Main.salesd.Enabled = True
    Main.purchased.Enabled = True
    Main.customerD.Enabled = True
    Main.stockd.Enabled = True
Unload Me
End Sub
Private Sub Timer2_Timer()
Label3.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Timer3_Timer()
If Label1.ForeColor = vbRed Then
 Label1.ForeColor = vbBlue
 ElseIf Label1.ForeColor = vbBlue Then
 Label1.ForeColor = vbGreen
 ElseIf Label1.ForeColor = vbGreen Then
 Label1.ForeColor = vbBlack
 Else: Label1.ForeColor = vbRed
 End If
End Sub


