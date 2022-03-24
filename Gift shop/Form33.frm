VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19080
   LinkTopic       =   "Form3"
   ScaleHeight     =   10170
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   480
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   840
      Top             =   4440
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   3480
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GIFT SHOP MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   18375
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   8520
      Picture         =   "Form33.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   12120
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING...."
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   11055
      Left            =   120
      Picture         =   "Form33.frx":2BA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Image2_Click()
Timer1.Enabled = True
Image2.Visible = False
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Value + 10
Label2.Visible = True
Label3.Visible = True
Label3.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then

    Form6.Show
    
Timer1.Enabled = False
End If
End Sub



Private Sub Timer2_Timer()
If Label1.ForeColor = vbRed Then
 Label1.ForeColor = vbBlue
 ElseIf Label1.ForeColor = vbBlue Then
 Label1.ForeColor = vbGreen
 ElseIf Label1.ForeColor = vbGreen Then
 Label1.ForeColor = vbYellow
 Else: Label1.ForeColor = vbRed
 End If
End Sub
