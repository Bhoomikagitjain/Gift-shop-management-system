VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8700
   ClientLeft      =   1545
   ClientTop       =   1890
   ClientWidth     =   17130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   17130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2160
      Top             =   1320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9899950001"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   6720
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9899950000"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   6720
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HAND MADE GIFTS               AVAILABLE"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   6615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CONCTACTS US :-"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   9960
      TabIndex        =   3
      Top             =   6000
      Width           =   6255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FOR MORE DEATILS "
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   6000
      Width           =   9135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form2.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   6720
      TabIndex        =   1
      Top             =   1320
      Width           =   10215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "HANDMADE GIFT AVAILABLE              FOR YOUR SPECIAL ONECES"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   12135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   9615
      Left            =   -720
      Picture         =   "Form2.frx":03A9
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   17895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If Label1.ForeColor = vbWhite Then
 Label1.ForeColor = vbBlue
 ElseIf Label1.ForeColor = vbBlue Then
 Label1.ForeColor = vbGreen
 ElseIf Label1.ForeColor = vbGreen Then
 Label1.ForeColor = vbYellow
 Else: Label1.ForeColor = vbWhite
 End If
End Sub
