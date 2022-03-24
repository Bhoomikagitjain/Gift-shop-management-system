VERSION 5.00
Begin VB.Form LoginFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Login"
   ClientHeight    =   9255
   ClientLeft      =   7950
   ClientTop       =   4635
   ClientWidth     =   18780
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "LoginFrm.frx":0000
   ScaleHeight     =   9255
   ScaleWidth      =   18780
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   6720
      Top             =   3960
   End
   Begin VB.TextBox Pass 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   9600
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6480
      Width           =   5775
   End
   Begin VB.TextBox UName 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   2
      Top             =   5520
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   0
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOGIN FORM"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   8640
      TabIndex        =   7
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GIFT SHOP MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   2880
      TabIndex        =   6
      Top             =   2880
      Width           =   16335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   5640
      Width           =   3615
   End
End
Attribute VB_Name = "LoginFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If UName = "admin" And Pass = "000" Then
Form3.Show
Else
    MsgBox ("done..!")
End If
End Sub

Private Sub Command2_Click()
If MsgBox("You are about to quit this application. Are you sure?", vbOKCancel + vbInformation, "Confirm Logoff") = vbOK Then
    End
Else
    Exit Sub
End If
End Sub


Private Sub Timer1_Timer()
If Label4.ForeColor = vbRed Then
 Label4.ForeColor = vbBlue
 ElseIf Label4.ForeColor = vbBlue Then
 Label4.ForeColor = vbGreen
 ElseIf Label4.ForeColor = vbGreen Then
 Label4.ForeColor = vbYellow
 Else: Label4.ForeColor = vbRed
 End If
End Sub
