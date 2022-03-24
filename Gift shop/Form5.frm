VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18165
   LinkTopic       =   "Form5"
   ScaleHeight     =   10335
   ScaleWidth      =   18165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "                 Exit"
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
      Height          =   1815
      Left            =   16080
      TabIndex        =   0
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   10935
      Left            =   0
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
MsgBox "Thank you"
End
Unload Me
End Sub



