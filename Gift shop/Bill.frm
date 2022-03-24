VERSION 5.00
Begin VB.Form BillFrm 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   6840
      Picture         =   "Bill.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1515
      TabIndex        =   23
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   240
      Picture         =   "Bill.frx":3AD0
      ScaleHeight     =   1035
      ScaleWidth      =   1515
      TabIndex        =   22
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Print 
      Caption         =   "Print"
      Height          =   375
      Left            =   7440
      TabIndex        =   17
      Top             =   7920
      Width           =   855
   End
   Begin VB.Line Line9 
      X1              =   2400
      X2              =   5535
      Y1              =   1080
      Y2              =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Samta Coloni , Raipur , C.G. )"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gift Shop"
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
      Left            =   5400
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      Height          =   255
      Left            =   5760
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   8760
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BHAVANA "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   18
      Top             =   0
      Width           =   3735
   End
   Begin VB.Line Line7 
      X1              =   6360
      X2              =   6360
      Y1              =   2880
      Y2              =   7800
   End
   Begin VB.Line Line6 
      X1              =   4200
      X2              =   4200
      Y1              =   2880
      Y2              =   7320
   End
   Begin VB.Line Line5 
      X1              =   2160
      X2              =   2160
      Y1              =   2880
      Y2              =   7320
   End
   Begin VB.Label Atot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tot"
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   8760
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   8760
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Tot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tot"
      Height          =   495
      Left            =   6960
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Rt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rt"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Qty 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label ItemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8760
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Nm 
      BackStyle       =   0  'Transparent
      Caption         =   "Nm"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Tm 
      BackStyle       =   0  'Transparent
      Caption         =   "Tm"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Dt 
      BackStyle       =   0  'Transparent
      Caption         =   "Dt"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label SrNo 
      BackStyle       =   0  'Transparent
      Caption         =   "SrNo"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sr. No. :"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "BillFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Print_Click()
PrintForm
End Sub
