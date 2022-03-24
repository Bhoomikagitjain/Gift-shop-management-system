VERSION 5.00
Begin VB.MDIForm Main 
   BackColor       =   &H8000000C&
   Caption         =   "Gift Shop"
   ClientHeight    =   8340
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16305
   LinkTopic       =   "MDIForm1"
   Picture         =   "Main.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   4200
   End
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   564
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   510
      ScaleWidth      =   16245
      TabIndex        =   0
      Top             =   0
      Width           =   16308
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   11775
         TabIndex        =   1
         Top             =   0
         Width           =   11775
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7440
            TabIndex        =   5
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   4
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Current Time"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9120
            TabIndex        =   3
            Top             =   120
            Width           =   2580
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Current Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1920
            TabIndex        =   2
            Top             =   120
            Width           =   2325
         End
      End
   End
   Begin VB.Menu ss 
      Caption         =   "Supplier"
      Begin VB.Menu S2 
         Caption         =   "Suppliers Details"
      End
      Begin VB.Menu SuplInfo 
         Caption         =   "Supplier Report"
      End
   End
   Begin VB.Menu customerDa 
      Caption         =   "Customer"
      Begin VB.Menu customerD 
         Caption         =   "Customer Details"
      End
      Begin VB.Menu cust 
         Caption         =   "Customer Report"
      End
   End
   Begin VB.Menu stockd 
      Caption         =   "Stock"
      Begin VB.Menu npe 
         Caption         =   "New Product Entry"
      End
      Begin VB.Menu stockDetails 
         Caption         =   "Stock Report"
      End
   End
   Begin VB.Menu purchased 
      Caption         =   "Purchase"
      Begin VB.Menu PPurchase 
         Caption         =   "Purchase"
      End
      Begin VB.Menu purchaseDetail 
         Caption         =   "Purchase report"
      End
   End
   Begin VB.Menu salesd 
      Caption         =   "Sales"
      Begin VB.Menu salesdet 
         Caption         =   "Sales"
      End
      Begin VB.Menu salesDetails 
         Caption         =   "Sales Report"
      End
   End
   Begin VB.Menu h 
      Caption         =   "Help"
      Begin VB.Menu h1 
         Caption         =   "About shop"
      End
      Begin VB.Menu h2 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cust_Click()
CustomerFrm.Show
End Sub

Private Sub customerD_Click()
Form4.Show
End Sub

Private Sub Exit_Click()
Form5.Show
End Sub



Private Sub lg_Click()
LoginFrm.Show
logout.Enabled = False
SuplInfo.Enabled = False
report.Enabled = False
salesd.Enabled = False
purchased.Enabled = False
customerD.Enabled = False
stockd.Enabled = False
End Sub

Private Sub h1_Click()
Form2.Show
End Sub

Private Sub h2_Click()
VBA.Shell "Notepad.exe"
End Sub

Private Sub npe_Click()
ItemEntryFrm.Show
End Sub

Private Sub PPurchase_Click()
PurchaseFrm.Show
End Sub

Private Sub Prj_Click()

End Sub

Private Sub stfJoin_Click()
AddEmpDet.Show
End Sub

Private Sub staffDetails_Click()

End Sub

Private Sub pr_Click()
PurchaseDetFrm.Show
End Sub

Private Sub pur_Click()
PurchaseFrm.Show
End Sub


Private Sub purchaseDetail_Click()
PurchaseDetFrm.Show
End Sub


Private Sub s2_Click()
Form1.Show
End Sub

Private Sub salesdet_Click()
SalesFrm.Show
End Sub

Private Sub salesDetails_Click()
SalesDetFrm.Show
End Sub

Private Sub sl_Click()
SalesFrm.Show
End Sub

Private Sub sr_Click()
SalesDetFrm.Show
End Sub



Private Sub stockDetails_Click()
StockDetFrm.Show
End Sub

Private Sub SuplInfo_Click()
SuppliarsFrm.Show
End Sub

Private Sub Timer1_Timer()
lblDate.Caption = Format(Date, "mmmm dd, yyyy")
lblTime.Caption = Format(Time, "hh:dd:ss am/pm")
End Sub

