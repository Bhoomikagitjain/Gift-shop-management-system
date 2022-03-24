VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ItemEntryFrm 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item_Name"
   ClientHeight    =   5385
   ClientLeft      =   5115
   ClientTop       =   3540
   ClientWidth     =   7590
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
   ScaleHeight     =   5385
   ScaleWidth      =   7590
   Begin VB.TextBox Qty 
      DataField       =   "Qty"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   4935
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1920
      Top             =   5880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   5880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Item_Name"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   360
         Top             =   480
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   735
         Left            =   960
         TabIndex        =   7
         Top             =   3840
         Width           =   5535
         Begin VB.CommandButton Delete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   4200
            TabIndex        =   11
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Cancle 
            Caption         =   "Cancle"
            Height          =   495
            Left            =   2880
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Save 
            Caption         =   "Save"
            Height          =   495
            Left            =   1560
            TabIndex        =   9
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton AddNew 
            Caption         =   "Add New"
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.TextBox Price 
         DataField       =   "PPU"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   2760
         Width           =   4935
      End
      Begin VB.TextBox IName 
         DataField       =   "ItemName"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock  Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Priece :"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty :"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
   End
End
Attribute VB_Name = "ItemEntryFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddNew_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
IName = ""
Qty = ""
Price = ""
End Sub

Private Sub Cancle_Click()
Adodc1.Refresh
End Sub

Private Sub Delete_Click()
Adodc1.Recordset.Delete
MsgBox ("1 Record Deleted")
End Sub



Private Sub First_Click()

End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * from Stock"
Adodc1.Refresh
End Sub






Private Sub Last_Click()

End Sub

Private Sub Next_Click()

End Sub

Private Sub Pre_Click()

End Sub

Private Sub Price_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
MsgBox "Please... Enter Numeric values"
KeyAscii = 0
End If
End Sub

Private Sub Qty_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
MsgBox "Please... Enter Numeric values"
KeyAscii = 0
End If
End Sub
Private Sub Save_Click()
Adodc1.Recordset.Update
MsgBox ("Saved!")
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
