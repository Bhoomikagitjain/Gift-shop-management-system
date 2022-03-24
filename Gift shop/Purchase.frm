VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchaseFrm 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase"
   ClientHeight    =   8040
   ClientLeft      =   4665
   ClientTop       =   1770
   ClientWidth     =   10920
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
   ScaleHeight     =   8040
   ScaleWidth      =   10920
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6360
      Top             =   7560
      Visible         =   0   'False
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
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4080
      Top             =   7560
      Visible         =   0   'False
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5160
      Top             =   7560
      Visible         =   0   'False
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
      Left            =   2880
      Top             =   7560
      Visible         =   0   'False
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
   Begin MSComCtl2.DTPicker dt 
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110821379
      CurrentDate     =   43227
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800080&
      Height          =   975
      Left            =   1560
      TabIndex        =   12
      Top             =   5760
      Width           =   7455
      Begin VB.CommandButton Cancle 
         Caption         =   "Cancle"
         Height          =   615
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit Stock Details"
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "New Entry"
         Height          =   615
         Left            =   2160
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   615
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox tot 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   4560
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Rate 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3840
      Width           =   8055
   End
   Begin VB.TextBox Qty 
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3000
      Width           =   8175
   End
   Begin VB.TextBox CompName 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   8175
   End
   Begin VB.ComboBox ItemName 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   8175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "                                   Purchase Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   480
         Top             =   240
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Details"
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
         Left            =   3600
         TabIndex        =   17
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate :"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name :"
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name :"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "PurchaseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancle_Click()
CompName.Text = ""
  Qty.Text = ""
 Rate.Text = ""
 Tot.Text = ""
 ItemName.Text = ""
  
End Sub



Private Sub Command1_Click()
Tot = Qty * Rate
End Sub

Private Sub Command2_Click()
ItemEntryFrm.Show
End Sub

Private Sub Command3_Click()
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc3.RecordSource = "Select * from Stock where ItemName='" + ItemName + "'"
Adodc3.RecordSource = "Select * from Purchase"
Adodc3.Refresh
Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields(0) = Dt
Adodc3.Recordset.Fields(1) = ItemName
Adodc3.Recordset.Fields(2) = CompName
Adodc3.Recordset.Fields(3) = Qty
Adodc3.Recordset.Fields(4) = Rate
Adodc3.Recordset.Fields(5) = Tot
Adodc3.Recordset.Update

Adodc4.RecordSource = "Select * from Stock where ItemName='" + ItemName + "'"
Adodc4.Refresh
i = Adodc4.Recordset.Fields("Qty")
Adodc4.Recordset.Fields("Qty") = i + Qty
Adodc4.Recordset.Update
MsgBox ("Success")
Adodc1.Refresh
Adodc2.Refresh
End Sub

Private Sub Command4_Click()
Adodc1.Refresh
Adodc3.Refresh
CompName.Text = ""
  Qty.Text = ""
 Rate.Text = ""
 Tot.Text = ""
 ItemName.Text = ""
End Sub



Private Sub Form_Load()
Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * from Stock"
Adodc1.Refresh
ItemName.Clear
Do Until Adodc1.Recordset.EOF
ItemName.AddItem Adodc1.Recordset.Fields("ItemName")
Adodc1.Recordset.MoveNext
Loop

Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc3.RecordSource = "Select * from Purchase"
Adodc3.Refresh

End Sub

Private Sub ItemName_Click()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc2.RecordSource = "Select * from Stock where ItemName='" + ItemName + "'"
Adodc2.Refresh
Rate = Adodc2.Recordset.Fields("PPU")
End Sub

Private Sub Text5_Change()
End Sub

Private Sub Qty_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
MsgBox "Please... Enter Numeric values"
KeyAscii = 0
End If
End Sub

Private Sub Rate_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
MsgBox "Please... Enter Numeric values"
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
If Label3.ForeColor = vbRed Then
 Label3.ForeColor = vbBlue
 ElseIf Label3.ForeColor = vbBlue Then
 Label3.ForeColor = vbGreen
 ElseIf Label3.ForeColor = vbGreen Then
 Label3.ForeColor = vbYellow
 Else: Label3.ForeColor = vbRed
 End If
End Sub
