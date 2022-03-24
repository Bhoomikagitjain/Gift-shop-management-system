VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalesFrm 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales"
   ClientHeight    =   7575
   ClientLeft      =   2355
   ClientTop       =   2400
   ClientWidth     =   11880
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
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   Begin MSComCtl2.DTPicker Dt 
      Height          =   510
      Left            =   6600
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   900
      _Version        =   393216
      Format          =   109641729
      CurrentDate     =   43958
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   735
      Left            =   2880
      TabIndex        =   16
      Top             =   6000
      Width           =   4935
      Begin VB.CommandButton Cancle 
         Caption         =   "Cancle"
         Height          =   495
         Left            =   3360
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Save 
         Caption         =   "Save"
         Height          =   495
         Left            =   1800
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton New 
         Caption         =   "New"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox Tot 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton CTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox CustName 
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   3720
      Width           =   6135
   End
   Begin VB.TextBox Qty 
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox AQty 
      DataField       =   "Qty"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   6135
   End
   Begin VB.TextBox Rate 
      DataField       =   "PPU"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox CompName 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3120
      Width           =   6135
   End
   Begin VB.TextBox ItemName 
      DataField       =   "ItemName"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   6375
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4920
      Top             =   7080
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
      Left            =   3720
      Top             =   7080
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
      Left            =   2520
      Top             =   7080
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Enter Sales Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton Pre 
         Caption         =   "<<"
         Height          =   735
         Left            =   2520
         TabIndex        =   24
         Top             =   5880
         Width           =   255
      End
      Begin VB.CommandButton Next 
         Caption         =   ">>"
         Height          =   735
         Left            =   7680
         TabIndex        =   23
         Top             =   5880
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Left            =   1080
         Top             =   120
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   22
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   240
         Left            =   4320
         TabIndex        =   20
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ammount :"
         Height          =   480
         Left            =   360
         TabIndex        =   15
         Top             =   5040
         Width           =   1800
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
         Height          =   480
         Left            =   480
         TabIndex        =   11
         Top             =   3720
         Width           =   1800
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty :"
         Height          =   360
         Left            =   240
         TabIndex        =   9
         Top             =   4560
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Available :"
         Height          =   360
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate :"
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name :"
         Height          =   360
         Left            =   480
         TabIndex        =   3
         Top             =   3120
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name :"
         Height          =   360
         Left            =   840
         TabIndex        =   1
         Top             =   1200
         Width           =   1320
      End
   End
End
Attribute VB_Name = "SalesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub Cancle_Click()
CompName.Text = ""
CustName.Text = ""
Qty.Text = ""
Tot.Text = ""
 ItemName.Text = ""
 AQty.Text = ""
 Rate.Text = ""
 
End Sub

Private Sub CTotal_Click()
If AQty >= Qty Then
Tot = Rate * Qty
Else
    MsgBox "Invalide"
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"

Adodc1.RecordSource = "Select * from Stock order by ItemName"
Adodc1.Refresh

Adodc2.RecordSource = "Select * from Sales order by ItemName Desc"
Adodc2.Refresh
i = Adodc2.Recordset.Fields("SrNo")
SrNo = i + 1
End Sub

Private Sub New_Click()
Adodc2.RecordSource = "Select * from Sales order by ItemName Desc"
Adodc2.Refresh
i = Adodc2.Recordset.Fields("SrNo")
SrNo = i + 1
Adodc1.Refresh
CompName.Text = ""
CustName.Text = ""
Qty.Text = ""
Tot.Text = ""
 ItemName.Text = ""
 AQty.Text = ""
 Rate.Text = ""
End Sub

Private Sub Pre_Click()
If Adodc1.Recordset.BOF Then
MsgBox "NO MORE RECORDS AVAILABLE"
Adodc1.Recordset.MoveFirst
Else
Adodc1.Recordset.MovePrevious
End If
End Sub
Private Sub Next_Click()
If Adodc1.Recordset.EOF Then
MsgBox "NO MORE RECORDS AVAILABLE"
Adodc1.Recordset.MoveLast
Else
Adodc1.Recordset.MoveNext
End If
End Sub
Private Sub Qty_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
MsgBox "Please... Enter Numeric values"
KeyAscii = 0
End If

End Sub

Private Sub Save_Click()
Adodc2.RecordSource = "Select * from Sales"
Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = Dt
Adodc2.Recordset.Fields(1) = SrNo
Adodc2.Recordset.Fields(2) = ItemName
Adodc2.Recordset.Fields(3) = Rate
Adodc2.Recordset.Fields(4) = CompName
Adodc2.Recordset.Fields(5) = CustName
Adodc2.Recordset.Fields(6) = Qty
Adodc2.Recordset.Fields(7) = Tot

Adodc3.RecordSource = "Select * from Stock where ItemName='" + ItemName + "'"
Adodc3.Refresh
i = Adodc3.Recordset.Fields("Qty")
Adodc3.Recordset.Fields("Qty") = i - Qty
Adodc3.Recordset.Update
MsgBox ("Success")
Adodc1.Refresh
Adodc2.Refresh
End Sub


Private Sub Timer1_Timer()
If Label9.ForeColor = vbRed Then
 Label9.ForeColor = vbBlue
 ElseIf Label9.ForeColor = vbBlue Then
 Label9.ForeColor = vbGreen
 ElseIf Label9.ForeColor = vbGreen Then
 Label9.ForeColor = vbYellow
 Else: Label9.ForeColor = vbRed
 End If
End Sub
