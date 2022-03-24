VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   5040
   ClientTop       =   3390
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8220
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Suppliars Details"
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   1080
         Top             =   120
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   735
         Left            =   960
         TabIndex        =   11
         Top             =   4200
         Width           =   5775
         Begin VB.CommandButton AddNew 
            Caption         =   "Add New"
            Height          =   495
            Left            =   240
            TabIndex        =   15
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Save 
            Caption         =   "Save"
            Height          =   495
            Left            =   1560
            TabIndex        =   14
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Cancle 
            Caption         =   "Cancle"
            Height          =   495
            Left            =   3000
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Delete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   4320
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox SID 
         DataField       =   "S_ID"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox SName 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   1440
         Width           =   4215
      End
      Begin VB.TextBox Add 
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox PhNO 
         DataField       =   "PhoneNo"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   2640
         Width           =   4215
      End
      Begin VB.TextBox Company 
         DataField       =   "Company"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   3240
         Width           =   4215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1680
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   3240
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Details"
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
         Left            =   2760
         TabIndex        =   16
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Suppliar ID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Suppliar Name :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddNew_Click()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc2.RecordSource = "Select * from Suppliar order by S_ID desc"
Adodc2.Refresh
i = Adodc2.Recordset.Fields("S_ID")
Adodc1.Recordset.AddNew
SID = i + 1
SName = ""
Add = ""
PhNO = ""
Company = ""
End Sub

Private Sub Cancle_Click()
Adodc1.Refresh
End Sub

Private Sub Delete_Click()
Adodc1.Recordset.Delete
MsgBox ("1 Record Deleted")
End Sub



Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * from Suppliar"
Adodc1.Refresh
End Sub





Private Sub PhNO_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
MsgBox "Please... Enter Numeric values"
KeyAscii = 0
End If
End Sub




Private Sub Save_Click()
Adodc1.Recordset.Update
MsgBox ("Saved")
End Sub

Private Sub Timer1_Timer()
If Label6.ForeColor = vbRed Then
 Label6.ForeColor = vbBlue
 ElseIf Label6.ForeColor = vbBlue Then
 Label6.ForeColor = vbGreen
 ElseIf Label6.ForeColor = vbGreen Then
 Label6.ForeColor = vbYellow
 Else: Label6.ForeColor = vbRed
 End If
End Sub
