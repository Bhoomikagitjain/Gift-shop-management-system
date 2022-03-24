VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SuppliarsFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplars Information"
   ClientHeight    =   5715
   ClientLeft      =   5055
   ClientTop       =   2415
   ClientWidth     =   13455
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
   ScaleHeight     =   5715
   ScaleWidth      =   13455
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4560
      Top             =   2520
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2520
      Top             =   2640
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Suppliars report"
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   735
         Left            =   5280
         TabIndex        =   5
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   240
         Top             =   120
      End
      Begin VB.TextBox SName 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00C0E0FF&
         Height          =   495
         Left            =   0
         TabIndex        =   0
         Top             =   5640
         Width           =   4215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Suppliers.frx":0000
         Height          =   4935
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   8705
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   615
         Left            =   3720
         TabIndex        =   4
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   2
         Top             =   0
         Width           =   2775
      End
   End
End
Attribute VB_Name = "SuppliarsFrm"
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

Private Sub First_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\GiftShop.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * from Suppliar"
Adodc1.Refresh
End Sub

Private Sub Last_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Next_Click()
If Adodc1.Recordset.EOF Then
MsgBox "NO MORE RECORDS AVAILABLE"
Adodc1.Recordset.MoveLast
Else
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub PhNO_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
MsgBox "Please... Enter Numeric values"
KeyAscii = 0
End If
End Sub

Private Sub Pre_Click()
If Adodc1.Recordset.BOF Then
MsgBox "NO MORE RECORDS AVAILABLE"
Adodc1.Recordset.MoveFirst
Else
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Save_Click()
Adodc1.Recordset.Update
MsgBox ("Saved")
End Sub

Private Sub Timer1_Timer()
If Label2.ForeColor = vbRed Then
 Label2.ForeColor = vbBlue
 ElseIf Label2.ForeColor = vbBlue Then
 Label2.ForeColor = vbGreen
 ElseIf Label2.ForeColor = vbGreen Then
 Label2.ForeColor = vbYellow
 Else: Label2.ForeColor = vbRed
 End If
End Sub
