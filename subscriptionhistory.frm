VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form subscriptionhistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUBSCRIPTION HISTORY"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   17865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   17535
      Begin VB.CommandButton cmdrefresh 
         BackColor       =   &H8000000E&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdfilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "subscriptionhistory.frx":0000
         Height          =   7455
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   13150
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
               LCID            =   16393
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
               LCID            =   16393
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3360
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"subscriptionhistory.frx":0015
      OLEDBString     =   $"subscriptionhistory.frx":00F8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from SubscriptionHistory;"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SUBSCRIPTION HISTORY"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "subscriptionhistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdfilter_Click()
a = InputBox(vbNewLine + vbNewLine + "Enter Vendor ID", "SEARCH VENDOR")
On Error GoTo errmsg
With Adodc1.Recordset
.MoveFirst
.Find "VendorID='" & a & "'"
Adodc1.RecordSource = "select * from SubscriptionHistory where VendorID='" & a & "'"
Adodc1.Refresh
If .EOF Then
MsgBox vbNewLine + "Enter Vendor ID is not found", vbCritical, "SEARCH VENDOR"
Adodc1.RecordSource = "select * from SubscriptionHistory"
Adodc1.Refresh
End If
End With
Exit Sub
errmsg:
MsgBox vbNewLine + "Enter Vendor ID is not found", vbCritical, "SEARCH VENDOR"
Adodc1.RecordSource = "select * from SubscriptionHistory"
Adodc1.Refresh
End Sub

Private Sub cmdrefresh_Click()
Adodc1.RecordSource = "select * from SubscriptionHistory"
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Me.Top = dashboard.Frame1.Height + 1000
Me.Left = 200
Me.Width = Screen.Width - 400
Me.Height = Screen.Height - 2000 - dashboard.Frame1.Height
Label1.Top = 200
Label1.Left = 800
Label1.Height = 600
Label1.Width = managevendors.Width - 1600
Frame1.Left = Me.Width / 2 - Frame1.Width / 2
End Sub

