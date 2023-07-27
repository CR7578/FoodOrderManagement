VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form managevendors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Vendors"
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   18855
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   2160
      Top             =   10080
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
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
      Connect         =   $"managevendors.frx":0000
      OLEDBString     =   $"managevendors.frx":00E3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"managevendors.frx":01C6
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   18615
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "Search"
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
         Left            =   13680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdactivate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Re-activate"
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
         Left            =   15360
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdsuspend 
         BackColor       =   &H000000FF&
         Caption         =   "Suspend"
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
         Left            =   13680
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtshop 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   120
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "managevendors.frx":025E
         Height          =   5775
         Left            =   720
         TabIndex        =   17
         Top             =   2160
         Width           =   17535
         _ExtentX        =   30930
         _ExtentY        =   10186
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
      Begin VB.TextBox txtexp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9840
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtcrt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtnum 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtplan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtemail 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9840
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtname 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtvid 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Food Shop Name:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Expries on :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Plan Created :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Number :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Subscription Plan :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   " Name :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "VendorID :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      Caption         =   "VendorID"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "MANAGE VENDORS"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "managevendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdactivate_Click()
Adodc1.Recordset.Fields("Status") = "Active"
Adodc1.Recordset.Update
MsgBox Adodc1.Recordset.Fields("VendorID") + " ( " + Adodc1.Recordset.Fields("OwnerName") + " ) has been Re-Activated", vbInformation, "Re-Activate"
End Sub

Private Sub cmdsearch_Click()
a = InputBox(vbNewLine + vbNewLine + "Enter Vendor ID", "SEARCH VENDOR")
On Error GoTo errmsg
With Adodc1.Recordset
.MoveFirst
.Find "VendorID='" & a & "'"
txtvid = Adodc1.Recordset.Fields("VendorID")
txtname = Adodc1.Recordset.Fields("OwnerName")
txtshop = Adodc1.Recordset.Fields("OwnerEmail")
txtnum = Adodc1.Recordset.Fields("OwnerNumber")
txtcrt = Adodc1.Recordset.Fields("SubscriptionCreated")
txtexp = Adodc1.Recordset.Fields("SubscriptionExpriesOn")
txtemail = Adodc1.Recordset.Fields("OwnerEmail")
txtplan = Adodc1.Recordset.Fields("SubscriptionPlan")
If .EOF Then
MsgBox vbNewLine + "Enter Vendor ID is not found", vbCritical, "SEARCH VENDOR"
End If
End With
Exit Sub
errmsg:
MsgBox vbNewLine + "Enter Vendor ID is not found", vbCritical, "SEARCH VENDOR"
End Sub

Private Sub cmdsuspend_Click()
Adodc1.Recordset.Fields("Status") = "Suspended"
Adodc1.Recordset.Update
MsgBox Adodc1.Recordset.Fields("VendorID") + " ( " + Adodc1.Recordset.Fields("OwnerName") + " ) has been Suspended", vbInformation, "SUSPEND"
End Sub

Private Sub DataGrid1_Click()
txtvid = Adodc1.Recordset.Fields("VendorID")
txtname = Adodc1.Recordset.Fields("OwnerName")
txtshop = Adodc1.Recordset.Fields("OwnerEmail")
txtnum = Adodc1.Recordset.Fields("OwnerNumber")
txtcrt = Adodc1.Recordset.Fields("SubscriptionCreated")
txtexp = Adodc1.Recordset.Fields("SubscriptionExpriesOn")
txtemail = Adodc1.Recordset.Fields("OwnerEmail")
txtplan = Adodc1.Recordset.Fields("SubscriptionPlan")
End Sub

Private Sub Form_Load()
txtvid = ""
txtname = ""
txtnum = ""
txtshop = ""
txtemail = ""
txtcrt = ""
txtexp = ""
txtplan = ""
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

