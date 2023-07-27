VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form orders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CUSTOMER ORDER"
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   20430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14640
      Top             =   10320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"orders.frx":0000
      OLEDBString     =   $"orders.frx":00E3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "OrderDetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8160
      Top             =   10320
      Width           =   6375
      _ExtentX        =   11245
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
      Connect         =   $"orders.frx":01C6
      OLEDBString     =   $"orders.frx":02A9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from VendorSignup where Status='Active';"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   0
      Top             =   1200
      Width           =   20175
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   1935
         Left            =   19800
         TabIndex        =   24
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   3413
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtid 
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
         Left            =   2040
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtamt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   18240
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   7920
         Width           =   1455
      End
      Begin MSComctlLib.ListView Order 
         Height          =   6375
         Left            =   13320
         TabIndex        =   19
         Top             =   1440
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   11245
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Selected Items"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Price Rs./-"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Amount Rs./-"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtnum 
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
         Left            =   6720
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "orders.frx":038C
         Left            =   6720
         List            =   "orders.frx":038E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdconfirm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Confirm"
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
         Left            =   15000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H000000FF&
         Caption         =   "Cancel"
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
         Left            =   16680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdqty 
         BackColor       =   &H0000FFFF&
         Caption         =   "Change Quantities"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtname 
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
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H8000000E&
         Caption         =   "Clear"
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
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin MSComctlLib.ListView rolls 
         Height          =   3255
         Left            =   4380
         TabIndex        =   4
         Top             =   4920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Chicken/Veg Rolls"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Starters 
         Height          =   3255
         Left            =   0
         TabIndex        =   5
         Top             =   4920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Starters"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Snacks 
         Height          =   3255
         Left            =   8760
         TabIndex        =   6
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Snacks"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Rice 
         Height          =   3255
         Left            =   4380
         TabIndex        =   7
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rice Items"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Burger 
         Height          =   3255
         Left            =   0
         TabIndex        =   8
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Burgers / Sandwichs"
            Object.Width           =   4852
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView icecream 
         Height          =   3255
         Left            =   8760
         TabIndex        =   13
         Top             =   4920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IceCreams / MilkShakes"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price Rs./-"
            Object.Width           =   2363
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Order ID :"
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
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16440
         TabIndex        =   21
         Top             =   7920
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "CUSTOMER NUMBER :"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Available Shops :"
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
         Left            =   4920
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "CUSTOMER NAME :"
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc ItemsDetails 
      Height          =   330
      Left            =   1200
      Top             =   10320
      Width           =   6735
      _ExtentX        =   11880
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
      Connect         =   $"orders.frx":0390
      OLEDBString     =   $"orders.frx":0473
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ItemsDetails"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   6240
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ORDER NOW"
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
      TabIndex        =   17
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vi As String
Dim vd As String

Private Sub cmdconfirm_Click()
If txtname = "" Or txtnum = "" Or Order.ListItems.Count = 0 Then
MsgBox vbNewLine + "Please select the items for order confirmation"
Else
Dim id As Integer
Dim id1 As String
On Error GoTo errmsg
Customerpay.Adodc1.Refresh
Customerpay.Adodc1.Recordset.MoveLast
id1 = Customerpay.Adodc1.Recordset("BillID")
id = Mid(id1, 2, 4) + 1
Customerpay.Label11.Caption = "B" & id
Customerpay.Label8.Caption = txtamt.Text
Customerpay.Label6.Caption = txtid.Text
Customerpay.Label13.Caption = txtname.Text
Customerpay.Label7.Caption = txtnum.Text
Customerpay.Label17.Caption = vi
Customerpay.Label18.Caption = vd
Customerpay.Label15.Caption = Combo1.Text
Customerpay.Show
Exit Sub
errmsg:
Customerpay.Label11.Caption = "B1001"
Customerpay.Label8.Caption = txtamt.Text
Customerpay.Label13.Caption = txtname.Text
Customerpay.Label7.Caption = txtnum.Text
Customerpay.Label6.Caption = txtid.Text
Customerpay.Label17.Caption = vi
Customerpay.Label18.Caption = vd
Customerpay.Label15.Caption = Combo1.Text
Customerpay.Show
End If
End Sub

Private Sub cmddelete_Click()
Unload Me
End Sub

Private Sub cmdqty_Click()
Order.Enabled = True
UpDown1.Enabled = True
End Sub

Private Sub Starters_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Call Order1
End Sub
Private Sub Rice_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Call Order1
End Sub

Private Sub Snacks_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Call Order1
End Sub
Private Sub Burger_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Call Order1
End Sub
Private Sub rolls_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Call Order1
End Sub
Private Sub icecream_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Call Order1
End Sub

Private Sub cmdclear_Click()
Call clear
End Sub

Private Sub Combo1_DropDown()
If txtname = "" Or txtnum = "" Then
MsgBox "Please enter your Name and Number"
txtname.SetFocus
Else
On Error GoTo errmsg
Combo1.clear
Adodc2.Refresh
With Adodc2.Recordset
.MoveFirst
While Not .EOF
If DateValue(Format(Now, "dd/MM/yyyy")) > DateValue(Format(.Fields("SubscriptionExpriesOn"), "dd/MM/yyyy")) Then
Else
Combo1.AddItem .Fields("FoodShopName")
.MoveNext
End If
Wend
End With
Exit Sub
errmsg:
MsgBox "No Food Shops are Available at this Moment"
End If
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
If Combo1 = "" Then
Else
ItemsDetails.RecordSource = "Select * from VendorSignup"
ItemsDetails.Refresh
ItemsDetails.Recordset.MoveFirst
ItemsDetails.Recordset.Find "FoodShopName='" & Combo1.Text & "'"
vi = ItemsDetails.Recordset.Fields("VendorID")
vd = ItemsDetails.Recordset.Fields("FoodShopAddress")
If KeyAscii = 13 Then
Call data
End If
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then
Else
ItemsDetails.RecordSource = "Select * from VendorSignup"
ItemsDetails.Refresh
ItemsDetails.Recordset.MoveFirst
ItemsDetails.Recordset.Find "FoodShopName='" & Combo1.Text & "'"
vi = ItemsDetails.Recordset.Fields("VendorID")
vd = ItemsDetails.Recordset.Fields("FoodShopAddress")
Call data
End If
End Sub

Private Sub Form_Load()
Call clear
Me.Top = dashboard.Frame1.Height + 1000
Me.Left = 200
Me.Width = Screen.Width - 400
Me.Height = Screen.Height - 2000 - dashboard.Frame1.Height
Label1.Top = 200
Label1.Left = 800
Label1.Height = 600
Label1.Width = managevendors.Width - 1600
Frame1.Left = Me.Width / 2 - Frame1.Width / 2
Order.Enabled = False
UpDown1.Enabled = False
Dim id As Integer
Dim id1 As String
On Error GoTo errmsg
Adodc1.Refresh
Adodc1.Recordset.MoveLast
id1 = Adodc1.Recordset("OrderID")
id = Mid(id1, 2, 4) + 1
Adodc1.Recordset.AddNew
txtid.Text = "O" & id
txtname = ""
txtnum = ""
txtid.Enabled = False
Exit Sub
errmsg:
orders.Adodc1.Recordset.AddNew
orders.txtid.Text = "O1001"
orders.txtname = ""
orders.txtnum = ""
orders.txtid.Enabled = False
End Sub
Function data()
Burger.ListItems.clear
Rice.ListItems.clear
Snacks.ListItems.clear
Starters.ListItems.clear
rolls.ListItems.clear
icecream.ListItems.clear
Dim lv As ListItem
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & vi & "' and MenuCategory='Burger'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Burger.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemName")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & vi & "' and MenuCategory='Rice'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Rice.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemName")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & vi & "' and MenuCategory='Snacks'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Snacks.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemName")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & vi & "' and MenuCategory='Starters'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Starters.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemName")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & vi & "' and MenuCategory='Rolls'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = rolls.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemName")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & vi & "' and MenuCategory='Icecream'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = icecream.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemName")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
End Function
Function clear()
txtname = ""
txtnum = ""
txtamt = ""
End Function
Function Order1()
Order.ListItems.clear
Dim i, j As Integer
i = 1
j = 1
While Not i = Burger.ListItems.Count + 1
If Burger.ListItems(i).Checked = True Then
Order.ListItems.Add , , Burger.ListItems(i).Text
Order.ListItems(j).ListSubItems.Add , , Burger.ListItems(i).SubItems(1)
Order.ListItems(j).ListSubItems.Add , , 1
Order.ListItems(j).ListSubItems.Add , , Burger.ListItems(i).SubItems(1)
j = j + 1
End If
i = i + 1
Wend
i = 1
While Not i = Rice.ListItems.Count + 1
If Rice.ListItems(i).Checked = True Then
Order.ListItems.Add , , Rice.ListItems(i).Text
Order.ListItems(j).ListSubItems.Add , , Rice.ListItems(i).SubItems(1)
Order.ListItems(j).ListSubItems.Add , , 1
Order.ListItems(j).ListSubItems.Add , , Rice.ListItems(i).SubItems(1)
j = j + 1
Else
End If
i = i + 1
Wend
i = 1
While Not i = Snacks.ListItems.Count + 1
If Snacks.ListItems(i).Checked = True Then
Order.ListItems.Add , , Snacks.ListItems(i).Text
Order.ListItems(j).ListSubItems.Add , , Snacks.ListItems(i).SubItems(1)
Order.ListItems(j).ListSubItems.Add , , 1
Order.ListItems(j).ListSubItems.Add , , Snacks.ListItems(i).SubItems(1)
j = j + 1
End If
i = i + 1
Wend
i = 1
While Not i = Starters.ListItems.Count + 1
If Starters.ListItems(i).Checked = True Then
Order.ListItems.Add , , Starters.ListItems(i).Text
Order.ListItems(j).ListSubItems.Add , , Starters.ListItems(i).SubItems(1)
Order.ListItems(j).ListSubItems.Add , , 1
Order.ListItems(j).ListSubItems.Add , , Starters.ListItems(i).SubItems(1)
j = j + 1
End If
i = i + 1
Wend
i = 1
While Not i = Starters.ListItems.Count + 1
If rolls.ListItems(i).Checked = True Then
Order.ListItems.Add , , rolls.ListItems(i).Text
Order.ListItems(j).ListSubItems.Add , , rolls.ListItems(i).SubItems(1)
Order.ListItems(j).ListSubItems.Add , , 1
Order.ListItems(j).ListSubItems.Add , , rolls.ListItems(i).SubItems(1)
j = j + 1
End If
i = i + 1
Wend
i = 1
While Not i = Starters.ListItems.Count + 1
If icecream.ListItems(i).Checked = True Then
Order.ListItems.Add , , icecream.ListItems(i).Text
Order.ListItems(j).ListSubItems.Add , , icecream.ListItems(i).SubItems(1)
Order.ListItems(j).ListSubItems.Add , , 1
Order.ListItems(j).ListSubItems.Add , , icecream.ListItems(i).SubItems(1)
j = j + 1
End If
i = i + 1
Wend
Dim sum As Integer
sum = 0
j = 1
For i = 1 To Order.ListItems.Count
sum = Order.ListItems(j).SubItems(3) + sum
j = j + 1
Next
txtamt.Text = sum
Order.Enabled = False
UpDown1.Enabled = False
If Order.ListItems.Count = 0 Then
cmdconfirm.Enabled = False
cmdqty.Enabled = False
Else
cmdconfirm.Enabled = True
cmdqty.Enabled = True
End If
End Function

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtname = "" Then
MsgBox "Please Enter your Name"
txtname.SetFocus
Else
txtnum.SetFocus
End If
End If
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtnum = "" Or Not (Len(txtnum) = 10) Then
txtnum.Text = ""
MsgBox "Please Enter 10 Digits Number Only"
txtnum.SetFocus
Else
Combo1.SetFocus
End If
End If
End Sub


Private Sub UpDown1_DownClick()
If Order.SelectedItem.ListSubItems(2).Text > 1 Then
Order.SelectedItem.ListSubItems(2).Text = Val(Order.SelectedItem.ListSubItems(2).Text) - 1
Else
End If
Order.SelectedItem.ListSubItems(3).Text = Order.SelectedItem.ListSubItems(2) * Order.SelectedItem.ListSubItems(1)
Dim sum As Integer
sum = 0
j = 1
For i = 1 To Order.ListItems.Count
sum = Order.ListItems(j).SubItems(3) + sum
j = j + 1
Next
txtamt.Text = sum
End Sub

Private Sub UpDown1_UpClick()
If Order.SelectedItem.ListSubItems(2).Text < 5 Then
Order.SelectedItem.ListSubItems(2).Text = Val(Order.SelectedItem.ListSubItems(2).Text) + 1
Else
End If
Order.SelectedItem.ListSubItems(3).Text = Order.SelectedItem.ListSubItems(2) * Order.SelectedItem.ListSubItems(1)
Dim sum As Integer
sum = 0
j = 1
For i = 1 To Order.ListItems.Count
sum = Order.ListItems(j).SubItems(3) + sum
j = j + 1
Next
txtamt.Text = sum
End Sub
