VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form home 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDER YOUR FOOD SYSTEM"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   18435
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Contact Admin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   8640
      TabIndex        =   1
      Top             =   2520
      Width           =   4215
      Begin VB.Label txthelp 
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VENDOR"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SignUp"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000E&
         BorderStyle     =   2  'Dash
         Height          =   1815
         Left            =   0
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.CommandButton txttrack 
      Caption         =   "Track Your Order"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   360
      Top             =   6360
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
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
      Connect         =   $"home.frx":0000
      OLEDBString     =   $"home.frx":00E3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from OrderDetails"
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
   Begin VB.Label txtorder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order Now"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5000
      Left            =   0
      Picture         =   "home.frx":01C6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10000
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
contact.Label1.Caption = "Hello , Dear Customer"
contact.Label2.Caption = "Happy to see you here ," + vbNewLine + "If you have any issues / Report on Vendor You can feel free to contact us."
contact.Label3.Caption = "Admin@gamil.com"
contact.Label4.Caption = "You can contact us throught email for this below Email Address"
contact.Show , home
End Sub

Private Sub Form_Load()
Image1.Top = 0
Image1.Left = 0
Image1.Width = Screen.Width
Image1.Height = Screen.Height
Frame1.Top = 750
Frame1.Left = Screen.Width - 5000
Command1.Top = 2800
Command1.Left = Screen.Width - 3000
txttrack.Left = Screen.Width - 10000
txtorder.Top = 8400
txtorder.Left = 2000
Adodc2.RecordSource = "select * from OrderDetails"
Adodc2.Refresh
End Sub

Private Sub txthelp_Click()
MsgBox "What is term Vendor in our System ?" + vbNewLine + "Vendors are the food-providers who can get orders from Customers by displaying their food list and prices in our System" + vbNewLine + "Willing to Join as Vendor ? Then click Signup to join us.", vbInformation
End Sub

Private Sub Label2_Click()
login.Show
login.txtvid.SetFocus
'Unload Me
End Sub

Private Sub Label3_Click()
Load vendorsignup
vendorsignup.Show , home
'home.Hide
End Sub
Private Sub txtorder_Click()
Dim id As Integer
Dim id1 As String
On Error GoTo errmsg
orders.Adodc1.Refresh
orders.Adodc1.Recordset.MoveLast
id1 = orders.Adodc1.Recordset("OrderID")
id = Mid(id1, 2, 4) + 1
orders.Adodc1.Recordset.AddNew
orders.txtid.Text = "O" & id
orders.txtname = ""
orders.txtnum = ""
orders.txtid.Enabled = False
Load orders
orders.Show , home
Exit Sub
errmsg:
orders.Adodc1.Recordset.AddNew
orders.txtid.Text = "O1001"
orders.txtname = ""
orders.txtnum = ""
orders.txtid.Enabled = False
Load orders
orders.Show , home
End Sub

Private Sub txttrack_Click()
Dim a As String
Dim b As String
a = InputBox("Please Enter your OrderID.")
With Adodc2.Recordset
.MoveFirst
.Find "OrderID='" & a & "'"
If .EOF Then
MsgBox "Entered OrderID is Not Exist.Please Check Your OrderID ", vbCritical
Else
b = InputBox("Please Enter your OrderKey.")
If b = .Fields("SecretKey") Then
MsgBox "Your Order is " & .Fields("OrderStatus") & " " & .Fields("OrderStatusTime") & "."
Else
MsgBox "You have Enter Wrong OrderKey, Please Check again ", vbCritical
End If
End If
End With
End Sub
