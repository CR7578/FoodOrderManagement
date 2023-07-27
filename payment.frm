VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form payment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " PAYMENT PAGE"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "payment.frx":0000
      Height          =   4215
      Left            =   10800
      TabIndex        =   16
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7435
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10800
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
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
      Connect         =   $"payment.frx":0015
      OLEDBString     =   $"payment.frx":00F8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from SubscriptionHistory"
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
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Google Pay"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   5040
      Width           =   2295
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H8000000E&
      Caption         =   "Paytm"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   5640
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H8000000E&
      Caption         =   "Net Banking"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   5040
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000E&
      Caption         =   "Phonepe"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmdpay 
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9960
      Top             =   240
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment ID :"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select your Payment Method"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   4440
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount :"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subscription Plan :"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor ID :"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   0
      Picture         =   "payment.frx":01DB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ta As String

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdpay_Click()
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False Then
MsgBox "Please select the Payment Method."
Else
Dim expr As String
Dim otp As String
If (vendorsignup.Combo1 = "Rs.2000/- Per Month") Then
expr = DateAdd("m", 1, Format(Now, "dd/MM/yyyy"))
Else
expr = DateAdd("yyyy", 1, Format(Now, "dd/MM/yyyy"))
End If
Dim pm As String
If Option1.Value = True Then
pm = "Google Pay"
ElseIf Option2.Value = True Then
pm = "Phonepe"
ElseIf Option3.Value = True Then
pm = "Net Banking"
ElseIf Option4.Value = True Then
pm = "Paytm"
End If
otp = InputBox(pm + " as sent OTP to your number: " & vendorsignup.txtnum & "  Please Enter the OTP Below To Complete the Transaction.", "Complete the Transaction", " HINT:Type OTP as 123456")
If Not (otp = "123456") Then
MsgBox "Entered otp Is invalid"
Else
vendorsignup.Adodc1.Recordset.AddNew
vendorsignup.Adodc1.Recordset.Fields("VendorID") = vendorsignup.txtvid
vendorsignup.Adodc1.Recordset.Fields("OwnerName") = vendorsignup.txtname
vendorsignup.Adodc1.Recordset.Fields("OwnerNumber") = vendorsignup.txtnum
vendorsignup.Adodc1.Recordset.Fields("OwnerEmail") = vendorsignup.txtemail
vendorsignup.Adodc1.Recordset.Fields("FoodShopName") = vendorsignup.txtshop
vendorsignup.Adodc1.Recordset.Fields("FoodShopAddress") = vendorsignup.txtadd
vendorsignup.Adodc1.Recordset.Fields("OwnerDOB") = vendorsignup.DTPicker1
vendorsignup.Adodc1.Recordset.Fields("SubscriptionPlan") = vendorsignup.Combo1.Text
vendorsignup.Adodc1.Recordset.Fields("Password") = vendorsignup.txtrepass
vendorsignup.Adodc1.Recordset.Fields("SubscriptionCreated") = Format(Now, "dd/MM/yyyy")
vendorsignup.Adodc1.Recordset.Fields("SubscriptionExpriesOn") = expr
vendorsignup.Adodc1.Recordset.Fields("Status") = "Active"
vendorsignup.Adodc1.Recordset.Update
payment.Adodc1.Recordset.AddNew
payment.Adodc1.Recordset.Fields("VendorID") = vendorsignup.txtvid
payment.Adodc1.Recordset.Fields("PaymentID") = payment.Label11.Caption
payment.Adodc1.Recordset.Fields("PaymentMethod") = pm
payment.Adodc1.Recordset.Fields("SubscriptionExpriesOn") = expr
payment.Adodc1.Recordset.Fields("SubscriptionID") = vendorsignup.txtsid.Text
payment.Adodc1.Recordset.Fields("SubscriptionPlan") = vendorsignup.Combo1.Text
payment.Adodc1.Recordset.Fields("TotalAmount") = payment.Label8.Caption
payment.Adodc1.Recordset.Fields("DateTime") = payment.Label9.Caption + " " + payment.Label4.Caption
payment.Adodc1.Recordset.Update
MsgBox "Completing your Transaction and Activating Your Account. please wait"
MsgBox "Your Account as been Activated. Now you can login your account"
Unload payment
Unload vendorsignup
login.Show
End If
End If
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = (Screen.Height) / 2 - Me.Height / 2
Label4.Caption = Format(Time, "HH:mm:ss AM/PM")
Label9.Caption = Date
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Format(Time, "HH:mm:ss AM/PM")
Label9.Caption = Date
End Sub

