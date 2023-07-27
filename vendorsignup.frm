VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form vendorsignup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VENDOR SIGNUP PAGE"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9060
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "vendorsignup.frx":0000
   ScaleHeight     =   10980
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "vendorsignup.frx":199DF
      Height          =   3975
      Left            =   9840
      TabIndex        =   21
      Top             =   2280
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7011
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
   Begin VB.CommandButton cmdsignup 
      Caption         =   "Signup"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   20
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   19
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   18
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton cmdjoin 
      BackColor       =   &H8000000A&
      Caption         =   "JOIN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   2280
      TabIndex        =   0
      Top             =   3240
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   124846081
         CurrentDate     =   44949
      End
      Begin VB.TextBox txtadd 
         Height          =   550
         Left            =   120
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   4560
         Width           =   3135
      End
      Begin VB.TextBox txtsid 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtrepass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   16
         Text            =   "Text7"
         Top             =   6360
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "vendorsignup.frx":199F4
         Left            =   120
         List            =   "vendorsignup.frx":199FE
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   5400
         Width           =   2415
      End
      Begin VB.TextBox txtpass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   14
         Text            =   "Text6"
         Top             =   6000
         Width           =   3135
      End
      Begin VB.TextBox txtshop 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   3960
         Width           =   3135
      End
      Begin VB.TextBox txtemail 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtnum 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtname 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtvid 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner DOB :"
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
         Left            =   120
         TabIndex        =   28
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Food Shop Address :"
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
         Left            =   120
         TabIndex        =   27
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2640
         Picture         =   "vendorsignup.frx":19A2C
         Top             =   5400
         Width           =   240
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "( Re-Enter ) "
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3360
         TabIndex        =   17
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
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
         Left            =   120
         TabIndex        =   7
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "New Vendor ID :"
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner Name :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Food Shop Name :"
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
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner Number :"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner Email :"
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
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9840
      Top             =   1800
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
      Connect         =   $"vendorsignup.frx":19C1D
      OLEDBString     =   $"vendorsignup.frx":19D00
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from VendorSignup"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Join us to grow your Business."
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Already have an Account?....Then"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Signin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "vendorsignup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
Call clear
End Sub

Private Sub cmdhome_Click()
home.Show
Unload Me
End Sub

Private Sub cmdjoin_Click()
Frame1.Enabled = True
txtpass = ""
txtrepass = ""
txtvid = ""
txtnum = ""
txtname = ""
txtshop = ""
txtadd = ""
txtemail = ""
Label12.Caption = "( Re-Enter )"
Label12.ForeColor = &H0&
Dim id As Integer
Dim id1 As String
Dim idd As Integer
Dim idd1 As String
On Error GoTo errmsg
Adodc1.Refresh
Adodc1.Recordset.MoveLast
id1 = Adodc1.Recordset("VendorID")
id = Mid(id1, 2, 4) + 1
'Adodc1.Recordset.AddNew
txtvid = "V" & id
payment.Adodc1.Refresh
payment.Adodc1.Recordset.MoveLast
idd1 = payment.Adodc1.Recordset("SubscriptionID")
idd = Mid(idd1, 2, 4) + 1
'payment.Adodc1.Recordset.AddNew
txtsid = "S" & id
txtname.SetFocus
Exit Sub
errmsg:
'Adodc1.Recordset.AddNew
'payment.Adodc1.Recordset.AddNew
txtvid = "V1001"
txtsid = "S1001"
txtname.SetFocus
End Sub

Private Sub cmdsignup_Click()
If (Combo1.Text = "Rs.24000/- Per Year") Then
ta = "24000"
ElseIf (Combo1.Text = "Rs.2000/- Per Month") Then
ta = "2000"
End If
If Not (txtpass = txtrepass) Then
MsgBox "Password are not matching, Please Check again"
ElseIf (txtvid = "" Or txtname = "" Or txtnum = "" Or txtemail = "" Or txtshop = "" Or txtadd = "" Or txtrepass = "" Or Combo1 = "") Then
MsgBox "Please Fill All the Required Details in the Signup form", vbExclamation
ElseIf DateValue(Format(Now, "dd/MM/yyyy")) < DateValue(DateAdd("yyyy", 18, DTPicker1)) Then
MsgBox "Please Enter an Valid Date of Birth, Vendor must be 18+."
DTPicker1.SetFocus
Else
Dim id As Integer
Dim id1 As String
On Error GoTo errmsg
payment.Adodc1.Refresh
payment.Adodc1.Recordset.MoveLast
id1 = payment.Adodc1.Recordset("PaymentID")
id = Mid(id1, 2, 4) + 1
'payment.Adodc1.Recordset.AddNew
payment.Label11.Caption = "P" & id
payment.Label8.Caption = ta
payment.Label6.Caption = txtvid.Text
payment.Label7.Caption = Combo1.Text
payment.Show
Exit Sub
errmsg:
'payment.Adodc1.Recordset.AddNew
payment.Label11.Caption = "P1001"
payment.Label8.Caption = ta
payment.Label6.Caption = txtvid.Text
payment.Label7.Caption = Combo1.Text
payment.Show
End If
End Sub



Private Sub Image1_Click()
MsgBox "Current Available Plans are:" + vbNewLine + "1. Rs.2000/- per Month  And" + vbNewLine + "2. Rs.24000/- per Year.", vbInformation
End Sub

Private Sub Label15_Click()
Hide
login.Show
login.txtvid.SetFocus
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtemail = "" Or Not (isEmail(txtemail.Text) = True) Then
txtemail = ""
MsgBox "Please Enter an Valid Email Address."
txtemail.SetFocus
Else
DTPicker1.SetFocus
End If
End If
End Sub
Private Sub DTPicker1_LostFocus()
If DateValue(Format(Now, "dd/MM/yyyy")) < DateValue(DateAdd("yyyy", 18, DTPicker1)) Then
MsgBox "Please Enter an Valid Date of Birth, Vendor must be 18+."
DTPicker1.SetFocus
Else
txtshop.SetFocus
End If
End Sub
Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtname = "" Then
MsgBox "Feild Name Can't be Empty"
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
txtemail.SetFocus
End If
End If
End Sub


Private Sub txtrepass_Click()
If (txtpass = "") Then
Label12.Caption = "( Re-Enter )"
Label12.ForeColor = &H0&
txtpass.SetFocus
End If
End Sub
Private Sub txtrepass_Change()
If (txtpass.Text = txtrepass.Text) Then
Label12.Caption = "Matching"
Label12.ForeColor = &HC000&
Else
Label12.Caption = "Not Matching"
Label12.ForeColor = &HFF&
End If
End Sub
Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = (Screen.Height) / 2 - Me.Height / 2
txtvid.Enabled = False
txtpass = ""
txtrepass = ""
txtvid = ""
txtsid = ""
txtnum = ""
txtname = ""
txtshop = ""
txtadd = ""
txtemail = ""
Label12.Caption = "( Re-Enter )"
Label12.ForeColor = &H0&
Frame1.Enabled = False
End Sub
Function clear()
txtpass = ""
txtrepass = ""
txtnum = ""
txtname = ""
txtshop = ""
txtadd = ""
txtemail = ""
Label12.Caption = "( Re-Enter )"
Label12.ForeColor = &H0&
Combo1.clear
Combo1.AddItem "Rs.2000/- Per Month"
Combo1.AddItem "Rs.24000/- Per Month"
End Function
Private Function ValidatePassword(ByVal sPass) As Boolean
Dim regEx
Set regEx = CreateObject("vbscript.regexp")
regEx.Pattern = "^.*(?=.{8,})(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[!@#$%^&+=]).*$"
ValidatePassword = regEx.test(sPass)
Set regEx = Nothing
End Function
Function isEmail(email As String) As Boolean
Dim At As Integer
Dim oneDot As Integer
Dim twoDots As Integer
isEmail = True
At = InStr(1, email, "@", vbTextCompare)
oneDot = InStr(At + 2, email, ".", vbTextCompare)
twoDots = InStr(At + 2, email, "..", vbTextCompare)
If At = 0 Or oneDot = 0 Or Not twoDots = 0 Or Right(email, 1) = "." Then isEmail = False
End Function

Private Sub txtrepass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdsignup_Click
End If
End Sub
Private Sub txtshop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtshop = "" Then
MsgBox "Field Shop Name can't be Empty"
Else
txtadd.SetFocus
End If
End If
End Sub
Private Sub txtadd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtadd = "" Then
MsgBox "Field Address can't be Empty"
Else
Combo1.SetFocus
End If
End If
End Sub
Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Combo1 = "" Then
MsgBox "Subscription plan field can't be Empty"
Else
txtpass.SetFocus
End If
End If
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtpass = "" Or Not (ValidatePassword(txtpass.Text) = True) Then
txtpass.Text = ""
MsgBox "Entered Password is Not Valid." + vbNewLine + "Please Make Sure the Entered password must Contain the following conditions" + vbNewLine + "Atleast 8 Characters" + vbNewLine + "Atleast 1 Number" + vbNewLine + "Atleast 1 Lowercase letter" + vbNewLine + "Atleast 1 Uppercase letter" + vbNewLine + "Atleast 1 Special Character", vbInformation
txtpass.SetFocus
Else
txtrepass.SetFocus
End If
End If
End Sub
