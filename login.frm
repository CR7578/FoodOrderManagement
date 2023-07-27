VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VENDOR LOGIN PAGE"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "login.frx":19F00
      Height          =   3135
      Left            =   9840
      TabIndex        =   12
      Top             =   2040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5530
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
      Left            =   2160
      TabIndex        =   7
      Top             =   7440
      Width           =   1455
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
      Left            =   3960
      TabIndex        =   6
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   2520
      TabIndex        =   0
      Top             =   3720
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "System Admin"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtvid 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtpass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Text6"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Click here"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   2760
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2160
         Picture         =   "login.frx":19F15
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Forget Password..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor ID :"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9960
      Top             =   1560
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
      Connect         =   $"login.frx":1A106
      OLEDBString     =   $"login.frx":1A1E9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select VendorID,Password,OwnerEmail,OwnerName,OwnerNumber,FoodShopName,SubscriptionExpriesOn,Status  from VendorSignup"
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
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Signup  Now"
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
      Left            =   6480
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't have an Account yet ?....Then"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3240
      Width           =   4335
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If (Check1.Value = 1) Then
Label4.Caption = "Admin :"
txtvid = "SYSTEM ADMIN"
txtvid.Enabled = False
Else
Label4.Caption = "Vendor ID :"
txtvid = ""
txtvid.Enabled = True
End If
End Sub

Private Sub cmdclear_Click()
txtvid = ""
txtpass = ""
Check1.Value = 0
End Sub

Private Sub cmdhome_Click()
home.Show
login.Hide
cmdclear_Click
End Sub

Private Sub cmdlogin_Click()
If (txtvid = "" Or txtpass = "") Then
MsgBox "Fields Can't be Empty"
Else
If (txtvid = "SYSTEM ADMIN") Then
If (txtpass = "Admin@01") Then
MsgBox "Hi SYSTEM ADMIN  You are successfully logged in"
dashboard.vendormenu.Visible = False
dashboard.adminmenu.Visible = True
dashboard.adminmenu.Enabled = True
dashboard.Label1.Visible = False
dashboard.Label2.Visible = False
dashboard.Show
txtpass = ""
login.Hide
Else
MsgBox "Wrong Password"
txtpass = ""
txtpass.SetFocus
End If
ElseIf Not (txtvid = "SYSTEM ADMIN") Then
Dim name As String
Adodc1.Refresh
With Adodc1.Recordset
.MoveFirst
.Find "VendorID='" & txtvid.Text & "'"
If .EOF = True Then
MsgBox "Sorry the VendorID is not found.. Please signup if you are New Vendor"
txtvid = ""
txtpass = ""
ElseIf (txtpass.Text = Adodc1.Recordset.Fields("Password")) Then
name = Adodc1.Recordset.Fields("OwnerName")
'str1 = Adodc1.Recordset.Fields("SubscriptionExpriesOn")
MsgBox "Hi " & name & " You are successfully logged in"
If DateValue(Format(Now, "dd/MM/yyyy")) > DateValue(Format(Adodc1.Recordset.Fields("SubscriptionExpriesOn"), "dd/MM/yyyy")) Then
'MsgBox (Format(Now, "dd/MM/yyyy")) & (Format(Adodc1.Recordset.Fields("SubscriptionExpriesOn"), "dd/MM/yyyy"))
Adodc1.Recordset.Fields("Status") = "InActive"
Adodc1.Recordset.Update
dashboard.manageitemsmenu.Enabled = False
dashboard.manageordersmenu.Enabled = False
dashboard.Label1.ForeColor = &HFF&
dashboard.Label1.Caption = "Subscription Has Been Expired ,Your Account As Been InActive."
dashboard.Label2.Caption = "Please Goto Profile page and Renew the Subscription."
ElseIf Adodc1.Recordset.Fields("Status") = "Suspended" Then
dashboard.manageitemsmenu.Enabled = False
dashboard.manageordersmenu.Enabled = False
dashboard.Label1.ForeColor = &HFF&
dashboard.Label1.Caption = "                                          Your Account As Been Suspended."
dashboard.Label2.Caption = "Please Goto Contact page and Contact admin throught email for enquiry."
Else
Adodc1.Recordset.Fields("Status") = "Active"
Adodc1.Recordset.Update
dashboard.Label2.Visible = False
End If
dashboard.adminmenu.Visible = False
dashboard.vendormenu.Enabled = True
dashboard.Show
txtpass = ""
login.Hide
Else
MsgBox "Wrong Password"
txtpass = ""
txtpass.SetFocus
End If
End With
End If
End If
End Sub

Private Sub Form_Load()
txtpass = ""
txtvid = ""
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = (Screen.Height) / 2 - Me.Height / 2
End Sub

Private Sub Label15_Click()
login.Hide
vendorsignup.Show
End Sub

Private Sub Label3_Click()
Dim vid As String
Dim num As String
Dim test As String
Dim test1 As String
Dim test2 As String
Dim newpass As String
Adodc1.Refresh
vid = InputBox("" + vbNewLine + vbNewLine + "Enter your Vendor ID", "FORGET PASSWORD")
With Adodc1.Recordset
.MoveFirst
.Find "VendorID='" & vid & "'"
If .EOF = True Then
MsgBox "" + vbNewLine + vbNewLine + "VendorID doesn't Exist, Please check your correct VendorID", , "FORGET PASSWORD"
Else
test = Adodc1.Recordset("OwnerNumber")
test1 = Mid(test, 8)
test2 = "*******" & test1
num = InputBox("" + vbNewLine + vbNewLine + "Enter your Full number " & test2 & "", "FORGET PASSWORD")
If (num = Adodc1.Recordset.Fields("OwnerNumber")) Then
newpass = InputBox("" + vbNewLine + vbNewLine + "Enter your New Password", "FORGET PASSWORD")
If Not (newpass = "") And ValidatePassword(newpass) = True Then
Adodc1.Recordset.Fields("Password") = newpass
Adodc1.Recordset.Update
MsgBox "" + vbNewLine + vbNewLine + "Your Password as been changed , please note your password " & newpass, , "FORGET PASSWORD"
Else
MsgBox "" + vbNewLine + vbNewLine + " Invalid Password, Please Enter an Vaild password", , "FORGET PASSWORD"
End If
Else
MsgBox "" + vbNewLine + vbNewLine + "Entered Owner Number is not Matching", , "FORGET PASSWORD"
End If
End If
End With
End Sub

Private Sub txtpass_Click()
If (txtvid = "") Then
txtvid.SetFocus
End If
End Sub

Private Sub txtvid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpass.SetFocus
End If
End Sub
Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdlogin_Click
End If
End Sub
Private Function ValidatePassword(ByVal sPass) As Boolean
Dim regEx
Set regEx = CreateObject("vbscript.regexp")
regEx.Pattern = "^.*(?=.{8,})(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[!@#$%^&+=]).*$"
ValidatePassword = regEx.test(sPass)
Set regEx = Nothing
End Function
