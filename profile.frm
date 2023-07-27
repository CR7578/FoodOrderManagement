VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form profile 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Profile"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtadd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   9120
      TabIndex        =   17
      Text            =   "Text5"
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Subscription Renewal"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   12
      Top             =   7080
      Width           =   5895
      Begin VB.CommandButton cmdbuy 
         Caption         =   "Purchase"
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
         Left            =   3720
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H0000C000&
      Caption         =   "Edit Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H0000FFFF&
      Caption         =   "Update Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12960
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
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
      Connect         =   $"profile.frx":0000
      OLEDBString     =   $"profile.frx":00E3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from VendorSignup;"
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
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox txtnum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   5880
      Width           =   3375
   End
   Begin VB.TextBox txtemail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   4680
      Width           =   3615
   End
   Begin VB.TextBox txtshop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5280
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   5880
      Width           =   3135
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   9120
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text6"
      Top             =   4665
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Food Shop Address :"
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
      Left            =   9120
      TabIndex        =   18
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Account is Active"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   6600
      Width           =   5775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Email :"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Number :"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Food Shop Name :"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name :"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
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
      Left            =   9120
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   4920
      Picture         =   "profile.frx":01C6
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbuy_Click()
If Combo1.Text = "" Then
MsgBox "Please Select Subscription Plan."
Combo1.SetFocus
Else
Dim ta As String
If (Combo1.Text = "Rs.24000/- Per Year") Then
ta = "24000"
ElseIf (Combo1.Text = "Rs.2000/- Per Month") Then
ta = "2000"
End If
Dim id As Integer
Dim id1 As String
renewal.Adodc1.Refresh
renewal.Adodc1.Recordset.MoveLast
id1 = renewal.Adodc1.Recordset("PaymentID")
id = Mid(id1, 2, 4) + 1
'renewal.Adodc1.Recordset.AddNew
renewal.Label11.Caption = "P" & id
renewal.Label8.Caption = ta
renewal.Label6.Caption = login.txtvid
renewal.Label7.Caption = Combo1.Text
renewal.Show
End If
End Sub

Private Sub cmdedit_Click()
txtname.Enabled = True
txtnum.Enabled = True
txtemail.Enabled = True
txtshop.Enabled = True
txtpass.Enabled = True
txtadd.Enabled = True
End Sub

Private Sub cmdupdate_Click()
Adodc1.RecordSource = "Select * from VendorSignup;"
Adodc1.Refresh
If txtname = "" Or txtnum = "" Or txtemail = "" Or txtshop = "" Or txtpass = "" Or txtadd = "" Then
MsgBox "Fields Can't Be Empty"
Else
Adodc1.Recordset.Find "VendorID='" & login.txtvid & "'"
Adodc1.Recordset.Fields("OwnerName") = txtname.Text
Adodc1.Recordset.Fields("OwnerNumber") = txtnum.Text
Adodc1.Recordset.Fields("OwnerEmail") = txtemail.Text
Adodc1.Recordset.Fields("FoodShopName") = txtshop.Text
Adodc1.Recordset.Fields("Password") = txtpass.Text
Adodc1.Recordset.Fields("FoodShopAddress") = txtadd.Text
Adodc1.Recordset.Update
Call data
MsgBox "Your Details Has Been Updated."
txtname.Enabled = False
txtnum.Enabled = False
txtemail.Enabled = False
txtshop.Enabled = False
txtpass.Enabled = False
txtadd.Enabled = False
End If
End Sub
Private Sub txtemail_LostFocus()
If isEmail(txtemail.Text) = True Then
Else
txtemail = ""
MsgBox "Please Enter an Valid Email Address."
txtemail.SetFocus
End If
End Sub

Private Sub txtnum_LostFocus()
If Len(txtnum) = 10 Then
Else
txtnum.Text = ""
MsgBox "Please Enter 10 Digits Number Only"
txtnum.SetFocus
End If
End Sub

Private Sub txtpass_LostFocus()
If ValidatePassword(txtpass.Text) = True Then
Else
txtpass.Text = ""
MsgBox "Entered Password is Not Valid." + vbNewLine + "Please Make Sure the Entered password must Contain the following conditions" + vbNewLine + "Atleast 8 Characters" + vbNewLine + "Atleast 1 Number" + vbNewLine + "Atleast 1 Lowercase letter" + vbNewLine + "Atleast 1 Uppercase letter" + vbNewLine + "Atleast 1 Special Character", vbInformation
txtpass.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = (Screen.Height) / 2 - Me.Height / 2
Call data
If DateValue(Format(Now, "dd/MM/yyyy")) > DateValue(Format(Adodc1.Recordset.Fields("SubscriptionExpriesOn"), "dd/MM/yyyy")) Then
 cmdbuy.Enabled = True
 Combo1.Enabled = True
 Label1.Caption = "Please Renew Your Subscription."
 Label2.Visible = False
 'Frame1.Visible=True
 Else
 Label2.Caption = "Expries on " & Adodc1.Recordset.Fields("SubscriptionExpriesOn")
cmdbuy.Enabled = False
Combo1.Enabled = False
'Frame1.Visible = False
End If
If Adodc1.Recordset.Fields("Status") = "Suspended" Then
 Label1.ForeColor = &HFF&
 Label1.Caption = "Your Account Has Been Suspended."
 Label2.Visible = False
End If
txtname.Enabled = False
txtnum.Enabled = False
txtemail.Enabled = False
txtshop.Enabled = False
txtpass.Enabled = False
txtadd.Enabled = False
Combo1.clear
Combo1.AddItem "Rs.2000/- Per Month"
Combo1.AddItem "Rs.24000/- Per Year"
End Sub
Function data()
Adodc1.RecordSource = "Select * from VendorSignup;"
Adodc1.Refresh
Adodc1.Recordset.Find "VendorID='" & login.txtvid & "'"
txtname.Text = Adodc1.Recordset.Fields("OwnerName")
txtnum.Text = Adodc1.Recordset.Fields("OwnerNumber")
txtemail.Text = Adodc1.Recordset.Fields("OwnerEmail")
txtshop.Text = Adodc1.Recordset.Fields("FoodShopName")
txtpass.Text = Adodc1.Recordset.Fields("Password")
txtadd.Text = Adodc1.Recordset.Fields("FoodShopAddress")
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

