VERSION 5.00
Begin VB.Form dashboard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDER YOUR FOOD"
   ClientHeight    =   9600
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   18600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   18600
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   18255
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6480
         TabIndex        =   2
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label Label1 
         Caption         =   "Your Account is in Good Condition, You Can Control Your Items and Orders"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   0
         Width           =   8295
      End
   End
   Begin VB.Image Image1 
      Height          =   16890
      Left            =   3240
      Picture         =   "dashboard.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   30000
   End
   Begin VB.Menu adminmenu 
      Caption         =   "ADMIN"
      Begin VB.Menu managevendorsmenu 
         Caption         =   "Manage Vendors"
      End
      Begin VB.Menu subscriptionhistorymenu 
         Caption         =   "Subscription Report"
      End
   End
   Begin VB.Menu vendormenu 
      Caption         =   "VENDOR"
      Begin VB.Menu manageitemsmenu 
         Caption         =   "Manage Items"
      End
      Begin VB.Menu manageordersmenu 
         Caption         =   "Manage Orders"
      End
      Begin VB.Menu viewbillsmenu 
         Caption         =   "Bills Report"
      End
      Begin VB.Menu mnucontact 
         Caption         =   "Contact Admin"
      End
      Begin VB.Menu mnuprofile 
         Caption         =   "Profile"
      End
   End
   Begin VB.Menu exitmenu 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exitmenu_Click()
MsgBox "Thank you for Your Session ," + vbNewLine + "Happy to see you again", vbInformation
Unload Me
login.Show
End Sub

Private Sub Form_Load()
Image1.Top = Frame1.Height
Image1.Left = 0
Image1.Width = Screen.Width
Image1.Height = Screen.Height - 1200 - Frame1.Height
Label1.Left = Screen.Width / 2 - Label1.Width / 2
Label2.Left = Screen.Width / 2 - Label2.Width / 2
End Sub

Private Sub manageitemsmenu_Click()
manageitems.Show , dashboard
manageitems.cmdadd.SetFocus
End Sub

Private Sub manageordersmenu_Click()
manageorders.Show , dashboard
End Sub

Private Sub managevendorsmenu_Click()
managevendors.Show , dashboard
End Sub

Private Sub mnucontact_Click()
contact.Label1.Caption = "Hello ," & profile.Adodc1.Recordset.Fields("Ownername")
contact.Label2.Caption = "Happy to see you here ," + vbNewLine + "If you have any issues / Account Suspension You can feel free to contact us and enquiry."
contact.Label3.Caption = "Admin@gamil.com"
contact.Label4.Caption = "You can contact us throught email for this below Email Address"
contact.Show , dashboard
End Sub

Private Sub mnuprofile_Click()
profile.Show , dashboard
End Sub

Private Sub subscriptionhistorymenu_Click()
subscriptionhistory.Show , dashboard
End Sub

Private Sub viewbillsmenu_Click()
bills.Show , dashboard
End Sub
