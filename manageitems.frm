VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form manageitems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Items"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   18345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Connect         =   $"manageitems.frx":0000
      OLEDBString     =   $"manageitems.frx":00E3
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   18615
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
         TabIndex        =   21
         Top             =   120
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
         Left            =   1920
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H0000FFFF&
         Caption         =   "Save"
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
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin MSComctlLib.ListView rolls 
         Height          =   3255
         Left            =   5940
         TabIndex        =   16
         Top             =   4920
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Chicken/Veg Rolls"
            Object.Width           =   4745
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Starters 
         Height          =   3255
         Left            =   360
         TabIndex        =   15
         Top             =   4920
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Starters"
            Object.Width           =   4745
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Snacks 
         Height          =   3255
         Left            =   11520
         TabIndex        =   14
         Top             =   1440
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Snacks"
            Object.Width           =   4745
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Rice 
         Height          =   3255
         Left            =   5940
         TabIndex        =   13
         Top             =   1440
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rice Items"
            Object.Width           =   4746
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin MSComctlLib.ListView Burger 
         Height          =   3255
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Burgers / Sandwichs"
            Object.Width           =   4745
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price Rs./-"
            Object.Width           =   2364
         EndProperty
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H0000C000&
         Caption         =   "Add"
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
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H000000FF&
         Caption         =   "Delete"
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
         Left            =   13125
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdmodify 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Modify"
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
         Left            =   11325
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   1455
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
         ItemData        =   "manageitems.frx":01C6
         Left            =   6480
         List            =   "manageitems.frx":01D0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtitemid 
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
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtprice 
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
         Left            =   6480
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   2175
      End
      Begin MSComctlLib.ListView icecream 
         Height          =   3255
         Left            =   11520
         TabIndex        =   17
         Top             =   4920
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IceCreams / MilkShakes"
            Object.Width           =   4745
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price Rs./-"
            Object.Width           =   2363
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "ITEM NAME :"
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
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "ITEM ID :"
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
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "ITEM PRICE :"
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
         Left            =   5040
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "MENU CATEGORY :"
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
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "MANAGE ITEMS"
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
      TabIndex        =   7
      Top             =   0
      Width           =   3255
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
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "manageitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
ItemsDetails.RecordSource = "Select * from ItemsDetails"
Dim id As String
Dim id1 As String
On Error GoTo errmsg
ItemsDetails.Refresh
ItemsDetails.Recordset.MoveLast
id1 = ItemsDetails.Recordset("ItemID")
id = Mid(id1, 2, 4) + 1
ItemsDetails.Recordset.AddNew
txtitemid = "I" & id
txtname = ""
txtprice = ""
Exit Sub
errmsg:
ItemsDetails.Recordset.AddNew
txtitemid = "I1001"
txtname = ""
txtprice = ""
End Sub

Private Sub cmdclear_Click()
Call clear
End Sub

Private Sub cmddelete_Click()
On Error GoTo errmsg
Dim id1 As String
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "'"
ItemsDetails.Refresh
id1 = InputBox(vbNewLine + vbNewLine + "Enter the Itemid to be deleted", "DELETE ITEM")
ItemsDetails.Recordset.MoveFirst
ItemsDetails.Recordset.Find "ItemID='" & id1 & "'"
If ItemsDetails.Recordset.EOF Then
MsgBox vbNewLine + "  ItemID doesn't Exist, Please check your correct ItemID", vbCritical, "DELETE ITEM"
Else
Dim wish As String
wish = MsgBox("do you really want to delete(y/n) ?", vbYesNoCancel + vbInformation, "DELETE ITEM")
If wish = 6 Then
ItemsDetails.Recordset.Delete
MsgBox "Item deleted successfully", vbInformation, "DELETE ITEM"
Call clear
Else
End If
End If
Call data
errmsg:
MsgBox Err.Description
End Sub

Private Sub cmdmodify_Click()
Dim id As String
id = InputBox(vbNewLine + vbNewLine + "Please Enter Item ID", "MODIFY ITEM")
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "'"
ItemsDetails.Refresh
With ItemsDetails.Recordset
.MoveFirst
.Find "ItemID='" & id & "'"
If .EOF = True Then
MsgBox vbNewLine + "  ItemID doesn't Exist, Please check your correct ItemID", vbCritical, "MODIFY ITEM"
Else
txtitemid = ItemsDetails.Recordset("ItemID")
txtname = ItemsDetails.Recordset("ItemName")
txtprice = ItemsDetails.Recordset("ItemPrice")
'Combo1.Style 0
Combo1 = ItemsDetails.Recordset("MenuCategory")
'Combo1.Style 2
End If
End With
End Sub

Private Sub cmdsave_Click()
If txtitemid = "" Or txtname = "" Or txtprice = "" Or Combo1 = "" Then
MsgBox "Fields Can't Be Empty, Please fill all the details"
Else
ItemsDetails.Recordset.Fields("ItemID") = txtitemid
ItemsDetails.Recordset.Fields("ItemName") = txtname
ItemsDetails.Recordset.Fields("MenuCategory") = Combo1.Text
ItemsDetails.Recordset.Fields("ItemPrice") = txtprice
ItemsDetails.Recordset.Fields("VendorID") = login.txtvid
ItemsDetails.Recordset.Update
MsgBox "Item " & txtname & " Details Saved Successfully"
End If
Call data
Call clear
End Sub


Private Sub Combo1_DropDown()
If txtitemid = "" Then
MsgBox "Please click on ADD Button"
cmdadd.SetFocus
Else
End If
End Sub

Private Sub Form_Load()
Call clear
Call data
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
Function data()
Burger.ListItems.clear
Rice.ListItems.clear
Snacks.ListItems.clear
Starters.ListItems.clear
rolls.ListItems.clear
icecream.ListItems.clear
Dim lv As ListItem
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "' and MenuCategory='Burger'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Burger.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemID")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemName"))
lv.SubItems(2) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "' and MenuCategory='Rice'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Rice.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemID")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemName"))
lv.SubItems(2) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "' and MenuCategory='Snacks'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Snacks.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemID")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemName"))
lv.SubItems(2) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "' and MenuCategory='Starters'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = Starters.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemID")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemName"))
lv.SubItems(2) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "' and MenuCategory='Rolls'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = rolls.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemID")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemName"))
lv.SubItems(2) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
ItemsDetails.RecordSource = "select * from ItemsDetails where VendorID='" & login.txtvid & "' and MenuCategory='Icecream'"
ItemsDetails.Refresh
If (ItemsDetails.Recordset.BOF = False) Then
ItemsDetails.Recordset.MoveFirst
For i = 1 To ItemsDetails.Recordset.RecordCount
Set lv = icecream.ListItems.Add(, , CStr(ItemsDetails.Recordset.Fields("ItemID")))
lv.SubItems(1) = CStr(ItemsDetails.Recordset.Fields("ItemName"))
lv.SubItems(2) = CStr(ItemsDetails.Recordset.Fields("ItemPrice"))
ItemsDetails.Recordset.MoveNext
Next
End If
End Function
Function clear()
txtitemid = ""
txtname = ""
txtprice = ""
Combo1.clear
Combo1.AddItem "Rice"
Combo1.AddItem "Burger"
Combo1.AddItem "Snacks"
Combo1.AddItem "Starters"
Combo1.AddItem "Rolls"
Combo1.AddItem "Icecream"
End Function


Private Sub txtname_Click()
If txtitemid = "" Then
MsgBox "Please click on ADD Button"
cmdadd.SetFocus
Else
End If
End Sub
Private Sub txtprice_click()
If txtitemid = "" Then
MsgBox "Please click on ADD Button"
cmdadd.SetFocus
Else
End If
End Sub

