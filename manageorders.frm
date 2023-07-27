VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form manageorders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Orders"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   18570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   8040
      Top             =   120
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
      Connect         =   $"manageorders.frx":0000
      OLEDBString     =   $"manageorders.frx":00E3
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
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
      Connect         =   $"manageorders.frx":01C6
      OLEDBString     =   $"manageorders.frx":02A9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from OrderDetails"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   18135
      Begin VB.CommandButton cmdcancel 
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
         Left            =   16560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdcomplete 
         BackColor       =   &H0000C000&
         Caption         =   "Complete"
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
         Left            =   16560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdready 
         BackColor       =   &H0000FFFF&
         Caption         =   "Ready"
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
         Left            =   16560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdfilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "View"
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
         Left            =   16560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3600
         Width           =   1455
      End
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
         Left            =   16560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "manageorders.frx":038C
         Height          =   7935
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   13996
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "MANAGE ORDERS"
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
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "manageorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdcomplete_Click()
Dim sk As String
With Adodc2.Recordset
.MoveFirst
.Find "OrderID='" & Adodc1.Recordset.Fields("OrderID") & "'"
'MsgBox .Fields("Secretkey")
sk = InputBox("Enter Order Key to Complete the Order")
If .EOF Then
MsgBox "Entered OrderID is Not vaild."
ElseIf sk = Adodc2.Recordset.Fields("Secretkey") Then
Adodc1.Recordset.Fields("OrderStatus") = "Completed"
Adodc1.Recordset.Fields("OrderStatusTime") = Time
Adodc1.Recordset.Update
MsgBox "Order ID: " & Adodc1.Recordset.Fields("OrderID") & vbNewLine + "Has been Updated to Complete Status."
Else
MsgBox "Invalid OrderKey, Please Check Your OrderKey"
End If
End With
Adodc1.Refresh
Call ref
End Sub

Private Sub cmdfilter_Click()
With Adodc2.Recordset
.MoveFirst
.Find "OrderID='" & Adodc1.Recordset.Fields("OrderID") & "'"
MsgBox .Fields("OrderItems")
End With
End Sub

Private Sub cmdready_Click()
Adodc1.Recordset.Fields("OrderStatus") = "Ready"
Adodc1.Recordset.Fields("OrderStatusTime") = Time
Adodc1.Recordset.Update
MsgBox "Order ID: " & Adodc1.Recordset.Fields("OrderID") & vbNewLine + "Has been Updated to Ready Status."
Adodc1.Refresh
Call ref
End Sub

Private Sub cmdrefresh_Click()
Call ref
End Sub

Private Sub DataGrid1_Click()
If Adodc1.Recordset.Fields("OrderStatus") = "Pending" Then
cmdready.Enabled = True
cmdcomplete.Enabled = False
'cmdcancel.Enabled = True
ElseIf Adodc1.Recordset.Fields("OrderStatus") = "Ready" Then
cmdready.Enabled = False
cmdcomplete.Enabled = True
'cmdcancel.Enabled = False
ElseIf Adodc1.Recordset.Fields("OrderStatus") = "Completed" Then
cmdready.Enabled = False
cmdcomplete.Enabled = False
'cmdcancel.Enabled = False
End If
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
Adodc1.RecordSource = "select OrderID,BillID,CustomerName,CustomerNumber,OrderTime,TotalAmount,PaymentMethod,OrderStatus,OrderStatusTime from OrderDetails where VendorID='" & login.txtvid & "'"
Adodc1.Refresh
Adodc2.RecordSource = "select * from OrderDetails"
Adodc2.Refresh
End Sub
Function ref()
Adodc1.RecordSource = "select OrderID,BillID,CustomerName,CustomerNumber,OrderTime,TotalAmount,PaymentMethod,OrderStatus,OrderStatusTime from OrderDetails where VendorID='" & login.txtvid & "'"
Adodc1.Refresh
End Function

