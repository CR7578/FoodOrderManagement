VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Customerpay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PAYMENT GATEWAY"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9960
      Top             =   240
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
      TabIndex        =   6
      Top             =   6360
      Width           =   1695
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
      TabIndex        =   5
      Top             =   6360
      Width           =   1815
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
      TabIndex        =   4
      Top             =   5640
      Width           =   2535
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
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
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
      TabIndex        =   2
      Top             =   5640
      Width           =   1935
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
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10800
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   582
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
      Connect         =   $"Customerpay.frx":0000
      OLEDBString     =   $"Customerpay.frx":00E3
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Customerpay.frx":01C6
      Height          =   4215
      Left            =   10800
      TabIndex        =   0
      Top             =   1080
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
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
      Left            =   8160
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
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
      TabIndex        =   23
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label16 
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
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label14 
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
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name :"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID :"
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
      Left            =   1320
      TabIndex        =   19
      Top             =   2280
      Width           =   735
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
      Left            =   600
      TabIndex        =   18
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Number :"
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
      Left            =   3720
      TabIndex        =   17
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID :"
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
      Left            =   1080
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
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
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
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
      Left            =   5880
      TabIndex        =   14
      Top             =   3360
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
      TabIndex        =   13
      Top             =   4560
      Width           =   4575
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
      TabIndex        =   12
      Top             =   2640
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
      Left            =   5880
      TabIndex        =   11
      Top             =   3720
      Width           =   2415
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
      TabIndex        =   10
      Top             =   4080
      Width           =   2415
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
      Left            =   6600
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
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
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
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
      Left            =   8640
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   0
      Picture         =   "Customerpay.frx":01DB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "Customerpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ta As String
Dim sk As String
Dim list As String

Function orderitems()
list = ""
Dim j As Integer
j = 1
For i = 1 To orders.Order.ListItems.Count
list = list + vbNewLine + CStr(orders.Order.ListItems(j).Text & " - " & orders.Order.ListItems(j).SubItems(2))
j = j + 1
Next
Adodc1.Recordset.Fields("OrderItems") = CStr(list)
End Function
Function random()
random = CLng((1000 - 9999) * Rnd + 9999)
sk = random
On Error GoTo errmsg
Adodc1.Refresh
With Adodc1.Recordset
.MoveFirst
For i = 1 To .RecordCount
If sk = .Fields("SecretKey") Then
random = CLng((1000 - 9999) * Rnd + 9999)
sk = random
.MoveNext
End If
Next
End With
errmsg:
sk = random
End Function

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdpay_Click()
Call random
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False Then
MsgBox "Please select the Payment Method."
Else
Dim otp As String
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
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("BillID") = Label11.Caption
Adodc1.Recordset.Fields("CustomerName") = Label13.Caption
Adodc1.Recordset.Fields("CustomerNumber") = Label7.Caption
Adodc1.Recordset.Fields("VendorID") = Label17.Caption
Adodc1.Recordset.Fields("OrderID") = Label6.Caption
Call orderitems
'Adodc1.Recordset.Fields("OrderItems") = orderitems
Adodc1.Recordset.Fields("FoodSHopName") = Label15.Caption
Adodc1.Recordset.Fields("PaymentMethod") = pm
Adodc1.Recordset.Fields("OrderTime") = Format(Now, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:ss AM/PM")
Adodc1.Recordset.Fields("TotalAmount") = Label8.Caption
Adodc1.Recordset.Fields("OrderStatus") = "Pending"
Adodc1.Recordset.Fields("OrderStatusTime") = "---------"
Adodc1.Recordset.Fields("SecretKey") = sk
Adodc1.Recordset.Update
MsgBox "Completing your Transaction please wait"
MsgBox "Your Transaction has been completed." + vbNewLine + "Note : You Will be getting An Order Key for your Order ," + vbNewLine + "You need to provide Order ID and Order Key only with Vendor at time of taking your Order "
MsgBox ("Your Order ID : " & Label6.Caption + vbNewLine + "Your Order Key: " & sk + vbNewLine + "Order has been placed Successfully")
Call printf
home.Show
Unload Me
Unload orders
orders.Show
End If
End If
End Sub

Private Sub Command1_Click()
Call printf
'Dim a As String
'a = InputBox("Enter order id")
'Adodc1.Recordset.Find "OrderID='" & a & "'"
'MsgBox Adodc1.Recordset.Fields("OrderItems")
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
Function printf()
CommonDialog1.ShowPrinter
Printer.PaperSize = 7

Printer.FontName = "Arial"
Printer.FontSize = 15
Printer.FontBold = True
Printer.CurrentX = 1000
Printer.CurrentY = 1250
Printer.Print "Bill ID: "
Printer.CurrentX = 3900
Printer.CurrentY = 1250
Printer.Print "Date: "
Printer.CurrentX = 6700
Printer.CurrentY = 1250
Printer.Print "Time: "
Printer.CurrentX = 1800
Printer.CurrentY = 4100
Printer.Print "Item Names"
Printer.CurrentX = 4800
Printer.CurrentY = 4100
Printer.Print "Price /-"
Printer.CurrentX = 6000
Printer.CurrentY = 4100
Printer.Print "Quantity"
Printer.CurrentX = 7700
Printer.CurrentY = 4100
Printer.Print "Amount /-"
Printer.CurrentX = 5000
Printer.CurrentY = 13380
Printer.Print "Total Amount /-"
Printer.CurrentX = 4000
Printer.CurrentY = 1900
Printer.Print "Order Details"
Printer.FontBold = False
Printer.FontSize = 12
Printer.CurrentX = 4800
Printer.CurrentY = 1300
Printer.Print Date
Printer.CurrentX = 7600
Printer.CurrentY = 1300
Printer.Print Format(Time, "HH:mm:ss AM/PM")
Printer.CurrentX = 2100
Printer.CurrentY = 1300
Printer.Print Label11.Caption
Printer.CurrentX = 6100
Printer.CurrentY = 2400
Printer.Print Label17.Caption
Printer.CurrentX = 4800
Printer.CurrentY = 3600
Printer.Print Label18.Caption
Printer.CurrentX = 7000
Printer.CurrentY = 2800
Printer.Print Label15.Caption
Printer.FontSize = 12
Printer.CurrentX = 2500
Printer.CurrentY = 2400
Printer.Print Label13.Caption
Printer.CurrentX = 2800
Printer.CurrentY = 2800
Printer.Print Label7.Caption
Printer.CurrentX = 1700
Printer.CurrentY = 3200
Printer.Print Label6.Caption
Printer.CurrentX = 1800
Printer.CurrentY = 3600
Printer.Print sk
Printer.CurrentX = 7700
Printer.CurrentY = 13380
Printer.Print "Rs." & Label8.Caption & " /-"

CurrentY = 4200
For i = 1 To orders.Order.ListItems.Count

    CurrentY = CurrentY + 500
    Printer.CurrentX = 600
    Printer.CurrentY = CurrentY
    Printer.Print CStr(orders.Order.ListItems(i).Text)
    
    Printer.CurrentX = 5100
    Printer.CurrentY = CurrentY
    Printer.Print CStr(orders.Order.ListItems(i).SubItems(1))
    
    Printer.CurrentX = 6500
    Printer.CurrentY = CurrentY
    Printer.Print CStr(orders.Order.ListItems(i).SubItems(2))
    
    Printer.CurrentX = 8100
    Printer.CurrentY = CurrentY
    Printer.Print CStr(orders.Order.ListItems(i).SubItems(3))
    
Next

Printer.FontBold = True
Printer.FontSize = 12
Printer.CurrentX = 4800
Printer.CurrentY = 2400
Printer.Print "Vendor ID: "
Printer.CurrentX = 4800
Printer.CurrentY = 2800
Printer.Print "Food Shop Name: "
Printer.CurrentX = 4800
Printer.CurrentY = 3200
Printer.Print "Food Shop Address: "
Printer.FontSize = 12
Printer.CurrentX = 500
Printer.CurrentY = 2400
Printer.Print "Customer Name: "
Printer.CurrentX = 500
Printer.CurrentY = 2800
Printer.Print "Customer Number: "
Printer.CurrentX = 500
Printer.CurrentY = 3200
Printer.Print "Order ID: "
Printer.CurrentX = 500
Printer.CurrentY = 3600
Printer.Print "Order Key: "

'Header
Printer.FontName = "Arial"
Printer.FontSize = 30
Printer.FontBold = True
Printer.CurrentX = 1500
Printer.CurrentY = 0
Printer.Print "Order Your Food System"

'Sleeping Lines
Printer.DrawWidth = 15
Printer.Line (350, 1000)-(9300, 1000)
Printer.Line (350, 1800)-(9300, 1800)
Printer.Line (350, 2300)-(9300, 2300)
Printer.Line (350, 4000)-(9300, 4000)
Printer.Line (350, 4500)-(9300, 4500)
Printer.Line (350, 13280)-(9300, 13280)
Printer.Line (350, 13780)-(9300, 13780)

'Standing lines
Printer.Line (350, 1000)-(350, 13780)
Printer.Line (7400, 4000)-(7400, 13780)
Printer.Line (5900, 4000)-(5900, 13280)
Printer.Line (4600, 4000)-(4600, 13280)
Printer.Line (9300, 1000)-(9300, 13780)

Printer.EndDoc
End Function



