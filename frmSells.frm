VERSION 5.00
Begin VB.Form frmSell 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Sell Items"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   13500
   Begin VB.ListBox listDummy 
      Height          =   450
      Left            =   120
      TabIndex        =   36
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbBranch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      TabIndex        =   34
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   6255
      Left            =   7920
      TabIndex        =   16
      Top             =   360
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   360
         Picture         =   "frmSells.frx":0000
         ScaleHeight     =   855
         ScaleWidth      =   855
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "PAWNSTAR PAWNSHOP"
         BeginProperty Font 
            Name            =   "Blippo Light SF"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "416-9531 | pawnstars@pawn.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   32
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "PAWNSTAR PAWNSHOP"
         BeginProperty Font 
            Name            =   "Blippo Light SF"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   31
         Top             =   -1200
         Width           =   4695
      End
      Begin VB.Label lblBranch 
         BackStyle       =   0  'Transparent
         Caption         =   "N. Bacalso Ave., Cebu City"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   30
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer Name:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Item To Retrieve:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Money Paid:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Transaction:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Remaining:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label lblMoney 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblDateTrans 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblBalance 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "In Pawnstar Pawnshop, we put your money where you need it!"
         BeginProperty Font 
            Name            =   "bromello"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   600
         TabIndex        =   19
         Top             =   4560
         Width           =   4455
      End
      Begin VB.Line Line3 
         X1              =   360
         X2              =   5040
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line4 
         X1              =   360
         X2              =   5040
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line6 
         X1              =   480
         X2              =   5160
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "[THIS IS A SAMPLE SLIP]"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   5760
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      Picture         =   "frmSells.frx":3F47
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox listItem 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2520
      TabIndex        =   7
      Top             =   4560
      Width           =   4935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSearchBidder 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "Sell"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox txtBrokerID 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox txtAmount 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   5160
      Width           =   4935
   End
   Begin VB.TextBox txtBidder 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Branch:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   33
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line5 
      X1              =   7680
      X2              =   7680
      Y1              =   120
      Y2              =   6960
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7440
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Caption         =   "N. Bacalso Ave., Cebu City"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Amount :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Item Name :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Bidder ID :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PAWNSTARS PAWNSHOP"
      BeginProperty Font 
         Name            =   "Blippo Light SF"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   12
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pawn with us at Tel No. 416-9531"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Caption         =   "or email us at pawnstars@pawn.com"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Broker ID :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
End
Attribute VB_Name = "frmSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim ten As String
Dim rsBroker As Recordset
Dim rsBuyer As Recordset
Dim rsItem As Recordset
Dim rsSells As Recordset


Private Sub cmbBranch_Click()
    SQL = "select BrokerID from broker where branchaddress = '" & cmbBranch.Text & "'"
    Set rsBroker = db.OpenRecordset(SQL)
    txtBrokerID.Text = rsBroker.Fields(0)
     cmdSearchBidder.Enabled = True
     txtBidder.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearchBidder_Click()

SQL = "select * from Buyer where CustomerID = '" & txtBidder.Text & "'"
Set rsBuyer = db.OpenRecordset(SQL)
If rsBuyer.BOF = True Then
    MsgBox "ID number does not exists", vbCritical, "Error"
Else
    While (rsItem.EOF = False)
        If rsItem.Fields("BrokerID") = txtBrokerID.Text Then
            listItem.AddItem rsItem.Fields("ItemStored")
            listDummy.AddItem rsItem.Fields("itemid")
        End If
        rsItem.MoveNext
    Wend
    If listItem.ListCount = 0 Then
        MsgBox "No items remaining from the broker."
    Else
        cmdSell.Enabled = True
        listItem.Enabled = True
    End If
    
End If
End Sub

Private Sub cmdSearchBroker_Click()
SQL = "select * from Broker where BrokerID = '" & txtBroker.Text & "'"
Set rsBroker = db.OpenRecordset(SQL)
If rsBroker.BOF = True Then
MsgBox "ID number does not exists"
Else
cmdSearchBidder.Enabled = True
End If
End Sub

Private Sub cmdSell_Click()
    If rsBuyer.Fields("BidMoney") < Val(txtAmount.Text) Then
        MsgBox "Bidder doesn't have enough money. Try again."
    Else
        SQL = "update buyer set bidmoney = '" & rsBuyer.Fields("BidMoney") - Val(txtAmount.Text) & "'"
        db.Execute (SQL)
        SQL = "insert into sells values ('" & txtBidder.Text & "', '" & txtBrokerID.Text & "','" & listItem.Text & "')"
        db.Execute (SQL)
        SQL = "delete from itemstock where itemid = '" & ten & "'"
        db.Execute (SQL)
        MsgBox "'Item Sold!"
        lblBranch.Caption = cmbBranch.Text
        SQL = "select name, bidmoney from buyer where customerid = '" & txtBidder.Text & "'"
        Set rsBuyer = db.OpenRecordset(SQL)
        lblName.Caption = rsBuyer.Fields("name")
        lblItem.Caption = listItem.Text
        lblMoney.Caption = Format(txtAmount.Text, "Currency")
        lblDateTrans.Caption = Format(Now(), "mm/dd/yy")
        lblBalance = Format(rsBuyer.Fields("bidmoney"), "Currency")
        txtBidder.Text = ""
        listItem.Clear
        txtAmount.Text = ""
        listItem.Text = ""
        cmbBranch.Text = ""
        txtBrokerID.Text = ""
    End If
End Sub

Private Sub Form_Load()
txtBidder.Text = ""
listItem.Text = ""
txtAmount.Text = ""
txtBrokerID.Text = ""

Set db = OpenDatabase(App.Path & "\PawnstarPawnshop.mdb")
Set rsBuyer = db.OpenRecordset("Buyer")
Set rsBroker = db.OpenRecordset("Broker")
Set rsItem = db.OpenRecordset("Itemstock")
Set rsSells = db.OpenRecordset("Sells")
txtBrokerID.Enabled = False
txtBidder.Enabled = False
cmdSearchBidder.Enabled = False
listItem.Enabled = False
txtAmount.Enabled = False
cmdSell.Enabled = False
Set db = OpenDatabase(App.Path & "\PawnstarPawnshop.mdb")
    Set rsClient = db.OpenRecordset("Client")
    Set rsBroker = db.OpenRecordset("Broker")
    Set rsItemsClient = db.OpenRecordset("itemsclient")
    While (rsBroker.EOF = False)
        cmbBranch.AddItem rsBroker.Fields("branchaddress")
        rsBroker.MoveNext
    Wend

End Sub

Private Sub listItem_DblClick()
    SQL = " select itemid, itemstored as UniqueItems, price from itemstock where itemstored = '" & listItem.Text & "'"
    Set rsItem = db.OpenRecordset(SQL)
    txtAmount.Text = rsItem.Fields("price")
    ten = listDummy.List(listItem.ListIndex)
End Sub

