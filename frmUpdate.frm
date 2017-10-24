VERSION 5.00
Begin VB.Form frmUpdate 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Update Transactions"
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13395
   LinkTopic       =   "Form4"
   ScaleHeight     =   9810
   ScaleWidth      =   13395
   Begin VB.OptionButton rdbExtend 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Extend deadline to another 30 days"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   42
      Top             =   7080
      Width           =   3135
   End
   Begin VB.OptionButton rdbFullPay 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fully Pay the Money"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   41
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   6855
      Left            =   7920
      TabIndex        =   20
      Top             =   360
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   360
         Picture         =   "frmUpdate.frx":0000
         ScaleHeight     =   855
         ScaleWidth      =   855
         TabIndex        =   21
         Top             =   240
         Width           =   855
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
         TabIndex        =   40
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
         Left            =   1320
         TabIndex        =   39
         Top             =   240
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
         TabIndex        =   38
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID:"
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
         TabIndex        =   37
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name:"
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
         TabIndex        =   36
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label14 
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
         TabIndex        =   35
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Money To Reimburse:"
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
         TabIndex        =   34
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label12 
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
         TabIndex        =   33
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date:"
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
         TabIndex        =   32
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Days Until Due:"
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
         TabIndex        =   31
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label lblID 
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
         TabIndex        =   30
         Top             =   1680
         Width           =   2655
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
         TabIndex        =   29
         Top             =   2040
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
         TabIndex        =   28
         Top             =   2400
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
         TabIndex        =   27
         Top             =   3360
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
         TabIndex        =   26
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblDue 
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
         Left            =   2760
         TabIndex        =   25
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label lblDays 
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
         Left            =   2760
         TabIndex        =   24
         Top             =   4440
         Width           =   2295
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
         Left            =   480
         TabIndex        =   23
         Top             =   5280
         Width           =   4455
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   5040
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line4 
         X1              =   360
         X2              =   5040
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line6 
         X1              =   360
         X2              =   5040
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label2 
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
         Left            =   1680
         TabIndex        =   22
         Top             =   6480
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      Picture         =   "frmUpdate.frx":3F47
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   15
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtBranch 
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3240
      Width           =   4815
   End
   Begin VB.TextBox txtTransactID 
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
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtName 
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox txtItem 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   4680
      Width           =   4815
   End
   Begin VB.CommandButton cmdTransSearch 
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
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtDaysRemaining 
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
      Left            =   3960
      TabIndex        =   3
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox txtMoneyPaid 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1560
      TabIndex        =   1
      Top             =   8880
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   585
      Left            =   4440
      TabIndex        =   0
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   7440
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7440
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   43
      Top             =   7800
      Width           =   5895
   End
   Begin VB.Line Line7 
      X1              =   7680
      X2              =   7680
      Y1              =   120
      Y2              =   7800
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Transaction ID:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7440
      Y1              =   2280
      Y2              =   2280
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
      TabIndex        =   18
      Top             =   1560
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
      TabIndex        =   17
      Top             =   1200
      Width           =   5295
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
      TabIndex        =   16
      Top             =   840
      Width           =   5295
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
      TabIndex        =   14
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Item Handed:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Loan Money (Php):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Days Remaining:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   2775
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsTransaction As Recordset
Dim rsClient As Recordset
Dim rsBroker As Recordset
Dim rsItemstock As Recordset
Dim dTransact As Date

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdTransSearch_Click()
    Dim BrokerID As String
    Dim ClientID As String
    Dim Days As Integer
    Dim ItemID As String
    rdbFullPay.Enabled = True
    rdbExtend.Enabled = True
    dTransact = Now()
    SQL = "select * from transaction where transactionID = '" & txtTransactID.Text & "'"
    Set rsTransaction = db.OpenRecordset(SQL)
    If rsTransaction.BOF = True Then
        MsgBox "Transaction is not found"
    Else
        Randomize
        ItemID = "ITEM" & CStr(Int(1000 * Rnd) + 1)
        Days = DateDiff("d", dTransact, rsTransaction.Fields("DateTransact") + rsTransaction.Fields("AllowableDays"))
        'MsgBox Days
        BrokerID = rsTransaction.Fields("brokerid")
        ClientID = rsTransaction.Fields("clientid")
        If Days < 0 Then
            SQL = "select count(*) as count from itemstock where itemid = '" & ItemID & "'"
            Set rsItemstock = db.OpenRecordset(SQL)
            While rsItemstock.Fields("count") > 0
                SQL = "select count(*) from transaction where itemid = '" & ItemID & "'"
                Set rsItemstock = db.OpenRecordset(SQL)
            Wend
            MsgBox "The customer has passed the deadline. His/Her item will be stored by the broker to be sold."
            SQL = "insert into itemstock (ItemID, BrokerID, ItemStored, Price) values ('" & ItemID & "', '" & BrokerID & "', '" & rsTransaction.Fields("Item") & "', '" & rsTransaction.Fields("MoneyLoaned") & "')"
            db.Execute (SQL)
            SQL = "delete from transaction where transactionid = '" & txtTransactID.Text & "'"
            db.Execute (SQL)
            rdbFullPay.Enabled = False
            rdbExtend.Enabled = False
            
        Else
            SQL = "select branchaddress from broker where brokerid = '" & BrokerID & "'"
            Set rsBroker = db.OpenRecordset(SQL)
            txtBranch.Text = rsBroker.Fields(0)
            SQL = "select name from client where customerid = '" & ClientID & "'"
            Set rsClient = db.OpenRecordset(SQL)
            txtName.Text = rsClient.Fields("name")
            SQL = "select itemstored from itemstock where ItemID = '" & ItemID & "'"
            Set rsClient = db.OpenRecordset(SQL)
            txtItem.Text = rsTransaction.Fields("item")
            txtDaysRemaining.Text = Days
            txtMoneyPaid.Text = Format(rsTransaction.Fields("moneyloaned"), "currency")
            cmdUpdate.Enabled = True
        End If
    End If
End Sub

Private Sub cmdUpdate_Click()
dTransact = Now()
SQL = "select * from transaction where transactionID = '" & txtTransactID.Text & "'"
Set rsTransaction = db.OpenRecordset(SQL)
If rsTransaction.Fields("MoneyLoaned") = 0 Then
    MsgBox "You have fully paid your loaned money."
    SQL = "delete from transaction where transactionID = '" & txtTransactID.Text & "'"
    db.Execute (SQL)
Else
    If rdbFullPay.Value = True Then
        SQL = "update transaction set Moneyloaned = 0 where transactionID = '" & txtTransactID.Text & "'"
        db.Execute (SQL)
        MsgBox "You have fully paid your loaned money."
        SQL = "delete from transaction where transactionID = '" & txtTransactID.Text & "'"
        db.Execute (SQL)
        cmdUpdate.Enabled = False
    ElseIf rdbExtend.Value = True Then
        SQL = "update transaction set Moneyloaned = '" & txtMoneyPaid.Text & "', allowabledays = '" & txtDaysRemaining.Text & "', datetransact = '" & Format(dTransact, "mm/dd/yy") & "' where transactionID = '" & txtTransactID.Text & "'"
        db.Execute (SQL)
        MsgBox "Transaction is updated."
        cmdUpdate.Enabled = False
        lblBranch.Caption = txtBranch.Text
        lblID.Caption = txtTransactID.Text
        lblName.Caption = txtName.Text
        lblItem.Caption = txtItem.Text
        lblMoney.Caption = Format(txtMoneyPaid.Text, "Currency")
        lblDateTrans.Caption = dTransact
        lblDue = txtDaysRemaining.Text
        lblDays = Format(dTransact + txtDaysRemaining.Text, "mm/dd/yy")
    End If
    txtTransactID.Text = ""
    txtBranch.Text = ""
    txtName.Text = ""
    txtItem.Text = ""
    txtDaysRemaining.Text = ""
    txtMoneyPaid = ""
    rdbFullPay.Value = False
    rdbExtend.Value = False
    rdbFullPay.Enabled = False
    rdbExtend.Enabled = False
    lblInfo.Caption = ""
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\PawnstarPawnshop.mdb")
Set rsClient = db.OpenRecordset("Client")
Set rsTransaction = db.OpenRecordset("Transaction")
Set rsBroker = db.OpenRecordset("Broker")
Set rsItemstock = db.OpenRecordset("Itemstock")
cmdUpdate.Enabled = False
txtMoneyPaid.Enabled = False
txtName.Enabled = False
txtBranch.Enabled = False
txtDaysRemaining.Enabled = False
txtItem.Enabled = False
rdbFullPay.Enabled = False
rdbExtend.Enabled = False
End Sub


Private Sub rdbExtend_Click()
    cmdUpdate.Enabled = True
    lblInfo.Caption = "Loan money with be added 20% and the due date will be extended to another 30 days."
    txtDaysRemaining.Text = txtDaysRemaining.Text + 30
    txtMoneyPaid.Text = Format(rsTransaction.Fields("MoneyLoaned") + (rsTransaction.Fields("MoneyLoaned") * 0.2), "Currency")
End Sub

Private Sub rdbFullPay_Click()
    cmdUpdate.Enabled = True
    lblInfo.Caption = "The item will be retrieved and the debt will be fully paid."
    txtDaysRemaining.Text = 0
    txtMoneyPaid.Text = Format(0, "Currency")
End Sub

Private Sub cmbBranch_Click()
    cmdCustSearch.Enabled = True
End Sub

