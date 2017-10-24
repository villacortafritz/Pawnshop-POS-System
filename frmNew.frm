VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Add a Transaction"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   LinkTopic       =   "Form4"
   ScaleHeight     =   10935
   ScaleWidth      =   13440
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
      Left            =   4080
      TabIndex        =   46
      Top             =   9600
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Height          =   6855
      Left            =   7920
      TabIndex        =   25
      Top             =   360
      Width           =   5295
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   360
         Picture         =   "frmNew.frx":0000
         ScaleHeight     =   855
         ScaleWidth      =   855
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label42 
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
         TabIndex        =   45
         Top             =   6480
         Width           =   3855
      End
      Begin VB.Line Line9 
         X1              =   360
         X2              =   5040
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line8 
         X1              =   360
         X2              =   5040
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         X1              =   360
         X2              =   5040
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label41 
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
         TabIndex        =   44
         Top             =   5280
         Width           =   4455
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
         TabIndex        =   43
         Top             =   4440
         Width           =   2295
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
         TabIndex        =   42
         Top             =   4080
         Width           =   2295
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
         TabIndex        =   41
         Top             =   3720
         Width           =   2175
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
         TabIndex        =   40
         Top             =   3360
         Width           =   2175
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
         TabIndex        =   39
         Top             =   2400
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
         TabIndex        =   38
         Top             =   2040
         Width           =   2655
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
         TabIndex        =   37
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label33 
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
         TabIndex        =   36
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label32 
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
         TabIndex        =   35
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label31 
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
         TabIndex        =   34
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label30 
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
         TabIndex        =   33
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label29 
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
         TabIndex        =   32
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label28 
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label27 
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
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
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
         TabIndex        =   29
         Top             =   600
         Width           =   3255
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
         TabIndex        =   28
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label9 
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
         TabIndex        =   27
         Top             =   840
         Width           =   4575
      End
   End
   Begin VB.TextBox txtNewValue 
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
      TabIndex        =   24
      Top             =   8520
      Width           =   4815
   End
   Begin VB.TextBox txtValue 
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
      TabIndex        =   23
      Top             =   7800
      Width           =   4815
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      TabIndex        =   22
      Top             =   6480
      Width           =   4815
   End
   Begin VB.CommandButton cmdCustSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   18
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox cmbCategory 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2520
      TabIndex        =   17
      Top             =   5880
      Width           =   4815
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
      Left            =   2520
      TabIndex        =   16
      Top             =   4440
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
      Left            =   2520
      TabIndex        =   15
      Top             =   5160
      Width           =   4815
   End
   Begin VB.TextBox txtCustID 
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
      TabIndex        =   14
      Top             =   3720
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      Picture         =   "frmNew.frx":3F47
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   10
      Top             =   120
      Width           =   1815
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
      TabIndex        =   8
      Top             =   3000
      Width           =   4815
   End
   Begin VB.ComboBox cmbBranch 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2520
      TabIndex        =   7
      Top             =   2400
      Width           =   4815
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "Transact"
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
      Left            =   720
      TabIndex        =   0
      Top             =   9600
      Width           =   2895
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7320
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Caption         =   "New Value:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Item Value:"
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
      Left            =   120
      TabIndex        =   20
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Description:"
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
      Left            =   120
      TabIndex        =   19
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Line Line7 
      X1              =   7680
      X2              =   7680
      Y1              =   120
      Y2              =   10800
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7320
      Y1              =   2160
      Y2              =   2160
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
      TabIndex        =   13
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
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
      TabIndex        =   12
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label16 
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
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Broker ID:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Customer ID:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Item Handed:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      TabIndex        =   2
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Item Category:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   2415
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsBroker As Recordset
Dim rsClient As Recordset
Dim itemsClient As Recordset
Dim dTransact As Date
Dim pawnMoney As Double
Dim rate As Double

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmbBranch_Click()
    cmdCustSearch.Enabled = True
    SQL = "select BrokerID from broker where branchaddress = '" & cmbBranch.Text & "'"
    Set rsBroker = db.OpenRecordset(SQL)
    txtBrokerID.Text = rsBroker.Fields(0)
    txtCustID.Enabled = True
End Sub

Private Sub cmbCategory_Click()
 If cmbCategory.ListIndex >= 0 Then
     txtValue.Enabled = True
     txtDescription.Enabled = True
End If
End Sub

Private Sub cmdCustSearch_Click()
SQL = "select * from client where CustomerID = '" & txtCustID.Text & "'"
Set rsClient = db.OpenRecordset(SQL)

If rsClient.BOF = True Then
    MsgBox "Customer ID Number does not exist."
Else
    txtCustID.Text = rsClient.Fields("CustomerID")
    txtName.Text = rsClient.Fields("Name")
    txtItem.Enabled = True
End If
End Sub

Private Sub cmdTransact_Click()
    Dim random As Integer

    pawnMoney = 0
    dTransact = Now()

    Randomize
    random = Int(10000 * Rnd) + 1
    pawnMoney = Val(txtNewValue.Text) * 1.1
    'MsgBox pawnMoney
    SQL = "select count(*) as count from transaction where transactionid = '" & "TR" & CStr(random) & "'"
    Set rsTransaction = db.OpenRecordset(SQL)
    While rsTransaction.Fields("count") > 0
        SQL = "select count(*) from transaction where transactionid = '" & "TR" & CStr(random) & "'"
        Set rsTransaction = db.OpenRecordset(SQL)
    Wend
    SQL = "select count(item) from itemsclient where customerid = '" & txtCustID.Text & "'"
    Set rsItemsClient = db.OpenRecordset(SQL)
    If rsItemsClient.Fields(0) > 3 Then
        MsgBox "The costumer cant loan using more than three items!"
    Else
        pawnMoney = Val(txtNewValue.Text) * 1.1
        SQL = "insert into Transaction values('" & "TR" & CStr(random) & "', '" & txtBrokerID.Text & "' ,'" & txtCustID.Text & "', '" & txtItem.Text & "','" & pawnMoney & "',30 , '" & Format(dTransact, "mm/dd/yy") & "', '" & cmbCategory.Text & "')"
        db.Execute (SQL)
        SQL = "insert into itemsClient values('" & txtCustID.Text & "', '" & txtItem.Text & "', '" & txtDescription.Text & "')"
        db.Execute (SQL)
        MsgBox "Your transaction has been successfully added."
        lblBranch.Caption = cmbBranch.Text
        lblID.Caption = "TR" + CStr(random)
        lblName.Caption = txtName.Text
        lblItem.Caption = txtItem.Text
        lblMoney.Caption = Format(pawnMoney, "Currency")
        lblDateTrans.Caption = Format(dTransact, "mm/dd/yy")
        lblDays = 30
        lblDue = Format(dTransact + 30, "mm/dd/yy")
        
        cmdTransact.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\PawnstarPawnshop.mdb")
    Set rsClient = db.OpenRecordset("Client")
    Set rsBroker = db.OpenRecordset("Broker")
    Set rsItemsClient = db.OpenRecordset("itemsclient")
    cmdCustSearch.Enabled = False
    cmdTransact.Enabled = False
    While (rsBroker.EOF = False)
        cmbBranch.AddItem rsBroker.Fields("branchaddress")
        rsBroker.MoveNext
    Wend
    cmbCategory.AddItem ("Jewelry")
    cmbCategory.AddItem ("Electronics")
    cmbCategory.AddItem ("Tools")
    cmbCategory.AddItem ("Antique Items")
    cmbCategory.AddItem ("Instruments")
    cmbCategory.AddItem ("Others")
    txtBrokerID.Enabled = False
    txtCustID.Enabled = False
    txtName.Enabled = False
    txtItem.Enabled = False
    cmbCategory.Enabled = False
    txtValue.Enabled = False
    txtNewValue.Enabled = True
    txtDescription.Enabled = False

End Sub

Private Sub Label20_Click()

End Sub

Private Sub txtBrokerID_Change()
txtCustID.Enabled = True
End Sub

Private Sub txtItem_Change()
cmbCategory.Enabled = True
End Sub


Private Sub txtValue_Change()
    pawnMoney = 0
    If cmbCategory.ListIndex = 0 Then
        rate = 0.01
    ElseIf cmbCategory.ListIndex = 1 Then
        rate = 0.02
    ElseIf cmbCategory.ListIndex = 2 Then
        rate = 0.07
    ElseIf cmbCategory.ListIndex = 3 Then
        rate = 0.1
    ElseIf cmbCategory.ListIndex = 4 Then
        rate = 0.15
    ElseIf cmbCategory.ListIndex = 5 Then
        rate = 0.2
    End If
    pawnMoney = Val(txtValue.Text) - (Val(txtValue.Text) * rate)
    txtNewValue.Text = pawnMoney
    cmdTransact.Enabled = True
End Sub
