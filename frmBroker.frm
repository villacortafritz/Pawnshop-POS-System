VERSION 5.00
Begin VB.Form frmBroker 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Broker Data"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdSearch 
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
      Left            =   600
      TabIndex        =   11
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      Picture         =   "frmBroker.frx":0000
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrokerClose 
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
      Left            =   5160
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdBrokerAdd 
      Caption         =   "Add"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txtBrokerAddress 
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
      Top             =   3240
      Width           =   4935
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
      TabIndex        =   1
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7440
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Address:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
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
      TabIndex        =   8
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
      TabIndex        =   9
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ID:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "frmBroker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsBroker As Recordset
Dim ctr As Integer

Private Sub cmdBrokerAdd_Click()
    txtBrokerAddress.Text = ""
    txtBrokerID.Text = ""
    txtBrokerAddress.Enabled = True
    txtBrokerAddress.SetFocus
End Sub

Private Sub cmdBrokerClose_Click()
    Unload Me
End Sub


Private Sub cmdSearch_Click()
    txtBrokerID.Enabled = True
    txtBrokerID.SetFocus
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\PawnstarPawnshop.mdb")
    Set rsBroker = db.OpenRecordset("Broker")
    txtBrokerID.Enabled = False
    txtBrokerAddress.Enabled = False
End Sub

Private Sub txtBrokerAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SQL = "select * from Broker where BrokerID = '" & txtBrokerID.Text & "'"
        Set rsBroker = db.OpenRecordset(SQL)
        If rsBroker.BOF = False Then
            MsgBox "Already inputted the same ID"
        Else
            If txtBrokerAddress.Text = "" Then
                MsgBox "Invalid Input"
            Else
                ctr = ctr + 1
                txtBrokerID.Text = CStr(ctr)
                    'txtID.FontSize = 16
                txtBrokerID.FontBold = True
                SQL = "insert into Broker values('" & txtBrokerID.Text & "','" & txtBrokerAddress.Text & "')"
                db.Execute (SQL)
                End If
        End If
    End If
End Sub


Private Sub txtBrokerID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SQL = "select * from Broker where BrokerID = '" & txtBrokerID.Text & "'"
        Set rsBroker = db.OpenRecordset(SQL)
        If rsBroker.BOF = True Then
            MsgBox "ID Number does not exist"
        Else
            txtBrokerAddress.Text = rsBroker.Fields("BranchAddress")
        End If
    End If
End Sub
