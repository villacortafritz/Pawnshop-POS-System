VERSION 5.00
Begin VB.Form frmBuyer 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Buyer Data"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   LinkTopic       =   "Form3"
   ScaleHeight     =   9000
   ScaleWidth      =   7650
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2640
      TabIndex        =   22
      Top             =   8160
      Width           =   2175
   End
   Begin VB.TextBox txtID 
      Height          =   615
      Left            =   2520
      TabIndex        =   13
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmdAdd 
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
      Index           =   1
      Left            =   6000
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtContact 
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   6240
      Width           =   5055
   End
   Begin VB.TextBox txtMTB 
      Height          =   615
      Left            =   3720
      TabIndex        =   10
      Top             =   7320
      Width           =   3375
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtProvince 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2520
      TabIndex        =   8
      Top             =   4800
      Width           =   5055
   End
   Begin VB.TextBox txtCity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2520
      TabIndex        =   7
      Top             =   4080
      Width           =   5055
   End
   Begin VB.TextBox txtZIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2520
      TabIndex        =   6
      Top             =   5520
      Width           =   5055
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   3000
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      Picture         =   "frmBidder.frx":0000
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   7560
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label lblMTB 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MONEY TO BID:"
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
      Left            =   1080
      TabIndex        =   21
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Contact No.:"
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
      TabIndex        =   20
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Province:"
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
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   18
      Top             =   6240
      Width           =   60
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7440
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ZIP:"
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
      TabIndex        =   17
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label4 
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
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "City:"
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
      TabIndex        =   15
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblID 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Customer ID:"
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
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
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
      TabIndex        =   3
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
      TabIndex        =   4
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7440
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmBuyer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsBuyer As Recordset

Private Sub cmdAdd_Click(Index As Integer)
txtName.Enabled = True
    txtName.SetFocus
End Sub


Private Sub cmdClear_Click()
    txtID.Text = " "
    txtName.Text = " "
    txtContact.Text = " "
    txtCity.Text = " "
    txtProvince.Text = " "
    txtZIP.Text = " "
    txtMTB.Text = " "
End Sub

Private Sub cmdClose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    txtID.Text = " "
    txtName.Text = " "
    txtContact.Text = " "
    txtCity.Text = " "
    txtProvince.Text = " "
    txtZIP.Text = " "
    txtMTB.Text = " "
    txtID.Enabled = True
    txtID.SetFocus
End Sub

Private Sub Form_Load()
      Set db = OpenDatabase(App.Path & "\PawnstarPawnshop.mdb")
    Set rsBuyer = db.OpenRecordset("Buyer")
    txtID.Enabled = False
    txtName.Enabled = False
    txtCity.Enabled = False
    txtProvince.Enabled = False
    txtZIP.Enabled = False
    txtContact.Enabled = False
    txtMTB.Enabled = False
End Sub

Private Sub Label3_Click()

End Sub

Private Sub txtCity_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtProvince.Enabled = True
        txtProvince.SetFocus
    End If
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMTB.Enabled = True
        txtMTB.SetFocus
    End If
End Sub


Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SQL = "select * from buyer where CustomerID = '" & txtID.Text & " ' "
        Set rsBuyer = db.OpenRecordset(SQL)
    
        If rsBuyer.BOF = True Then
            MsgBox "ID number does not exist"
            Call cmdClear_Click
        Else
            txtID.Text = rsBuyer.Fields("customerid")
            txtName.Text = rsBuyer.Fields("Name")
            txtContact.Text = rsBuyer.Fields("ContactNo")
            txtCity.Text = rsBuyer.Fields("city")
            txtProvince.Text = rsBuyer.Fields("province")
            txtZIP.Text = rsBuyer.Fields("zip")
            txtMTB.Text = Format(rsBuyer.Fields("BidMoney"), "Currency")
        End If
    End If
End Sub


Private Sub txtName_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCity.Enabled = True
        txtCity.SetFocus
    End If
End Sub

Private Sub txtMTB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Randomize
        random = Int(10000 * Rnd) + 1
        SQL = "select count(*) as count from buyer where customerid = '" & "BUY" & CStr(random) & "'"
        Set rsBuyer = db.OpenRecordset(SQL)
        While rsBuyer.Fields("count") > 0
            SQL = "select count(*) from buyer where customerid = '" & "BUY" & CStr(random) & "'"
            Set rsBuyer = db.OpenRecordset(SQL)
        Wend
        txtID.Text = "BUY" + CStr(random)
        'txtID.FontSize = 16
        txtID.FontBold = True
        SQL = "insert into buyer values('" & txtID.Text & "' , '" & txtName.Text & "' , '" & txtCity.Text & "' , '" & txtProvince.Text & "', '" & txtZIP.Text & "' , '" & txtContact.Text & "', '" & txtMTB.Text & "')"
        db.Execute (SQL)
    End If
End Sub



Private Sub txtProvince_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtZIP.Enabled = True
        txtZIP.SetFocus
    End If
End Sub

Private Sub txtZIP_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtContact.Enabled = True
        txtContact.SetFocus
    End If
End Sub

