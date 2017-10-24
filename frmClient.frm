VERSION 5.00
Begin VB.Form frmClient 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   Caption         =   "Client Data"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   FillStyle       =   3  'Vertical Line
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   7635
   Begin VB.TextBox txtContactNo 
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
      TabIndex        =   16
      Top             =   6480
      Width           =   3255
   End
   Begin VB.TextBox txtProvince 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2640
      TabIndex        =   15
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox txtCity 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2640
      TabIndex        =   14
      Top             =   4320
      Width           =   4815
   End
   Begin VB.TextBox txtZIP 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2640
      TabIndex        =   13
      Top             =   5760
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
      Height          =   645
      Left            =   2640
      TabIndex        =   10
      Top             =   3600
      Width           =   4815
   End
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
      Left            =   6000
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtCID 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2640
      TabIndex        =   7
      Top             =   2520
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      Picture         =   "frmClient.frx":0000
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   120
      Width           =   1815
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
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   7560
      Width           =   2295
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
      Left            =   6000
      TabIndex        =   0
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   7440
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7440
      Y1              =   3360
      Y2              =   3360
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
      TabIndex        =   21
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label8 
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
      TabIndex        =   20
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   1800
      TabIndex        =   19
      Top             =   6480
      Width           =   60
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
      TabIndex        =   18
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label4 
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
      TabIndex        =   17
      Top             =   4440
      Width           =   975
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
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   11
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID No.:"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   1020
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
      TabIndex        =   5
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
      TabIndex        =   6
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7440
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsClient As Recordset
Dim ClientID As String

Private Sub cmdAdd_Click()
    txtCID.Text = ""
    txtName.Text = ""
    txtCity.Text = ""
    txtProvince.Text = ""
    txtZIP.Text = ""
    txtContactNo.Text = ""
    txtName.Enabled = True
    txtName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
txtCID.Text = ""
    txtName.Text = ""
    'txtLname.Text = ""
    txtCity.Text = ""
    txtProvince.Text = ""
    txtZIP.Text = ""
    txtContactNo.Text = ""
    txtCID.Enabled = True
    txtName.Enabled = False
    'txtLname.Enabled = False
    txtCity.Enabled = False
    txtProvince.Enabled = False
    txtZIP.Enabled = False
    txtContactNo.Enabled = False
    txtCID.SetFocus
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\PawnstarPawnshop.mdb")
    Set rsClient = db.OpenRecordset("Client")
    txtCID.Enabled = False
    txtName.Enabled = False
    txtCity.Enabled = False
    txtProvince.Enabled = False
    txtZIP.Enabled = False
    txtContactNo.Enabled = False
End Sub
 
Private Sub txtCID_Keypress(KeyAscii As Integer)
    
    SQL = "select * from client where customerid = '" & txtCID.Text & "'"
    Set rsClient = db.OpenRecordset(SQL)
    If KeyAscii = 13 Then
        If rsClient.BOF = True Then
            MsgBox "id number does not exist"
        Else
            
            txtName.Text = rsClient.Fields("name")
            'txtLname.Text = rsClient.Fields("name")
            txtCity.Text = rsClient.Fields("city")
            txtProvince.Text = rsClient.Fields("province")
            txtZIP.Text = rsClient.Fields("zip")
            txtContactNo.Text = rsClient.Fields("contactno")
        End If
    End If
End Sub

Private Sub txtCity_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtProvince.Enabled = True
        txtProvince.SetFocus
    End If
End Sub

Private Sub txtContactNo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Randomize
        random = Int(10000 * Rnd) + 1
        SQL = "select count(*) as count from client where customerid = '" & "CL" & CStr(random) & "'"
        Set rsClient = db.OpenRecordset(SQL)
        While rsClient.Fields("count") > 0
            SQL = "select count(*) from client where customerid = '" & "CL" & CStr(random) & "'"
            Set rsTransaction = db.OpenRecordset(SQL)
        Wend
        txtCID.Text = "CL" & CStr(random)
        txtCID.FontBold = True
        SQL = "insert into client (customerid, name, city, province, zip, contactno) values ('" & txtCID.Text & "', '" & txtName.Text & "', '" & txtCity.Text & "', '" & txtProvince.Text & "', '" & txtZIP.Text & "', '" & txtContactNo.Text & "')"
        db.Execute (SQL)
        cmdClose.SetFocus
    End If
End Sub

Private Sub txtName_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCity.Enabled = True
        txtCity.SetFocus
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
        txtContactNo.Enabled = True
        txtContactNo.SetFocus
    End If
End Sub
