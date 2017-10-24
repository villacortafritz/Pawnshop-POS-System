VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "LOG-OUT AND EXIT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   14
      Top             =   7680
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2520
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "SELL UNRETRIEVED ITEMS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE TRANSACTION"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "PAWN AN ITEM"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton cmdBuyer 
      Caption         =   "BUYER"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "CLIENT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdBroker 
      Caption         =   "BROKER"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Manage"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      TabIndex        =   6
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4320
      TabIndex        =   7
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Manage Transactions"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   3855
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*MUST BE AN ADMIN"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   3240
         Width           =   1935
      End
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8040
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PAWNSHOP"
      BeginProperty Font 
         Name            =   "Blippo Light SF"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   12
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PAWNSTARS"
      BeginProperty Font 
         Name            =   "Blippo Light SF"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   11
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "bromello"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   9
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub cmdAdd_Click()
frmNew.Show
End Sub

Private Sub cmdBroker_Click()
frmBroker.Show
End Sub

Private Sub cmdBuyer_Click()
frmBuyer.Show
End Sub

Private Sub cmdClient_Click()
frmClient.Show
End Sub

Private Sub cmdSell_Click()
frmSell.Show
End Sub

Private Sub cmdUpdate_Click()
frmUpdate.Show
End Sub

