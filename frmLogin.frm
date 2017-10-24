VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login Form"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9885
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShow 
      BackColor       =   &H0080FF80&
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      TabIndex        =   0
      Top             =   6600
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   8115
      ScaleWidth      =   9915
      TabIndex        =   2
      Top             =   0
      Width           =   9975
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "LOG-IN ADMINISTRATOR PASSWORD"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   6240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkShow_Click()
    If chkShow.Value = 1 Then
        txtPassword.PasswordChar = ""
        txtPassword.SetFocus
    Else
        txtPassword.PasswordChar = "*"
    End If
End Sub

Private Sub Form_Load()
    txtPassword.PasswordChar = "*"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPassword.Text = "YouShallPass" Then
            Form1.Show
            txtPassword.Text = ""
            Unload Me
        Else
            MsgBox "You shall not pass!"
        End If
    End If
End Sub
