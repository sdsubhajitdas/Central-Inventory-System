VERSION 5.00
Begin VB.Form LoginForm 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login to Central Inventory System"
   ClientHeight    =   5415
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3199.358
   ScaleMode       =   0  'User
   ScaleWidth      =   10802.57
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1380
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2760
      Width           =   3525
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7800
      TabIndex        =   1
      Top             =   1680
      Width           =   3525
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FF80&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   240
      Picture         =   "LoginForm.frx":0000
      Top             =   480
      Width           =   4500
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim user As String
Dim pass As String

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure to close this Application?", vbQuestion + vbYesNo, "System") = vbYes Then
    'If user decides not to close
        LoginSucceeded = True
        Unload Me
    Else
        LoginSucceeded = False
        Me.Show
    End If
    'If user decides to close
    
End Sub

Private Sub cmdOK_Click()
    user = "admin"
    pass = "12345"

    If txtUserName.Text = user Then
        If txtPassword.Text = pass Then

                'MsgBox "Username and Password Accepted!", vbInformation, "Login"
                HomeForm.Show
                Unload Me
                
        ElseIf txtPassword.Text = "" Then
            MsgBox "Password Field Empty!", vbExclamation, "Login"
        Else
     
            MsgBox "Username and Password not Matched!", vbExclamation, "Login"
        End If
    ElseIf txtUserName.Text = "" Then
        MsgBox "Username Field Empty!", vbExclamation, "Login"
    Else
        MsgBox "Invalid Username, try again!", , "Login"
        txtPassword.SetFocus
    End If

End Sub

