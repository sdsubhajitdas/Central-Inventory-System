VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1560
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   921.699
   ScaleMode       =   0  'User
   ScaleWidth      =   3647.805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
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

