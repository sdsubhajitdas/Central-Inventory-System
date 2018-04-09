VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   0  'None
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "LOGIN PAGE"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton Command2 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8520
         TabIndex        =   4
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         MaskColor       =   &H00FFFF80&
         TabIndex        =   3
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   6480
         TabIndex        =   2
         Top             =   3720
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   6480
         TabIndex        =   1
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Image Image1 
         Height          =   3345
         Left            =   480
         Picture         =   "LoginForm.frx":0000
         Top             =   1920
         Width           =   3795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   4800
         TabIndex        =   7
         Top             =   3960
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   4800
         TabIndex        =   6
         Top             =   2400
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME  TO  THE  DAILY  NEEDS  SHOP"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Width           =   8655
      End
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
