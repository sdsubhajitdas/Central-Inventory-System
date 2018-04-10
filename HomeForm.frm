VERSION 5.00
Begin VB.Form HomeForm 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Inventory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton addProductButton 
      Caption         =   "Add Product"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "HomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addProductButton_Click()
    AddProductFrom.Show
    Unload Me
End Sub

Private Sub ExitButton_Click()
    Unload Me
End Sub
