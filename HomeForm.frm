VERSION 5.00
Begin VB.Form HomeForm 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Central Inventory System"
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
   Begin VB.CommandButton makeBillButton 
      Caption         =   "Make a Bill "
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
      Left            =   7080
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton addcustomer 
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   5520
      Width           =   2295
   End
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
   Begin VB.CommandButton viewInventoryButton 
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
   Begin VB.Image Image1 
      Height          =   5760
      Left            =   1800
      Picture         =   "HomeForm.frx":0000
      Top             =   720
      Width           =   5760
   End
End
Attribute VB_Name = "HomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addcustomer_Click()
    customerdetailsform.Show
    Unload Me
End Sub

Private Sub addProductButton_Click()
    AddProductForm.Show
    Unload Me
End Sub

Private Sub ExitButton_Click()
    Unload Me
End Sub

Private Sub makeBillButton_Click()
    BillForm.Show
    Unload Me
End Sub

Private Sub viewInventoryButton_Click()
    ViewInventoryForm.Show
    Unload Me
End Sub
