VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ViewInventoryForm 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Inventory"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox searchIdText 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   7800
      TabIndex        =   9
      Text            =   "Search by Product Id"
      ToolTipText     =   "Enter for search"
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton searchIdButton 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton resetButton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton searchNameButton 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox searchNameText 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   7800
      TabIndex        =   5
      Text            =   "Search by Name"
      ToolTipText     =   "Enter for search"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton delButton 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Delete Data"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Bindings        =   "ViewInventoryForm.frx":0000
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   7858
      _Version        =   393216
      AllowArrows     =   -1  'True
      BackColor       =   16777152
      HeadLines       =   1
      RowHeight       =   28
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Products in Database"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "pid"
         Caption         =   "Product Id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "qty"
         Caption         =   "Quantity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "price"
         Caption         =   "Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "tprice"
         Caption         =   "Total Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5999.812
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2399.811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   495
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"ViewInventoryForm.frx":0014
      OLEDBString     =   $"ViewInventoryForm.frx":00B5
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "product_details_table"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton backButton 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin MSAdodcLib.Adodc searchAdodc 
      Height          =   495
      Left            =   2160
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"ViewInventoryForm.frx":0156
      OLEDBString     =   $"ViewInventoryForm.frx":01F7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "product_details_table"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label updateLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Click and update any data in the table."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10320
      TabIndex        =   4
      Top             =   3120
      Width           =   3645
   End
   Begin VB.Label delLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a product to delete it."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   2700
   End
End
Attribute VB_Name = "ViewInventoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Adodb.Recordset

Private Sub Form_Load()
    Set rs = Adodc.Recordset
End Sub

Private Sub delButton_Click()
    If MsgBox("Condirm Delete Action?", vbQuestion + vbYesNo, "Alert") = vbYes Then
        'If user decides to delete data
        rs.Delete
    End If
    'rs.Delete
End Sub

Private Sub backButton_Click()
    HomeForm.Show
    Unload Me
End Sub

Private Sub resetButton_Click()
    Set DataGrid.DataSource = Adodc
    DataGrid.Refresh
End Sub


Private Sub searchNameButton_Click()
    Set DataGrid.DataSource = searchAdodc
    If Len(searchNameText.Text) = 0 Then
        MsgBox "Enter something to search"
    Else
        searchAdodc.Recordset.Filter = "name Like '*" & searchNameText.Text & "*'"
    End If
    DataGrid.Refresh
End Sub

Private Sub searchNameText_GotFocus()
    searchNameText.Text = ""
End Sub

Private Sub searchIdButton_Click()
    Set DataGrid.DataSource = searchAdodc
    If Len(searchIdText.Text) = 0 Then
        MsgBox "Enter something to search"
    Else
        searchAdodc.Recordset.Filter = "pid Like *" & searchIdText.Text & "*"
    End If
    DataGrid.Refresh
End Sub

Private Sub searchIdText_GotFocus()
    searchIdText.Text = ""
End Sub

'Private Sub searchText_LostFocus()
'    If Len(searchText.Text) = 0 Then
'        searchText.Text = "Search by Name"
'    End If
'End Sub
