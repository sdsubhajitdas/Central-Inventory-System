VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AddProductForm 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Product"
   ClientHeight    =   6780
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   13
      Top             =   0
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Bindings        =   "AddProductForm.frx":0000
      Height          =   5535
      Left            =   6840
      TabIndex        =   12
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
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
         Size            =   8.25
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1454.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   5520
      Top             =   5400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   $"AddProductForm.frx":0014
      OLEDBString     =   $"AddProductForm.frx":00B5
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
   Begin VB.TextBox pPrice 
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
      Left            =   1680
      TabIndex        =   7
      Top             =   3960
      Width           =   5055
   End
   Begin VB.TextBox pQty 
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
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   5055
   End
   Begin VB.TextBox pName 
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
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   5055
   End
   Begin VB.TextBox pId 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton addProduct 
      Caption         =   "Add New Product To Database"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   6615
   End
   Begin VB.Label pTotalPriceHolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs. "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Label pTotalPriceLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label pPriceLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label pQuantityLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label pNameLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label pIdLable 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Product Id"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label topLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD NEW PRODUCT DETAILS BELOW"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14055
   End
End
Attribute VB_Name = "AddProductForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim singleDecimal As Boolean    'To keep track of single decimal point in price.



Private Sub backButton_Click()
    'Navigation purpose returning back to home screen
    HomeForm.Show
    Unload Me
End Sub

Private Sub Form_Load()
    'Initializing some database parameters during form load.
    Set rs = Adodc.Recordset
    singleDecimal = False
End Sub

Private Sub addProduct_Click()
    'Checking of data.
    If Len(pId.Text) <> 0 Then              'Checking if there is no id
        If Len(pName.Text) <> 0 Then        'Checking if there is no name
            If Len(pPrice.Text) <> 0 Then   'Checking if there is no price
                
                If Len(pQty.Text) = 0 Then 'If no qty is given then the default is 1.
                    pQty.Text = "1"
                End If
                
                'Adding the data to the database
                rs.MoveLast                     'Moving the DB cursor to last row for adding data.
                rs.AddNew                       'Now adding a new data with the below fields.
                rs.Fields("pid") = pId.Text
                rs.Fields("name") = pName.Text
                rs.Fields("qty") = pQty.Text
                rs.Fields("price") = pPrice.Text
                rs.Fields("tprice") = Val(pPrice.Text) * Val(pQty.Text)
                
                'Data is updated along with the grid.
                DataGrid.Refresh
                
                'All fields are cleared.
                pId.Text = ""
                pName.Text = ""
                pQty.Text = ""
                pPrice.Text = ""
                pTotalPriceHolder.Caption = "Rs. "
                
            Else
                MsgBox "Product Price must be provided"
            End If
        Else
            MsgBox "Product Name must be provided"
        End If
    Else
        MsgBox "Product Id must be provided"
    End If
End Sub



Private Sub pPrice_Change()
    'Updating the total price according to the qty.
    Dim totalPrice As Double
    Dim qty As Long
    If Len(pQty.Text) = 0 Then      'Setting default qty to 1
        qty = 1
    Else
        qty = Val(pQty.Text)
    End If
    
    'Total price of the product is shown
    totalPrice = qty * Val(pPrice.Text)
    pTotalPriceHolder.Caption = "Rs. " & totalPrice
    
End Sub

Private Sub pPrice_KeyPress(KeyAscii As Integer)
    'Filtering the price
    If KeyAscii <> 8 And KeyAscii <> 46 Then    'Ommiting "backspace" . and
        'Length problems unclear code dont remove cause it works
        If Len(pPrice.Text) <> 1 Or Len(pPrice.Text) <> 0 Then
            If KeyAscii < 48 Or KeyAscii > 57 Then  'Price is between 0-9
                MsgBox "Price should be a digit"
                KeyAscii = 0
            End If
        End If
    End If
    
    If KeyAscii = 46 And singleDecimal = True Then  'Filtering 2nd "."
        KeyAscii = 0
    End If
    
    If KeyAscii = 46 Then       'Counter change on 1st . encounter
        singleDecimal = True
    End If
        
End Sub

Private Sub pQty_Change()
    'Update price on qty change
    Call pPrice_Change
End Sub

Private Sub pQty_KeyPress(KeyAscii As Integer)
    'Filtering the price
    If KeyAscii <> 8 Then       'Ommiting "backspace"
        'Length problems unclear code dont remove cause it works
        If Len(pPrice.Text) <> 1 Or Len(pPrice.Text) <> 0 Then
            If KeyAscii < 48 Or KeyAscii > 57 Then  'Qty is between 0-9
                MsgBox "Price should be a digit"
                KeyAscii = 0
            End If
        End If
    End If
End Sub

