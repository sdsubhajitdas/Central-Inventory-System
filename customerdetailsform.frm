VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form customerdetailsform 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "customer details"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   18180
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "customerdetailsform.frx":0000
      Height          =   4335
      Left            =   10200
      TabIndex        =   12
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "customer details"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "cid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10560
      Top             =   5640
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
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Central-Inventory-System\database\cis-database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Central-Inventory-System\database\cis-database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton next 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   11
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton back 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   10
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton delete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   9
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton add 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   8
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox cmailid 
      Height          =   735
      Left            =   4680
      TabIndex        =   7
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox cnumber 
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox cname 
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox cid 
      Height          =   855
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label mailid 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Email d"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   2520
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label contactnumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   2400
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label customername 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label customerid 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Id"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "customerdetailsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub add_Click()

                
                'Adding the data to the database
               ' rs.MoveLast                     'Moving the DB cursor to last row for adding data.
               ' rs.AddNew                       'Now adding a new data with the below fields.
                'rs.Fields("cid") = cid.Text
               ' rs.Fields("cname") = cname.Text
               ' rs.Fields("cnumber") = cnumber.Text
               ' rs.Fields("cmailid") = cmailid.Tex
On Error GoTo aerr
rs.AddNew

rs(0).Value = cid.Text
rs(1).Value = cname.Text
rs(2).Value = cnumber.Text
rs(3).Value = cmailid.Text
rs.UpdateBatch
cid.Text = ""
cname.Text = ""
cnumber.Text = ""
cmailid.Text = ""
aerr:
Err.Clear
Exit Sub
                
End Sub

Private Sub back_Click()
HomeForm.Show
    Unload Me
End Sub

Private Sub cid_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 46 Then    'Ommiting "backspace" . and
        'Length problems unclear code dont remove cause it works
        If Len(cid.Text) <> 1 Or Len(cid.Text) <> 0 Then
            If KeyAscii < 48 Or KeyAscii > 57 Then  'Price is between 0-9
                MsgBox "Id should be a digit"
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

Private Sub delete_Click()
 Adodc1.Recordset.delete
End Sub

Private Sub Form_Load()
Set con = New customerdetailsform
End Sub
