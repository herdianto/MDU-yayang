VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form TransaksiMasukRpt 
   Caption         =   "Laporan Transaksi Masuk"
   ClientHeight    =   7110
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   5530
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
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
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu SignOut 
         Caption         =   "Sign Out"
      End
      Begin VB.Menu mainMen 
         Caption         =   "MainMenu"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "TransaksiMasukRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBCon As ADODB.Connection
Dim Cmd As ADODB.Command
Dim Rs As ADODB.recordSet

Private Sub Command1_Click()
Cmd.CommandText = _
    "SELECT t.code as normalisasi, m.name, date as tanggal, qty as jumlah, t.condition as kondisi, pklgno, Remark " _
    & "From material m, transaction t " _
    & "Where t.condition <> 4 And t.Code = m.Code"
    
    
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records

    Set DataGrid1.DataSource = Rs
    
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Rs.Sort = DataGrid1.Columns(ColIndex).Caption + " Asc"
End Sub

Private Sub Form_Load()
    'Create a connection to the database
    Set DBCon = New ADODB.Connection
    DBCon.CursorLocation = adUseClient
    'This is a connectionstring to a local MySQL server
    DBCon.Open "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=MDU;User=root;Password=;Option=3;"

    'Create a new command that will execute the query
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = DBCon
    Cmd.CommandType = adCmdText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close your database connection
    DBCon.Close

    'Delete all references
    Set Rs = Nothing
    Set Cmd = Nothing
    Set DBCon = Nothing
    
End Sub

Private Sub mainMen_Click()
    Unload Me
    mainMenu.Show
End Sub

Private Sub SignOut_Click()
    Unload Me
    login.Show
End Sub
