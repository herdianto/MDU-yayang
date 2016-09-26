VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RekapStok 
   Caption         =   "Laporan Persediaan"
   ClientHeight    =   6285
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "-Choose-"
      Top             =   480
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
   Begin VB.Label Label1 
      Caption         =   "Kondisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu signOut 
         Caption         =   "Sign Out"
      End
      Begin VB.Menu mMenu 
         Caption         =   "Main Menu"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "RekapStok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBCon As ADODB.Connection
Dim Cmd As ADODB.Command
Dim Cmd2 As ADODB.Command
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Private Sub About_Click()
    Aboutform.Show
End Sub

Private Sub Combo1_Click()
    Call getAllData(Combo1.Text)
End Sub

Private Sub Command1_Click()
    CrystalReport1.ReportFileName = App.Path + "\rpt\tStok.rpt"
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
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
    
    Set Cmd2 = New ADODB.Command
    Cmd2.ActiveConnection = DBCon
    Cmd2.CommandType = adCmdText
    
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    Call getAllData("")
    
    'Combo1.AddItem ("1")
    'Combo1.AddItem ("2")
    'Combo1.AddItem ("3")
    Combo1.AddItem ("Baru")
    Combo1.AddItem ("Baik")
    Combo1.AddItem ("Rusak")
End Sub
Private Sub getAllData(kon As String)
    If kon = "Baru" Then
        kon = "1"
    ElseIf kon = "Baik" Then
        kon = "2"
    ElseIf kon = "Rusak" Then
        kon = "3"
    End If
    Cmd.CommandText = _
    "select t.code, m.name, sum(t.qty) as Qty FROM transaction t, material m " _
    & "WHERE t.Condition like '%" & kon & "%' and t.Code = m.Code " _
    & "GROUP BY t.code"
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records
    
    Set DataGrid1.DataSource = Rs
    
    Cmd2.CommandText = "drop view if exists print_b"
    Cmd2.Execute
    Cmd2.CommandText = "create view print_b as " & Cmd.CommandText
    Cmd2.Execute
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Delete all references
    Set Rs = Nothing
    Set Cmd = Nothing
    Set DBCon = Nothing
    'Close your database connection
    'DBCon.Close
End Sub

Private Sub mMenu_Click()
    Unload Me
    MainMenu.Show
End Sub

Private Sub SignOut_Click()
    Unload Me
    login.Show
End Sub
