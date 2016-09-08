VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form RekapStok 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9855
      _ExtentX        =   17383
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
    
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
    getAllData
    DataGrid1.SetFocus
End Sub
Private Sub getAllData()
Cmd.CommandText = _
    "SET @rownr=0;"
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records

    Cmd.CommandText = _
    "select @rownr:=@rownr+1 AS No, s.* from " _
    & "(select t.code, m.name, sum(t.qty) FROM transaction t, material m " _
    & "WHERE t.Condition = '2' and t.Code = m.Code " _
    & "GROUP BY t.code) s " _
    & "CROSS JOIN (SELECT @cnt := 0) AS dummy"
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records
    
    Set DataGrid1.DataSource = Rs
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close your database connection
    DBCon.Close

    'Delete all references
    Set Rs = Nothing
    Set Cmd = Nothing
    Set DBCon = Nothing
End Sub

Private Sub mMenu_Click()
    Unload Me
    MainMenu.Show
End Sub

Private Sub signOut_Click()
    Unload Me
    login.Show
End Sub
