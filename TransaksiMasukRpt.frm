VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form TransaksiMasukRpt 
   Caption         =   "Laporan Transaksi Masuk"
   ClientHeight    =   7110
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   17565
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   17565
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   13680
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   6
      Text            =   "Month"
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Text            =   "Year"
      Top             =   360
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Descending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ascending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   16455
      _ExtentX        =   29025
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
      Height          =   135
      Left            =   3120
      TabIndex        =   7
      Top             =   6360
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Sorting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   0
      Width           =   855
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
Dim Cmd2 As ADODB.Command
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim year As String
Dim month As String
Dim selectedRecred As String
Dim sort As String

Private Sub getAllData()
Cmd.CommandText = _
    "SELECT t.code as normalisasi, m.name, date as tanggal, qty as jumlah, t.condition as kondisi, pklgno, Remark " _
    & "From material m, transaction t " _
    & "Where t.condition <> 4 And t.Code = m.Code"
    
    
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records

    Set DataGrid1.DataSource = Rs
    
End Sub

Private Sub getSelectedData()
Cmd.CommandText = _
    "SELECT t.code as normalisasi, m.name, date as tanggal, qty as jumlah, t.condition as kondisi, pklgno, Remark " _
    & "From material m, transaction t " _
    & "Where t.condition <> 4 And t.Code = m.Code " _
    & "and extract(year from date) like '%" & year & "%' " _
    & "and extract(month from date) like '%" & month & "%' " _
    
    'Label1.Caption = Cmd.CommandText
    'Label1.BackColor = vbRed
    
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records

    Set DataGrid1.DataSource = Rs
    
End Sub

Private Sub Combo1_Click()
    Call getMonth(Combo1.Text)
    year = Combo1.Text
    month = ""
    getSelectedData
End Sub

Private Sub Combo2_Click()
    month = Combo2.Text
    getSelectedData
End Sub

Private Sub Command1_Click()
    Cmd.CommandText = _
    "SELECT t.code as normalisasi, m.name, date as tanggal, qty as jumlah, t.condition as kondisi, pklgno, Remark " _
    & "From material m, transaction t " _
    & "Where t.condition <> 4 And t.Code = m.Code " _
    & "and extract(year from date) like '%" & year & "%' " _
    & "and extract(month from date) like '%" & month & "%' " _
    & selectedRecred & " " _
    & sort
    
    'Label1.Caption = Cmd.CommandText
    'Label1.BackColor = vbRed
    
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records
    Set DataGrid1.DataSource = Rs
    
    Cmd2.CommandText = "drop view if exists print_a"
    Cmd2.Execute
    Cmd.CommandText = "create view print_a as " + Cmd.CommandText
    Cmd.Execute
    
    CrystalReport1.ReportFileName = App.Path + "\rpt\tMasuk.rpt"
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If Option1.Value = True Then
    Rs.sort = DataGrid1.Columns(ColIndex).Caption + " Asc"
End If
If Option2.Value = True Then
    Rs.sort = DataGrid1.Columns(ColIndex).Caption + " Desc"
End If
selectedRecred = DataGrid1.Columns(ColIndex).Caption
If selectedRecred = "name" Then
    selectedRecred = "order by m." + selectedRecred
Else
    selectedRecred = "order by " + selectedRecred
End If
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
    
    Call getAllData
    Call getYear
    Call getMonth(Combo1.Text)
    
        With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub
Private Sub getYear()
    Cmd.CommandText = _
    "select distinct(extract(year from date)) as year from transaction"
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs2 = Cmd.Execute
    'Loop through the results of your recordset until there are no more records
        Do While Not Rs2.EOF
        'Put the value of field 'Name' into string variable 'Name'
        Combo1.AddItem Rs2("year")
        'Move to the next record in your resultset
        Rs2.MoveNext
    Loop
End Sub

Private Sub getMonth(year As String)
    Cmd.CommandText = _
    "select distinct(extract(month from date)) as month from transaction where extract(year from date) ='" & year & "'"
    'Executes the query-command and puts the result into Rs (recordset)
    Combo2.Clear
    Combo2.Text = "Month"
    Set Rs2 = Cmd.Execute
    'Loop through the results of your recordset until there are no more records
        Do While Not Rs2.EOF
        'Put the value of field 'Name' into string variable 'Name'
        Combo2.AddItem Rs2("month")
        'Move to the next record in your resultset
        Rs2.MoveNext
    Loop
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
    MainMenu.Show
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Or Option2.Value = True Then
        sort = "asc"
    Else
        sort = ""
    End If
End Sub

Private Sub Option2_Click()
    If Option1.Value = True Or Option2.Value = True Then
        sort = "desc"
    Else
        sort = ""
    End If
End Sub

Private Sub SignOut_Click()
    Unload Me
    login.Show
End Sub
