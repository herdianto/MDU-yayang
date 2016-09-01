VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3240
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3625
      _Version        =   393216
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
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim DBCon As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim strName As String
    Dim NIP As String

    'Create a connection to the database
    Set DBCon = New ADODB.Connection
    DBCon.CursorLocation = adUseClient
    'This is a connectionstring to a local MySQL server
    DBCon.Open "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=pln;User=root;Password=;Option=3;"

    'Create a new command that will execute the query
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = DBCon
    Cmd.CommandType = adCmdText
    'This is your actual MySQL query
    Cmd.CommandText = "SELECT NAMA, nip from user WHERE nip = 11111"

    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute

    Set DataGrid1.DataSource = Rs
    DataGrid1.Refresh
    'Loop through the results of your recordset until there are no more records
    Do While Not Rs.EOF
        'Put the value of field 'Name' into string variable 'Name'
        strName = Rs("nama")
        NIP = Rs("nip")
        
        Text1.Text = strName
        Text2.Text = NIP
        
        'Move to the next record in your resultset
        Rs.MoveNext
    Loop

    'Close your database connection
  '  DBCon.Close

    'Delete all references
   ' Set Rs = Nothing
  '  Set Cmd = Nothing
  '  Set DBCon = Nothing

End Sub

Private Sub Text3_Change()
    getData (Text3.Text)
End Sub

Sub getData(A As String)
    Dim DBCon As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim strName As String
    Dim NIP As String

    'Create a connection to the database
    Set DBCon = New ADODB.Connection
    DBCon.CursorLocation = adUseClient
    'This is a connectionstring to a local MySQL server
    DBCon.Open "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=pln;User=root;Password=;Option=3;"

    'Create a new command that will execute the query
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = DBCon
    Cmd.CommandType = adCmdText
    'This is your actual MySQL query
    Cmd.CommandText = "SELECT NAMA, nip from user WHERE nip like '%" & A & "%'"

    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute

    Set DataGrid1.DataSource = Rs
    DataGrid1.Refresh
    'Loop through the results of your recordset until there are no more records
    Do While Not Rs.EOF
        'Put the value of field 'Name' into string variable 'Name'
        strName = Rs("nama")
        NIP = Rs("nip")
        
        Text1.Text = strName
        Text2.Text = NIP
        
        'Move to the next record in your resultset
        Rs.MoveNext
    Loop

    'Close your database connection
  '  DBCon.Close

    'Delete all references
   ' Set Rs = Nothing
  '  Set Cmd = Nothing
  '  Set DBCon = Nothing
End Sub
