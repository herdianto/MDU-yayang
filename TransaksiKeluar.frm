VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TransaksiKeluar 
   Caption         =   "Transaksi Keluar"
   ClientHeight    =   7560
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus Baris"
      Height          =   495
      Left            =   9120
      TabIndex        =   11
      Top             =   4440
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   405
      Left            =   2760
      TabIndex        =   10
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   714
      _Version        =   393216
      Format          =   90243073
      CurrentDate     =   42621
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan Transaksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   8
      Top             =   5280
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2778
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      AllowDelete     =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2760
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   2778
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      AllowAddNew     =   -1  'True
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
   Begin VB.Label Label4 
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Cari Material"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "No TUG 9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu SignOut 
         Caption         =   "Sign Out"
      End
      Begin VB.Menu Main 
         Caption         =   "Main Menu"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "TransaksiKeluar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ctlDynamic As VBControlExtender
Attribute ctlDynamic.VB_VarHelpID = -1
Dim WithEvents ctlText As VB.TextBox
Attribute ctlText.VB_VarHelpID = -1
Dim WithEvents ctlCommand As VB.CommandButton
Attribute ctlCommand.VB_VarHelpID = -1
Dim WithEvents CtlLabel As VB.Label
Attribute CtlLabel.VB_VarHelpID = -1
Dim count1 As Integer
Dim rst As ADODB.Recordset
Public username As String

Private Sub getData(code As String)
    Dim DBCon As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim Rs As ADODB.Recordset

    'Create a connection to the database
    Set DBCon = New ADODB.Connection
    DBCon.CursorLocation = adUseClient
    'This is a connectionstring to a local MySQL server
    DBCon.Open "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=MDU;User=root;Password=;Option=3;"

    'Create a new command that will execute the query
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = DBCon
    Cmd.CommandType = adCmdText
    'This is your actual MySQL query
    Cmd.CommandText = _
    "SELECT transaction.Code, material.Name, SUM( qty ) AS Balance " _
    & "FROM transaction, material " _
    & "WHERE transaction.code = material.code and ( transaction.condition = '2' or transaction.condition = '4') and transaction.code like '%" & code & "%'" _
    & "GROUP BY transaction.code"
    
    
    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records

    Set DataGrid2.DataSource = Rs
    
    'Close your database connection
    'DBCon.Close
    'Delete all references
    'Set Rs = Nothing
    'Set Cmd = Nothing
    'Set DBCon = Nothing
End Sub

Private Sub about_Click()
Aboutform.Show
End Sub

Private Sub Command1_Click()
Dim status As Boolean
Dim a As Integer

Do Until a = rst.RecordCount
    DataGrid1.Row = a
    If DataGrid1.Columns(0).Value = DataGrid2.Columns(0).Value Then
        status = True
        Exit Do
    End If
    a = a + 1
Loop

If Not status Then
    If DataGrid2.Columns(2).Value <= 0 Then
    
    Else
        ' Add data to the Recordset
   rst.AddNew Array("Code", "Name", "Qty"), _
      Array(DataGrid2.Columns(0).Value, DataGrid2.Columns(1).Value, 0)
   ' Populate the Data in the DataGrid
   Set DataGrid1.DataSource = rst
    End If
End If

End Sub

Private Sub Command2_Click()
    Call saveData
End Sub

Private Sub saveData()
    Dim DBCon As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim condition As Integer

    'Create a connection to the database
    Set DBCon = New ADODB.Connection
    DBCon.CursorLocation = adUseClient
    'This is a connectionstring to a local MySQL server
    DBCon.Open "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=MDU;User=root;Password=;Option=3;"

    'Create a new command that will execute the query
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = DBCon
    Cmd.CommandType = adCmdText
    
        
        Do Until a = rst.RecordCount
        DataGrid1.Row = a
        
        Cmd.CommandText = "insert into transaction values (NULL, '" & username & "', '" & DataGrid1.Columns(0).Value & "', '" & DTPicker2.year & "-" & DTPicker2.month & "-" & DTPicker2.Day & "', '" & DataGrid1.Columns(2).Value * -1 & "', NULL, '" & Text1.Text & "', '4', NULL, NULL, NULL)"
        
        'Executes the query-command and puts the result into Rs (recordset)
        Cmd.Execute
        a = a + 1
    Loop
        
    
    'Close your database connection
    DBCon.Close

    'Delete all references
    Set Rs = Nothing
    Set Cmd = Nothing
    Set DBCon = Nothing

    MsgBox "data saved", vbInformation, "Success"
    Call getData(Text2.Text)
    
    Set rst = New ADODB.Recordset
rst.CursorLocation = adUseClient
   
' Add columns to the Recordset
rst.Fields.Append "Code", adVarChar, 40, adFldIsNullable
rst.Fields.Append "Name", adVarChar, 40, adFldIsNullable
rst.Fields.Append "Qty", adInteger

' Open the Recordset
rst.Open , , adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = rst
   
End Sub


Private Sub Command3_Click()
rst.Delete
End Sub

Private Sub DataGrid1_AfterUpdate()
    Label4.Caption = Label4.Caption + "a"
End Sub




Private Sub Form_Load()

DataGrid2.Columns.Add 1
DataGrid2.Columns(0).Caption = "Code"
DataGrid2.Columns(1).Caption = "Name"
DataGrid2.Columns(2).Caption = "Balance"



Set rst = New ADODB.Recordset
rst.CursorLocation = adUseClient
   
' Add columns to the Recordset
rst.Fields.Append "Code", adVarChar, 40, adFldIsNullable
rst.Fields.Append "Name", adVarChar, 40, adFldIsNullable
rst.Fields.Append "Qty", adInteger

' Open the Recordset
rst.Open , , adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = rst

    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
End Sub

Private Sub Main_Click()
    Unload Me
    MainMenu.Show
End Sub

Private Sub SignOut_Click()
    Unload Me
    login.Show
End Sub

Private Sub Text2_Change()
    If Text2.Text <> "" Then
        Call getData(Text2.Text)
    Else
        Call getData("#$")
    End If
End Sub
