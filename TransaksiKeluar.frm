VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TransaksiKeluar 
   Caption         =   "Transaksi Keluar"
   ClientHeight    =   6720
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   106561537
      CurrentDate     =   42618
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   975
      Left            =   6120
      TabIndex        =   10
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Cari Material"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "No TUG 9"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   375
      Left            =   240
      TabIndex        =   1
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

Private Sub getData(Code As String)
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
    & "WHERE transaction.code = material.code and transaction.code like '%" & Code & "%'" _
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
    ' Add data to the Recordset
   rst.AddNew Array("Code", "Name", "Qty"), _
      Array(DataGrid2.Columns(0).Value, DataGrid2.Columns(1).Value, 0)
   ' Populate the Data in the DataGrid
   Set DataGrid1.DataSource = rst
End If

End Sub

Private Sub Form_Load()
' parameter of the Controls.Add to specify the container.
   Set ctlDynamic = Controls.Add("MSComctlLib.TreeCtrl", _
                    "myctl", TransaksiKeluar)
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
