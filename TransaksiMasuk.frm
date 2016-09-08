VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TransaksiMasuk 
   Caption         =   "Transaksi Masuk"
   ClientHeight    =   5730
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10335
   LinkTopic       =   "Form3"
   ScaleHeight     =   5730
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   107151361
      CurrentDate     =   42621
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
      Left            =   1800
      TabIndex        =   24
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox MarkRusak 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   23
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox PKLGRusak 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   22
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox MarkBaik 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   21
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox PKLGBaik 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   20
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox GoodIssueNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   15
      Top             =   1200
      Width           =   2655
   End
   Begin VB.OptionButton Condition3 
      Caption         =   "Retur Kondisi Rusak"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   2880
      Width           =   2535
   End
   Begin VB.OptionButton Condition2 
      Caption         =   "Retur Kondisi Baik"
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
      Left            =   5040
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.OptionButton Condition1 
      Caption         =   "Retur Kondisi Baru"
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
      Left            =   5040
      TabIndex        =   11
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Unit 
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
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Qty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox MaterialName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox TUG10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
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
      Left            =   8400
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Kondisi Barang"
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
      Left            =   5160
      TabIndex        =   25
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Keterangan"
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
      Left            =   5400
      TabIndex        =   19
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "No PK/LG"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Keterangan"
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
      Left            =   5400
      TabIndex        =   17
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "No PK/LG"
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
      Left            =   5400
      TabIndex        =   16
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "No Good Issue"
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
      Left            =   5400
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Nama Material"
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
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "No Normalisasi"
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
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "No TUG 10"
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
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Unit"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   735
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
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Exit 
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
Attribute VB_Name = "TransaksiMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public username As String

Private Sub selectMaterialName(code As String, parameter As String)
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
    Cmd.CommandText = "SELECT name, unit from material where code = '" & code & "'"

    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records
    If Not Rs.EOF Then
        'Put the value of field 'Name' into string variable 'Name'
        If parameter = 1 Then
            MaterialName.Text = Rs("name")
        Else
            Unit.Text = Rs("unit")
        End If
        
        'Move to the next record in your resultset
        Rs.MoveNext
    End If

    'Close your database connection
    DBCon.Close

    'Delete all references
    Set Rs = Nothing
    Set Cmd = Nothing
    Set DBCon = Nothing
End Sub

Private Sub About_Click()
Aboutform.Show
End Sub

Private Sub Combo1_Change()
    Call selectMaterialName(Combo1.Text, 1)
    Call selectMaterialName(Combo1.Text, 0)
    Call clearCombo
    Call getMaterialCode(Combo1.Text)
End Sub
Sub clearCombo()
    Dim a As Integer
    a = Combo1.ListCount
    For i = 1 To a
        Combo1.RemoveItem 0
    Next
End Sub
Private Sub Combo1_Click()
    Call selectMaterialName(Combo1.Text, 1)
    Call selectMaterialName(Combo1.Text, 0)
End Sub

Private Sub Command1_Click()
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
    

    'This is your actual MySQL query
    If Condition1.Value = True Then
        condition = 1
        Cmd.CommandText = "INSERT INTO transaction VALUES (NULL, '" & username & "', '" & Combo1.Text & "', '" & DTPicker2.year & "-" & DTPicker2.month & "-" & DTPicker2.Day & "'," & Qty & ",'" & TUG10.Text & "', NULL , '" & condition & "','" & GoodIssueNo.Text & "', NULL, NULL)"
    End If
    If Condition2.Value = True Then
        condition = 2
        Cmd.CommandText = "INSERT INTO transaction VALUES (NULL, '" & username & "', '" & Combo1.Text & "', '" & DTPicker2.year & "-" & DTPicker2.month & "-" & DTPicker2.Day & "'," & Qty & ",'" & TUG10.Text & "', NULL , '" & condition & "', NULL, '" & PKLGBaik.Text & "', '" & MarkBaik.Text & "')"
    End If
    If Condition3.Value = True Then
        condition = 3
        Cmd.CommandText = "INSERT INTO transaction VALUES (NULL, '" & username & "', '" & Combo1.Text & "', '" & DTPicker2.year & "-" & DTPicker2.month & "-" & DTPicker2.Day & "'," & Qty & ",'" & TUG10.Text & "', NULL , '" & condition & "', NULL, '" & PKLGRusak.Text & "', '" & MarkRusak.Text & "')"
    End If
'    On Error GoTo lalala
    'Executes the query-command and puts the result into Rs (recordset)
    Cmd.Execute
    'Close your database connection
    DBCon.Close

    'Delete all references
    Set Rs = Nothing
    Set Cmd = Nothing
    Set DBCon = Nothing
    
'lalala:
'    If Len(Errors) > 0 Then
'        MsgBox Error
'    End If

    MsgBox "data saved", vbInformation, "Success"
End Sub


Private Sub Condition1_Click()
    If Condition1.Value = True Then
        GoodIssueNo.Enabled = True
        GoodIssueNo.BackColor = &H80000005
        
        PKLGBaik.Enabled = False
        PKLGBaik.BackColor = &H8000000F
        
        PKLGRusak.Enabled = False
        PKLGRusak.BackColor = &H8000000F
        
        MarkBaik.Enabled = False
        MarkBaik.BackColor = &H8000000F
        
        MarkRusak.Enabled = False
        MarkRusak.BackColor = &H8000000F
        
    End If
End Sub

Private Sub Condition2_Click()
    If Condition2.Value = True Then
        PKLGBaik.Enabled = True
        MarkBaik.Enabled = True
        PKLGBaik.BackColor = &H80000005
        MarkBaik.BackColor = &H80000005
    
        GoodIssueNo.Enabled = False
        PKLGRusak.Enabled = False
        MarkRusak.Enabled = False
        
        GoodIssueNo.Enabled = False
        GoodIssueNo.BackColor = &H8000000F
                
        PKLGRusak.Enabled = False
        PKLGRusak.BackColor = &H8000000F
                                
        MarkRusak.Enabled = False
        MarkRusak.BackColor = &H8000000F
    End If
End Sub

Private Sub Condition3_Click()
 If Condition3.Value = True Then
        PKLGRusak.Enabled = True
        MarkRusak.Enabled = True
        PKLGRusak.BackColor = &H80000005
        MarkRusak.BackColor = &H80000005
    
        GoodIssueNo.Enabled = False
        PKLGBaik.Enabled = False
        PKLGBaik.BackColor = &H8000000F
        MarkBaik.Enabled = False
        MarkBaik.BackColor = &H8000000F
        
        GoodIssueNo.Enabled = False
        GoodIssueNo.BackColor = &H8000000F
                
    End If
End Sub

Private Sub Exit_Click()
Unload Me
login.Show
End Sub

Private Sub Form_Load()
With Me
   .Top = (Screen.Height - .Height) / 2
   .Left = (Screen.Width - .Width) / 2
End With

GoodIssueNo.Enabled = False
PKLGBaik.Enabled = False
PKLGRusak.Enabled = False
MarkBaik.Enabled = False
MarkRusak.Enabled = False

GoodIssueNo.Enabled = False
GoodIssueNo.BackColor = &H8000000F
        
PKLGBaik.Enabled = False
PKLGBaik.BackColor = &H8000000F
        
PKLGRusak.Enabled = False
PKLGRusak.BackColor = &H8000000F
        
MarkBaik.Enabled = False
MarkBaik.BackColor = &H8000000F
        
MarkRusak.Enabled = False
MarkRusak.BackColor = &H8000000F

MaterialName.Enabled = False
MaterialName.BackColor = &H8000000F

Unit.Enabled = False
Unit.BackColor = &H8000000F

Call getMaterialCode(Combo1.Text)

End Sub
Private Sub getMaterialCode(code As String)
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
    Cmd.CommandText = "SELECT code from material where code like '%" & code & "%'"

    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    'Loop through the results of your recordset until there are no more records
    Do While Not Rs.EOF
        'Put the value of field 'Name' into string variable 'Name'
        Combo1.AddItem Rs("code")
        'Move to the next record in your resultset
        Rs.MoveNext
    Loop

    'Close your database connection
    DBCon.Close

    'Delete all references
    Set Rs = Nothing
    Set Cmd = Nothing
    Set DBCon = Nothing
End Sub

Private Sub Main_Click()
Unload Me
MainMenu.Show
End Sub

Private Sub Qty_Change()
 textval = Qty.Text
  If IsNumeric(textval) Then
    Qty.Text = textval
  Else
    Qty.Text = CStr(numval)
  End If
End Sub
