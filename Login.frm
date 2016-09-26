VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Menu"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Masukkan NIP"
      Top             =   2010
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Masukkan password"
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   600
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PASSWORD"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "USERNAME"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "RETURN MANAGEMENT SYSTEM"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "PT PLN (Persero)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call login(Text1.Text, Text2.Text)
End Sub

Private Sub login(username As String, password As String)
Dim DBCon As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim accessRight As String

    accessRight = -1
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
    Cmd.CommandText = "SELECT accessRight from user where userID = '" & username & "' and password = '" & password & "'"

    'Executes the query-command and puts the result into Rs (recordset)
    Set Rs = Cmd.Execute
    If Not Rs.EOF Then
        accessRight = Rs("accessRight")
    End If
    
    If accessRight = 1 Then
        MainMenu.Show
        TransaksiMasuk.username = Text1.Text
        TransaksiKeluar.username = Text1.Text
        Unload Me
    ElseIf accessRight = 0 Then
        AdminPage.Show
        Unload Me
    Else
        MsgBox "Invalid username / password", vbOKOnly, "Invalid credential"
        
    End If
    
End Sub

Private Sub Form_Load()
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
        .Visible = True
    End With
    Text1.SetFocus
End Sub
