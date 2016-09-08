VERSION 5.00
Begin VB.Form MainMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main Menu"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainMenu.frx":0000
   ScaleHeight     =   2000
   ScaleMode       =   0  'User
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selamat Datang di Aplikasi MDU"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   6735
   End
   Begin VB.Image Image4 
      Height          =   1155
      Left            =   4560
      Picture         =   "MainMenu.frx":1A6BA
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan Pergerakan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan Persediaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   613
      Left            =   6480
      TabIndex        =   2
      Top             =   3120
      Width           =   1230
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   6480
      Picture         =   "MainMenu.frx":1EF2E
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   2760
      Picture         =   "MainMenu.frx":23A0A
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaksi Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   613
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaksi Masuk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   613
      Left            =   840
      TabIndex        =   0
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   840
      Picture         =   "MainMenu.frx":28016
      Top             =   1920
      Width           =   1155
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

Private Sub Image1_Click()
TransaksiMasuk.Show
Unload Me
End Sub
Private Sub Image2_Click()
TransaksiKeluar.Show
Unload Me
End Sub

Private Sub Image3_Click()
RekapStok.Show
Unload Me
End Sub

Private Sub Image4_Click()
TransaksiMasukRpt.Show
Unload Me
End Sub
