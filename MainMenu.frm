VERSION 5.00
Begin VB.Form MainMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Selamat Datang di Aplikasi MDU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   6735
   End
   Begin VB.Image Image4 
      Height          =   1155
      Left            =   4680
      Picture         =   "MainMenu.frx":0000
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   6600
      Picture         =   "MainMenu.frx":0E91
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   2880
      Picture         =   "MainMenu.frx":1D22
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   960
      Picture         =   "MainMenu.frx":2BB3
      Top             =   1680
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

Private Sub Image4_Click()
TransaksiMasukRpt.Show
Unload Me
End Sub
