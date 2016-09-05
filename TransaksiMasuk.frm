VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TransaksiMasuk 
   Caption         =   "Form3"
   ClientHeight    =   7710
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8940
   LinkTopic       =   "Form3"
   ScaleHeight     =   7710
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7080
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox MarkRusak 
      Height          =   285
      Left            =   3720
      TabIndex        =   25
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox PKLGRusak 
      Height          =   285
      Left            =   3720
      TabIndex        =   24
      Top             =   5880
      Width           =   2655
   End
   Begin VB.TextBox MarkBaik 
      Height          =   285
      Left            =   3720
      TabIndex        =   23
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox PKLGBaik 
      Height          =   285
      Left            =   3720
      TabIndex        =   22
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox GoodIssueNo 
      Height          =   285
      Left            =   3720
      TabIndex        =   17
      Top             =   3960
      Width           =   2655
   End
   Begin VB.OptionButton Condition3 
      Caption         =   "Retur Kondisi Rusak"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   5640
      Width           =   2535
   End
   Begin VB.OptionButton Condition2 
      Caption         =   "Retur Kondisi Baik"
      Height          =   195
      Left            =   1800
      TabIndex        =   14
      Top             =   4560
      Width           =   1935
   End
   Begin VB.OptionButton Condition1 
      Caption         =   "Retur Kondisi Baru"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   3480
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   104464385
      CurrentDate     =   42618
   End
   Begin VB.TextBox Unit 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Qty 
      Height          =   405
      Left            =   1800
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox MaterialName 
      Height          =   405
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Normalisasi 
      Height          =   405
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox TUG10 
      Height          =   405
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Keterangan"
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "No PK/LG"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Keterangan"
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "No PK/LG"
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "No Good Issue"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Nama Material"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "No Normalisasi"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "No TUG 10"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Unit"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   975
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
Private Sub Condition1_Click()
    'Text1.Text = Condition1.Value
    If Condition1.Value = True Then
        GoodIssueNo.Enabled = True
        GoodIssueNo.BackColor = &H80000005
    Else
        GoodIssueNo.Enabled = False
        GoodIssueNo.BackColor = &H808080
    End If
End Sub

Private Sub Exit_Click()
Unload Me
Login.Show
End Sub

Private Sub Form_Load()
 'Set Cmd1 = Controls.Add("vb.commandbutton", "Cmd1")
 'Cmd1.Caption = "Dynamic Button"
 'Cmd1.Visible = True
 
 'Set txt = Me.Controls.Add("VB.Textbox", "Text11")
  '  txt.Move 360, 240, 2000, 360
  '  txt.Text = "Hello World!"
   ' txt.Visible = True
GoodIssueNo.Enabled = False
PKLGBaik.Enabled = False
PKLGRusak.Enabled = False
MarkBaik.Enabled = False
MarkRusak.Enabled = False

GoodIssueNo.BackColor = &H808080

End Sub

Private Sub Main_Click()
Unload Me
MainMenu.Show
End Sub

