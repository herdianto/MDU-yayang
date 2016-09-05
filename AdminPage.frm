VERSION 5.00
Begin VB.Form AdminPage 
   Caption         =   "Form3"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form3"
   ScaleHeight     =   4320
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "AdminPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox ":P", vbCritical
End Sub
