VERSION 5.00
Begin VB.Form Old 
   Caption         =   "Form2"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   10230
   Begin VB.Frame Frame2 
      Caption         =   "Barang Masuk"
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   9975
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   855
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaksi"
      ClipControls    =   0   'False
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9975
      Begin VB.CommandButton Command3 
         Caption         =   "Create PDF"
         Height          =   735
         Left            =   3480
         TabIndex        =   12
         Top             =   4680
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hi"
         Height          =   735
         Left            =   1200
         TabIndex        =   3
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah"
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "No Normalisasi"
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.PictureBox TabStrip1 
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8835
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 MsgBox "Hai Yangg ;)", vbInformation, "Haii"
End Sub

Private Sub Command3_Click()
Dim dblElapsed As Double
  
  Dim clPDF As New Class1
  Dim strFile As String
  Dim i As Single
  
  ' Imposta il file di output
  strFile = App.Path & "\Demo.pdf"
  
  With clPDF
    .Title = "Demo clsPDFCreator"       ' Titolo
    .ScaleMode = pdfCentimeter          ' Unità di misura
    .PaperSize = pdfA4                  ' Formato pagina
    .Margin = 0                         ' Margine
    .Orientation = pdfPortrait          ' Orientamento
    
   ' .EncodeASCII85 = (chkASCII85.Value = Checked)
    
    .InitPDFFile strFile                ' inizializza il file
    
    ' Definisce le risorse relative ai font
    .LoadFont "Fnt1", "Times New Roman"                       ' Tipo TrueType
    .LoadFont "Fnt2", "Arial", pdfItalic                      ' Tipo TrueType
    .LoadFont "Fnt3", "Courier New"                           ' Tipo TrueType
    .LoadFontStandard "Fnt4", "Courier New", pdfBoldItalic    ' Tipo Type1
    
    ' Definisce le risorse relative alle immagini
    ' .LoadImgFromBMPFile "Img1", App.Path & "\img\20x20x24.bmp" ', pdfGrayScale
    ' .LoadImgFromBMPFile "Img2", App.Path & "\img\200x200x24.bmp" ', pdfGrayScale

    ' watermark
    .StartObject "Item1", pdfAllPages ' , pdfEvenPages
      .SetColorFill -240
      .SetTextHorizontalScaling 120
      .DrawText 4, 4, "Test Watermark", "Fnt2", 80, , 90
      .SetColorFill 0
    .EndObject

'     Inizializza la prima pagina
    .BeginPage


      '.DrawText 19, 1.5, "Test Footer", "Fnt1", 12, pdfAlignRight
      .DrawObject "Footers"
      .DrawText 10.5, 27, "ini Judul", "Fnt1", 48, pdfCenter

      .SetTextHorizontalScaling 100
      .DrawText 1, 25, "left", "Fnt2", 24, pdfAlignLeft
      .DrawText 4, 25, "left 2", "Fnt2", 24, pdfAlignLeft
      'left cmfrombot text font size alignment rotation
      '.SetTextHorizontalScaling 100

      For i = 0 To 10 Step 1
        .DrawText 0.1, 13 + i, "creating new line with logic" & CStr(i), "Fnt2", 12
      Next

    ' Chiude la prima pagina
    .EndPage


    ' Definisce una risorsa da stampare su tutte le pagine
    .StartObject "Footers", pdfAllPages
      .DrawText 20, 1.5, "footer test", "Fnt1", 12, pdfAlignRight
    .EndObject
    
    
    ' Chiude il documento
    .ClosePDFFile
    
  End With
    
  Call Shell("rundll32.exe url.dll,FileProtocolHandler " & (strFile), vbMaximizedFocus)
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs(1).Selected = True Then
    Frame1.Visible = True
    Frame2.Visible = False
Else
    Frame1.Visible = False
    Frame2.Visible = True
End If

End Sub
