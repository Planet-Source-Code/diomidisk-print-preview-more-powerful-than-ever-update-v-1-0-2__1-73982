VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Char Spacing"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   300
      Left            =   150
      TabIndex        =   3
      Top             =   75
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   3330
      TabIndex        =   1
      Text            =   "10"
      Top             =   75
      Width           =   615
   End
   Begin PrintPreviewVB.VBPrintPreview VBPrintPreview1 
      Height          =   5250
      Left            =   105
      TabIndex        =   0
      Top             =   390
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   9260
      CurrentX        =   2
      CurrentY        =   27,7
      FontName        =   "Lucida Sans Unicode"
      FontCharSet     =   161
      Zoom            =   5
      PageBorder      =   3
      FromPage        =   1
      ToPage          =   1
      NavBar          =   1
      PageWidth       =   21
      PageHeight      =   29,7
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Extra char spacing"
      Height          =   195
      Left            =   1935
      TabIndex        =   2
      Top             =   105
      Width           =   1320
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim g_CharSpacing%
' note: this API is declared incorrectly in the VB API Viewer.
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long

Private Sub Command1_Click()
Dim i%
With VBPrintPreview1
    MousePointer = 11
    .Clear
    ' note: SetTextCharacterExtra takes a spacing in pixels, which we get by converting from twips.
    g_CharSpacing = Val(Text1) / Printer.TwipsPerPixelX
    .FontBold = False
    .FontSize = 18
    .FontName = "Arial"
    .StartDoc
    
    For i = 1 To 50
        .Paragraph "This is a test, just a little test. " & _
                   "This is a test, just a little test. " & _
                   "This is a test, just a little test. " & _
                   "This is a test, just a little test: Paragraph " & i & vbCrLf
    Next
    .EndDoc
    MousePointer = 0
End With
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   With VBPrintPreview1
       .Move .Left, .Top, ScaleWidth - .Left * 2, ScaleHeight - .Top - .Left
   End With
End Sub

Private Sub VBPrintPreview1_PageNew()
    'note: this needs to be called at the NewPage event because each page has its own hDC.
    SetTextCharacterExtra VBPrintPreview1.hdc, g_CharSpacing

End Sub
