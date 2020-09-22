VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "External Controls"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin PrintPreviewVB.VBPrintPreview VBPrintPreview1 
      Height          =   4500
      Left            =   390
      TabIndex        =   11
      Top             =   765
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   7938
      CurrentX        =   2
      CurrentY        =   27,7
      FontBold        =   -1  'True
      FontSize        =   12
      FontCharSet     =   161
      Zoom            =   5
      FromPage        =   1
      ToPage          =   1
      PageWidth       =   21
      PageHeight      =   29,7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   11625
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      Begin VB.CommandButton cmdFirst 
         Height          =   500
         Left            =   0
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton CmdZoom 
         Height          =   500
         Left            =   4665
         Picture         =   "Form1.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   15
         Width           =   800
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3135
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   75
         Width           =   1485
      End
      Begin VB.CommandButton cmdGoTo 
         Height          =   500
         Left            =   1215
         MaskColor       =   &H80000005&
         Picture         =   "Form1.frx":0FA4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmd_quit 
         Cancel          =   -1  'True
         Height          =   510
         Left            =   6360
         Picture         =   "Form1.frx":170E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   15
         Width           =   810
      End
      Begin VB.CommandButton cmd_print 
         Height          =   500
         Left            =   5460
         Picture         =   "Form1.frx":23D8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   15
         Width           =   800
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   500
         Index           =   0
         Left            =   615
         Picture         =   "Form1.frx":2B42
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   500
         Index           =   1
         Left            =   1815
         Picture         =   "Form1.frx":2ECC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdLast 
         Height          =   500
         Left            =   2415
         Picture         =   "Form1.frx":3256
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   240
         Left            =   8370
         TabIndex        =   10
         Top             =   135
         Width           =   2325
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Combo1.AddItem "50%"
    Combo1.AddItem "75%"
    Combo1.AddItem "100%"
    Combo1.AddItem "150%"
    Combo1.AddItem "200%"
    Combo1.AddItem "WholePage"
    Combo1.AddItem "PageWidth"
    Combo1.AddItem "ThumbNail"
    
    Combo1.ListIndex = VBPrintPreview1.Zoom
    DoParagraphs
End Sub
Private Sub CmdZoom_Click()
     If Combo1.ListIndex + 1 <> Combo1.ListCount Then
       Combo1.ListIndex = Combo1.ListIndex + 1
     Else
       Combo1.ListIndex = 0
     End If
End Sub

Private Sub cmd_print_Click()
   VBPrintPreview1_PagePrint
End Sub

Private Sub cmd_quit_Click()
      Unload Me
End Sub

Private Sub cmdFirst_Click()
    VBPrintPreview1.PageFirst
End Sub

Private Sub cmdGoTo_Click()
    VBPrintPreview1.PageGoTo
End Sub

Private Sub cmdLast_Click()
    VBPrintPreview1.PageLast
End Sub

Private Sub cmdPrevious_Click(Index As Integer)
    If Index = 0 Then
       VBPrintPreview1.PagePreview
    Else
       VBPrintPreview1.PageNext
    End If
    
End Sub
Private Sub Combo1_Click()
    VBPrintPreview1.Zoom = Combo1.ListIndex
End Sub

Private Sub Form_Resize()
     If Me.WindowState = 1 Then Exit Sub
     If Me.ScaleHeight < Picture1.Height Then Exit Sub
     'If Me.ScaleWidth < List1.Width Then Exit Sub
     
     VBPrintPreview1.Move 0, Picture1.Height, Me.ScaleWidth, Me.ScaleHeight - Picture1.Height
End Sub

Private Sub VBPrintPreview1_AfterUserScroll()
Combo1.ListIndex = VBPrintPreview1.Zoom
End Sub

Private Sub VBPrintPreview1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim tX As String, ty As String
       tX = Format(Round(X, 2), "0.00")
       ty = Format(Round(Y, 2), "0.00")
       Label1.Caption = "X:" + tX + "- Y:" + ty
End Sub

Private Sub VBPrintPreview1_PagePrint()
       If VBPrintPreview1.DialogPrint(pdPrint) Then
           VBPrintPreview1.SendToPrinter = True
           DoParagraphs 'print
           VBPrintPreview1.SendToPrinter = False
           DoParagraphs 'preview
        End If
End Sub

Private Sub DoParagraphs()
    
    Dim s$
    Dim iChapter%, iParagraph%
    
    With VBPrintPreview1
        .Clear
        .Zoom = zmWholePage
        .TextAlign = taJustifyTop
        .PageBorder = pbTopBottom
        SetPages "PrintPreview|Paragraph"
        
       .StartDoc
        .FontSize = 12
         s = "Test paragraph to show the potential of 'Function Paragrafh'." + _
             " Before adopting the text, you can register your Property TextAlign," + _
             " LineSpace, IndentFirst, IndentLeft and Fonts property."
         
       .TextAlign = taJustifyTop
       .LineSpace = lsSpaceSingle ' lsSpaceLine15
       
        For iChapter = 1 To 3
            .FontSize = 24
            .FontBold = True
            .FontItalic = True
            .FontUnderline = True
             SetSubTitle "Chapter " & iChapter
            .IndentLeft = "10mm"
            .IndentFirst = "5mm"
            .IndentRight = "10mm"
            .FontSize = 12
            .FontBold = False
            .FontItalic = False
            .FontUnderline = False
           .LineSpace = lsSpaceLine15
            For iParagraph = 1 To 3 + Rnd * 5
                .Paragraph s
            Next
            .IndentLeft = 0
            .IndentFirst = 0
        Next
       .EndDoc
       
    End With

End Sub

Sub SetPages(Optional Header As String = "PrintPreview", _
             Optional Footer As String = "", _
             Optional PageOrientation As PageOrientationConstants = smCentimeters)

   With VBPrintPreview1
         If Footer <> "" Then Footer = Footer + "|"
         .Footer = Format(Now, "dddd dd mmmm yyyy Hh:Nn:Ss") + "|" + Footer + "Page p$"
         .HdrFontName = "Times New Roman"
         .HdrFontSize = 12
         .HdrFontBold = True
         .HdrFontItalic = False
         .HdrColor = vbBlue
         .Header = Header
    End With
    
End Sub

Sub SetSubTitle(s$)
   With VBPrintPreview1
   
        .FontName = "Arial"
        .Paragraph ""
        .FontBold = True
        .FontItalic = True
        .FontUnderline = True
        .FontSize = .FontSize * 1.5
        .ForeColor = RGB(255, 126, 64)
        .Paragraph s$
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontSize = .FontSize / 1.5
        .ForeColor = 0
        .FillColor = vbBlack
        .FillStyle = vbFSTransparent
        .DrawWidth = 0
   End With
End Sub

Private Sub VBPrintPreview1_PageView()
     Me.Caption = "PrintPreview " + Str(VBPrintPreview1.CurrentPage) + "/" + Str(VBPrintPreview1.PageCount)

End Sub
