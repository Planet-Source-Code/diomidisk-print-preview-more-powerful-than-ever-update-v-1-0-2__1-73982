VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MsFlexGrid"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   2430
   End
   Begin PrintPreviewVB.VBPrintPreview VBPrintPreview1 
      Height          =   6870
      Left            =   4215
      TabIndex        =   1
      Top             =   465
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   12118
      CurrentX        =   2
      CurrentY        =   27,7
      FontName        =   "Lucida Sans Unicode"
      FontCharSet     =   161
      Zoom            =   5
      FromPage        =   1
      ToPage          =   1
      NavBar          =   1
      PageWidth       =   21
      PageHeight      =   29,7
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6810
      Left            =   30
      TabIndex        =   0
      Top             =   525
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   12012
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
      PrintMsFlexGrid
End Sub

Private Sub PrintMsFlexGrid()
  Dim FormatCol$, Header$, sBody, rAlign As Integer, wAlign As Integer, i As Integer
      
  With MSHFlexGrid1
      
          .Rows = 11
          .Cols = 4
          .Row = 0
          .Col = 0
          .RowSel = 10
          .ColSel = 3
          .Clip = "" & vbTab & "Col 1" & vbTab & "Col 2" & vbTab & "Col 3" & vbCr & _
                "Row 1" & vbTab & "RC 1_1" & vbTab & "RC 1_2" & vbTab & "RC 1_3" & vbCr & _
                "Row 2" & vbTab & "RC 2_1" & vbTab & "RC 2_2" & vbTab & "RC 2_3" & vbCr & _
                "Row 3" & vbTab & "RC 3_1" & vbTab & "RC 3_2" & vbTab & "RC 3_3" & vbCr & _
                "Row 4" & vbTab & "RC 4_1" & vbTab & "RC 4_2" & vbTab & "RC 4_3" & vbCr & _
                "Row 5" & vbTab & "RC 5_1" & vbTab & "RC 5_2" & vbTab & "RC 5_3" & vbCr & _
                "Row 6" & vbTab & "RC 6_1" & vbTab & "RC 6_2" & vbTab & "RC 6_3" & vbCr & _
                "Row 7" & vbTab & "RC 7_1" & vbTab & "RC 7_2" & vbTab & "RC 7_3" & vbCr & _
                "Row 8" & vbTab & "RC 8_1" & vbTab & "RC 8_2" & vbTab & "RC 8_3" & vbCr & _
                "Row 9" & vbTab & "RC 9_1" & vbTab & "RC 9_2" & vbTab & "RC 9_3" & vbCr & _
                "Row 10" & vbTab & "RC 10_1" & vbTab & "RC 10_2" & vbTab & "RC 10_3"
          .Row = 0
          .Col = 0
          .RowSel = 0
          .ColSel = .Cols - 1
          Header$ = .Clip
          
          .Row = 1
          .Col = 0
          .RowSel = .Rows - 1
          .ColSel = .Cols - 1
          sBody = .Clip
          
    End With
    
    With VBPrintPreview1
        .Clear
        
        .FontName = "Arial"
        .Footer = "Page p$"
        .Zoom = zmPageWidth
        .HdrFontName = "Times New Roman"
        .HdrFontSize = 12
        .HdrFontBold = True
        .HdrFontItalic = False
        .HdrColor = vbRed
        .Header = "PrintPreview|Demo MSFlexGrid"
        
        .StartDoc
         
        .FontSize = 10
        
        .CurrentY = .MarginTop
        .TextAlign = taLeftTop
        .TableBorder = tbAll
        
        .FontName = MSHFlexGrid1.Font.Name
        .FontSize = MSHFlexGrid1.Font.Size
        .FontItalic = MSHFlexGrid1.Font.Italic
        .FontStrikethru = MSHFlexGrid1.Font.Strikethrough
        
        .StartTable
        FormatCol$ = ""
        For i = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.Col = i
           FormatCol$ = FormatCol$ + Format$(ScaleX(MSHFlexGrid1.CellWidth, vbTwips, .ScaleMode), "0") + "|"
        Next
        'REMOVE LAST '|'
        FormatCol$ = Mid(FormatCol$, 1, Len(FormatCol$) - 1)
         'FormatCol$ = "3|3|3|3;"
         
         Header$ = Replace(Header$, vbTab, "|")
         Header$ = Replace(Header$, vbCr, ";")
         Header$ = Header$ + ";"
          
         sBody = Replace(sBody, vbTab, "|")
         sBody = Replace(sBody, vbCr, ";")
         sBody = sBody + ";"
         
         .Table FormatCol$, Header$, sBody, MSHFlexGrid1.BackColorFixed, MSHFlexGrid1.ForeColorSel, , MSHFlexGrid1.GridLineWidth, , "1mm"
         .TableCell tcBackColor, , 1, MSHFlexGrid1.BackColorFixed
         
         For i = 0 To MSHFlexGrid1.Cols - 1
            rAlign = MSHFlexGrid1.ColAlignment(i)
            Select Case rAlign
            Case flexAlignLeftTop '= 0 The column content is aligned left, top.
                 wAlign = taLeftTop
            Case flexAlignLeftCenter '= 1  Default for strings. The column content is aligned left, center.
                 wAlign = taLeftMiddle
            Case flexAlignLeftBottom '= 2 The column content is aligned left, bottom.
                 wAlign = taLeftBottom
            Case flexAlignCenterTop '= 3 The column content is aligned center, top.
                 wAlign = taCenterTop
            Case flexAlignCenterCenter '= 4 The column content is aligned center, center.
                 wAlign = taCenterMiddle
            Case flexAlignCenterBottom '= 5 The column content is aligned center, bottom.
                 wAlign = taCenterBottom
            Case flexAlignRightTop '= 6 The column content is aligned right, top.
                 wAlign = taRightTop
            Case flexAlignRightCenter '= 7 Default for numbers. The column content is aligned right, center.
                 wAlign = taRightMiddle
            Case flexAlignRightBottom '= 8 The column content is aligned right, bottom.
                 wAlign = taRightBottom
            Case Else
                wAlign = taLeftTop
            End Select
            
            .TableCell tcTextAling, , i + 1, wAlign
         Next
       
         .EndTable
         .EndDoc
      End With
    
End Sub
 

