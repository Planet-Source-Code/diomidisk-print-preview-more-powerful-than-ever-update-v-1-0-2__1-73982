VERSION 5.00
Begin VB.Form FormDemo 
   Caption         =   "Demo Print Preview"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   Icon            =   "FormDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin PrintPreviewVB.VBPrintPreview VBPrintPreview1 
      Height          =   5025
      Left            =   3450
      TabIndex        =   2
      Top             =   75
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   8864
      CurrentX        =   2
      CurrentY        =   2,2
      FontBold        =   -1  'True
      FontSize        =   36
      FontCharSet     =   161
      Zoom            =   5
      PageBorder      =   3
      ToPage          =   1
      NavBar          =   3
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
      NavBarLabels    =   "Page 1"
   End
   Begin VB.PictureBox PicLoad 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   9750
      ScaleHeight     =   1350
      ScaleWidth      =   1605
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   3300
   End
End
Attribute VB_Name = "FormDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function AbortPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Sub DoTable()
    
    Dim FormatCol$, Header$, Body$, B As Variant
    
    With VBPrintPreview1
    
        .Clear
        .Zoom = zmWholePage
         SetPages "PrintPreview|Tables"
         
        .StartDoc
         SetTitle "Tables function"
        
        .LineSpace = lsSpaceSingle
        .Paragraph
        .Paragraph "Renders a table on the page."
        .Paragraph
        .FontBold = True
        .Paragraph "Syntax:"
        
        .Paragraph
        .FontSize = .FontSize - 2
        .Paragraph "[form.]VBPPreview.Table FormatCols As String, Header As String, Body As String,"
        .IndentLeft = .TextWidth("[form.]VBPPreview.Table ")
        .Paragraph "[HeaderShade As Long], [BodyShade As Long],"
        .Paragraph "[LineColor As Long], [LineWidth As Integer],"
        .Paragraph "[Wrap As Boolean], [Indent As Single],"
        .Paragraph "[WordWrap As Boolean], [Indent As Single]"
        .FontSize = .FontSize + 2
        .FontBold = False
        .IndentLeft = 0
        .LineSpace = lsSpaceSingle
        .TextAlign = taJustifyTop
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "FormatCols$"
        .FontBold = False: .FontItalic = False
        
        .Paragraph "This parameter contains formatting information and is not printed. " + _
                   "The formatting information describes each column using a sequence of " + _
                   "formatting characters followed by the column width. The information for " + _
                   "each column is delimited by the column separator character (by default a pipe (''|'')."
        
        .Paragraph
        .IndentFirst = "1cm"
        .Paragraph "For example, the following string defines a table with four center-aligned, " + _
                    "two-inches wide columns:"
        .Paragraph
        .Paragraph "s$ = ''^+2in|^+2in|^+2in|^+2in''"
        .Paragraph
        .Paragraph "The following lists shows all valid formatting characters:"
        .Paragraph
        
            
      '"<|Align column contents to the left top;" + _
      '">|Align column contents to the right top;" + _
      '"^|Align column contents to the center top;" + _
      '"=|Align column contents to the justify top;" + _
      '"<+|Align column contents to the left Middle;" + _
      '">+|Align column contents to the right Middle;" + _
      '"^+|Align column contents to the center Middle;" + _
      '"=+|Align column contents to the justify Middle;" + _
      '"<_|Align column contents to the left Bottom;" + _
      '">_|Align column contents to the right Bottom;" + _
      '"^_|Align column contents to the center Bottom;" + _
      '"=_|Align column contents to the justify Bottom;"
        
        .IndentFirst = 0
        .StartTable
          .Table "^0.8in|<4in;", "Character|Effect;", _
            "<|Align column contents to the left top;" + _
            ">|Align column contents to the right top;" + _
            "^|Align column contents to the center top;" + _
            "=|Align column contents to the justify top;" + _
            "<+|Align column contents to the left middle;" + _
            ">+|Align column contents to the right middle;" + _
            "^+|Align column contents to the center middle;" + _
            "=+|Align column contents to the justify middle;" + _
            "<_|Align column contents to the left bottom;" + _
            ">_|Align column contents to the right bottom;" + _
            "^_|Align column contents to the center bottom;" + _
            "=_|Align column contents to the justify bottom;", , , vbWhite, , False, "2mm"
          .TableCell tcFontBold, 1, , True
        .EndTable
        
        .IndentFirst = "1cm"
       
        .Paragraph
        .Paragraph "Column widths may be specified in twips, inches, points, millimeters, centimeters, pixel, or as a percentage of the width of age. If the units are not provided, Scalemode used. For details on using unit aware measurements, see Using Unit Properties."
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Header$"
        .FontBold = False: .FontItalic = False
        .Paragraph "This parameter contains the text to be printed on the first row of the table " + _
                   "and after each column or page break (the header row). The text for each cell " + _
                   "in the header row is delimited by the column separator character (by default a pipe (''|'')."
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Body$"
        .FontBold = False: .FontItalic = False
        .Paragraph "This parameter contains the text for the table body. Cells are delimited by column" + _
                   " separator characters, and rows are delimited by row separator characters. By default," + _
                   " cells pare separated by pipes (''|'') and rows by semi-colons ('';'')."
        .Paragraph "Instead of supplying the table data as a string, you may create the table based on data" + _
                   " from a Variant array. To do this, use the TableArray method."
        .Paragraph "You may also choose to supply data for individual cells separately. To do this, " + _
                    "use the TableCell property."
'        .NewPage
        .FontBold = True: .FontItalic = True
        .Paragraph "HeaderShade, BodyShade(optional)"
        .FontBold = False: .FontItalic = False
        .Paragraph "These parameters specify colors to be used for shading the cells in the header " + _
                   "and in the body of the table. If omitted or set to zero, the cells are not shaded. " + _
                   "If you want to use black shading use a very dark shade of gray instead (e.g. 1 or RGB(1,1,1))."
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "LineColor, LineWidth(optional)"
        .FontBold = False: .FontItalic = False
        .Paragraph " These parameters specify colors and width to be used for line in table." + _
                   " Defaults color is black and line width=1. Also see TableBorder."

        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Wrap"
        .FontBold = False: .FontItalic = False
        .Paragraph "Specifies whether text should be allowed to wrap within the box. Optional, defaults to True."

        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Indent"
        .FontBold = False: .FontItalic = False
        .Paragraph "To indent a table by a specified amount of left and right of the column."
        .FontBold = False: .FontItalic = False

        .Paragraph
        .FontBold = True: .FontItalic = True: .FontUnderline = True
        .Paragraph "Aligning and indenting tables:"
        .FontBold = False: .FontItalic = False: .FontUnderline = False
        .Paragraph "Tables may be aligned to the left, center, or right of the page depending on the " + _
                   "setting of the TextAlign property."
        .Paragraph
        .Paragraph "Note: The table is using StartTable, Table or TableArray and finally EndTable. " + _
                   "Set properties of a table cell with a TableCell and for border with TableBorder. " + _
                   "To read the dimensions of the table on each page and table CalcTable call the function and read the values in the X1, Y1, X2, Y2"

        .Paragraph
        .FontBold = True
        .Paragraph "Example"
        .FontBold = False
        .ForeColor = vbBlue
        .Paragraph "With PrintPreview"
        .IndentLeft = 1
        .Paragraph ".StartDoc"
        .Paragraph ".TextAlign = taLeftTop"
        .Paragraph ".TableBorder = tbAll"
        .Paragraph ".StartTable"
        .Paragraph "FormatCol$ = ''^+3.5cm|<+3.5cm|<+3.5cm|^+3.5cm|>+3.5cm''"
        .Paragraph "Header$ = ''Col 1 Row 1|Col 2 Row 1|Col 3 Row 1|Col 4 Row 1|Col 5 Row 1''"
        .Paragraph "Body$ = ''Col 1 Row 2|Col 2 Row 2|Col 3 Row 2|Col 4 Row 2|Col 5 Row 2;'' + _"
        .IndentLeft = .TextWidth("Body$ = ") + 1
        .Paragraph "''Col 1 Row 3|Col 2 Row 3|Col 3 Row 3|Col 4 Row 3|Col 5 Row 3;'' + _"
        .Paragraph "''Col 1 Row 4|Col 2 Row 4|Col 3 Row 4|Col 4 Row 4|Col 5 Row 4;''"
        .IndentLeft = 1
        .Paragraph
        .Paragraph ".Table FormatCol$, Header$, Body$, vbYellow, &HFFFFC0, vbBlue, 2, , ''1mm''"
        .Paragraph ".EndTable"
        .Paragraph ".EndDoc"
        .IndentLeft = 0
        .Paragraph "End With"
        .ForeColor = vbBlack
        .IndentFirst = 0
        .Paragraph
        .TextAlign = taLeftTop
        .TableBorder = tbAll
        .StartTable
         FormatCol$ = "^+3.5cm|<+3.5cm|<+3.5cm|^+3.5cm|>+3.5cm"
            Header$ = "Col 1 Row 1|Col 2 Row 1|Col 3 Row 1|Col 4 Row 1|Col 5 Row 1"
              Body$ = "Col 1 Row 2|Col 2 Row 2|Col 3 Row 2|Col 4 Row 2|Col 5 Row 2;" & _
                      "Col 1 Row 3|Col 2 Row 3|Col 3 Row 3|Col 4 Row 3|Col 5 Row 3;" & _
                      "Col 1 Row 4|Col 2 Row 4|Col 3 Row 4|Col 4 Row 4|Col 5 Row 4;"
           .Table FormatCol$, Header$, Body$, vbYellow, &HFFFFC0, vbBlue, 2, , "0.8mm"
           .TableCell tcRowHeight, , , "10mm"
        .EndTable
       .CalcTable
       .Paragraph "Position Table - X1:" + Format(.X1, "0.00") + " - Y1:" + Format(.Y1, "0.00") + _
                                    " - X2:" + Format(.X2, "0.00") + " - Y2:" + Format(.Y2, "0.00")
                                    
        Debug.Print "Position Table", .X1, .Y1, .X2, .Y2
     
     .EndDoc
     
     End With
End Sub

Sub DoTableCell()
     Dim Body  As String
     With VBPrintPreview1
        .Clear
        
        .Zoom = zmWholePage
         SetPages "PrintPreview|TableCell"
        .PageBorder = pbTopBottom
        .StartDoc
         SetTitle "TableCell function"
        .ForeColor = vbBlack
        .Paragraph
        .FontSize = 12
        .Paragraph "Returns or sets properties of a table cell or range."
        .Paragraph
        .FontBold = True
        .Paragraph "Syntax:"
 
        .Paragraph
        .Paragraph "[form.]VBPPreview.TableCell (Settings As TableSettingConstants,"
         .IndentLeft = .TextWidth("[form.]VBPPreview.TableCell ")
         .Paragraph "[Row As Variant], [Col As Variant],"
         .Paragraph "[Value As Variant] = Variant)"
         .FontBold = False
         .IndentLeft = 0
        .Paragraph
        
        .Paragraph "The TableCell property is used to build and format tables. Using TableCell requires four steps:"
        .IndentLeft = "5mm"
        .IndentFirst = "-5mm"
        .Paragraph "1) Start a table definition with the StartTable method."
        .Paragraph "2) Create the table using the AddTable or AddTableArray methods, or using the TableCell property by itself."
        .Paragraph "3) Format the table using the TableCell property."
         .Paragraph "4) Close the table definition with the EndTable method. This will render the table."
         .IndentLeft = 0
        .IndentFirst = 0
         .Paragraph
'         .IndentLeft = "20mm"
'         .IndentFirst = "-20mm"
        .Paragraph "The parameters for the TableCell property are described below:"
        .FontUnderline = True
'        .Paragraph "Parameter Description"
        .FontUnderline = False
         
        Body = "Setting|Determines which property of the table, row, column, or cell you want to set or retrieve. " + _
                   " The list of valid settings is given below.;"
        Body = Body + "Row|Row in the range. The header row has index zero. Body rows are one-based.;"
        Body = Body + "Col|Column in the range. The column has index one.;"
         
         .StartTable
           .Table "38mm|14cm", "Parameter|Description", Body, , , vbWhite
           .TableCell tcFontBold, 1, , True
         .EndTable
         
        
        .Paragraph ""
        .Paragraph "Table Properties: These settings affect the entire table."
        .Paragraph ""
    
'        .IndentLeft = "34mm"
'         .IndentFirst = "-34mm"
          
         Body = "tcCols|Returns or sets the number of columns on the table. If you change the number of columns, columns are added or deleted from the right of the table.;"
         Body = Body + "tcRows|Returns or sets the number of rows on the table. " + _
                       "The header row is not included in this count. If you change the number " + _
                       "of rows, rows are added or deleted from the bottom of the table.;"
         Body = Body + "tcColWidth |Returns or sets the width of the columns. You may specify units with this value.;"
         Body = Body + "tcRowHeight|Returns or sets the height of the rows. The header row has index zero. You may specify units with this value.;"
         Body = Body + "tcIndent|Returns or sets the indent for the table. You may specify units with this value.;"
         Body = Body + "tcText|Returns or sets the cell text. If the table is bound to an array, setting the property will change the table.;"
         Body = Body + "tcBackColor|Returns or sets the background color for the cell.;"
         Body = Body + "tcForeColor|Returns or sets the foreground (text) color for the cell.;"
         Body = Body + "tcFontName |Returns or sets the name of the cell font.;"
         Body = Body + "tcFontSize |Returns or sets the size of the cell font.;"
         Body = Body + "tcFontCharSet|Returns or sets the character set for font.;"
         Body = Body + "tcFontBold |Returns or sets the bold attribute of the cell font.;"
         Body = Body + "tcFontItalic|Returns or sets the italic attribute of the cell font.;"
         Body = Body + "tcFontUnderline|Returns or sets the underline of the cell font.;"
         Body = Body + "tcFontStrikethru|Returns or sets the strikethrough attribute of the cell font.;"
         Body = Body + "tcFontTransparent|Returns or sets the FontTransparent attribute of the cell font.;"
         Body = Body + "tcPicture|Returns or sets the cell picture.;"
         Body = Body + "tcColSpan|Returns or sets the number of rows that the cell should span (col merging).;"
         Body = Body + "tcRowSpan|Returns or sets the number of rows that the cell should span (row merging).;"
         Body = Body + "tcTextAling|Returns or sets the alignment of text in the cells. Valid settings are the same used with the TextAlign property.;"
         .StartTable
           .Table "38mm|13cm", "Setting|Description", Body, , , vbWhite
           .TableCell tcFontBold, 1, , True
         .EndTable
         .IndentLeft = 0
         .IndentFirst = 0
         .Paragraph
         .Paragraph "The Table Formatting example shows how you can use the TableCell property to create and format a table."
    .EndDoc
    End With
End Sub

Sub DoTextBoxes()
    Dim s$, X!, Y!, wid!, i%, fs!
 
    With VBPrintPreview1
    .Clear
    
     SetPages "PrintPreview|TextBox"
    .PageBorder = pbTopBottom
        
    .StartDoc
     SetTitle "TableBox function"
    .ForeColor = vbBlack
    .Paragraph
    .FontSize = 12
    .Paragraph "Draws text within a rectangle."
    .Paragraph
    .FontBold = True
    .Paragraph "Syntax:"
    .Paragraph
    .Paragraph "[form.]VBPPreview.TextBox Text As String, X As Variant, Y As Variant, _"
    .IndentLeft = .TextWidth("[form.]VBPPreview.TextBox ")
    .Paragraph "Width As Variant, Height As Variant, _"
    .Paragraph "[Align As TextAlignConstants], _"
    .Paragraph "[BackShade As Boolean ], _"
    .Paragraph "[BoxShade As Boolean ]"
    .IndentLeft = 0
    .Paragraph
    .FontBold = False
    .IndentLeft = "14mm"
    .IndentFirst = "-14mm"
        s = "Text    : Text to be drawn in the text box." + vbCr
    s = s + "X, Y    : Coordinates of the upper left corner of the box. You may specify the units for these parameters. See ''Using Unit Property''." + vbCr
    s = s + "Width : The width of the text box, in explicit units. See ''Using Unit Property''." + vbCr
    s = s + "Height: The height of the text box, in  explicit units. If this parameter is set to zero, the height is calculated automatically so the text fits within the box. See ''Using Unit Property''." + vbCr
    s = s + "Aling   : Sets the alignment of printed textboxes. Optional, defaults to LeftTop." + vbCr
    s = s + "BackShade: Specifies whether the text box will be outlined with the ForeColor and shaded with the FillColor. Optional, defaults to True." + vbCr
    s = s + "BoxShade: Specifies whether the box will contain shadow. Optional, defaults to False."
    .Paragraph s
    .Paragraph
    .IndentLeft = 0
    .IndentFirst = 0
    .TextAlign = taJustifyTop
    .Paragraph vbTab + "This is something that comes up every once in a while: you want to draw text at a certain spot " + _
               "on the page, with shading and a border."
    .Paragraph vbTab + "You can use the TextBox method to build sophisticated reports with total control over text positioning and page breaks, or forms with arbitrary field positioning."
    .Paragraph vbTab + "So here's an example that draws a bunch of random boxes, with random colors, and random font sizes."
    
    fs = .FontSize
     .FillStyle = vbFSSolid
    
    s = "Silly text in a box, drawn with the 'TextBox' method."

    For i = 0 To 20
        .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
        X = Format(.MarginLeft + Rnd() * 10, "0.00")        'unit is cm
        Y = Format(.MarginTop + 14 + Rnd() * 8, "0.00")     'unit is cm
        wid = Format(4 + Rnd() * 2, "0.00")                 'unit is cm
        If X + wid > .PageWidth - .MarginRight Then
            X = .PageWidth - wid
            wid = Abs(wid - .MarginRight)
        End If
        .FontSize = 6 + Rnd() * 10
        .TextBox s, X, Y, wid, 0, taLeftTop, True
    Next
    .FontSize = fs
    .EndDoc
    End With
    
End Sub



Private Sub Form_Load()
    List1.Clear
    
    List1.AddItem "• VB Print Preview"
    List1.AddItem "  Text"
    List1.AddItem "  TextRTF"
    List1.AddItem "  Paragraph"
    List1.AddItem "  Text Boxes"
    List1.AddItem "  Zoom mode"
    List1.AddItem "  Using Unit Properties"
            
    List1.AddItem "• Document Layout"
    List1.AddItem "  Page Border"
    List1.AddItem "  PaperSize"
    List1.AddItem "  Orientation"
    List1.AddItem "  Margins"
    List1.AddItem "  GetMargins"
    List1.AddItem "  Header & Footers"
    List1.AddItem "  Page Setup Dialog"
    List1.AddItem "  Print Setup"
    List1.AddItem "  Print Dialog"
    
    List1.AddItem "• Document Navigation"
    List1.AddItem "  Navigation bar"
    List1.AddItem "  Moving thru Pages"
    
    List1.AddItem "• Tables"
    List1.AddItem "  Tables"
    List1.AddItem "  Tables Array"
    List1.AddItem "  Tables Border"
    List1.AddItem "  Tables Cell"
    
    List1.AddItem "• Drawings"
    List1.AddItem "  Picture"
    List1.AddItem "  Line"
    List1.AddItem "  Rectangle"
    List1.AddItem "  Polygon"
    List1.AddItem "  Circe, Arc"
    List1.AddItem "  Ellipse"
    
    List1.AddItem "• Object Formatting"
    List1.AddItem "  Setting Fonts"
   'List1.AddItem "  Rotating Text"
    List1.AddItem "  LineSpace"
    List1.AddItem "  Alignment"
    List1.AddItem "  Justification"
    List1.AddItem "  Indent Property"
    
    List1.AddItem "• Object Measuring"
    List1.AddItem "  Measurement Units"
    List1.AddItem "  Measuring Text"
    
    List1.AddItem "• Example"
    List1.AddItem "  Paragraphs"
    List1.AddItem "  Margins demo"
    List1.AddItem "  Tables With Ado"
    List1.AddItem "  Tables Formatting"
    List1.AddItem "  Pictures"
    List1.AddItem "  Calendar month"
    List1.AddItem "  Calendar year"
    List1.AddItem "  OutLine"
    List1.AddItem "  MsFlexGrid"
    List1.AddItem "  Export as Picture"
    List1.AddItem "  Char Spacing"
    List1.AddItem "  External Controls"
    List1.AddItem "  Picture Background"
    List1.AddItem "  Fax form"
    List1.AddItem "  Invoice"
    List1.AddItem "  Labels A4 70x37mm"
    VBPrintPreview1.SendToPrinter = False
   
    List1.ListIndex = 0

End Sub

Private Sub Form_Resize()
     If Me.WindowState = 1 Then Exit Sub
      'If Me.ScaleHeight < Picture1.Height Then Exit Sub
      'If Me.ScaleWidth < List1.Width Then Exit Sub
      
      'List1.Move 0, Picture1.Height, 3000, Me.ScaleHeight - Picture1.Height
      'VBPrintPreview1.Move List1.Left + List1.Width, Picture1.Height, Me.ScaleWidth - List1.Width, Me.ScaleHeight - Picture1.Height
      List1.Move 0, 0, 3000, Me.ScaleHeight
      VBPrintPreview1.Move List1.Left + List1.Width, 0, Me.ScaleWidth - List1.Width, Me.ScaleHeight
End Sub

Private Sub List1_Click()
     Screen.MousePointer = 11
     
     SetOriginalSettings
     
        Select Case Trim(List1.Text)
        Case "• VB Print Preview":    DoDocument
        Case "Text":                  DoText
        Case "TextRTF":               MsgBox "Not yet implementation.", vbInformation 'DoTextRTF
        Case "Paragraph":             DoParagraphs
        Case "Text Boxes":            DoTextBoxes
        Case "Zoom mode":             DoZoom
        Case "Using Unit Properties": DoUsingUnit
        
        Case "• Document Layout":     DoPageTitle "Document Layout"
        Case "Page Border":           DoPageBorder
        Case "PaperSize":             DoPaperSize
        Case "Margins":               DoMargins
        Case "GetMargins":            DoGetMargins
        Case "Orientation":           DoOrientation
        Case "Header & Footers":      DoHeaderFooter
        Case "Page Setup Dialog":     DoDialog: DoPageSetup
        Case "Print Setup":           DoDialog: DoPrintSetup
        Case "Print Dialog":          DoDialog: DoPrintDialog
        
        Case "• Document Navigation": DoPageTitle "Document Navigation"
        Case "Navigation bar":        DoNavBar
        Case "Moving thru Pages":     DoMovingThruPages
    
        Case "• Tables":              DoPageTitle "Tables"
        Case "Tables":                DoTable
        Case "Tables Cell":           DoTableCell
        Case "Tables Array":          DoTableArray 'AdoData
        Case "Tables Border":         DoTableBorder
        
        Case "• Drawings":            DoPageTitle "Drawings"
        Case "Picture":               DoDrawPicture
        Case "Line":                  DoLine
        Case "Rectangle":             DoRectangle
        Case "Polygon":               DoPolygon
        Case "Circe, Arc":            DoCircle
        Case "Ellipse":               DoElipse
        
        Case "• Object Formatting":   DoPageTitle "Object Formatting"
        Case "Setting Fonts":         DoFont
        Case "LineSpace":             DoLineSpace
        Case "Alignment":             DoAlignment
        Case "Justification":         DoJustification
        Case "Indent Property":       DoIndent
        
        Case "• Object Measuring":    DoPageTitle "Object Measuring"
        Case "Measurement Units":     DoMeasureUnits
        Case "Measuring Text":        DoMeasureText

        Case "• Example":             DoPageTitle "Example"
        Case "Paragraphs":            DoParagraphs True
        Case "Margins demo":          DoMargins True
        Case "Tables With Ado":       DoAdoData
        Case "Tables Formatting":     DoTableFormat
        Case "Pictures":              DoPicture
        Case "Calendar month":        DoCalendar
        Case "Calendar year":         DoCalendarYear
        Case "OutLine":               DoOutline
        Case "MsFlexGrid":            Form2.Show 1: Exit Sub
        Case "Export as Picture":     DoExport
        Case "Char Spacing":          Form3.Show 1
        Case "External Controls":      Form1.Show 1
        Case "Picture Background":    DoGetMargins True
        Case "Fax form":              Form4.Show 1
        Case "Invoice":               DoInvoice
        Case "Labels A4 70x37mm":     DoLabels
        Case Else
           VBPrintPreview1.Clear
           VBPrintPreview1.StartDoc
           VBPrintPreview1.EndDoc
      End Select
     
     Screen.MousePointer = 0
End Sub

Private Sub VBPrintPreview1_Error(ByVal id As Long, ByVal ErrorDescription As Variant)
    MsgBox "Error:" + Str(id) + "-" + ErrorDescription, vbCritical
End Sub

Private Sub VBPrintPreview1_PageEndDoc()
     
     Me.Caption = "PrintPreview " + Str(VBPrintPreview1.CurrentPage) + "/" + Str(VBPrintPreview1.PageCount)
     
End Sub

'Press Command Print from NavBar
Private Sub VBPrintPreview1_PagePrint()
        If VBPrintPreview1.DialogPrint(pdPrint) Then
           VBPrintPreview1.SendToPrinter = True
           List1_Click 'print
           VBPrintPreview1.SendToPrinter = False
           List1_Click 'preview
        End If
End Sub

Private Sub VBPrintPreview1_PageView()
        Me.Caption = "PrintPreview " + Str(VBPrintPreview1.CurrentPage) + "/" + Str(VBPrintPreview1.PageCount)
End Sub


Private Sub DoInvoice()
   
   With VBPrintPreview1
        .PaperSize = vbPRPSA4
        .PageBorder = pbNone
        .Zoom = zmWholePage
        .TableBorder = tbBoxColumns
        .StartDoc
        
        .MarginBottom = 0
        .MarginTop = 0
        
        ' Set Title
        .FontSize = 24
        .ForeColor = vbBlue
        .FontBold = True
        .FontName = "Arial"
        .CurrentX = "6.25in"
        .CurrentY = "1in"
        .Text "Invoice"
        .FontSize = 10
        .ForeColor = vbBlack
        .TextAlign = taLeftTop '
        .Paragraph
        
        'set Logo
        PicLoad.Picture = LoadPicture(App.Path + "\Library\Logo.jpg")
        .DrawPicture PicLoad.Picture, "0.8in", "0.5in", "100%", "100%"
        .FontBold = False
        
        'Set boxes
        DoBox "5.75in", "1.5in", "^0.75in|^1in", "DATE|PO NUMBER;", vbCrLf & vbCrLf
        DoBox "1in", "2.5in", "<3in", "BILL TO", vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        DoBox "4.5in", "2.5in", "<3in", "SHIP TO", vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        DoBox "1in", "4in", "^0.75in|^0.75in|^1in|^1in|^1in|^1in|^1in;", "P.O. NO.|TERMS|REP|SHIP DATE|SHIP VIA|FOB|PROJECT;", vbCrLf & vbCrLf
        DoBox "1in", "4.75in", "^0.75in|^2.75in|^1in|^1in|^1in;", "ITEM|DESCRIPTION|QTY|RATE|AMOUNT;", vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        
        'set total table
        .TableBorder = tbAll
        .MarginLeft = "5.5in"
        .StartTable
            .Table "<1in|<1in;", "Subtotal|" & vbCrLf & ";", "Tax|" & vbCrLf & ";Total|" & vbCrLf & ";" ', , , , 2
            .TableCell tcFontSize, 3, 1, 18
            .TableCell tcFontBold, 3, 1, True
        .EndTable
        .EndDoc
        
    End With
End Sub

Private Sub DoBox(X, Y, BoxWidth, Header, Body)
    
    With VBPrintPreview1
        .MarginLeft = X
        .CurrentY = Y
        .StartTable
            .Table BoxWidth & ";", Header & ";", Body & ";", 1
            .TableCell tcForeColor, 1, , &HFFFFFF
            .TableCell tcFontBold, , , True
        .EndTable
    End With

End Sub

Private Sub DoOutline()

    Dim hBrush As Long, oldBrush As Long

   With VBPrintPreview1
     .Zoom = zmWholePage
    .StartDoc
    
        .FontSize = 36
        .ForeColor = RGB(0, 0, 0)
        
        .Paragraph "Hello, World!"
        .Paragraph
        
        .FontName = "Tahoma"
        
        'penwidth outline
        .DrawWidth = 1
        
        '----------------------------------
        'color outline
        .ForeColor = RGB(255, 0, 0)
        ' opens a path bracket on the specified hDC
        BeginPath .hdc
        ' draw text and objects as usual
        .Paragraph "Hello, World!"
        'closes path bracket on the specified hDC
        EndPath .hdc
        'outlines the current path using the current pen
        StrokePath .hdc
        ' discards the path
        AbortPath .hdc
        
        'create a new, white brush
        hBrush = CreateSolidBrush(vbGreen)
        'replace the current brush with the new white brush
        oldBrush = SelectObject(.hdc, hBrush)
        '----------------------------------
        'penwidth outline
        .DrawWidth = 2
        
        .ForeColor = vbBlue
        ' opens a path bracket on the specified hDC
        BeginPath .hdc
        ' draw text and objects as usual
        .Paragraph "Hello, World!"
        'closes path bracket on the specified hDC
        EndPath .hdc
        'outlines the current path using the current pen
        StrokePath .hdc
        '----------------------------------
        'penwidth outline
        .DrawWidth = 1
        .ForeColor = vbBlue
        .FillColor = vbGreen
       ' opens a path bracket on the specified hDC
         BeginPath .hdc
        .Paragraph "Hello, World!"
       'close the path bracket
         EndPath .hdc
        'render the specified path by using the current pen
        
         StrokeAndFillPath .hdc
         
         SelectObject .hdc, oldBrush
        'delete our white brush
        DeleteObject hBrush
        
    .EndDoc
End With

End Sub
Private Sub DoJustification()
    
    Dim i%, mCurrY#, s$
    s = "Talking to yourself is the first sign of insanity."
 
    With VBPrintPreview1
         .StartDoc
            i = .TextAlign
            
            .FontSize = 11
            SetSubTitle "Understanding Justification"
            .Paragraph "The TextAlign property is the one that let you set the justification for text in paragraph, " + _
                       "table cells and textboxes.This property allow you to set both the horizontal justification " + _
                       "of these objects. Following is a description of how justification affects each object."
  
             SetSubTitle "Full Justification in Paragraphs"
            .Paragraph "Paragraphs can only be justified horizontally.  Vertical settings are ignored.  In order to fully justified paragraphs, you should use only " & _
                        "taJustifyTop."
            
            .TextAlign = taJustifyTop
            .FontSize = 18
            .Paragraph "Sample:"
            .Paragraph s & s & s
            .Paragraph " "
            SetNormal
            
            ' Aligning Table Cells and TextBoxes
            SetSubTitle "Full justification in TextBoxes"
            .Paragraph "The text in textboxes and/or tablecells can be justified horizontally by setting the TextAlign property to taJustifyTop,taJustifyMiddle and taJustifyBottom ."
            .Paragraph " "
            
            .IndentLeft = "2px"
            .IndentRight = "2px"
            .FillStyle = vbFSSolid
             mCurrY = .CurrentY + .ScaleX("0.5in", .ScaleMode, .ScaleMode)
            .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
            .TextBox s, .CurrentX, mCurrY, "1.8in", "0.8in", taJustifyTop, True
            .CalcTextBox
             Debug.Print .X1, .Y1, .X2 - .X1, .Y2 - .Y1
            
            .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
            .TextBox s, .X1 + .ScaleX("1in", .ScaleMode, .ScaleMode) * 2, .Y1, "1.8in", "0.8in", taJustifyMiddle, True
            .CalcTextBox
             Debug.Print .X1, .Y1, .X2 - .X1, .Y2 - .Y1
            
            .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
            .TextBox s, .CurrentX + .ScaleX("1in", .ScaleMode, .ScaleMode) * 4, mCurrY, "1.8in", "0.8in", taJustifyBottom, True
            .CalcTextBox
             Debug.Print .X1, .Y1, .X2 - .X1, .Y2 - .Y1
            
            ' restore defaults
            .TextAlign = i
        .EndDoc
    End With

End Sub

Private Sub DoDrawPicture()
     Dim s As String, Body As String
     
     With VBPrintPreview1
         .Clear
         .DocName = "DemoPicture"
         .StartDoc
         .Zoom = zmWholePage
         SetPages "PrintPreview|DrawPicture"
        .TextAlign = taJustifyTop
        
        .StartDoc
        
            SetTitle "DrawPicture function"
            .Paragraph
            .Paragraph "Draws a picture." + vbCr
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .Paragraph "[form.]VBPPreview.DrawPicture Picture As StdPicture, _"
            .IndentLeft = .TextWidth("[form.]VBPPreview.DrawPicture ")
            .Paragraph " Left As Variant, Top As Variant, _"
            .Paragraph "[ Width As Variant ], [ Height As Variant ], _"
            .Paragraph "[ Opcode As RasterOpConstants ]"
            .IndentLeft = 0
            .FontBold = False
            
            Body = "Picture|Picture to draw. The StdPicture object enables you to manipulate bitmaps, icons, metafiles enhanced metafiles, GIF, and JPEG images assigned to objects having a Picture property.;" + _
                   "Left,Top|Position of the top left corner of the picture on the page. You may specify the units for these parameters. See ''Using Unit Property''.;" + _
                   "Width,Height|The width and height of the picture box, in explicit units. See ''Using Unit Property''.;" + _
                   "Opcode|Optional. Long value or code that is used only with bitmaps. It defines a bit-wise operation " + _
                   "(such as vbMergeCopy or vbSrcAnd) that is performed on picture as it's drawn on object. For a complete " + _
                   "list of bit-wise operator constants, see the RasterOp Constants topic in Visual Basic Help. " + vbCr + _
                   "There are some limitations in the usage of opcodes.  For example, you can't use any opcode other than " + _
                   "vbSrcCopy if the source is an icon or metafile, and the opcodes that interact with the pattern (or ''brush'' " + _
                   "in SDK terms) such as MERGECOPY, PATCOPY, PATPAINT, and PATINVERT actually interact with the FillStyle property " + _
                   "of the destination." + vbCr + _
                   "Note: Opcode is used to pass a bitwise operation on a bitmap. Placing a value in this argument when passing " + _
                   "other image types will cause an ''Invalid procedure call or argument'' error. This is by design. To avoid this error," + _
                   " leave the Opcode argument blank for any image other than a bitmap.;"
            .StartTable
            .Table "3cm|15cm", "|", Body, , , vbWhite
            .EndTable
        
            .Paragraph
            .Paragraph "For details on using unit-aware measurements, see the 'Using Unit Properties' topic."
            .Paragraph "See example 'Pictures' topic."
            .Paragraph
         
         .EndDoc

         .HdrFontTransparent = True
        
     End With
     
End Sub

Sub DoExport()
    Dim FullPath As String, Drive As String, Path As String, FileName As String, File As String, Extension As String
    Dim i As Integer, savedata As Boolean
    
        'random print
        Select Case Int((6 * Rnd) + 1)
        Case 1: DoParagraphs True
        Case 2: DoMargins True
        Case 3: DoAdoData
        Case 4: DoTableFormat
        Case 5: DoPicture
        Case 6: DoCalendar
        End Select
         
         With VBPrintPreview1
              
              FullPath = .ShowSaveFile(, , App.Path + "\", .DocName)
              If FullPath <> "" Then
                 SplitPath FullPath, Drive, Path, FileName, File, Extension
                 Debug.Print FullPath, Drive, Path, FileName, File, Extension
                 For i = 1 To .PageCount
                    If .SetDataClipboard(i) = True Then
                       savedata = True
                       PicLoad.Picture = Clipboard.GetData(vbCFBitmap)
                       SavePicture PicLoad, Path + "\" + File + Trim(Str(i)) + "." + Extension
                       PicLoad.Picture = LoadPicture()
                    End If
                 Next
               End If
         End With
         If savedata Then
            MsgBox "Export Complete.", vbInformation
         Else
            MsgBox "Export export is not completed.", vbCritical
         End If
End Sub

Sub DoPicture()
     PicLoad.Picture = LoadPicture(App.Path + "\Library\Angel.jpg")
     PicLoad.AutoSize = True
     
     With VBPrintPreview1
         .Clear
         .StartDoc
         .Zoom = zmWholePage
         
        .StartDoc
            SetPages "PrintPreview|Draw Pictures"
            .TextAlign = taJustifyTop
            .Paragraph
            .TextAlign = taCenterTop
            .FontBold = True
            .FontSize = 10
            .TextAlign = taLeftTop
            
            .Paragraph "Size 10%"
            '.DrawPicture Pic2, 0.5, .CurrentY, "50%", "50%"  'or
            .DrawPicture PicLoad, .MarginLeft, .CurrentY, "10%", "10%"
            .CalcPicture
            .CurrentY = .Y1 - .TextHeight
            .CurrentX = .X2 + .ScaleX("10px", .ScaleMode, .ScaleMode)

            .FillStyle = vbFSTransparent
           .TextBox "Size 25%", .CurrentX, .CurrentY, "1in", 0, taLeftTop, False
           .DrawPicture PicLoad, .CurrentX, .CurrentY, "25%", "25%"

           'Get Position
           .CalcPicture
           .CurrentY = .Y1 - .TextHeight '.Y1 + .Y2
           .CurrentX = .X2 + .ScaleX("10px", .ScaleMode, .ScaleMode)
           .TextBox "Size 50%", .CurrentX, .CurrentY, "1in", 0, taLeftTop, False
            'Pic2.AutoSize = True
            .DrawPicture PicLoad, .CurrentX, .CurrentY, "50%", "50%"
            .PageBorder = pbNone
           .Header = ""
         .NewPage
            .DrawPicture PicLoad, "1cm", "1cm", "100%", "100%"
            .FontTransparent = False
            .Paragraph "Picture Size 100%"
            .HdrFontTransparent = False
         .NewPage
            .DrawPicture PicLoad, 0, 0, .PageWidth, .PageHeight
            .FontTransparent = False
            .ForeColor = vbRed
            .Paragraph " Full Page "
            .FontTransparent = True
         .EndDoc

         .HdrFontTransparent = True
        
     End With
     PicLoad.Picture = LoadPicture()
End Sub


Private Sub DoCalendarYear()
        
    Dim R%, C%, TheDate As Date, i As Integer, PxW As Single, PxH As Single, w As String, D As Integer, X As Integer, Y As Integer
    Dim TopCurrent As Single
    ReDim Cal(7, 1) As Variant
    
    Dim Mnt(1 To 12) As String
    Mnt(1) = "January"
    Mnt(2) = "February"
    Mnt(3) = "March"
    Mnt(4) = "April"
    Mnt(5) = "May"
    Mnt(6) = "June"
    Mnt(7) = "July"
    Mnt(8) = "August"
    Mnt(9) = "September"
    Mnt(10) = "October"
    Mnt(11) = "November"
    Mnt(12) = "December"
         
     With VBPrintPreview1
          .Clear
          'start document
          .PageBorder = pbNone
          .FillColor = vbCyan
          .FontName = "Tahoma"
          .FontItalic = True
          .MarginBottom = 0
          .MarginLeft = 2
          .StartDoc
                
            'set zoom mode
            .Zoom = zmWholePage
            
            .CurrentY = 2
            .FontSize = 28
            .FontBold = True
            .FontName = "Times New Roman"
            
             'load picture
             PicLoad.Picture = LoadPicture(App.Path + "\Library\" + Format(Rnd * 11 + 1, "00") + ".jpg")
             '.Draw Picture
             .DrawPicture PicLoad, "0.8in", "0.8in", (.PageWidth - .MarginLeft - .MarginRight) - 0.6, "70%"
             .CalcPicture
            .DrawRectangle "0.8in", "0.8in", (.PageWidth - .MarginLeft - .MarginRight) - 0.6, .Y2 - .Y1, , , vbBlue, , vbFSTransparent
            
            'write month-year
           .TextBox Trim(Str(Year(Now))), .MarginLeft, 10, (.PageWidth - .MarginLeft - .MarginRight) - 0.6, 0, taCenterMiddle, True, True
            .FontBold = False
            X = 1
            Y = 1
            .FontName = "Tahoma"
            TopCurrent = 8
            .TextAlign = taLeftTop
            For i = 1 To 12
                                
                 'build calendar
                 TheDate = "1/" + Trim(Str(i)) + "/" + Trim(Str(Year(Now)))
                 BuildCalendar TheDate, Cal, R, C
                  
                .FontBold = True
                .FontSize = 10
                
                .CurrentY = TopCurrent + 4 * Y
                '.TextAlign = taLeftTop
                .MarginLeft = (4.8 + 1) * (X - 1) + 2
                
                .FontBold = False
                
                 .StartTable
                 
                  'set header and bind to calendar array
                   w = .TextWidth("Wed")
                   If InStr(1, w, ",") Then w = Replace(w, ",", ".")
                  .TableArray "" + w + "|" + w + "|" + w + "|" + w + "|" + w + "|" + w + "|" + w + "", "Sun|Mon|Tue|Wed|Thu|Fri|Sat", Cal
                                        
                  'header align cells to center
                  .TableCell tcTextAling, , , taCenterMiddle
                  
                  'set body font
                  .TableCell tcFontName, , , "Times New Roman"
                  '.TableCell tcFontSize, , , 24
                  .TableCell tcFontItalic, , , True
                                       
                  'set Sunday font-back color
                  .TableCell tcBackColor, , 1, &HC0E0FF
                  .TableCell tcForeColor, , 1, vbRed
                  
                  'set RowHeight for dates
                  For D = 2 To .TableCell(tcRows)
                     .TableCell tcRowHeight, D, , .TextHeight
                  Next
                  
                  'format header row (week days)
                  .TableCell tcFontBold, 1, , True
                  .TableCell tcFontItalic, 1, , False
                  .TableCell tcBackColor, 1, , &HC0FFFF
                  
                  'render the table
                  .EndTable
                  
                  .CalcTable
                  .FontBold = True
                  .TextBox Mnt(i) + " " + Trim(Str(Year(Now))), .X1, .Y1 - .TextHeight, .X2 - .X1, 0, taCenterTop
                  .FontBold = False
                 If X = 3 Then
                   X = 1: Y = Y + 1
                 Else
                  X = X + 1
                 End If
            Next
          'finish the document
          .EndDoc
     End With
    
End Sub

Private Sub DoCalendar()
       
    Dim R%, C%, TheDate As Date, i As Integer, PxW As Single, PxH As Single, w As String, D As Integer
    ReDim Cal(7, 1) As Variant
    
    Dim Mnt(1 To 12) As String
    Mnt(1) = "January"
    Mnt(2) = "February"
    Mnt(3) = "March"
    Mnt(4) = "April"
    Mnt(5) = "May"
    Mnt(6) = "June"
    Mnt(7) = "July"
    Mnt(8) = "August"
    Mnt(9) = "September"
    Mnt(10) = "October"
    Mnt(11) = "November"
    Mnt(12) = "December"

     With VBPrintPreview1
            .Clear
            ' start document
            .PageBorder = pbNone
            .FillColor = vbCyan
            .FontName = "Tahoma"
            .FontItalic = True
            .MarginBottom = 0
            
            .StartDoc
                
            ' set zoom mode to Thumbnail
            .Zoom = zmThumbnail
           
            For i = 1 To 12
                 .TextAlign = taLeftTop
                 ' build calendar for the given date  return TheDate's row and column in r, c
                 TheDate = "1/" + Trim(Str(i)) + "/" + Trim(Str(Year(Now)))
                 BuildCalendar TheDate, Cal, R, C
                  
                .FontBold = True
                .FontSize = 30
                .CurrentY = "5in"
                
                 'load picture
                 PicLoad.Picture = LoadPicture(App.Path + "\Library\" + Format(i, "00") + ".jpg")
                 '.Draw Picture
                 .DrawPicture PicLoad, "0.8in", "0.8in", .PageWidth - .ScaleX("1.6in", .ScaleMode, .ScaleMode), "100%"
                 
                 'write month-year
                 .TextBox Mnt(i) + " " + Trim(Str(Year(Now))), .MarginLeft, .CurrentY, (.PageWidth - .MarginLeft - .MarginRight), 0, taCenterTop
                 
                .FontBold = False
                
                 .StartTable
                 
                 ' set header and bind to calendar array (use dummy column widths for now)
                  w = Format$((.PageWidth - .MarginLeft - .MarginRight) / 7, "0.0000")
                  If InStr(1, w, ",") Then w = Replace(w, ",", ".")
                  .TableArray "" + w + "|" + w + "|" + w + "|" + w + "|" + w + "|" + w + "|" + w + "", "Sun|Mon|Tue|Wed|Thu|Fri|Sat", Cal
                                        
                  'header align cells to center
                  .TableCell tcTextAling, , , taCenterMiddle
                  
                  ' set body font
                  .TableCell tcFontName, , , "Times New Roman"
                  .TableCell tcFontSize, , , 24
                  .TableCell tcFontItalic, , , True
                                       
                  'set Sunday font-back color
                  .TableCell tcBackColor, , 1, &HC0E0FF
                  .TableCell tcForeColor, , 1, vbRed
                  
                  'set RowHeight for dates
                  For D = 2 To .TableCell(tcRows)
                     .TableCell tcRowHeight, D, , "0.6in"
                  Next
                  
                  ' format header row (week days)
                  .TableCell tcFontBold, 1, , True
                  .TableCell tcFontItalic, 1, , False
                  .TableCell tcBackColor, 1, , &HC0FFFF
                  
'                  ' highlight today (note 1-based indices)
'                  .TableCell tcFontBold, R + 1, C + 1, True
'                  .TableCell tcBackColor, R + 1, C + 1, &HC0FFC0
                  
                  ' done, render the table
                  .EndTable
                  
                 If i <> 12 Then .NewPage
            Next
                ' done, finish the document
                .EndDoc
     End With
    
End Sub

Private Sub BuildCalendar(TheDay As Date, ByRef TheCal(), ByRef TheDayRow, ByRef TheDayCol)
    Dim dt As Date
    ' clear array
    ReDim TheCal(7, 1)
    ' initialize date to the first of the month
    dt = TheDay
    While Day(dt) > 1
        dt = dt - 1
    Wend
    ' fill array with dates for current month
    Dim R%, C%
    R = 0
    C = Weekday(dt) - 1
    While Month(dt) = Month(TheDay)
        ' add row if we have to
        If C >= 7 Then
            C = 0
            R = R + 1
            ReDim Preserve TheCal(7, R)
        End If
        ' save day value in the calendar
        TheCal(C, R) = Day(dt)
        ' return TheDate's row and column
        If dt = TheDay Then
            TheDayRow = R
            TheDayCol = C
        End If
        ' increment day
        dt = dt + 1
        C = C + 1
    Wend
End Sub

Private Sub DoDocument()
       
       With VBPrintPreview1
        .Clear
        .PageBorder = pbNone
        
        .Zoom = zmWholePage
        .BackColorPage = &HFFFFF0
        .StartDoc
        ''Rectangle Shadow
        .DrawRectangle 2.1, 2.1, .PageWidth - 4, .PageHeight - 4, 3, 3, &H808080, &H808080, vbFSSolid
        ''Rectangle white
        .DrawRectangle 2, 2, .PageWidth - 4.1, .PageHeight - 4.1, 3, 3, vbBlack, &HFFFFF0, vbFSSolid
        
        .FillColor = RGB(255, 126, 64)
        .FontSize = 20
        .FontBold = True
        
        .TextBox "VB Print Preview", "5cm", "6cm", "11.5cm", "2cm", taCenterMiddle, True, True
        .FontSize = 30
        .ForeColor = vbBlue
        .TextAlign = taCenterTop
        .Paragraph
        .Paragraph "ActiveX Control"
        .FontSize = 20
        .Paragraph "Version 1.1"
        .Paragraph
        .FontSize = 12
        .Paragraph "Print Preview more powerful than ever!"
        .Paragraph ""
        .ForeColor = 0
        .TextAlign = taJustifyTop
        
        'Print Preview; more; powerful; than; ever!
         'VbPrintPreview makes it easy to add robust View, Format, Export, and Print capabilities to your
         'Visual Basic applications. It replaces the Printer Object in Visual Basic to significantly expand,
         'yet simplify, printing functionality.
         'With VbPrintPreview, you can create your output quickly and easily; print paragraphs and text that
         'automatically wrap; or easily combine paragraphs, graphs, pictures, tables anywhere on the page.
         'A styles property separates format information from content generation, making it easy to create
         'documents with consistent formatting. You have full control of the formatting of paragraphs,tables,
         'fonts, colors, alignment, and justification.
         .CalcParagraph
         
        .TextBox vbTab + "The VB Print Preview control makes it easy to create documents and reports for printing and " + _
                "print previewing from your applications. It only takes one statement to print plain text, " + _
                "and a little more work to print graphics, tables, and formatted text. You have complete " + _
                "control over the printing device and document layout, including paper size and orientation, " + _
                "number of columns, headers and footers, page borders, shading, fonts, etc." + vbCr + vbCr + _
                vbTab + "The control has a rich set of properties, methods, and events." + vbCr, "5cm", .Y2, "12cm", 0, taJustifyTop, False
        .EndDoc
        .BackColorPage = &HFFFFFF
       End With
End Sub

Private Sub DoPageTitle(Title As String)
       
       With VBPrintPreview1
        .Clear
        .PageBorder = pbNone
        
        .Zoom = zmWholePage
        .BackColorPage = &HFFFFF0
        .StartDoc
        'Rectangle Shadow
        .DrawRectangle 2.1, 2.1, .PageWidth - 4, .PageHeight - 4, 3, 3, &H808080, &H808080, vbFSSolid
        'Rectangle white
        .DrawRectangle 2, 2, .PageWidth - 4.1, .PageHeight - 4.1, 3, 3, vbBlack, &HFFFFF0, vbFSSolid
        
        .FillColor = RGB(255, 126, 64)
        .FontSize = 20
        .FontBold = True

        .TextBox "VB Print Preview", "5cm", "6cm", "11.5cm", "2cm", taCenterMiddle, True, True
         .FontName = "Times New Roman"
        .FontSize = 30
        .ForeColor = vbBlue
        .TextAlign = taCenterTop
        .Paragraph
        .Paragraph "ActiveX Control"
        .FontSize = 20
        .Paragraph "Version 1.1"
        .Paragraph " "
        .FontSize = 30
        
        .Paragraph " "
        .Paragraph " "
        .FontBold = True
        .TextAlign = taCenterTop
        .Paragraph
        .ForeColor = vbBlue
       
        .FontItalic = True
        .FontUnderline = True
        .Paragraph Title
        .FontUnderline = False
        .FontItalic = False
        .Paragraph
        .FontSize = 12
        .FontBold = False
        .EndDoc
        .BackColorPage = &HFFFFFF
       End With
End Sub

Private Sub DoLabels()
    
    Dim db As New ADODB.Connection
    Dim rs As New Recordset
    Dim FormatCol$, Header$
    Dim Body As String, arr() As String, fc As Long
    Dim i As Integer, Y As Long, X As Long, ty As Long, ViewLine As Boolean
    
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Data\Nwind.mdb;Persist Security Info=False"
    db.Open
    
    rs.Open "SELECT Suppliers.CompanyName, Suppliers.ContactTitle, Suppliers.Address, Suppliers.Region, Suppliers.PostalCode FROM Suppliers;", db, adOpenKeyset, adLockOptimistic
    If rs.BOF = False Then
    If rs.RecordCount > 0 Then
       Body = ""
       rs.MoveFirst
       For i = 1 To rs.RecordCount
          Body = Body + rs.Fields("CompanyName") + vbCr + rs.Fields("ContactTitle") + vbCr + rs.Fields("Address") + vbCr + rs.Fields("PostalCode") + "|"
          rs.MoveNext
       Next
    End If
    End If
            
    rs.Close
    Body = Mid(Body, 1, Len(Body) - 1)
    
    arr = Split(Body, "|")
    
    With VBPrintPreview1
        'If .SendToPrinter = False Then ViewLine = True Else ViewLine = False
        .Clear
        .Orientation = PagePortrait
        .PageBorder = pbNone
        .Zoom = zmWholePage
        .MarginBottom = 0
        .MarginFooter = 0
        .MarginLeft = 0
        .MarginRight = 0
        .MarginHeader = 0
            
        .StartDoc
            fc = .FillColor
            .FillColor = RGB(255, 255, 255)
            .FontSize = 10
            .FontBold = True
            '.DrawStyle = vbDashDotDot ' vbDot
            '.DrawMode = vbCopyPen
            '.FillStyle = 1
            .DrawWidth = 1
            i = 0
            For ty = 0 To UBound(arr) \ 3 \ 7
                For Y = 0 To 7
                    For X = 0 To 2
                       If i <= UBound(arr) Then
                        .TextBox arr(i), X * 7, Y * 3.7, "7cm", "3.7cm", taCenterMiddle, False 'ViewLine
                        i = i + 1
                       Else
                          'If ViewLine Then .DrawRectangle X * 7, Y * 3.7, "7cm", "3.7cm"
                       End If
                    Next
                    If .CurrentY + 3.7 > .PageHeight And i < UBound(arr) Then
                        .NewPage
                    End If
                Next
            Next
            .FillColor = fc
        .EndDoc
     End With
     
End Sub

Private Sub DoAdoData()

    Dim db As New ADODB.Connection
    Dim rs As New Recordset
    Dim FormatCol$, Header$
    Dim Body() As Variant
    Dim i As Integer
    
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Data\Nwind.mdb;Persist Security Info=False"
    db.Open
    
    rs.Open "SELECT productid,productname,quantityperunit,unitprice FROM Products where productid>0", db, adOpenKeyset, adLockOptimistic
     If rs.RecordCount > 0 Then
         Body = rs.GetRows(rs.RecordCount)
     End If
    rs.Close
    
    With VBPrintPreview1
        .Clear
        .Orientation = PagePortrait
        
        .Zoom = zmWholePage
       
        SetPages "PrintPreview|TableArray"
        
        .StartDoc
        SetSubTitle "TableArray with ADO"
        .Paragraph
        .Paragraph "This example uses TableArray with no wrapping of the words and TableBorder = tbBoxRows."
        .Paragraph
        .ForeColor = vbBlack
        .FontSize = .FontSize + 2
        .TextAlign = taLeftTop
        
        .StartTable
            .TableBorder = tbBoxRows
            .FontSize = 12
            FormatCol$ = "^20mm|<60mm|<50mm|>30mm;"
            Header$ = "Id|Product Name|Quanti Type Unit|UnitPrice|"
            .TableArray FormatCol$, Header$, Body, 1, &HE0E0E0, , , False
            .TableCell tcFontBold, 1, , True  'Row 1 set FontBold = True
            .TableCell tcFontBold, , 1, True  'Col 1 set FontBold = True
            For i = 2 To .TableCell(tcRows) Step 2
                .TableCell tcBackColor, i, , &HFFFFFF  'Shade color line by line
            Next
            .TableCell tcForeColor, 1, , &HFFFFFF     'Header color
        .EndTable

        .NewPage
        SetSubTitle "TableArray with ADO"
        .Paragraph
        .Paragraph "This example uses TableArray to wrap the words and TableBorder = tbAll."
        .ForeColor = vbBlack
        .Paragraph
        .FontSize = .FontSize + 2
        .TextAlign = taLeftTop
        
        .StartTable
            .TableBorder = tbAll
            .FontSize = 12
            FormatCol$ = "^20mm|<60mm|<50mm|>30mm;"
            Header$ = "Id|Product Name|Quanti Type Unit|UnitPrice"
            .TableArray FormatCol$, Header$, Body, 1
            .TableCell tcFontBold, 1, , True 'Row 1 set FontBold = True
            .TableCell tcFontBold, , 1, True 'Col 1 set FontBold = True
            .TableCell tcForeColor, 1, , &HFFFFFF 'Header color
            .TableCell tcTextAling, , 1, taCenterMiddle
            .TableCell tcTextAling, , 2, taLeftMiddle
            .TableCell tcTextAling, , 3, taLeftMiddle
            .TableCell tcTextAling, , 4, taRightMiddle
        .EndTable
        
        .CalcTable
        Debug.Print "Position Table", .X1, .Y1, .X2, .Y2
        
        .EndDoc
     End With
     
End Sub


Private Sub DoTableFormat()

    Dim db As New ADODB.Connection
    Dim rs As New Recordset
    Dim FormatCol$, Header$
    Dim Body() As Variant, Body2() As Variant, Totals As String, Maxrecord As Integer
    Dim Pic() As Byte, sData As String, R As Integer, C As Integer
    
    Dim i As Long, p As Integer
    
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Data\Nwind.mdb;Persist Security Info=False"
    db.Open
    
    rs.Open " SELECT c.CompanyName, p.ProductName AS pName, p.ProductID AS pID, od.Quantity AS quantity" + _
            " FROM orders AS o, [order details] AS od, products AS p, customers AS c " + _
            " WHERE (((od.ProductID)=[p].[ProductID]) AND ((od.OrderID)=[o].[OrderID]) AND ((o.CustomerID)=([c].[CustomerID])) AND ((c.CustomerID)='ANTON'));", db, adOpenKeyset, adLockOptimistic
    If rs.BOF = False Then
    If rs.RecordCount > 0 Then
        Maxrecord = rs.RecordCount
        Body = rs.GetRows(rs.RecordCount)
    End If
    End If
    rs.Close
    
    rs.Open " SELECT '', '' AS pName, '' AS pID, Sum(od.Quantity) AS quantity" + _
            " FROM orders AS o, [order details] AS od, products AS p, customers AS c " + _
            " WHERE (((od.ProductID)=[p].[ProductID]) AND ((od.OrderID)=[o].[OrderID]) AND ((o.CustomerID)=([c].[CustomerID])) AND ((c.CustomerID)='ANTON'));", db, adOpenKeyset, adLockOptimistic
    If rs.BOF = False Then
    If rs.RecordCount > 0 Then
           Totals = rs.Fields("Quantity")
    End If
    End If
    rs.Close
    
    
    rs.Open " SELECT CategoryId,CategoryName,Description,'' FROM categories order by 1", db, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
              Body2 = rs.GetRows(rs.RecordCount)
        End If
    rs.Close
    
    With VBPrintPreview1
        .Clear
       .Zoom = zmWholePage
        SetPages "PrintPreview|TableArray"
                
        .StartDoc
            .FontSize = 14
             SetSubTitle "Merging Rows with TableArray and ADO"
            .FontSize = 12
            .ForeColor = vbBlack
            .TextAlign = taJustifyTop
            .Paragraph
            .Paragraph vbTab + "The following example is generated TableArray table and recordset." + _
                       " The table is created by cutting of the words declaring the variable Wrap = True." + _
                       " The configuration of cells done with tcColSpan, tcRowSpan, tcTextAling, tcRows, tcFontBold and tcText."
            .Paragraph
            .FontSize = .FontSize + 2
            .TextAlign = taLeftTop
            
            'Example Table 1
            .StartTable
                .TableBorder = tbAll
                .FontSize = 12
                FormatCol$ = "<6cm|<5cm|>3cm|>2cm"
                Header$ = "CompanyName|Product Name|Product ID|Quantity"
                .TableArray FormatCol$, Header$, Body, vbYellow, , , , False
                .TableCell tcFontBold, 1, , True
                .TableCell tcRowSpan, 2, 1, Maxrecord
                Maxrecord = Maxrecord + 1
                .TableCell tcRows, , , Maxrecord
                Maxrecord = Maxrecord + 1
                .TableCell tcTextAling, , 1, taCenterMiddle
                .TableCell tcColSpan, Maxrecord, 1, 3
                .TableCell tcText, Maxrecord, 1, "Totals"
                .TableCell tcText, Maxrecord, 4, Totals
                .TableCell tcTextAling, Maxrecord, 1, taRightTop
                .TableCell tcTextAling, Maxrecord, 4, taRightTop
                .TableCell tcFontBold, Maxrecord, , True
            .EndTable
            .CalcTable
            Debug.Print "Position Table:", .X1, .Y1, .X2, .Y2
            .FontSize = 10
            .FontItalic = True
            .Paragraph "Position Table - X1:" + Format(.X1, "0.00") + " - Y1:" + Format(.Y1, "0.00") + _
                                     " - X2:" + Format(.X2, "0.00") + " - Y2:" + Format(.Y2, "0.00")
            .FontSize = 12
            .FontItalic = False
            '---------------------------
            .Paragraph
            .FontSize = 16
            SetSubTitle "Merging Columns with TableCell"
            .FontSize = 12
            .Paragraph
            
            .FontBold = False
            .Paragraph
            .Paragraph "The following table make is span two columns. "
            
            'Example Table 2
            .StartTable
                ' create table with four rows and four columns
                .TableCell tcCols, , , 4
                .TableCell tcRows, , , 4
            
                ' set some column widths (default width is 0.5in)
                .TableCell tcColWidth, , 1, "1in"
                .TableCell tcColWidth, , 2, "1.3in"
            
                'set font size
                .TableCell tcFontSize, , 1, 10
                .TableCell tcFontSize, , 2, 10
                .TableCell tcFontSize, , 3, 10
                .TableCell tcFontSize, , 4, 10
            
                .TableCell tcIndent, , , "1mm"
            
                ' assign text to each cell
                For R = 1 To 4
                    For C = 1 To 4
                        .TableCell tcText, R, C, "Row " & Str(R) & " Col " & Str(C)
                    Next C
                Next R
            
                ' format cell (1,1): make it span two columns, with a blue  background, center alignment, and bold
                .TableCell tcColSpan, 1, 1, 2
                .TableCell tcBackColor, 1, 1, vbCyan
                .TableCell tcTextAling, 1, 1, taCenterMiddle
                .TableCell tcFontBold, 1, 1, True
            
                ' set row height for row 1 (default height is calculated to fit the contents)
                .TableCell tcRowHeight, 1, , "0.2in"
                  
                ' format cell (3,2): make is span two columns, with a yellow background, center alignment, and bold
                .TableCell tcColSpan, 3, 2, 2
                .TableCell tcBackColor, 3, 2, vbYellow
                .TableCell tcTextAling, 3, 2, taCenterTop
                .TableCell tcFontBold, 3, 2, True
    
                ' set row height for row 3
                .TableCell tcRowHeight, 3, "0.2in"
      
                ' set row borders all around
                .TableBorder = tbAll
            .EndTable
            
           'measure tables
           .CalcTable
           .FontSize = 10
           .FontItalic = True
           .Paragraph "Position Table - X1:" + Format(.X1, "0.00") + " - Y1:" + Format(.Y1, "0.00") + _
                                    " - X2:" + Format(.X2, "0.00") + " - Y2:" + Format(.Y2, "0.00")
           .FontSize = 12
           .FontItalic = False
           
          '--------------------------------------
          .NewPage
            .FontSize = 14
             SetSubTitle "TableArray and ADO with Pictures"
            .FontSize = 12
            .ForeColor = vbBlack
            .TextAlign = taJustifyTop
            .Paragraph
            .Paragraph vbTab + "The following table reads data from the database, formats it and displays the table. The example includes images." + _
                       " The configuration of images cells done with tcPicture, tcColWidth and tcRowHeight."
            .Paragraph
            
        'Example Table 3
        .StartTable
           'set data in table
            FormatCol$ = "^3cm|<30mm|<50mm|<47mm"
            Header$ = "CategoryId|CategoryName|Description|Picture"
            .TableArray FormatCol$, Header$, Body2, vbYellow ', , , , ,  'cm
        
            'read picture from database and insert to cells
            rs.Open " SELECT Picture FROM categories order by CategoryId", db, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                For p = 1 To rs.RecordCount
                    Dim ff As Integer
                    ff = FreeFile
                    Open App.Path + "\00" + Trim(Str(p)) + ".bmp" For Binary As #ff
                        Pic = rs!Picture.Value
                        Put #ff, , Pic
                    Close ff
                    PicLoad.Picture = LoadPicture(App.Path + "\00" + Trim(Str(p)) + ".bmp")
                    .TableCell tcTextAling, , 1, taCenterMiddle
                    .TableCell tcTextAling, , 2, taLeftMiddle
                    .TableCell tcTextAling, , 3, taLeftMiddle
                    .TableCell tcPicture, p + 1, 4, PicLoad
                    .TableCell tcColWidth, p + 1, 4, .ScaleX(PicLoad.Width, PicLoad.ScaleMode, .ScaleMode)
                    .TableCell tcRowHeight, p + 1, 4, .ScaleY(PicLoad.Height, PicLoad.ScaleMode, .ScaleMode)
                    Kill App.Path + "\00" + Trim(Str(p)) + ".bmp"
                    rs.MoveNext
                Next
            End If
            rs.Close
            .TableCell tcFontBold, 1, , True
        .EndTable
        
            'measure tables
            .CalcTable
            Debug.Print "Position Table2", .X1, .Y1, .X2, .Y2
            .FontSize = .FontSize - 2
            .FontItalic = True
            .Paragraph "Position Table - X1:" + Format(.X1, "0.00") + " - Y1:" + Format(.Y1, "0.00") + _
                                     " - X2:" + Format(.X2, "0.00") + " - Y2:" + Format(.Y2, "0.00")
            .FontSize = 12
            .FontItalic = False
        .EndDoc
     End With
     
End Sub

Function SaveBinaryData(FileName, ByteArray) As Boolean
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream As Object
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  If IsNull(ByteArray) = False Then
     BinaryStream.Open
     BinaryStream.Write ByteArray
     SaveBinaryData = True
  Else
    SaveBinaryData = False
    Exit Function
  End If
  'Save binary data To disk
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Public Function LoadPictureDB(rs As ADODB.Recordset) As Boolean

    On Error GoTo procNoPicture
    Dim strStream As Object
    
    'If Recordset is Empty, Then Exit
    If rs Is Nothing Then
        GoTo procNoPicture
    End If
    
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    strStream.Write rs.Fields("Picture").Value
    strStream.SaveToFile App.Path + "\Temp.bmp", adSaveCreateOverWrite
    PicLoad.Picture = LoadPicture(App.Path + "\Temp.bmp")
    PicLoad.AutoSize = True
    Kill (App.Path + "\Temp.bmp")
    LoadPictureDB = True

procExitFunction:
    Exit Function
    
procNoPicture:
    LoadPictureDB = False
    GoTo procExitFunction
    
End Function

Private Sub DoTableBorder()
    
       With VBPrintPreview1
        .Clear
        .Zoom = zmWholePage
        SetPages "PrintPreview|TableBorder"
        
        .StartDoc
         SetTitle "TableBorder Property"
         
        .ForeColor = vbBlack
        .Paragraph
        .FontSize = 12
        
        .Paragraph "Returns or sets the type of border for tables"
        .Paragraph
        .FontBold = True
        .Paragraph "Syntax:"
        .Paragraph
        .Paragraph "[form.]VBPPreview.TableBorder [= TableBorderConstants]"
        .FontBold = False
        .Paragraph
        .Paragraph vbTab + "The border is drawn using the DrawWidth, ForeColor properties. " + _
                   "You may override the pen width using the TablePen, TablePenLR, and TablePenTB properties."
        .Paragraph
        .Paragraph vbTab + "The picture below shows valid settings for the TableBorder property and their effect:"
        
        '.IndentLeft = 1
        '.Paragraph "tbNone"
        '.Paragraph "tbBottom "
        '.Paragraph "tbTop "
        '.Paragraph "tbTopBottom "
        '.Paragraph "tbBox "
        '.Paragraph "tbColums "
        '.Paragraph "tbColTopBottom "
        '.Paragraph "tbAll "
        '.Paragraph "tbBoxRows "
        '.Paragraph "tbBoxColumns "
        '.Paragraph "tbBelowHeader "
        PicLoad.Picture = LoadPicture(App.Path + "\Library\Tableborder.bmp")
        .DrawPicture PicLoad, .MarginLeft, .CurrentY  ', "100%", "100%"
        .EndDoc
     End With
     
End Sub


Private Sub DoMeasureText()

Dim s$, fmt$, hdr$, bdy$, mm$

    's = "This chapter deals with the problems inherent to youths in their early twenties."
    's = s & s & s
    s = vbTab + "Manage SQL Server 2008 R2 Express databases with SQL Server Management Studio Express." + _
        " Connect to local SQL Server 2008 R2 Express databases and manage objects with full " + _
        "Object Explorer integration. Write, execute, and test queries by using visual query " + _
        "plans that provide hints to tune queries and access management and maintenance options."
        
    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .Zoom = zmWholePage
        SetPages "PrintPreview|Measuring Text"
      
        .StartDoc
        
        ' introduction
        .FontSize = 14
        SetSubTitle "Measuring Text"
        .FontSize = 14
        .Paragraph "To measure the width, height of a specific " + _
                   "string you can use one of the following properties:"
        
        fmt = "^3000tw|<5500tw;"
        hdr = "Property|Description;"
        bdy = "CalcPicture|Calculates the size of a Picture.;" & _
              "CalcParagraph|Calculates the size of a paragraph string.;" & _
              "CalcTable|Calculates the size of a table string.;" & _
              "CalcTextBox|Calculates the size of a TextBox string.;"
              
        .StartTable
            .Table fmt, hdr, bdy, 1, , False
            .TableCell tcForeColor, 1, , &HFFFFFF
        .EndTable
        
        .Paragraph "All the above properties return their results in the following " + _
                   "properties X1, Y1, X2 and Y2."

        ' show sample
        SetSubTitle "Sample"
        .FontSize = 14
        .Paragraph "As an example, we will measure the following paragraph.  " + _
                   "The same technique works for regular text and tables."
        .Paragraph
        .FontItalic = True
        .TextAlign = taJustifyTop
        .Paragraph s
        .FontItalic = False
        
        'if set TextAlign = taJustifyTop the table aling to pagemargin
        'change the taLeftTop with taJustifyTop
        '.TextAlign = taJustifyTop
        .TextAlign = taLeftTop
        
        ' measure text
        .CalcParagraph
        
        ' Print results
        Dim arr(1, 5)
        fmt = "^3cm|<5cm;"
        hdr = "Property|Value;"
        If .ScaleMode = smCentimeters Then mm$ = " cm" Else mm$ = " In"
        arr(0, 0) = "X1": arr(1, 0) = Format$(.X1, "0.00") & mm$
        arr(0, 1) = "X2": arr(1, 1) = Format$(.X2, "0.00") & mm$
        arr(0, 2) = "Y1": arr(1, 2) = Format$(.Y1, "0.00") & mm$
        arr(0, 3) = "Y2": arr(1, 3) = Format$(.Y2, "0.00") & mm$
        
        SetSubTitle "Results"
        .FontSize = 14
        .StartTable
            .TableArray fmt, hdr, arr, 1, , True, , , "2mm"
            .TableCell tcForeColor, 1, , &HFFFFFF
        .EndTable
        
        .EndDoc
    End With

End Sub


Private Sub DoMeasureUnits()

Dim s$, fmt$, hdr$, bdy$

    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .Zoom = zmWholePage
        SetPages "PrintPreview|Measurement Units"
        
         .StartDoc
        .PageBorder = tbNone
        
        .FontSize = 18
        SetSubTitle "Measurement Units"
        .FontSize = 18
        .TextAlign = taJustifyTop
        .Paragraph vbTab + "VbPrintPreview control has unit aware measurements.  " + _
                   "When setting a measurement value to a VBPrintPrinter property, now you can specify the unit type of that value." + vbCr
        .Paragraph "The following are the supported types:" + vbCr
        
        
        fmt = "^2500tw|^5500tw;"
        hdr = "Type| Description;"
        bdy = "in| Inches;" & _
              "tw| Twips;" & _
              "cm| Centimeters;" & _
              "mm| Milimeters;" & _
              "ch| Characters;" & _
              "pt| Point;" & _
              "px| Pixel;" & _
              "%| Percent;"

        .StartTable
            .Table fmt, hdr, bdy, 1
            .TableCell tcForeColor, 1, , &HFFFFFF
        .EndTable
        
        SetSubTitle "Samples"
        .FontSize = 18
        .FontName = "Courier New"
        .Paragraph " "
        .Paragraph "VBPrintPreview.MarginLeft =  ''2in''"
        .Paragraph "VBPrintPreview.MarginRigth = ''1in''"
        .Paragraph " "
        .Paragraph "VBPrintPreview.CurrentX =  ''3.2 cm''"
        .Paragraph "VBPrintPreview.CurrentY =  ''5 cm''"
       .EndDoc
    End With

End Sub

Private Sub DoTableArray()

    Dim db As New ADODB.Connection
    Dim rs As New Recordset
    Dim FormatCol$, Header$
    Dim Body() As Variant
    
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Data\Nwind.mdb;Persist Security Info=False"
    db.Open
    rs.Open "SELECT productid,productname,quantityperunit,unitprice FROM Products", db, adOpenKeyset, adLockOptimistic
       If rs.RecordCount > 0 Then
           Body = rs.GetRows(rs.RecordCount)
       End If
    rs.Close
    
    With VBPrintPreview1
        
        .Clear
        .Zoom = zmWholePage
         SetPages "PrintPreview|TableArray"
         
        .StartDoc
         SetTitle "TableArray function"
        
        .LineSpace = lsSpaceSingle
        .Paragraph
        .Paragraph "Renders a variant array as a table with row headers and special formatting."
        .Paragraph
        .FontBold = True
        .Paragraph "Syntax:"
        .Paragraph
        .FontSize = .FontSize - 2
        .Paragraph "[form.]VBPPreview.TableArray FormatCols As String, Header As String, Body As String,"
        .IndentLeft = .TextWidth("[form.]VBPPreview.TableArray ")
        .Paragraph "[HeaderShade As Long], [BodyShade As Long],"
        .Paragraph "[LineColor As Long], [LineWidth As Integer],"
        .Paragraph "[Wrap As Boolean], [Indent As Single],"
        .Paragraph "[WordWrap As Boolean], [Indent As Single]"
        .FontSize = .FontSize + 1
        .FontBold = False
        .IndentLeft = 0
        .LineSpace = lsSpaceSingle
        .TextAlign = taJustifyTop
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "FormatCols$"
        .FontBold = False: .FontItalic = False
        .Paragraph "This parameter contains formatting information and is not printed. " + _
                   "The formatting information describes each column using a sequence of " + _
                   "formatting characters followed by the column width. The information for " + _
                   "each column is delimited by the column separator character (by default a pipe (''|'')."
        .Paragraph
        .IndentFirst = "1cm"
        .Paragraph "For example, the following string defines a table with four center-aligned, " + _
                    "two-centimeter wide columns:"
        .Paragraph
        .Paragraph "s$ = ''^+2cm|^+2cm|^+2cm|^+2cm''"
        .Paragraph
        .Paragraph "The following lists shows all valid formatting characters:"
        
        .IndentFirst = 0
        .StartTable
          .Table "^0.8in|<4in;", "Character|Effect;", _
            "<|Align column contents to the left top;" + _
            ">|Align column contents to the right top;" + _
            "^|Align column contents to the center top;" + _
            "=|Align column contents to the justify top;" + _
            "<+|Align column contents to the left Middle;" + _
            ">+|Align column contents to the right Middle;" + _
            "^+|Align column contents to the center Middle;" + _
            "=+|Align column contents to the justify Middle;" + _
            "<_|Align column contents to the left Bottom;" + _
            ">_|Align column contents to the right Bottom;" + _
            "^_|Align column contents to the center Bottom;" + _
            "=_|Align column contents to the justify Bottom;", , , vbWhite, , False, "2mm"
          .TableCell tcFontBold, 1, , True
        .EndTable
        .IndentFirst = "1cm"
        
        .Paragraph
        .Paragraph "Column widths may be specified in twips, inches, points, millimeters, centimeters, pixel, or as a percentage of the width of age. If the units are not provided, Scalemode used. For details on using unit aware measurements, see Using Unit Properties."
        .IndentFirst = 0
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Header$"
        .FontBold = False: .FontItalic = False
        .Paragraph "This parameter contains the text to be printed on the first row of the table " + _
                   "and after each column or page break (the header row). The text for each cell " + _
                   "in the header row is delimited by the column separator character (by default a pipe (''|'')."
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Body$"
        .FontBold = False: .FontItalic = False
        .Paragraph "The Body parameter, which must be a two-dimensional array of Variants. The first array dimension" + _
                   " contains the rows, and the second dimension contains the columns."

        .Paragraph "You may also choose to supply data for individual cells separately. To do this, " + _
                    "use the TableCell property."
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "HeaderShade, BodyShade(optional)"
        .FontBold = False: .FontItalic = False
        .Paragraph "These parameters specify colors to be used for shading the cells in the header " + _
                   "and in the body of the table. If omitted or set to zero, the cells are not shaded. " + _
                   "If you want to use black shading use a very dark shade of gray instead (e.g. 1 or RGB(1,1,1))."
         
         .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "LineColor, LineWidth(optional)"
        .FontBold = False: .FontItalic = False
        .Paragraph " These parameters specify colors and width to be used for line in table." + _
                   " Defaults color is black and line width 1. Also see TableBorder."
  
        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Wrap"
        .FontBold = False: .FontItalic = False
        .Paragraph "Specifies whether text should be allowed to wrap within the box. Optional, defaults to True."

        .Paragraph
        .FontBold = True: .FontItalic = True
        .Paragraph "Indent"
        .FontBold = False: .FontItalic = False
        .Paragraph "To indent a table by a specified amount of left and right of the column."
        .FontBold = False: .FontItalic = False

        .Paragraph
        .FontBold = True: .FontItalic = True: .FontUnderline = True
        .Paragraph "Aligning and indenting tables:"
        .FontBold = False: .FontItalic = False: .FontUnderline = False
        .Paragraph "Tables may be aligned to the left, center, or right of the page depending on the " + _
                   "setting of the TextAlign property."
        .Paragraph
        .Paragraph "Note: The table is using StartTable, Table or TableArray and finally EndTable. " + _
                   "Set properties of a table cell with a TableCell and for border with TableBorder." + _
                   "To read the dimensions of the table on each page and table CalcTable call the function and read the values in the X1, Y1, X2, Y2"
        .Paragraph

        .FontBold = True
        '.FontSize = .FontSize - 2
        .Paragraph "Example"
        .Paragraph
        .FontBold = False
        .ForeColor = vbBlue
        .TextAlign = taLeftTop
        .Paragraph "Dim db As New ADODB.Connection"
        .Paragraph "Dim rs As New Recordset"
        .Paragraph "Dim Body() As Variant"
        .Paragraph ""
        .Paragraph "db.ConnectionString = ''Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Nwind.mdb;Persist Security Info=False"
        .Paragraph ""
        .Paragraph "db.Open"
        .Paragraph ""
        .Paragraph "rs.Open ''SELECT productid,productname,quantityperunit,unitprice FROM Products'', db, adOpenKeyset, adLockOptimistic"
        .Paragraph ""
        .Paragraph "If rs.RecordCount > 0 Then"
        .Paragraph "   Body = rs.GetRows(rs.RecordCount)"
        .Paragraph "End If"
        .Paragraph ""
        .Paragraph "rs.Close"
        .Paragraph ""
        .Paragraph "With PrintPreview"
        .IndentLeft = 1
        .Paragraph ".TextAlign = taLeftTop"
        .Paragraph ".TableBorder = tbAll"
        .Paragraph ".StartDoc"
        .Paragraph ".FontSize = 10"
        .Paragraph ".StartTable"
        .Paragraph " FormatCol$ = ''^+2cm|<+2.9cm|^+4.0cm|>+4.0cm;''"
        .Paragraph " Header$ = ''Id|Product Name|Quanti Type Uunit|UnitPrice;'';"
        .Paragraph ".TableArray FormatCol$, Header$, Body, , , , , False"
        .Paragraph ".EndTable"
        .Paragraph ".EndDoc"
        .IndentLeft = 0
        .Paragraph "End With"
        .Paragraph
        .ForeColor = vbBlack
        '-----------------------
        .NewPage
        '-----------------------
        .Paragraph
        '.FontSize = .FontSize + 2
         .TextAlign = taLeftTop
        
         .TableBorder = tbAll
         .DrawWidth = 1
         '.FontCharSet = 0
         .StartTable
            FormatCol$ = "^+2.0cm|<+5cm|^+5.0cm|>+4.0cm"
            Header$ = "Id|Product Name|Quanti Type Uunit|UnitPrice;" '•
           .TableArray FormatCol$, Header$, Body, vbGreen, , , , , "2mm"
           .TableCell tcFontBold, 1, , True
           .TableCell tcFontBold, , 1, True
         .EndTable
         .CalcTable
         Debug.Print "Position Table", .X1, .Y1, .X2, .Y2
       .EndDoc

     End With
     
End Sub

Private Sub DoText()
    Dim s As String
    
    With VBPrintPreview1
       
        .Clear
        .Orientation = PagePortrait
        .Zoom = zmWholePage
        '.LineSpace = lsSpaceHalfline ' lsSpaceSingle ' lsSpaceLine15
         SetPages "PrintPreview|Text"
        .StartDoc
         SetTitle "Text function"
         .Paragraph
        .FontSize = 12
        .Paragraph "Renders a string on the page at the current cursor position." + vbCrLf
        .FontBold = True
        .Paragraph "Syntax:"
        .Paragraph
        .Paragraph "[form.]VBPPreview.Text (Value As String)"
        .FontBold = False
        .Paragraph
        .TextAlign = taJustifyTop
         s = "   This property renders text on the page, at the current cursor position " + _
             "(determined by the CurrentX and CurrentY properties). It is similar to the Paragraph property, " + _
             "except it does not terminate the string with a new line. Instead, it leaves the cursor  " + _
             "at the point where the text stopped, allowing you to print a paragraph in pieces. " + _
             "This is the most efficient way to generate paragraphs with mixed fonts and colors."
        .Paragraph s + vbCr
        
        .Paragraph "Example:" + vbCrLf
        .FontSize = 30
        .PageBorder = pbTopBottom
        
        ' start a line or paragraph with one font
        .ForeColor = vbBlack
        
        .LineSpace = lsSpaceSingle
        .Text " With VB PrintPrinter control, you can print"
        ' continue with a different font
        .FontName = "Times New Roman"
        .FontBold = True
        .ForeColor = RGB(0, 0, 125)
        .FontUnderline = True
        .FontItalic = True
        .Text " Line by Line"
        .FontUnderline = False
        .FontItalic = False
        .FontName = "Arial"
        .FontSize = 30
        ' and finish with the original font!
        .ForeColor = 0
        .FontBold = False
        .Text " and modify text attributes as you go. You can use"
        'Change Fonts property
        .FontBold = True
        .Text " FontBold,"
        .FontBold = False
        .FontItalic = True
        .Text " FontItalic, "
        .FontItalic = False
        .FontUnderline = True
        .Text " FontUnderline, "
        .FontUnderline = False
        
        If .TextWidth("FontTransparent,") + .CurrentX > .PageWidth - .MarginRight Then
            .DrawRectangle .MarginLeft, .CurrentY + .TextHeight, .TextWidth("FontTransparent,"), .TextHeight, , , vbBlack, vbBlack, vbFSSolid
        Else
            .DrawRectangle .CurrentX, .CurrentY, .TextWidth("FontTransparent,"), .TextHeight, , , vbBlack, vbBlack, vbFSSolid
        End If
        .ForeColor = vbWhite
        .Text " FontTransparent, "
        .ForeColor = vbBlack
        
        .Text " and "
        .FontStrikethru = True
        .Text "FontStrikethru."
        .FontStrikethru = False
        .Text " The text wraps automatically, so your life becomes easier."
        .FontSize = 11
       
      .EndDoc
      
    End With

End Sub

Private Sub DoParagraphs(Optional Example As Boolean = False)
    
    Dim s$
    Dim iChapter%, iParagraph%
    
    With VBPrintPreview1
        .Clear
        .Zoom = zmWholePage
        .TextAlign = taJustifyTop
        SetPages "PrintPreview|Paragraph"
        
       .StartDoc
        .FontSize = 12
        
       If Example = False Then
         SetTitle "Paragraph function"
         .Paragraph
          
        .Paragraph "Renders a paragraph on the page at the current cursor position."  '+ vbCrLf
        .Paragraph
        .FontBold = True
        .Paragraph "Syntax:" + vbCr
        .Paragraph "[form.]VBPPreview.Paragraph [Value As String]"
        .FontBold = False
        .Paragraph
        .Paragraph " This is the main function for placing text on a page. The control takes care of justification, word wrapping, and page breaks." + _
                   " Before rendering the text, the Paragraph property checks the current cursor position. If the cursor is free (that is, " + _
                   "was not set by the user or as a result of rendering text with the Text property), then its horizontal position is reset to " + _
                   "the left margin and Indent (for left and right) and its vertical position is incremented by the amount specified with the CurrentY property. " + vbCr + _
                   "After rendering the text, the Paragraph property resets the horizontal cursor position to the left margin."
        .Paragraph vbCr
        
        .Paragraph "Example:" + vbCrLf
         s = vbTab + "The example 'Paragraphs' is a test paragraph to show the potential of 'Function Paragrafh'." + _
             " Before adopting the text, you can register your Property TextAlign, LineSpace, IndentFirst, IndentLeft and Fonts property."
         .Paragraph s
         
       Else
         
        '.NewPage
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
       End If
       .EndDoc
       
    End With

End Sub

Private Sub DoLine()
     Dim Counter As Integer
     Dim R, G, B
    Dim X, Y


     With VBPrintPreview1
          .Clear
          
          .Zoom = zmWholePage
          .TextAlign = taJustifyTop
          SetPages "PrintPreview|Draw Line"
          .StartDoc
            SetTitle "Line function"
            .Paragraph
        
            .Paragraph "Draws a line segment."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:" + vbCr
            .Paragraph "[form.]VBPPreview.Line X1 As Variant, ByVal Y1 As Variant, _"
            .IndentLeft = .TextWidth("[form.]VBPPreview.Line ")
                                  .Paragraph " X2 As Variant, ByVal Y2 As Variant, _"
                                  .Paragraph " [LineWidth As Integer], [ColorLine As Long]"
            .IndentLeft = 0
            .FontBold = False
            .Paragraph
            .Paragraph vbTab + "The DrawLine method draws a line between point (X1, Y1) and (X2, Y2). " + _
                       "If (X2, Y2) are omitted, the line is drawn from the last point used in a " + _
                       "call to DrawLine and (X1, Y1)."
            .Paragraph vbTab + "If you have not set ''LineWith'' and ''ColorLine'', the line will get the properties " + _
                       "to the current pen, as defined by the properties ForeColor, DrawStyle and DrawWidth."
            .Paragraph vbTab + "The X1, Y1, X2, and Y2 parameters may be specified with units (inches, " + _
                       "twips, cm, mm, or pixels). The default unit is Scalemode. For details on using " + _
                       "unit-aware measurements, see the Using Unit Properties."

            .Paragraph vbTab + "To draw complex shapes, you may prefer to use the Polygon or PolyLine properties instead."

            .ForeColor = vbRed
            
             For Counter = 1 To 100
                R = Rnd * 255
                G = Rnd * 255
                B = Rnd * 255 + 50
                .DrawStyle = Rnd * 6
                If Counter = 1 Then
                   .DrawLine Rnd * 10 + 5, Rnd * 10 + 15, Rnd * 10 + 5, Rnd * 10 + 15, 1, RGB(R, G, B)
                Else
                  .DrawLine Rnd * 10 + 5, Rnd * 10 + 15, , , 1, RGB(R, G, B)
                End If
             Next
            
            .FontBold = False
            
            .Paragraph
          .EndDoc
     End With
End Sub

Sub DoRectangle()
  Dim Counter As Integer
  Dim R As Integer, G As Integer, B As Integer, RF As Integer, GF As Integer, BF As Integer
  Dim X As Single, Y As Single, w As Single, H As Single ', R As Single
  Dim Body As String
  
     With VBPrintPreview1
          .Clear
          
          .Zoom = zmWholePage
          .TextAlign = taJustifyTop
          SetPages "PrintPreview|DrawRectangle"
          .StartDoc
            SetTitle "DrawRectangle function"
            .Paragraph
        
            .Paragraph "Draws a rectangle."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:" + vbCr
            .Paragraph "[form.]VBPPreview.DrawRectangle X As Variant, ByVal Y As Variant, _"
            .IndentLeft = .TextWidth("[form.]VBPPreview.DrawRectangle ")
            .Paragraph " Width As Variant, ByVal Height As Variant, _"
            .Paragraph " [Radius1 As Variant], [Radius2 As Variant], _"
            .Paragraph " [ColorLine As Long], [ColorFill As Long], _"
            .Paragraph " [FilledBox As FillStyleConstants]"
            .IndentLeft = 0
            .FontBold = False
            .FontSize = 11
            .Paragraph
       
            Body = "X, Y|First corner of the rectangle.;" + _
                   "Width, Height| Second corner of the rectangle.;" + _
                   "Radius1|Optional parameter that specifies the radius of a rounded rectangle's corner, in the horizontal direction.;" + _
                   "Radius2|Optional parameter that specifies the radius of a rounded rectangle's corner, in the vertical direction.;" + _
                   "ColorLine|Optional parameter that specifies the Color of line, if there is then we use it ForeColor;" + _
                   "ColorFill|Optional parameter that specifies the Color of backround, if there is then we use it FillColor;" + _
                   "FilledBox|Optional parameter that specifies the fiil style, default is vbFSTransparent;"
            .StartTable
            .Table "2.5cm|12.5cm", "Parameter|Description", Body, , , vbWhite
            .TableCell tcFontBold, 1, , True
            .EndTable
            
            .Paragraph vbTab + "The X, Y, Width, Height, Radius1 and Radius2 parameters may be specified with units (inches, " + _
                       "twips, cm, mm, or pixels). The default unit is Scalemode. For details on using " + _
                       "unit-aware measurements, see the Using Unit Properties."

            .Paragraph vbTab + "To draw complex shapes, you may prefer to use the Polygon or PolyLine properties instead."

            .ForeColor = vbRed
            Randomize
             For Counter = 1 To 20
                R = Rnd * 255
                G = Rnd * 255
                B = Rnd * 255
                RF = Rnd * 255
                GF = Rnd * 255
                BF = Rnd * 255
                X = Rnd * 10 + 2
                Y = Rnd * 10 + 15.5
                w = Rnd * 4 + 5
                H = Rnd * 3 + 2
                R = Rnd * 2
                If Y + H > .PageHeight - .MarginBottom Then
                   H = .PageHeight - .MarginBottom - Y
                End If
                .DrawStyle = Rnd * 6
                .DrawRectangle X, Y, _
                               w, H, _
                               R, R, _
                               RGB(R, G, B), RGB(RF, GF, BF), _
                               Rnd * 6
             Next
            
            .FontBold = False
            
            .Paragraph
          .EndDoc
     End With
End Sub

Private Sub DoPolygon()
    Dim Body As String
    With VBPrintPreview1
          .Clear
          
          .Zoom = zmWholePage
          .TextAlign = taJustifyTop
          SetPages "PrintPreview|Draw Polygon"
          .StartDoc
            SetTitle "Polygon function"
            .Paragraph
        
            .Paragraph "Draws a polygon defined by a string of X,Y coordinates."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:" + vbCr
            .Paragraph "[form.]VBPPreview.DrawPolygon Points As String, _"
            .IndentLeft = .TextWidth("[form.]VBPPreview.DrawPolygon ")
            .Paragraph " [ColorLine As Long], [ColorFill As Long], _"
            .Paragraph " [FilledPolygon As FillStyleConstants], _"
            .Paragraph " [FillPolyMode As PolyFillMode]"
            .IndentLeft = 0
            .FontBold = False
            '.FontSize = 10
            .Paragraph
       
            Body = "Points|The string assigned to the Polygon property contains a sequence of coordinates, separated by spaces or commas;" + _
                   "ColorLine|Optional parameter that specifies the Color of line, if there is then we use it ForeColor;" + _
                   "ColorFill|Optional parameter that specifies the Color of backround, if there is then we use it FillColor;" + _
                   "FilledPolygon|Optional parameter that specifies the fill style, default is vbSolid;" + _
                   "FillPolyMode|Optional parameter that specifies the fill mode, default is WINDING.;"
            .StartTable
               .Table "3cm|12cm", "Parameter|Description", Body, , , vbWhite
               .TableCell tcFontBold, 1, , True
            .EndTable
            
            .Paragraph
            .Paragraph vbTab + "The ''Points'' parameters may be specified with units (inches, " + _
                       "twips, cm, mm, or pixels). The default unit is Scalemode. For details on using " + _
                       "unit-aware measurements, see the Using Unit Properties."
             .Paragraph
  
            .ForeColor = vbRed
            .DrawPolygon "6.98 20,7.96 23.96,5 21.32, 8.96 21.32,6.00 23.96", vbBlue, vbGreen, vbFSSolid
            
            .DrawPolygon "11.48 20,12.46 23.96,9.49 21.32,13.46 21.32,10.50 23.96", vbBlue, vbGreen, vbDiagonalCross, ALTERNATE
            
            .TextBox "WINDING", 6.2, 24, .TextWidth("WINDING"), 0, taCenterTop, False
            .TextBox "ALTERNATE", 10.2, 24, .TextWidth("ALTERNATE"), 0, taCenterTop, False
            .FontBold = False
          .EndDoc
    End With
End Sub

Private Sub DoCircle()
    Dim Body As String
    Dim Counter As Integer
    Dim R As Integer, G As Integer, B As Integer, RF As Integer, GF As Integer, BF As Integer
    Const pi = 3.14159265
    With VBPrintPreview1
          .Clear
          
          .Zoom = zmWholePage
          .TextAlign = taJustifyTop
          SetPages "PrintPreview|Draw Circle Arc"
          .StartDoc
            SetTitle "Circle function"
            .Paragraph
        
            .Paragraph "Draws a circle, circular wedge, or circular arc."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:" + vbCr
            
            
            .Paragraph "[form.]VBPPreview.DrawCircle CenterX As Variant, CenterY As Variant, _"
            .IndentLeft = .TextWidth("[form.]VBPPreview.DrawCircle ")
            .Paragraph " Radius As Variant, _"
            .Paragraph " [cStart As Single], [cEnd As Single], _"
            .Paragraph " [ColorLine As Long], [ColorFill As Long], _"
            .Paragraph " [DrawStyle As DrawStyleConstants], _"
            .Paragraph " [FillStyle As FillStyleConstants]"
            .IndentLeft = 0
            .FontBold = False
            '.FontSize = 10
            .Paragraph
             .TextAlign = taJustifyTop
            Body = "CenterX|X coordinate of the center of the circle, arc, or wedge.;" + _
                   "CenterY|Y coordinate of the center of the circle, arc, or wedge.;" + _
                   "Radius|Radius of the circle, arc, or wedge.;" + _
                   "cStart,cEnd|Optional. Single-precision values. When an arc or a partial circle or ellipse is drawn, start and end specify (in radians) the beginning and end positions of the arc. The range for both is –2 pi radians to 2 pi radians. The default value for start is 0 radians, the default for end is 2 * pi radians;" + _
                   "ColorLine|Optional parameter that specifies the Color of line, if there is then we use it ForeColor;" + _
                   "ColorFill|Optional parameter that specifies the Color of backround, if there is then we use it FillColor;" + _
                   "DrawStyle|Optional parameter that specifies the draw fill , default is vbSolid;" + _
                   "FillStyle|Optional parameter that specifies the fill style, default is vbFSTransparent;"
            .StartTable
                .Table "3cm|12cm", "Parameter|Description", Body, , , vbWhite
                .TableCell tcFontBold, 1, , True
            .EndTable
            .Paragraph
            .Paragraph vbTab + "The Variant parameters may be specified with units (inches, " + _
                       "twips, cm, mm, or pixels). The default unit is Scalemode. For details on using " + _
                       "unit-aware measurements, see the Using Unit Properties."
             .Paragraph
  
            .ForeColor = vbRed

              For Counter = 1 To 30
                  R = Rnd * 255
                  G = Rnd * 255
                  B = Rnd * 255
                  RF = Rnd * 255
                  GF = Rnd * 255
                  BF = Rnd * 255
                 .DrawCircle Rnd * 18 + 2, Rnd * 5 + 18, (Rnd * 1) + 0.5, , , RGB(R, G, B), RGB(RF, GF, BF), Rnd * 6, Rnd * 7
             Next
             
             For Counter = 1 To 20
                  R = Rnd * 255
                  G = Rnd * 255
                  B = Rnd * 255
                  RF = Rnd * 255
                  GF = Rnd * 255
                  BF = Rnd * 255
                 .DrawCircle Rnd * 18 + 2, Rnd * 5 + 18, (Rnd * 1) + 0.5, Rnd * -2 * pi, Rnd * 2 * pi, RGB(R, G, B), RGB(RF, GF, BF), Rnd * 6, Rnd * 7
             Next
             
            .FontBold = False
          .EndDoc
    End With
   
End Sub

Private Sub DoElipse()
    Dim Body As String
    Dim Counter As Integer
    Dim R As Integer, G As Integer, B As Integer, RF As Integer, GF As Integer, BF As Integer
    Const pi = 3.14159265
     
    With VBPrintPreview1
          .Clear
          
          .Zoom = zmWholePage
          .TextAlign = taJustifyTop
          SetPages "PrintPreview|Draw Elipse"
          .StartDoc
            SetTitle "Elipse function"
            .Paragraph
        
            .Paragraph "Draws an ellipse, wedge, or arc."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:" + vbCr
            
            'Left As Variant, Top As Variant, _
             Width As Variant,Height As Variant, _
             [cStart As Single], [cEnd As Single], _
             [ColorLine As Long], [ColorFill As Long], _
             [DrawStyle As DrawStyleConstants], _
             [FillStyle As FillStyleConstants]
            .Paragraph "[form.]VBPPreview.DrawElipse Left As Variant, Top As Variant, _"
            .IndentLeft = .TextWidth("[form.]VBPPreview.DrawElipse ")
            .Paragraph " Width As Variant, Height As Variant, _"
            .Paragraph " [ColorLine As Long], [ColorFill As Long], _"
            .Paragraph " [cStart As Single], [cEnd As Single], _"
            .Paragraph " [DrawStyle As DrawStyleConstants], _ "
            .Paragraph " [FillStyle As FillStyleConstants]"
            .IndentLeft = 0
            .FontBold = False
            '.FontSize = 10
            .Paragraph
             .TextAlign = taJustifyTop
            Body = "Left,Top|First corner of the rectangle that encloses the ellipse.;" + _
                   "Width,Height|Second corner of the rectangle that encloses the ellipse.;" + _
                   "cStart,cEnd|Optional. Single-precision values. When an arc or a partial circle or ellipse is drawn, start and end specify (in radians) the beginning and end positions of the arc. The range for both is –2 pi radians to 2 pi radians. The default value for start is 0 radians, the default for end is 2 * pi radians;" + _
                   "ColorLine|Optional parameter that specifies the Color of line, if there is then we use it ForeColor;" + _
                   "ColorFill|Optional parameter that specifies the Color of backround, if there is then we use it FillColor;" + _
                   "DrawStyle|Optional parameter that specifies the draw fill , default is vbSolid;" + _
                   "FillStyle|Optional parameter that specifies the fill style, default is vbFSTransparent;"
            .StartTable
            .Table "3cm|12cm", "Parameter|Description", Body, , , vbWhite
            .TableCell tcFontBold, 1, , True
            .TableCell tcTextAling, , 2, taJustifyTop
            .EndTable
            .Paragraph
            .Paragraph vbTab + "The Variant parameters may be specified with units (inches, " + _
                       "twips, cm, mm, or pixels). The default unit is Scalemode. For details on using " + _
                       "unit-aware measurements, see the Using Unit Properties."
             .Paragraph
  
            .ForeColor = vbRed

              For Counter = 1 To 30
                  R = Rnd * 255
                  G = Rnd * 255
                  B = Rnd * 255
                  RF = Rnd * 255
                  GF = Rnd * 255
                  BF = Rnd * 255
                 .DrawEllipse Rnd * 15 + 2, Rnd * 10 + 17, Rnd * 2 + 1, Rnd * 2 + 1, , , _
                              RGB(R, G, B), RGB(RF, GF, BF), Rnd * 6, Rnd * 7
             Next
             
              For Counter = 1 To 20
                  R = Rnd * 255
                  G = Rnd * 255
                  B = Rnd * 255
                  RF = Rnd * 255
                  GF = Rnd * 255
                  BF = Rnd * 255
                 .DrawEllipse Rnd * 15 + 2, Rnd * 10 + 17, Rnd * 2 + 1, Rnd * 2 + 1, Rnd * -2 * pi, Rnd * 2 * pi, _
                              RGB(R, G, B), RGB(RF, GF, BF), Rnd * 6, Rnd * 7
             Next
            .FontBold = False
          .EndDoc
    End With
   
End Sub

Private Sub DoFont()

    Dim i As Integer, s As String
 
    s = "This is a Printer font "
    
    With VBPrintPreview1
        .Clear
        .PaperSize = vbPRPSA4
        .Orientation = PagePortrait
        .Zoom = zmWholePage
        SetPages "PrintPreview|Fonts"
        .StartDoc
        ' Paragraph/Text
        
        SetSubTitle "Paragraph/Text Fonts"
        SetNormal
        .LineSpace = lsSpaceLine15
        .Paragraph "The following are the properties you can use to customize the look and feel of the font in your text:"

        .IndentFirst = "1cm"
        .FontBold = True
        .Paragraph "•" + Chr(9) + "FontName" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "FontSize" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "FontCharSet" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "FontBold" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "FontItalic" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "FontStrikethru" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "FontUnderline" + Chr(9) + "•"
        .IndentFirst = 0
        SetNormal
        .Paragraph "Take a look at this sample code that sets the paragraph to Arial, 14 pts., bold, underline, and italic."
        SetCode
        .Paragraph "VBPPreview.FontName = ''Arial''"
        .Paragraph "VBPPreview.FontBold = True"
        .Paragraph "VBPPreview.FontCharSet = 0"
        .Paragraph "VBPPreview.FontItalic = True"
        .Paragraph "VBPPreview.FontUnderline = True"
        .Paragraph "VBPPreview.FontSize = 14"
        SetNormal

        '------------------------------------------------------------------------------
        ' Headers/Footers
        '------------------------------------------------------------------------------
        SetSubTitle "Headers/Footers Fonts"

        .Paragraph "The following are the properties you can use to customize the look and feel of the font in your text:"
        .IndentFirst = "10mm"
        .FontBold = True
        .Paragraph "•" + Chr(9) + "HdrFontName" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "HdrFontSize" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "HdrFontBold" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "HdrFontItalic" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "HdrFontStrikethru" + Chr(9) + "•"
        .Paragraph "•" + Chr(9) + "HdrFontUnderline" + Chr(9) + "•"
        .IndentFirst = 0
        SetNormal

        '------------------------------------------------------------------------------
        ' Sample Fonts
        '------------------------------------------------------------------------------
        .NewPage
        
        SetSubTitle "Sample Fonts Printer"
        .Paragraph
        .FontSize = 12
    
        .StartTable
        .Table "3000tw|5500tw", "Font Name|Sample", "", vbYellow, , , , True
         For i = 1 To Screen.FontCount
             .TableCell tcRows, , , i + 1
             .TableCell tcFontName, i + 1, 2, Screen.Fonts(i - 1)
             .TableCell tcText, i + 1, 1, Screen.Fonts(i - 1)
             .TableCell tcText, i + 1, 2, s
         Next
        
        .TableCell tcFontBold, 1, , True
        .EndTable
    
      .EndDoc
  End With

End Sub

Sub DoLineSpace()
 Dim Body As String
 
 With VBPrintPreview1
    .Clear
    
     SetPages "PrintPreview|LineSpace"
    .PageBorder = pbTopBottom
        
    .StartDoc
     SetTitle "LineSpace property"
    .ForeColor = vbBlack
    .Paragraph
    .FontSize = 12
    .Paragraph "Returns or sets the line spacing."
    .Paragraph
    .FontBold = True
    .Paragraph "Syntax:"
    .Paragraph
    .Paragraph "[form.]VBPPreview.LineSpace [ = LineSpaceConstants]"
    .Paragraph
    .FontBold = False
    .TextAlign = taJustifyTop
    .Paragraph vbTab + "If no units are supplied, this value is interpreted as a percentage of the " + _
               "single line spacing for the current font (font height plus external leading.) " + _
               "In this case, the spacing changes automatically when you change the font size."
     .Paragraph vbCr + vbTab + "The table below shows some common settings for the LineSpacing property:"
     .Paragraph
           .TextAlign = taJustifyTop
            Body = "lsSpaceSingle|0|Single-line spacing (default value);" + _
                   "lsSpaceLine15|1|1.5 line spacing.;" + _
                   "lsSpaceDoubleline|2|Double line spacing.;" + _
                   "lsSpaceHalfline|3|Half line spacing;"
            .StartTable
                .Table "<4cm|^2.5cm|<12cm", "Constant|Value|Description", Body, , , vbWhite
                .TableCell tcFontBold, 1, , True
            .EndTable
      .Paragraph
     
     .Paragraph "Default Value : lsSpaceSingle"
     
    .EndDoc
    End With
End Sub

Sub DoUsingUnit()
    Dim Body As String
    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .Zoom = zmWholePage
        SetPages "PrintPreview|Unit property"
        .LineSpace = lsSpaceLine15
        
       .StartDoc
        SetSubTitle "Using Unit"
        SetNormal
       .TextAlign = taJustifyTop
       .FontName = "Tahoma"
       .Paragraph
       .IndentFirst = "10mm"
        .Paragraph "VBPPreview allows you to specify virtually all measurements and positioning " + _
                   "properties and parameters as a value followed by a unit (To allow this type of assignment, all " + _
                   "unit-aware properties and method parameters are of type Variant.)"
        .IndentFirst = 0
        .Paragraph
        .Paragraph "For example, the MarginLeft property may be assigned in several ways:"
        SetCode
        .Paragraph "VBPPreview.MarginLeft = 14        ' no units, assume ScaleMode"
        .Paragraph "VBPPreview.MarginLeft = ''1in''   ' one inch"
        .Paragraph "VBPPreview.MarginLeft = ''62mm''  ' 62 millimeter"
        .Paragraph "VBPPreview.MarginLeft = ''2.3cm'' ' 2.3 centimeters"
        SetNormal
        .Paragraph
        .Paragraph "When the assignment is made, VBPrintPrinter converts the given measurement into the default " + _
                   "units. This is done to allow the property to be used in mathematical expressions. For example:"
         SetCode
        .Paragraph "VBPPreview.ScaleMode = smCentimeters"
        .Paragraph "VBPPreview.MarginLeft = ''1in''   ' one inch"
        .Paragraph "VBPPreview.MarginLeft = 2 * vp.MarginLeft"
        .Paragraph "Debug.Print vp.MarginLeft; ''(default unit is the choice of ScaleMode)''"
         SetNormal
        .Paragraph ""
        .Paragraph "The table below shows the units recognized by the VBPPreview control:"

        .Paragraph
        SetCode
        
        '.Paragraph "Symbol unit"
        '.Paragraph "None        Default unit is the choice of ScaleMode."
        '.Paragraph "in          Inches."
        '.Paragraph "tw          Twips."
        '.Paragraph "cm          Centimeters."
        '.Paragraph "mm          Millimeters."
        '.Paragraph "px          Printer pixels."
        '.Paragraph "%           Percentage."
        
        Body = "None|Default unit is the choice of ScaleMode.;" + _
                "in|Inches.;" + _
                "tw|Twips.;" + _
                "cm |Centimeters.;" + _
                "mm|Millimeters.;" + _
                "cr|Characters.;" + _
                "pt|Point.;" + _
                "px|Printer pixels.;" + _
                "%|Percentage.;"
         .StartTable
            .TableBorder = tbNone
            .Table "<3cm|<10cm", "Symbol unit|Description", Body
            .TableCell tcFontUnderline, 1, , True
        .EndTable
        SetNormal
        .Paragraph ""
        .Paragraph "When you use percentage units, the meaning is context-dependent." + _
                   " For rendering pictures, 100% is actual size. For horizontal margins and " + _
                   "table column widths, 100% is the page width."
      .EndDoc
   End With
End Sub

Sub DoZoom()
   Dim Body As String
   With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .Zoom = zmWholePage
        SetPages "PrintPreview|Zoom"
        .LineSpace = lsSpaceSingle
       .StartDoc
             SetTitle "Zoom Mode"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            .FontSize = 12
            .Paragraph "Sets or returns the zoom mode."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .Paragraph "[form.]VBPPreview.Zoom [ = ZoomModeConstants]"
            .FontBold = False
            .Paragraph
            .FontSize = 10
            .Paragraph "The settings for the Zoom property are described below:"
            .Paragraph ""
            
            Body = "zmRation50|0|The page appears 50%;" + _
                    "zmRation75|1|The page appears 75%;" + _
                    "zmRation100|2|The page appears 100%;" + _
                    "zmRation150|3|The page appears 150%;" + _
                    "zmRation200|4|The page appears 200%;" + _
                    "zmWholePage|5|Show a whole page.;" + _
                    "zmPageWidth|6|Show page so that it fits horizontally within the control.;" + _
                    "zmThumbnail|7|Show wide preview pages as will fit on the control;"
         .StartTable
            .TableBorder = tbNone
            .Table "<3cm|^2cm|10cm<", "Constant|Value|Description", Body
            .TableCell tcFontUnderline, 1, , True
        .EndTable
        
        '.Paragraph ""
        PicLoad.Picture = LoadPicture(App.Path + "\Library\WholePage.bmp")
        .DrawPicture PicLoad, .MarginLeft, .CurrentY, "100%", "100%"
        .CalcPicture
        .CurrentY = .Y2
        .Paragraph "Show WholePage"
        
        .NewPage
        .Paragraph ""
        PicLoad.Picture = LoadPicture(App.Path + "\Library\PageWidth.bmp")
        .DrawPicture PicLoad, .MarginLeft, .CurrentY, "100%", "100%"
        .CalcPicture
        .CurrentY = .Y2
        .Paragraph "Show PageWidth"
        
        .Paragraph ""
        PicLoad.Picture = LoadPicture(App.Path + "\Library\ThumbNail.bmp")
        .DrawPicture PicLoad, .MarginLeft, .CurrentY, "100%", "100%"
        .CalcPicture
        .CurrentY = .Y2
        .Paragraph "Show ThumbNail"
        
       .EndDoc
   End With
End Sub

Sub DoNavBar()
    Dim Format  As String, Body As String
    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .Zoom = zmWholePage
        SetPages "PrintPreview|Navigation Bar"
        
       .StartDoc
             SetTitle "Navigation Bar"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            .FontSize = 11
            .Paragraph "Returns or sets whether to display the document navigation bar."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .Paragraph "[form.]VBPPreview.NavBar [ = NavBarSettings ]"
            .FontBold = False
            .FontSize = 9
            .Paragraph
            .Paragraph "The NavBar property controls the position and configuration of the optional " + _
                       "built-in preview navigation bar. Valid settings are:"
             .Paragraph
              'Format = "Constant|Value|Description"
              Body = "nbNone|0|Do not display the navigation bar.This is the default setting.;" + _
                     "nbTop|1|Display a simple navigation bar at the top of the control.;" + _
                     "nbBottom|2|Display a simple navigation bar at the bottom of the control.;" + _
                     "nbTopPrint|3|Display a complete navigation bar at the top of the control (including a print button).;" + _
                     "nbBottomPrint|4|Display a complete navigation bar at the bottom of the control (including a print button).;"
             .StartTable
               .Table "<3cm|^2cm|<15", "Constant|Value|Description", Body, , , vbWhite
               .TableCell tcFontBold, 1, , True
             .EndTable
             .Paragraph
            .Paragraph "Display property NavBar controls to manage Preview."
            '.Paragraph ""
            PicLoad.Picture = LoadPicture(App.Path + "\Library\NavBar.bmp")
            .DrawPicture PicLoad, .MarginLeft, .CurrentY, "100%", "100%"
            .CalcPicture
            .CurrentY = .Y2
            .Paragraph "Show Navigation Bar"
        
       .EndDoc
   End With
End Sub

Sub DoPageBorder()
    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .Zoom = zmWholePage
        SetPages "PrintPreview|Page Border"
        .PageBorderColor = vbBlue
        .PageBorderWidth = 2
       .StartDoc
             SetTitle "Page Border"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            .FontSize = 12
            .Paragraph "Returns or sets the type of border to draw around each page."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .Paragraph "[form.]VBPrintPreview.PageBorder[ = PageBorderSettings ]"
            .Paragraph
            .FontBold = False
            .IndentFirst = "10mm"
            .IndentRight = 0
            .Paragraph "The border drawn using color and thickness are by the PageBorderColor, " + _
                       "and PageBorderWidth properties. The distance between the border and the " + _
                       "edges of the page is determined by the MarginLeft,MarginRight, MarginTop " + _
                       "and MarginBottom properties."

            .Paragraph
            .Paragraph "The picture below shows valid settings for the PageBorder property and their effect:"
            .Paragraph
            
            PicLoad.Picture = LoadPicture(App.Path + "\Library\PageBorder.bmp")
            .DrawPicture PicLoad, .MarginLeft, .CurrentY, "100%", "100%"
            .CalcPicture
            .CurrentY = .Y2
            .FontSize = .FontSize - 2
            .Paragraph "Show Page Border Settings"

       .EndDoc
        .PageBorderColor = 0
        .PageBorderWidth = 1
   End With
End Sub

Sub DoIndent()
       
    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .TextAlign = taLeftTop
       .Zoom = zmWholePage
        SetPages "PrintPreview|Indent"
       .StartDoc
             SetTitle "Indent Property"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            .FontSize = 12
            .Paragraph "Returns or sets an additional indent each paragraphs."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .Paragraph "[form.]VBPrintPreview.IndentFirst[ = value As Variant ]"
            .Paragraph "[form.]VBPrintPreview.IndentLeft [ = value As Variant ]"
            .Paragraph "[form.]VBPrintPreview.IndentRight[ = value As Variant ]"
            .Paragraph
            .FontBold = False
            .Paragraph "This property sets the distance between the text and the left margin of the page (MarginLeft)."
            .Paragraph "The IndentFirst returns or sets an additional left indent for the first line of each paragraphs."
            .Paragraph
            .Paragraph "For details on using unit-aware measurements, see the 'Using Unit Properties' topic. For example see 'Paragraph'."
            .Paragraph
             PicLoad.Picture = LoadPicture(App.Path + "\Library\Indent.bmp")
            .DrawPicture PicLoad, .MarginLeft, .CurrentY, "100%", "100%"
            .CalcPicture
            .CurrentY = .Y2
            .Paragraph ""
            .FontSize = .FontSize - 2
            .Paragraph "The diagram below shows the effect of the IndentLeft, IndentFirst, and IndentRight properties:"
                        
            .FontSize = .FontSize + 2
            .Paragraph
            .Paragraph
            .Paragraph "Example:"
            .Paragraph "the next page show the potential of Indent Property."
'            .TextAlign = taJustifyTop
'            .IndentLeft = "15mm"
'            .IndentFirst = "10mm"
'            .IndentRight = "15mm"
'            .LineSpace = lsSpaceLine15
'            .Paragraph "Test paragraph to show the potential of 'Indent Property'." + _
'                       " Before adopting the text, you can register your Property TextAlign = taJustifyTop," + _
'                       " LineSpace = lsSpaceLine15, IndentFirst = ''10mm'', IndentLeft = ''15mm'', IndentRight = ''15mm'' and Fonts property."
            .IndentLeft = 0
            .IndentFirst = 0
            .NewPage
            DoIndentation
       .EndDoc
   End With
End Sub


Private Sub DoIndentation()

    With VBPrintPreview1
            .FontSize = 18
            
            ' indent left, right, or center
            SetSubTitle "Indent First Line"
            .FontSize = 18
            .IndentFirst = "0.5in"
            .Paragraph "You can automatically indent the FIRST LINE of a paragraph."
            .Paragraph " "
                    
             SetSubTitle "Indent Left"
            .FontSize = 18
            .IndentFirst = 0
            .IndentLeft = "0.5in"
            .Paragraph "You can automatically indent from the LEFT margin of the page."
            .Paragraph " "
          
             SetSubTitle "Indent Right"
            .FontSize = 18
            .IndentLeft = 0
            .IndentRight = "0.5in"
            .Paragraph "You can automatically indent from the RIGHT margin of the page."
            .Paragraph " "
        
             SetSubTitle "Hanging Indents"
            .FontSize = 18
            .IndentLeft = "0.5in"
            .IndentFirst = "-0.5in"
            .Paragraph "•" & vbTab & "You can create HANGING INDENTS by setting IndentFirst to a negative value."
            .Paragraph " "
        
            ' restore defaults
            .IndentLeft = 0
            .IndentFirst = 0
            .IndentRight = 0
            
    End With

End Sub

Sub DoPaperSize()
    
    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .TextAlign = taLeftTop
       .Zoom = zmWholePage
        SetPages "PrintPreview|PaperSize"
       .StartDoc
          .FontSize = 12
         
             SetTitle "PaperSize"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            
            .Paragraph "Returns or sets a standard paper size."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
           ' .FontName = "Courier New"
            '.FontSize = .FontSize - 1
            .Paragraph "[form.]VBPrintPreview.PaperSize [ = PaperSizeConstans ]"
            .FontSize = 10
            '.FontName = "Arial"
            .Paragraph
            .FontBold = False
            .Paragraph "The settings for the PaperSize property are described below:"
            .Paragraph
            
            .IndentFirst = 0
        .StartTable
          .Table "<2in|^1in|<4in;", "Constant|Value|Description;", _
             "vbPRPSLetter|1|Letter, 8 1/2 x 11 in.;vbPRPSLetterSmall|2|Letter Small, 8 1/2 x 11 in.;" + _
             "vbPRPSTabloid|3|Tabloid, 11 x 17 in.;vbPRPSLedger|4|Ledger, 17 x 11 in.;" + _
             "vbPRPSLegal|5|Legal, 8 1/2 x 14 in.;vbPRPSStatement|6|Statement, 5 1/2 x 8 1/2 in.;" + _
             "vbPRPSExecutive|7|Executive, 7 1/2 x 10 1/2 in.;vbPRPSA3|8|A3, 297 x 420 mm;" + _
             "vbPRPSA4|9|A4, 210 x 297 mm;vbPRPSA4Small|10|A4 Small, 210 x 297 mm;" + _
             "vbPRPSA5|11|A5, 148 x 210 mm;vbPRPSB4|12|B4, 250 x 354 mm;" + _
             "vbPRPSB5|13|B5, 182 x 257 mm;vbPRPSFolio|14|Folio, 8 1/2 x 13 in.;" + _
             "vbPRPSQuarto|15|Quarto, 215 x 275 mm;vbPRPS10x14|16|10 x 14 in.;" + _
             "vbPRPS11x17|17|11 x 17 in.;vbPRPSNote|18|Note, 8 1/2 x 11 in.;" + _
             "vbPRPSEnv9|19|Envelope #9, 3 7/8 x 8 7/8 in.;vbPRPSEnv10|20|Envelope #10, 4 1/8 x 9 1/2 in.;" + _
             "vbPRPSEnv11|21|Envelope #11, 4 1/2 x 10 3/8 in.;vbPRPSEnv12|22|Envelope #12, 4 1/2 x 11 in.;" + _
             "vbPRPSEnv14|23|Envelope #14, 5 x 11 1/2 in.;" + _
             "vbPRPSEnvDL|27|Envelope DL, 110 x 220 mm;vbPRPSEnvC3|29|Envelope C3, 324 x 458 mm;" + _
             "vbPRPSEnvC4|30|Envelope C4, 229 x 324 mm;vbPRPSEnvC5|28|Envelope C5, 162 x 229 mm;" + _
             "vbPRPSEnvC6|31|Envelope C6, 114 x 162 mm;vbPRPSEnvC65|32|Envelope C65, 114 x 229 mm;" + _
             "vbPRPSEnvB4|33|Envelope B4, 250 x 353 mm;vbPRPSEnvB5|34|Envelope B5, 176 x 250 mm;" + _
             "vbPRPSEnvB6|35|Envelope B6, 176 x 125 mm;vbPRPSEnvItaly|36|Envelope, 110 x 230 mm;" + _
             "vbPRPSEnvMonarch|37|Envelope Monarch, 3 7/8 x 7 1/2 in.;vbPRPSEnvPersonal|38|Envelope, 3 5/8 x 6 1/2 in.;" + _
             "vbPRPSFanfoldUS|39|U.S. Standard Fanfold, 14 7/8 x 11 in.;vbPRPSFanfoldStdGerman|40|German Standard Fanfold, 8 1/2 x 12 in.;" + _
             "vbPRPSFanfoldLglGerman|41|German Legal Fanfold, 8 1/2 x 13 in.;vbPRPSUser|256|User-defined;", , , vbWhite, , False, "2mm"
            
            .TableCell tcFontBold, 1, , True
        .EndTable
            
            .Paragraph "For details on using unit-aware measurements, see the 'Using Unit Properties' topic and the 'Indent Properties'."
            .Paragraph
            
       .EndDoc
   End With
End Sub

Sub DoMargins(Optional Example As Boolean = False)
    Dim mBorder%
    Dim s$
    
    With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .TextAlign = taLeftTop
       .Zoom = zmWholePage
        SetPages "PrintPreview|Margin"
       .StartDoc
          .FontSize = 12
          If Example = False Then
             SetTitle "Margin"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            
            .Paragraph "Returns or sets the margin."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .FontName = "Courier New"
            .FontSize = .FontSize - 1
            .Paragraph "[form.]VBPrintPreview.MarginLeft  [ = value As Variant ]"
            .Paragraph "[form.]VBPrintPreview.MarginRight [ = value As Variant ]"
            .Paragraph "[form.]VBPrintPreview.MarginTop   [ = value As Variant ]"
            .Paragraph "[form.]VBPrintPreview.MarginBottom[ = value As Variant ]"
            '.Paragraph
            .Paragraph "[form.]VBPrintPreview.MarginHeader[ = value As Variant ]"
            .Paragraph "[form.]VBPrintPreview.MarginFooter[ = value As Variant ]"
            .FontSize = .FontSize + 1
            .FontName = "Arial"
            .Paragraph
            .FontBold = False
            .Paragraph vbTab + "The MarginLeft property measures the space between the text and the left edge of the page."
            .Paragraph vbTab + "The MarginRight property measures the space between the text and the right edge of the page."
            .Paragraph vbTab + "The MarginTop property measures the space between the text and the top of the page."
            .Paragraph vbTab + "The MarginBottom property measures the space between the text and the bottom of the page."
            .Paragraph vbTab + "The MarginHeader property measures the space between the top of the page and the top of the header text."
            .Paragraph vbTab + "The MarginFooter property measures the space between the top of the footer text and the bottom of the page."
            
            .Paragraph
            .Paragraph "For details on using unit-aware measurements, see the 'Using Unit Properties' topic and the 'Indent Properties'."
            .Paragraph
            
            PicLoad.Picture = LoadPicture(App.Path + "\Library\Margin.bmp")
            .DrawPicture PicLoad, .MarginLeft, .CurrentY, "100%", "100%"
            .Paragraph
            
        Else
            
            s = "Take advantage of existing Transact-SQL skills, and incorporate technologies, including the " + _
                "Microsoft ADO.NET Entity Framework and LINQ. Develop applications faster through deep " + _
                "integration with Visual Studio 2008, Visual Web Developer 2008, and SQL Server Management " + _
                "Studio. Model data by using the ADO.NET Entity Framework to hide database schema details and " + _
                "access data by using entities that closely resemble business logic. Take advantage of support for " + _
                "LINQ, including LINQ to SQL and LINQ to Entities, which allows data to be retrieved from entities " + _
                "natively from any Microsoft .NET language. Manage SQL Server 2008 R2 Express databases with " + _
                "SQL Server Management Studio Express. Connect to local SQL Server 2008 R2 Express databases and " + _
                "manage objects with full Object Explorer integration. Write, execute, and test queries by using " + _
                "visual query plans that provide hints to tune queries and access management and maintenance options."
    
            .PageBorder = pbNone
            .TextAlign = taJustifyTop
            
            ' introduction
            SetSubTitle "Changing Margins"
            .Paragraph "Set the MarginLeft, MarginRight, MarginTop, MarginBottom to adjust the margins of each page."
             
            .MarginLeft = "2in"
            SetSubTitle "MarginLeft to 2 inches"
            .Paragraph s
        
            .MarginRight = "2in"
            SetSubTitle "MarginLeft and MarginRight to 2 inches"
            .Paragraph s
        
            .MarginLeft = "1in"
            .MarginRight = "1in"
            SetSubTitle "MarginLeft and MarginRight to 1 inch"
            .Paragraph s
        End If
       .EndDoc
   End With
End Sub

Sub DoGetMargins(Optional Example As Boolean = False)
 With VBPrintPreview1
       .Clear
       .Orientation = PagePortrait
       .TextAlign = taLeftTop
       .Zoom = zmWholePage
        SetPages "PrintPreview|GetMargin"
       .StartDoc
       If Example = False Then
          .FontSize = 12
             SetTitle "GetMargin"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            .Paragraph "Returns the printable area, excluding margins, in the X1, Y1, X2, and Y2 properties."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
        .Paragraph "[form.]VBPrintPreview.GetMargins"
        .FontBold = False
        .Paragraph "This method provides a convenient way to determine the rectangle defined by the current page size, margins."
        .Paragraph
        .Paragraph "For example, to draw a line across the page (from margin to margin), you could use code such as:"
        .Paragraph
        .ForeColor = vbBlue
        .Paragraph "With VBPrintPreview1"
        .IndentLeft = "0.5in"
        .Paragraph ".X1 = .MarginLeft"
        .Paragraph ".X2 = .PageWidth - .MarginRight"
        .Paragraph ".DrawLine .X1, .Y1, .X2, .Y1"
        .IndentLeft = 0
        .Paragraph "End With"
        .Paragraph
        .ForeColor = 0
        .Paragraph "Or you could use the GetMargins method and write:"
        .ForeColor = vbBlue
        .Paragraph
        .Paragraph "With VBPrintPreview1"
        .IndentLeft = "0.5in"
        .Paragraph ".GetMargins"
        .Paragraph ".DrawLine .X1, .Y1, .X2, .Y1"
        .IndentLeft = 0
        .Paragraph "End With"
        .ForeColor = vbBlack
        .Paragraph "See example ''Picture Background''."
        .Paragraph
       Else
       
        If .FileExists(App.Path + "\Library\PrintWM.wmf") Then
            PicLoad.Picture = LoadPicture(App.Path + "\Library\PrintWM.wmf")
        End If
       ' place background on the page
        .GetMargins
        .DrawPicture PicLoad, .X1, .Y1, .PageWidth - .X2 - .X1, .PageHeight - .Y2 - .Y1
    
        ' place picture on top left
         If .FileExists(App.Path + "\Library\Printer.wmf") Then
            PicLoad.Picture = LoadPicture(App.Path + "\Library\Printer.wmf")
        End If
        .DrawPicture PicLoad, "1in", "4in", "5in", "5in"
        
        ' add some text to the document
        .FontName = "Tahoma"
        .FontSize = 24
        .IndentLeft = 0
        .Paragraph "VBPrintPrinter handles text over pictures." & vbLf
        .FontSize = 10
        .IndentLeft = "0.5in"
        .Paragraph "This page consists of the following elements:" & vbLf & vbLf & _
                   "1. A background picture from an image control" & vbLf & vbLf & _
                   "2. A picture (using a metafile for transparency)" & vbLf & vbLf & _
                   "3. Plain text over the graphics."
         .CalcParagraph
         Debug.Print .X1, .Y1, .X2, .Y2
         
         .NewPage
         
        If .FileExists(App.Path + "\Library\watermark.JPG") Then
            PicLoad.Picture = LoadPicture(App.Path + "\Library\watermark.JPG")
        End If
       ' place background on the page
        .GetMargins
        .DrawPicture PicLoad, .X1, .Y1, .PageWidth - .X2 - .X1, .PageHeight - .Y2 - .Y1
    
        ' place picture on top left
         If .FileExists(App.Path + "\Library\printer.jpg") Then
            PicLoad.Picture = LoadPicture(App.Path + "\Library\printer.jpg")
        End If
        .DrawPicture PicLoad, "1.5in", "4in", "4in", 0, vbSrcAnd
        
        ' add some text to the document
        .FontName = "Tahoma"
        .FontSize = 24
        .IndentLeft = 0
        .Paragraph "VBPrintPrinter handles text over pictures." & vbLf
        .FontSize = 10
        .IndentLeft = "4in"
        .Paragraph "This page consists of the following elements:" & vbLf & vbLf & _
                   "1. A background picture from an image control" & vbLf & vbLf & _
                   "2. A picture (using a image for transparency)" & vbLf & vbLf & _
                   "3. Plain text over the graphics."
         .CalcParagraph
         Debug.Print .X1, .Y1, .X2, .Y2
         
       End If
        
        .EndDoc
     End With
End Sub
Sub DoOrientation()
    Dim i%
    With VBPrintPreview1
    .Clear
    ' start in portrait mode
    '.Orientation = PagePortrait
    .Zoom = zmThumbnail
    .StartDoc
    .Orientation = PagePortrait
    .FontSize = 36
    .FontBold = True
    ' print some text in portrait mode
    For i = 1 To 10
        .Paragraph i & ": This is portrait mode."
    Next
    
    ' print graphics in landscape mode
    If .SendToPrinter Then .Orientation = PageLandscape
    .NewPage
    If .SendToPrinter = False Then .Orientation = PageLandscape

    .Paragraph "This is landscape mode."
    .FillStyle = vbFSTransparent
    .DrawCircle .PageWidth / 2, .PageHeight / 2, .PageHeight / 4
    
    ' print more text in portrait mode
    If .SendToPrinter Then .Orientation = PagePortrait
        .NewPage
    If .SendToPrinter = False Then .Orientation = PagePortrait
     For i = 1 To 10
        .Paragraph Str(i) & ": This is portrait mode."
    Next
    
    .EndDoc
    End With

End Sub

Private Sub DoAlignment()

  Dim i%
  Dim s$, mColor#, mCurrY#
  
   s = "This is same sample text for the textbox."

   With VBPrintPreview1
          SetPages "PrintPreview|Alignment"
          .StartDoc
            mColor = .ForeColor
            .FontSize = 11
            ' Understanding Alignment
            SetSubTitle "Understanding Alignment"
            .Paragraph "The TextAlign property is the one that let you set the alignment for text in paragraph, table cells and textboxes." & _
                        "This property allow you to set both the horizontal alignment of these objects." & _
                        "Following is a description of how alignment affects each object."
             .Paragraph
            
            ' Aligning Paragraphs
            SetSubTitle "Alignment in Paragraphs"
            .Paragraph "Paragraphs can only be aligned horizontally.  Vertical settings are ignored.  In order to align paragraphs, you should use only " & _
                      "taLeftTop, taCenterTop, taRightTop, or taJustifyTop."
            .Paragraph
            .FontSize = 18
            
            .TextAlign = taCenterTop
            .ForeColor = vbRed
            .Paragraph "Align Center"
            .ForeColor = mColor
            .Paragraph "This paragraph is aligned to the center.  Just set TextAlign property to taCenterTop."
            .Paragraph
            .TextAlign = taRightTop
            .ForeColor = vbRed
            .Paragraph "Align Right"
            .ForeColor = mColor
            .Paragraph "This paragraph is aligned to the right.  Just set TextAlign property to taRightTop."
            .Paragraph
            .TextAlign = taLeftTop
            .ForeColor = vbRed
            .Paragraph "Align Left"
            .ForeColor = mColor
            .Paragraph "This paragraph is aligned to the left.  Just set TextAlign property to taLeftTop."
            '.Paragraph
            
             SetNormal

            ' Aligning Table Cells and TextBoxes
             SetSubTitle "Aligninment in TextBoxes"
             mCurrY = .CurrentY + .ScaleX("0.5in", .ScaleMode, .ScaleMode)
             .FillStyle = vbFSSolid
             
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX, mCurrY, "1.8in", "0.8in", taLeftTop, True
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX + .ScaleX("1in", .ScaleMode, .ScaleMode) * 2, mCurrY, "1.8in", "0.8in", taCenterTop, True
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX + .ScaleX("1in", .ScaleMode, .ScaleMode) * 4, mCurrY, "1.8in", "0.8in", taRightTop, True
             
             mCurrY = .CurrentY + .ScaleX("0.5in", .ScaleMode, .ScaleMode)
             
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX, mCurrY, "1.8in", "0.8in", taLeftMiddle, True
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX + .ScaleX("1in", .ScaleMode, .ScaleMode) * 2, mCurrY, "1.8in", "0.8in", taCenterMiddle, True
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX + .ScaleX("1in", .ScaleMode, .ScaleMode) * 4, mCurrY, "1.8in", "0.8in", taRightMiddle, True
             
             
             mCurrY = .CurrentY + .ScaleX("0.5in", .ScaleMode, .ScaleMode)
             
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX, mCurrY, "1.8in", "0.8in", taLeftBottom, True
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX + .ScaleX("1in", .ScaleMode, .ScaleMode) * 2, mCurrY, "1.8in", "0.8in", taCenterBottom, True
             .FillColor = RGB(125 + Rnd * 125, 125 + Rnd * 125, 125 + Rnd * 125)
             .TextBox s, .CurrentX + .ScaleX("1in", .ScaleMode, .ScaleMode) * 4, mCurrY, "1.8in", "0.8in", taRightBottom, True
             
             
             '.FillStyle = vbFSTransparent
         .EndDoc
    End With
End Sub

Sub DoDialog()
    With VBPrintPreview1
       .Clear
       '.Orientation = PagePortrait
       .TextAlign = taLeftTop
       .Zoom = zmWholePage
        SetPages "PrintPreview|Print dialog"
       .StartDoc
            .Paragraph
             SetTitle "Print dialog"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            .FontSize = 12
            
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .Paragraph "[form.]VBPrintPreview.DialogPrint(pdPageSetup) [ = {True | False} ]"
            .Paragraph
            .Paragraph "[form.]VBPrintPreview.DialogPrint(pdPrinterSetup) [ = {True | False} ]"
            .Paragraph
            .Paragraph "[form.]VBPrintPreview.DialogPrint(pdPrint) [ = {True | False} ]"
            .FontBold = False
            .Paragraph
            .Paragraph "if DialogPrint(pdPrint) = True, then you can start the process for printing."
            .Paragraph
            .Paragraph "if select button print from NavBar then RaiseEvent PagePrint."
       .EndDoc
   End With

End Sub

Sub DoPageSetup()
     With VBPrintPreview1
          
          If .DialogPrint(pdPageSetup) Then
             DoDialog
             MsgBox "Select Button OK.", vbInformation
          Else
             MsgBox "Select Button Cancel.", vbInformation
          End If
     End With
End Sub

Sub DoPrintDialog()
     
     With VBPrintPreview1
          If .DialogPrint(pdPrint) Then
             DoDialog
             MsgBox "Select Button Print.", vbInformation
          Else
             MsgBox "Select Button Cancel.", vbInformation
          End If
     End With

End Sub

Sub DoPrintSetup()
     With VBPrintPreview1
          If .DialogPrint(pdPrinterSetup) Then
              DoDialog
              MsgBox "Select Button OK.", vbInformation
          Else
             MsgBox "Select Button Cancel.", vbInformation
          End If
     End With

End Sub
        
Sub DoHeaderFooter()
    With VBPrintPreview1
        .Clear
        .Zoom = zmWholePage
        .Orientation = PagePortrait
        .DocName = "Header_Footer Example"
         SetPages "PrintPreview|Header & Footer", .DocName
         Printer.NewPage
        .StartDoc
             SetTitle "Header & Footer"
            .TextAlign = taJustifyTop
            .ForeColor = vbBlack
            .Paragraph
            
            .Paragraph "Returns or sets the header and footer text."
            .Paragraph
            .FontBold = True
            .Paragraph "Syntax:"
            .Paragraph
            .Paragraph "[form.]VBPrintPreview.Header [ = value As String ]"
            .Paragraph
            .Paragraph "[form.]VBPrintPreview.Footer [ = value As String ]"
             .Paragraph
            '.FontName = "Courier New"
            .FontBold = False
            .FontItalic = False
            .FontSize = 10
            .Paragraph "The Header and Footer is composed of one to three sections, separated by pipe characters (''|'')." + _
                        " The first section is left-justified, the second is centered, and the third is right-justified."
             .Paragraph
            .Paragraph "You may include a page number field by embedding a ''p&'' code in the string."
            .Paragraph "For example, the following footer would print the file name and page number on the left and right " + _
                       "corners of every page:"
            .Paragraph
            
            .Paragraph "VBPrintPreview.Footer = VBPrintPreview.DocName & ''||Page %d''"
            .Paragraph
            .Paragraph "You may create multi-line footers by embedding line-feed characters within the Footer string. For example:"
            .Paragraph
            .Paragraph "VBPrintPreview1.Footer = ''Document:'' & vbCr & VBPrintPreview1.DocName & ''||Page'' & vbCr & ''%d''"
            .Paragraph
            .Paragraph "The color and font used to print the header and footer are defined by the HdrFont and HdrColor properties. " + _
                       "The position of the header is defined by the MarginHeader property." + _
                       "The position of the footer is defined by the MarginFooter property."
        .EndDoc
     End With
End Sub

Private Sub DoMovingThruPages()
Dim s$, fmt$, hdr$, bdy$

    With VBPrintPreview1
         
        .PageBorder = tbNone
        .StartDoc
        .FontSize = 16
        SetSubTitle "Navigating thru Pages"
        .Paragraph
        
        .Paragraph "VB PrintPreview allows you to move thru the document using the following property and methods:"
        .Paragraph
        .TextAlign = taLeftTop
        .FontSize = 14
        fmt = "^+2500tw|>+4000tw;"
        fmt = "^+4cm|>+8cm;"
        hdr = "Type|Description"
        bdy = "PageFirst|Move to the first page;" & _
              "PageNext|Move to the next page;" & _
              "PagePreview|Set the current preview page " + vbCr + "(first page is 1).;" & _
              "PageLast|Move to the last page;" & _
              "PageGoTo|Select the preview page number;" & _
              "CurrentPage|Returns the number of the page being printed.;" & _
              "PageCount|Returns the number of pages in the current document.;"
              
        .StartTable
            .Table fmt, hdr, bdy, 1, , , , , "1mm"
            .TableCell tcForeColor, 1, , &HFFFFFF
        .EndTable
        .EndDoc
    End With
    
End Sub

Sub SetCode()
    With VBPrintPreview1
        .FontName = "Courier New"
        .FontBold = False
        .FontItalic = False
        .FontSize = 10
        .ForeColor = vbBlue
     End With
End Sub
Sub SetNormal()
    
    With VBPrintPreview1
        .FontName = "Arial"
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontSize = 11
        .ForeColor = 0
        .FillColor = vbBlack
        .PageBorder = pbTopBottom
        .FillStyle = vbFSTransparent
        .DrawWidth = 0
    End With
    
End Sub


Private Sub SetOriginalSettings()
    
    With VBPrintPreview1
    
        .PaperSize = vbPRPSA4
        .NavBarMenu = "Whole Page|Page Width|ThunbNail"
        .ScaleMode = smCentimeters
        .Orientation = PagePortrait
        '.ToolTipText = ""
        ' font
        .FontName = "Arial"
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontSize = 11
        .FontTransparent = True
        
        ' text
        .ForeColor = 0
        .TextAlign = taLeftTop
        
        'spacing
        .LineSpace = lsSpaceSingle
        .IndentLeft = 0
        .IndentFirst = 0
        .IndentRight = 0
        
        'Margins
        .MarginRight = "20mm"
        .MarginLeft = "20mm"
        .MarginTop = "20mm"
        .MarginBottom = "20mm"
        .MarginHeader = "20mm"
        .MarginFooter = "20mm"
        
        'layout
        .PageBorder = pbTopBottom
        .Header = ""
        .HdrColor = &H0&
        .Footer = ""
        '.Columns = 1

        'drawing
        .DrawStyle = vbSolid
        .DrawWidth = 1
        
        ' table
        .TableBorder = tbAll
    
    End With

    List1.Visible = True

End Sub

Sub SetSubTitle(s$)
    Dim oF As Single
   With VBPrintPreview1
   
        .FontName = "Arial"
        .Paragraph ""
        .FontBold = True
        .FontItalic = True
        .FontUnderline = True
        oF = .FontSize
        .FontSize = 14
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
        .FontSize = oF
   End With
End Sub

Sub SetTitle(s$)

    With VBPrintPreview1
        .FontName = "Times New Roman"
        .FontBold = True
        .FontUnderline = True
        .FontItalic = True
        .FontSize = 18
        .ForeColor = vbBlue
        .Paragraph s
        .FontName = "Arial"
        .FontItalic = False
        .FontBold = False
        .FontUnderline = False
        .FontSize = 12
        .ForeColor = vbBlack
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

Public Sub SplitPath(FullPath As String, _
                     Optional Drive As String, _
                     Optional Path As String, _
                     Optional FileName As String, _
                     Optional File As String, _
                     Optional Extension As String)
 Dim nPos As Integer
 nPos = InStrRev(FullPath, "\")
 
 If nPos > 0 Then
    If Left$(FullPath, 2) = "\\" Then
        If nPos = 2 Then
            Drive = FullPath: Path = vbNullString: FileName = vbNullString: File = vbNullString
            Extension = vbNullString
            Exit Sub
        End If
    End If
    
    Path = Left$(FullPath, nPos - 1)
    FileName = Mid$(FullPath, nPos + 1)
    nPos = InStrRev(FileName, ".")
    
    If nPos > 0 Then
        File = Left$(FileName, nPos - 1)
        Extension = Mid$(FileName, nPos + 1)
    Else
        File = FileName
        Extension = vbNullString
    End If
 Else
    nPos = InStrRev(FullPath, ":")
    If nPos > 0 Then
        Path = Mid(FullPath, 1, nPos - 1): FileName = Mid(FullPath, nPos + 1)
        nPos = InStrRev(FileName, ".")
        If nPos > 0 Then
            File = Left$(FileName, nPos - 1)
            Extension = Mid$(FileName, nPos + 1)
        Else
            File = FileName
            Extension = vbNullString
        End If
    Else
        Path = vbNullString: FileName = FullPath
        nPos = InStrRev(FileName, ".")
        If nPos > 0 Then
            File = Left$(FileName, nPos - 1)
            Extension = Mid$(FileName, nPos + 1)
        Else
            File = FileName
            Extension = vbNullString
        End If
   End If
 End If
 
 If Left$(Path, 2) = "\\" Then
    nPos = InStr(3, Path, "\")
    If nPos Then
        Drive = Left$(Path, nPos - 1)
    Else
        Drive = Path
    End If
 Else
    If Len(Path) = 2 Then
        If Right$(Path, 1) = ":" Then
            Path = Path & "\"
        End If
    End If
    If Mid$(Path, 2, 2) = ":\" Then
        Drive = Left$(Path, 2)
    End If
 End If
 
End Sub
