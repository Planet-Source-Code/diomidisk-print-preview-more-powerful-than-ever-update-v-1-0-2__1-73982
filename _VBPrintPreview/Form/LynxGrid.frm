VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LynxGrid Tester"
   ClientHeight    =   8940
   ClientLeft      =   1215
   ClientTop       =   2445
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "With Colors "
      Height          =   225
      Left            =   3105
      TabIndex        =   3
      Top             =   195
      Width           =   1200
   End
   Begin PrintPreviewVB.VBPrintPreview VBPrintPreview1 
      Height          =   4005
      Left            =   3360
      TabIndex        =   2
      Top             =   675
      Visible         =   0   'False
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   7064
      CurrentX        =   2
      CurrentY        =   2,2
      FontCharSet     =   161
      Orientation     =   2
      Zoom            =   5
      Header          =   "VB Print Preview|Demo|LynxGrid"
      PageBorder      =   3
      ToPage          =   1
      NavBar          =   3
      PageWidth       =   29,66
      PageHeight      =   20,955
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      NavBarLabels    =   "Page 1"
   End
   Begin PrintPreviewVB.LynxGrid LynxGrid1 
      Height          =   2205
      Left            =   15
      TabIndex        =   1
      Top             =   600
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   3889
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12563634
      ForeColorSel    =   0
      CustomColorFrom =   16512244
      CustomColorTo   =   9601666
      GridColor       =   11246491
      FocusRectColor  =   4406585
      ThemeColor      =   1
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Preview"
      Height          =   465
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   2220
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum NameTypeEnum
   ntRandom = 0
   ntMale = 1
   ntFemale = 2
End Enum

'These are for generating the demo data
Private Const M_FORENAMES = "Alan,Alfie,Andrew,Ben,Bill,Bob,Boris,Brian,Charles,Chris,David,Gavin,Geoff,Grant,Harry,Ian," & _
                            "James,Jon,Mark,Matthew,Michael,Patrick,Paul,Peter,Richard,Robert,Samuel,Simon,Tony,Trevor,William"
Private Const F_FORENAMES = "Alicia,Alison,Amanda,Barbara,Caroline,Charlotte,Dawn,Hannah,Harriet,Hayley,Jane,Jennifer,Karen," & _
                            "Katie,Kerry,Kim,Lara,Laura,Lucy,Mary,Mellisa,Patricia,Paula,Rachel,Sarah,Stephanie,Susan,Tracy,Vanessa"
Private Const SURNAMES = "Anderson-Allen,Black,Bloggs,Brown,Clarke,Cole,Davis,Dawson,Evans,Gate,Johnson,Jones,Lawson,Lee," & _
                         "Richards,Ryan,Smith,Stephens,Temple,Turner,Wallace,White,Williams"

Private Const JOBS = "Accountant,Architect,Artist,Banker,Builder,Carpenter,Dentist,Director,Doctor,Engineer,Estate Agent," & _
                     "Fire Fighter,Gardener,Manager,Mechanic,Miner,Nurse,Optician,Pilot,Plumber,Police,Programmer,Scientist,Secretary,Shop" & _
                     " Assistant,Solicitor,Surgeon,Teacher,Truck Driver,Vet"

Private mCalled As Boolean

Private mMF() As String
Private mFF() As String
Private mSurnames() As String

Private mJobs() As String

Private gclrBack As OLE_COLOR

Private Sub CreateGrid()

   With LynxGrid1
      'Set ImageList to provide Item Images
      '.ImageList = ImageList1

      'Create the Columns
      .AddColumn "Code", 1000, , , ">"
      .AddColumn "G", 250
      .AddColumn "Forename", 1500
      .AddColumn "Surname", 1500, , , ">" '// Allow Only UPPERCASE
      .AddColumn "Job Title", 800, , , , , , True, , , True  '// This column is locked

      .AddColumn "Pension", 1000, lgAlignCenterCenter, lgBoolean
      .AddColumn "DOB", 1000, lgAlignCenterCenter, lgDate, "mm/dd/yyyy"
      .AddColumn "Premium Dollars and cents", 1600, lgAlignRightCenter, lgNumeric, "$#,#.00"
      .AddColumn "Notes", 5000
      .AddColumn "Button", 800, lgAlignCenterCenter, lgButton

      .ColImageAlignment(2) = lgAlignRightCenter

      .TotalsLineCaption(7) = "Total:"
      .TotalsLineShow = True

   End With

End Sub


Private Sub Check1_Click()
         Call PrintGrid(LynxGrid1)
End Sub

Private Sub cmdPrint_Click()
   If LynxGrid1.Visible = True Then
     LynxGrid1.Visible = False
     VBPrintPreview1.Visible = True
     cmdPrint.Caption = "LynxGrid"
     Call PrintGrid(LynxGrid1)
   Else
     LynxGrid1.Visible = True
     VBPrintPreview1.Visible = False
     cmdPrint.Caption = "Print Preview"
   End If
End Sub

Private Sub Form_Load()

   CreateGrid
   LoadDemoData

End Sub

Private Sub Form_Resize()

   If Not Me.WindowState = vbMinimized Then
      LynxGrid1.Height = Me.ScaleHeight - LynxGrid1.Top
      LynxGrid1.Width = Me.ScaleWidth - LynxGrid1.Left
      VBPrintPreview1.Left = LynxGrid1.Left
      VBPrintPreview1.Top = LynxGrid1.Top
      VBPrintPreview1.Height = LynxGrid1.Height
      VBPrintPreview1.Width = LynxGrid1.Width
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Form5 = Nothing

End Sub

Private Sub LoadDemoData()

  Dim lCount As Long
  Dim lRow   As Long
  Dim sForename As String
  Dim sGender   As String

   With LynxGrid1
      .Redraw = False

      'Add some random data

      For lCount = 1 To 50

         'Simple method to specify Gender!

         If RandomInt(0, 1) = 0 Then
            sGender = "M"
            sForename = GetForename(1) 'male
         Else
            sGender = "F"
            sForename = GetForename(2) 'female
         End If

         '// Add data to grid and return row number
         lRow = .AddItem(Format$("XD" & Format$(.ItemCount, "000")) & vbTab & _
                         sGender & vbTab & sForename & vbTab & _
                         GetSurname() & vbTab & _
                         GetJobName() & vbTab & _
                         (RandomInt(0, 1) = 0) & vbTab & _
                         DateSerial(RandomInt(1930, 1990), RandomInt(1, 12), RandomInt(1, 28)) & vbTab & _
                         Round(100 + (Rnd * 100), 2) & vbTab & _
                         vbTab & _
                         sGender)

         If sGender = "M" Then
            'Set the Key for the ImageList Image (can use text Key or Index)
            '.RowImage(lRow) = "MALE" & RandomInt(1, 3)
            .CellForeColor(lRow, 1) = vbBlue

         Else
            '.RowImage(lRow) = RandomInt(3, 6)
            .RowForeColor(lRow) = vbRed
            .CellForeColor(lRow, 1) = vbGreen
         End If

         '.CellImage(lRow, 2) = RandomInt(7, 9)

      Next lCount

      '// Lock Row #5
      .RowLocked(5) = True
      .CellText(5, 8) = "This Row is Locked" '// value change
      '.CellImage(5, 9) = RandomInt(7, 9)

      '.CellImage(8, 9) = RandomInt(7, 9)
      '.ColImageAlignment(9) = lgAlignCenterCenter

      'The grid supports per cell formatting but provides Item
      'formatting options for simplicity when only per Row formatting
      'is required (Row formatting reformats all Cells in the Row).
      .RowBackColor(5) = &H95E0F1
      .RowForeColor(5) = &H1F488A
       
      'Tell the grid to Draw
      .Redraw = True
   End With

End Sub

Private Sub LynxGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

   'Is the Edit allowed?
   Select Case Col
   Case 1 'Gender Column
      Cancel = True
   End Select

End Sub

Private Sub LynxGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

   MsgBox "clicked button on row#" & CStr(Row) & ", col#" & CStr(Col)

End Sub

Private Sub LynxGrid1_Click()

   If LynxGrid1.RowLocked(LynxGrid1.Row) Then
      MsgBox "This row is locked"

   ElseIf LynxGrid1.ColLocked(LynxGrid1.Col) Then
      MsgBox "This column is locked"
   End If

End Sub

Private Function GetForename(Optional nType As NameTypeEnum) As String

   Initialise

   Select Case nType
   Case 0 'ntRandom

      If RandomInt(0, 1) = 0 Then
         GetForename = mMF(RandomInt(LBound(mMF), UBound(mMF)))
      Else
         GetForename = mFF(RandomInt(LBound(mFF), UBound(mFF)))
      End If

   Case 1 'ntMale
      GetForename = mMF(RandomInt(LBound(mMF), UBound(mMF)))

   Case 2 'ntFemale
      GetForename = mFF(RandomInt(LBound(mFF), UBound(mFF)))

   End Select

End Function

Private Function GetJobName(Optional Index As Long = -1) As String

   Initialise

   If Index = -1 Then
      GetJobName = mJobs(RandomInt(LBound(mJobs), UBound(mJobs)))
   Else
      GetJobName = mJobs(Index)
   End If
   
End Function

Private Function GetSurname() As String

   Initialise

   GetSurname = mSurnames(RandomInt(LBound(mSurnames), UBound(mSurnames)))

End Function

Private Sub Initialise()

   If Not mCalled Then
      mCalled = True
      Randomize Timer

      mMF() = Split(M_FORENAMES, ",")
      mFF() = Split(F_FORENAMES, ",")
      mSurnames() = Split(SURNAMES, ",")

      mJobs() = Split(JOBS, ",")
   End If

End Sub

Private Function JobCount() As Long

   Initialise

   JobCount = UBound(mJobs)

End Function


Private Function RandomInt(lowerbound As Long, upperbound As Long) As Long

   RandomInt = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

End Function

Private Function rVal(ByVal vString As String) As Double

   '// MLH - New
   '// Returns the numbers contained in a string as a numeric value
   '// The Val function recognizes only the period (.) as a valid decimal separator.
   '// The CDbl errors on empty strings or values containing non-numeric values
   '// Returns the numbers contained in a string as a numeric value

  Dim lngI     As Long
  Dim lngS     As Long
  Dim bytAscV  As Byte
  Dim strTemp  As String
  
  On Error Resume Next

   vString = Trim$(UCase$(vString))
   
   If Left$(vString, 4) = "TRUE" Then
      rVal = True
      
   ElseIf Left$(vString, 5) = "FALSE" Then
      rVal = False
   
   Else
      Select Case Left$(vString, 2) '// Hex or Octal?
      Case Is = "&H", Is = "&O"
         lngS = 3
         strTemp = Left$(vString, 2)
      Case Else
         lngS = 1
      End Select
      
      For lngI = lngS To Len(vString)
         bytAscV = Asc(Mid$(vString, lngI, 1))
         Select Case bytAscV
         Case 48 To 57 '// 1234567890
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 44, 45, 46 '// , - .
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 36 '// $
            '// Ignore
            
         Case Is > 57, Is < 44
            If Left$(strTemp, 2) = "&H" Then '// Hex Values ?
               Select Case bytAscV
               Case 65 To 70 '// ABCDEF
                  strTemp = strTemp & Mid$(vString, lngI, 1)
               Case Else
                  Exit For
               End Select
            Else
               Exit For
            End If
         End Select
      Next lngI
      
      If LenB(strTemp) Then
         rVal = CDbl(strTemp)
      End If
   End If
   
   On Error GoTo 0

End Function


Public Sub PrintGrid(ByRef Grid As LynxGrid)
  
  Dim FormatCol As String, Header As String, Body As String
  Dim CA As String
  Dim r        As Long
  Dim C        As Long
  
   On Local Error GoTo ERR_Proc

   'READ DATA FROM LynxGrid
    FormatCol = ""
    For C = 0 To Grid.Cols - 1
        CA = ""
        Select Case Grid.ColAlignment(C)
        Case lgAlignLeftTop
              CA = "<"
        Case lgAlignLeftCenter
             CA = "<+"
        Case lgAlignLeftBottom
             CA = "<_"
        Case lgAlignCenterTop
             CA = "^"
        Case lgAlignCenterCenter
             CA = "^+"
        Case lgAlignCenterBottom
             CA = "^_"
        Case lgAlignRightTop
            CA = ">"
        Case lgAlignRightCenter
            CA = ">+"
        Case lgAlignRightBottom
            CA = ">_"
        End Select
        If FormatCol = "" Then
            FormatCol = CA + Str(Grid.ColWidth(C)) + "tw"
        Else
            FormatCol = FormatCol + "|" + CA + Str(Grid.ColWidth(C)) + "tw"
        End If
    Next
    
    Header = ""
    For C = 0 To Grid.Cols - 1
        If Header = "" Then
           Header = Grid.ColHeading(C)
        Else
        Header = Header + "|" + Grid.ColHeading(C)
        End If
    Next
    Body = ""
    For r = 0 To Grid.Rows - 1
      For C = 0 To Grid.Cols - 1
          If Body = "" Then
            Body = Grid.CellText(r, C)
          Else
            Body = Body + "|" + Grid.CellText(r, C)
          End If
      Next
      Body = Body + ";"
    Next

   With VBPrintPreview1
        .MarginBottom = "2.5cm"
        .MarginFooter = "2cm"
        .MarginHeader = "2cm"
        .MarginLeft = "2cm"
        .MarginRight = "2cm"
        .MarginTop = "2.5cm"
        .PaperSize = vbPRPSA4
        .PageBorder = pbTopBottom
        .Zoom = zmRation100
        Set .HdrFont = LynxGrid1.Font
        .HdrFontSize = 12
        .HdrFontBold = True
        .Header = "VbPrintPreview|Demo|LynxGrid"
        .Footer = "||Page p$"
        .Orientation = PageLandscape
        'Set .Font = LynxGrid1.Font
        
        .StartDoc
            .Paragraph
            .FontSize = Grid.Font.Size
            .StartTable
              .TableBorder = tbAll
              .Table FormatCol, Header, Body, , , , , False
            
               If Check1.Value = 1 Then
                  For r = 0 To Grid.Rows - 1
                     For C = 0 To Grid.Cols - 1
                        .TableCell tcBackColor, r + 1, , Grid.CellBackColor(r, C)
                        .TableCell tcForeColor, r + 1, C + 1, Grid.CellForeColor(r, C)
                     Next
                  Next
                  For r = 0 To Grid.Rows Step 2
                     .TableCell tcBackColor, r + 1, , Grid.GridColor
                  Next
              End If
              
              If .SendToPrinter = True Then
                'printer
                .TableCell tcRowHeight, , , .ScaleX(Grid.RowHeight(1), vbPixels, .ScaleMode) * 6.3
              Else
                'screen
                .TableCell tcRowHeight, , , .ScaleX(Grid.RowHeight(1), vbPixels, .ScaleMode)
              End If
              .EndTable
        .EndDoc
   End With
   Exit Sub
   
ERR_Proc:
   MsgBox "Error# " & Err.Number & vbNewLine & Err.Description, vbCritical, "LynxGrid.Export"
   Close

End Sub

Private Sub VBPrintPreview1_PagePrint()
       If VBPrintPreview1.DialogPrint(pdPrint) Then
           VBPrintPreview1.SendToPrinter = True
           Call PrintGrid(LynxGrid1)
           VBPrintPreview1.SendToPrinter = False
           cmdPrint_Click
        End If
End Sub
