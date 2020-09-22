VERSION 5.00
Begin VB.UserControl VBPrintPreview 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "PrintPreview.ctx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   7905
   ToolboxBitmap   =   "PrintPreview.ctx":0023
   Begin VB.PictureBox PicNaV 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   7905
      TabIndex        =   5
      Top             =   0
      Width           =   7905
      Begin VB.CommandButton CmdNav 
         Height          =   285
         Index           =   6
         Left            =   3570
         Picture         =   "PrintPreview.ctx":0335
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   30
         Width           =   465
      End
      Begin VB.CommandButton CmdNav 
         Height          =   285
         Index           =   5
         Left            =   3045
         Picture         =   "PrintPreview.ctx":06BF
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   25
         Width           =   315
      End
      Begin VB.CommandButton CmdNav 
         Height          =   290
         Index           =   4
         Left            =   2505
         Picture         =   "PrintPreview.ctx":0A49
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   25
         Width           =   525
      End
      Begin VB.CommandButton CmdNav 
         Height          =   285
         Index           =   3
         Left            =   2025
         Picture         =   "PrintPreview.ctx":0DD3
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   25
         Width           =   315
      End
      Begin VB.CommandButton CmdNav 
         Height          =   285
         Index           =   2
         Left            =   1710
         Picture         =   "PrintPreview.ctx":115D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   25
         Width           =   315
      End
      Begin VB.CommandButton CmdNav 
         Height          =   285
         Index           =   1
         Left            =   360
         Picture         =   "PrintPreview.ctx":14E7
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   25
         Width           =   315
      End
      Begin VB.CommandButton CmdNav 
         Height          =   285
         Index           =   0
         Left            =   60
         Picture         =   "PrintPreview.ctx":1871
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   25
         Width           =   315
      End
      Begin VB.Label NavBarLabel 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   4080
         TabIndex        =   20
         Top             =   60
         Width           =   2100
      End
      Begin VB.Label position 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6900
         TabIndex        =   17
         Top             =   15
         Width           =   45
      End
      Begin VB.Label LabelPages 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   675
         TabIndex        =   12
         Top             =   75
         Width           =   1020
      End
   End
   Begin VB.PictureBox picFullPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   210
      ScaleHeight     =   900
      ScaleWidth      =   765
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picPrintPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   150
      ScaleHeight     =   900
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   2175
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox PicViewPort 
      BackColor       =   &H00808080&
      Height          =   5400
      Left            =   1080
      ScaleHeight     =   5340
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   1080
      Width           =   4620
      Begin VB.PictureBox ThumbNail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   0
         Left            =   45
         ScaleHeight     =   900
         ScaleWidth      =   735
         TabIndex        =   16
         Top             =   30
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Left            =   -45
         Max             =   100
         TabIndex        =   15
         Top             =   5055
         Visible         =   0   'False
         Width           =   4290
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   4980
         LargeChange     =   10
         Left            =   4275
         Max             =   100
         TabIndex        =   14
         Top             =   30
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicBoxCorner 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   4275
         Picture         =   "PrintPreview.ctx":1BFB
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   5055
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox PagePicture 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4395
         Left            =   345
         MouseIcon       =   "PrintPreview.ctx":1F3D
         ScaleHeight     =   291
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   3405
      End
      Begin VB.PictureBox PicBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3885
         Left            =   150
         ScaleHeight     =   3885
         ScaleWidth      =   2880
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   2880
      End
   End
   Begin VB.PictureBox PiNAV 
      Height          =   270
      Left            =   30
      ScaleHeight     =   210
      ScaleWidth      =   5700
      TabIndex        =   18
      Top             =   45
      Visible         =   0   'False
      Width           =   5760
      Begin VB.Image Image3 
         Height          =   105
         Index           =   4
         Left            =   4260
         Picture         =   "PrintPreview.ctx":2247
         Top             =   75
         Width           =   105
      End
      Begin VB.Image Image3 
         Height          =   105
         Index           =   3
         Left            =   4005
         Picture         =   "PrintPreview.ctx":2309
         Top             =   60
         Width           =   105
      End
      Begin VB.Image Image3 
         Height          =   105
         Index           =   2
         Left            =   3810
         Picture         =   "PrintPreview.ctx":23CB
         Top             =   60
         Width           =   105
      End
      Begin VB.Image Image3 
         Height          =   105
         Index           =   1
         Left            =   3600
         Picture         =   "PrintPreview.ctx":248D
         Top             =   60
         Width           =   105
      End
      Begin VB.Image Image3 
         Height          =   105
         Index           =   0
         Left            =   3420
         Picture         =   "PrintPreview.ctx":254F
         Top             =   60
         Width           =   105
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   2880
         Picture         =   "PrintPreview.ctx":2611
         Top             =   -15
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   2520
         Picture         =   "PrintPreview.ctx":299B
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   2250
         Picture         =   "PrintPreview.ctx":2D25
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   2055
         Picture         =   "PrintPreview.ctx":30AF
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1830
         Picture         =   "PrintPreview.ctx":3439
         Top             =   15
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   1350
         Picture         =   "PrintPreview.ctx":37C3
         Top             =   -15
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   1050
         Picture         =   "PrintPreview.ctx":3B4D
         Top             =   15
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   750
         Picture         =   "PrintPreview.ctx":3ED7
         Top             =   -30
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "PrintPreview.ctx":4261
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   165
         Picture         =   "PrintPreview.ctx":45EB
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Menu mnuzoom 
      Caption         =   "Zoom"
      Begin VB.Menu mnzoom 
         Caption         =   "50 %"
         Index           =   0
      End
      Begin VB.Menu mnzoom 
         Caption         =   "75 %"
         Index           =   1
      End
      Begin VB.Menu mnzoom 
         Caption         =   "100 %"
         Index           =   2
      End
      Begin VB.Menu mnzoom 
         Caption         =   "150 %"
         Index           =   3
      End
      Begin VB.Menu mnzoom 
         Caption         =   "200 %"
         Index           =   4
      End
      Begin VB.Menu mnzoom 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnzoom 
         Caption         =   "Whole Page"
         Index           =   6
      End
      Begin VB.Menu mnzoom 
         Caption         =   "Page Width"
         Index           =   7
      End
      Begin VB.Menu mnzoom 
         Caption         =   "ThumbNail"
         Index           =   8
      End
   End
End
Attribute VB_Name = "VBPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mouseIsDown As Boolean
Dim cx As Single
Dim cy As Single

'===============
Public Enum PrinterConstants
    PD_ALLPAGES = &H0
    PD_COLLATE = &H10
    PD_DISABLEPRINTTOFILE = &H80000
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
    PD_NOPAGENUMS = &H8
    PD_NOSELECTION = &H4
    PD_NOWARNING = &H80
    PD_PAGENUMS = &H2
    PD_PRINTSETUP = &H40
    PD_PRINTTOFILE = &H20
    PD_RETURNDC = &H100
    PD_RETURNDEFAULT = &H400
    PD_RETURNIC = &H200
    PD_SELECTION = &H1
    PD_SHOWHELP = &H800
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    DLG_PRINT = 0
    DLG_PRINTSETUP = &H40
End Enum

Public Enum EPageSetup
    PSD_Defaultminmargins = &H0 ' Default (printer's)
    PSD_InWinIniIntlMeasure = &H0
    PSD_MINMARGINS = &H1
    PSD_MARGINS = &H2
    PSD_INTHOUSANDTHSOFINCHES = &H4
    PSD_INHUNDREDTHSOFMILLIMETERS = &H8
    PSD_DISABLEMARGINS = &H10
    PSD_DISABLEPRINTER = &H20
    PSD_NoWarning = &H80
    PSD_DISABLEORIENTATION = &H100
    PSD_ReturnDefault = &H400
    PSD_DISABLEPAPER = &H200
    PSD_ShowHelp = &H800
    PSD_EnablePageSetupHook = &H2000
    PSD_EnablePageSetupTemplate = &H8000
    PSD_EnablePageSetupTemplateHandle = &H20000
    PSD_EnablePagePaintHook = &H40000
    PSD_DisablePagePainting = &H80000
End Enum

Private Type PRINTER_DEFAULTS
    'Note:The definition of Printer_Defaults in the VB5 API viewer is incorrect.
    '      Below, pDevMode has been corrected to LONG.
    pDataType       As String
    pDevMode        As Long
    DesiredAccess   As Long
End Type

Private Type PRINTER_INFO_2
    pServerName     As Long
    pPrinterName    As Long
    pShareName      As Long
    pPortName       As Long
    pDriverName     As Long
    pComment        As Long
    pLocation       As Long
    pDevMode        As Long
    pSepFile        As Long
    pPrintProcessor As Long
    pDataType       As Long
    pParameters     As Long
    pSecurityDescriptor As Long
    Attributes      As Long
    Priority        As Long
    DefaultPriority As Long
    StartTime       As Long
    UntilTime       As Long
    Status          As Long
    cJobs           As Long
    AveragePPM      As Long
End Type

' --- API CONSTANTS
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_COPIES = &H100&
Private Const DM_ORIENTATION = &H1&

'Constants used to make changes to the values contained in the DevMode
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
'Private Const DM_IN_BUFFER      As Long = 8
'Private Const DM_OUT_BUFFER     As Long = 2
Private Const NULLPTR           As Long = 0&

Private Const PRINTER_ACCESS_ADMINISTER  As Long = &H4
Private Const PRINTER_ACCESS_USE         As Long = &H8
Private Const STANDARD_RIGHTS_REQUIRED   As Long = &HF0000
Private Const PRINTER_ALL_ACCESS         As Long = (STANDARD_RIGHTS_REQUIRED Or _
                                                   PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
       Left As Long
       Top As Long
       Right As Long
       Bottom As Long
End Type

Private Type RectSingle
       Left As Single
       Top As Single
       Right As Single
       Bottom As Single
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Type PageSetupDialog
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * 32 'CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32 'CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private InitialPaperSize As Long
Private InitialPaperOrit As Long
Private CallingHwnd As Long

' --- API DECLARATIONS
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDialog) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
'Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Any, ByVal pDevModeInput As Any, ByVal fMode As Long) As Long

'===============
Public MaxPageNumber As Integer
Private ViewPage As Integer
Private TempDir As String
' Storage for the Printer's original scale mode
Private pSM As Integer
' Storage for the Object's original scale mode
Private oSM As Integer
' Object used for Print Preview
Private ObjPrint As Control
Private oDW As Integer, oFC As Long, oFS As Integer, oCX As Single, oCy As Single, oDS As Integer

Private SetUnitPicture As RectSingle   'dimensions last Picture
Private SetUnitParagraph As RectSingle 'dimensions last Paragraph
Private SetUnitTable As RectSingle     'dimensions last Table
Private SetUnitTextBox As RectSingle   'dimensions last TextBox
Private TmpUnit As RectSingle          'tmp dimensions last print

'-------------
Private Type LogBrush ' 12 bytes
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
' Pen styles
Private Const PS_SOLID As Long = 0
Private Const PS_DASH As Long = 1
Private Const PS_DOT As Long = 2
Private Const PS_NULL As Long = 5
Private Const PS_INSIDEFRAME As Long = 6
Private Const PS_USERSTYLE As Long = 7
Private Const PS_ALTERNATE As Long = 8
Private Const PS_STYLE_MASK As Long = &HF

Private Const PS_ENDCAP_ROUND As Long = &H0
Private Const PS_ENDCAP_SQUARE As Long = &H100
Private Const PS_ENDCAP_FLAT As Long = &H200
Private Const PS_ENDCAP_MASK As Long = &HF00

Private Const PS_JOIN_ROUND As Long = &H0
Private Const PS_JOIN_BEVEL As Long = &H1000
Private Const PS_JOIN_MITER As Long = &H2000
Private Const PS_JOIN_MASK As Long = &HF000&

Private Const PS_COSMETIC As Long = &H0
Private Const PS_GEOMETRIC As Long = &H10000

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, RC As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As RECT) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LogBrush, ByVal dwStyleCount As Long, ByRef lpStyle As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LogBrush) As Long

Private iBKMode As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

'Font Character
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_IDEFAULTANSICODEPAGE = &H1004&

Private Type PanState
   X As Long
   Y As Long
End Type
Dim PanSet As PanState

Public Enum PaperSizeConstans
        vbPRPSLetter = 1 'Letter, 8 1/2 x 11 in.
        vbPRPSLetterSmall = 2 'Letter Small, 8 1/2 x 11 in.
        vbPRPSTabloid = 3 'Tabloid, 11 x 17 in.
        vbPRPSLedger = 4 'Ledger, 17 x 11 in.
        vbPRPSLegal = 5 'Legal, 8 1/2 x 14 in.
        vbPRPSStatement = 6 'Statement, 5 1/2 x 8 1/2 in.
        vbPRPSExecutive = 7 'Executive, 7 1/2 x 10 1/2 in.
        vbPRPSA3 = 8 'A3, 297 x 420 mm
        vbPRPSA4 = 9 'A4, 210 x 297 mm
        vbPRPSA4Small = 10 'A4 Small, 210 x 297 mm
        vbPRPSA5 = 11 'A5, 148 x 210 mm
        vbPRPSB4 = 12 'B4, 250 x 354 mm
        vbPRPSB5 = 13 'B5, 182 x 257 mm
        vbPRPSFolio = 14 'Folio, 8 1/2 x 13 in.
        vbPRPSQuarto = 15 'Quarto, 215 x 275 mm
        vbPRPS10x14 = 16 '10 x 14 in.
        vbPRPS11x17 = 17 '11 x 17 in.
        vbPRPSNote = 18 'Note, 8 1/2 x 11 in.
        vbPRPSEnv9 = 19 'Envelope #9, 3 7/8 x 8 7/8 in.
        vbPRPSEnv10 = 20 ' Envelope #10, 4 1/8 x 9 1/2 in.
        vbPRPSEnv11 = 21 'Envelope #11, 4 1/2 x 10 3/8 in.
        vbPRPSEnv12 = 22 'Envelope #12, 4 1/2 x 11 in.
        vbPRPSEnv14 = 23 'Envelope #14, 5 x 11 1/2 in.
        'vbPRPSCSheet = 24 'C size sheet
        'vbPRPSDSheet = 25 'D size sheet
        'vbPRPSESheet = 26 'E size sheet
        vbPRPSEnvDL = 27 'Envelope DL, 110 x 220 mm
        vbPRPSEnvC3 = 29 'Envelope C3, 324 x 458 mm
        vbPRPSEnvC4 = 30 'Envelope C4, 229 x 324 mm
        vbPRPSEnvC5 = 28 'Envelope C5, 162 x 229 mm
        vbPRPSEnvC6 = 31 'Envelope C6, 114 x 162 mm
        vbPRPSEnvC65 = 32 'Envelope C65, 114 x 229 mm
        vbPRPSEnvB4 = 33 'Envelope B4, 250 x 353 mm
        vbPRPSEnvB5 = 34 'Envelope B5, 176 x 250 mm
        vbPRPSEnvB6 = 35 'Envelope B6, 176 x 125 mm
        vbPRPSEnvItaly = 36 'Envelope, 110 x 230 mm
        vbPRPSEnvMonarch = 37 'Envelope Monarch, 3 7/8 x 7 1/2 in.
        vbPRPSEnvPersonal = 38 'Envelope, 3 5/8 x 6 1/2 in.
        vbPRPSFanfoldUS = 39 'U.S. Standard Fanfold, 14 7/8 x 11 in.
        vbPRPSFanfoldStdGerman = 40 'German Standard Fanfold, 8 1/2 x 12 in.
        vbPRPSFanfoldLglGerman = 41 'German Legal Fanfold, 8 1/2 x 13 in.
        vbPRPSUser = 256 'User-defined
End Enum

Public Enum LineSpaceConstants
    lsSpaceSingle = 0       'Single-line spacing (default value).
    lsSpaceLine15 = 1       '1.5 line spacing.
    lsSpaceDoubleline = 2   'Double line spacing.
    lsSpaceHalfline = 3     'Half line spacing
End Enum

Public Enum PrinterColorModeTypes
    cmMonochrome = vbPRCMMonochrome
    cmColor = vbPRCMColor
End Enum

Public Enum TextAlignConstants
          taLeftTop = 0  'Left top
         taRightTop = 1  'right top
        taCenterTop = 2  'Center top
       taJustifyTop = 3  'Justify top
       taLeftMiddle = 4  'Left Middle
      taRightMiddle = 5  'right Middle
     taCenterMiddle = 6  'Center Middle
    taJustifyMiddle = 7  'Justify Middle
       taLeftBottom = 8  'Left Bottom
      taRightBottom = 9  'right Bottom
     taCenterBottom = 10 'Center Bottom
    taJustifyBottom = 11 'Justify Bottom
End Enum

Public Enum PageOrientationConstants
    'PageOrientUndefined = 0
    PagePortrait = vbPRORPortrait
    PageLandscape = vbPRORLandscape
End Enum

Public Enum PrinterScaleMode
    smCentimeters = 7
   'smMillimeters = 6
    sminches = 5
End Enum

Public Enum ZoomModeConstants
    zmRation50 = 0
    zmRation75 = 1
    zmRation100 = 2
    zmRation150 = 3
    zmRation200 = 4
    zmWholePage = 5
    zmPageWidth = 6
    zmThumbnail = 7
    'zmTwoPages = 8
End Enum

Public Enum PageBorderConstants
     pbNone = 0
     pbBottom = 1
     pbTop = 2
     pbTopBottom = 3
     pbBox = 4
End Enum

Public Enum TableBorderConstants
     tbNone = 0
     tbBottom = 1
     tbTop = 2
     tbTopBottom = 3
     tbBox = 4
     tbColums = 5
     tbColTopBottom = 6
     tbAll = 7
     tbBoxRows = 8
     tbBoxColumns = 9
     tbBelowHeader = 10
End Enum

Public Enum NavBarSetting
     nbNone = 0        'Do not display the navigation bar.
     nbTop = 1         'Display a simple navigation bar at the top of the
     nbBottom = 2      'Display a simple navigation bar at the bottom of the control.
     nbTopPrint = 3     'Display a complete navigation bar at the top of the control (including a print button). This is the default setting.
     nbBottomPrint = 4 'Display a complete navigation bar at the bottom of the control (including a print button).
End Enum

Public Enum PrintDialogSettings
     pdPrinterSetup = 0  'Displays a Printer Setup dialog. This dialog is normally used before the document is created,
                         ' to select the target printer, paper size, and orientation.
     pdPageSetup = 1     'Displays a Page Setup dialog. This dialog is normally used before the document is created,
                         ' to select margins, paper size, and orientation.
     pdPrint = 2         'Displays a Print dialog. This dialog is used after the document is ready, to select the
                         ' range of pages to print and number of copies. (If the document is empty, this parameter
                         ' displays a Printer Setup dialog instead).
End Enum
Public Enum TableSettingConstants
     tcText = 0
     tcBackColor = 1
     tcForeColor = 2
     tcFontName = 3
     tcFontSize = 4
     tcFontCharSet = 5
     tcFontBold = 6
     tcFontItalic = 7
     tcFontUnderline = 8
     tcFontStrikethru = 9
     tcFontTransparent = 10
     tcPicture = 11
     tcColSpan = 12
     tcRowSpan = 13
     tcTextAling = 14
     tcCols = 15
     tcRows = 16
     tcColWidth = 17
     tcRowHeight = 18
     tcIndent = 19
End Enum
Public Enum PictureAlignConstants
    pLeftTop = 0       'Align to the left top corner of the rectangle.
    paCenterTop = 1    'Align to the center and to the top of the rectangle.
    paRightTop = 2     'Align the right top corner of the rectangle.
    paLeftBottom = 3   'Align to the left bottom corner of the rectangle.
    paCenterBottom = 4 'Align to the center and to the bottom of the rectangle.
    paRightBottom = 5  'Align to the right bottom corner of the rectangle.
    paLeftMiddle = 6   'Align to the left and to the middle of the rectangle.
    paCenterMiddle = 7 'Align to the center and to the middle of the rectangle.
    paRightMiddle = 8  'Align to the right and to the middle of the rectangle.
    paClip = 9         'Align to the center and to the middle of the rectangle (same as vppaCenterMiddle)
    paZoom = 10        'Fit rectangle while preserving aspect ratio.
    paStretch = 11     'Fill rectangle, stretching the picture as needed.
    paTile = 12        'Fill rectangle by tiling copies of the picture (useful for rendering backgrounds).
End Enum

Public Enum PolyFillMode
      ALTERNATE = 1
      WINDING = 2
End Enum

'Default Property Values:
Const m_def_NavBarMenu = "Whole Page|Page Width|ThunbNail"
Const m_def_NavBarLabels = ""
Const m_def_About = 0
Const m_def_DrawMode = vbCopyPen
Const m_def_PaperSize = 9
Const m_def_DocName = "Document"
Const m_def_HdrFontTransparent = True
Const m_def_HdrDrawWidth = 1
Const m_def_HdrDrawStyle = 0
Const m_def_PageBorderWidth = 1
Const m_def_PageBorderColor = 0
Const m_def_PageWidth = 0
Const m_def_PageHeight = 0
'Const m_def_PhysicalPage = False
Const m_def_X1 = 0
Const m_def_X2 = 0
Const m_def_Y1 = 0
Const m_def_Y2 = 0
Const m_def_IndentFirst = 0
Const m_def_IndentLeft = 0
Const m_def_IndentRight = 0
Const m_def_FillStyle = vbFSTransparent
Const m_def_FillColor = 0
Const m_def_HdrFontUnderline = 0
Const m_def_HdrFontStrikethru = 0
Const m_def_BackColor = &HFFFFFF
Const m_def_BackColorPage = &HFFFFFF
Const m_def_NavBar = 0
Const m_def_FromPage = 0
Const m_def_ToPage = 0
Const m_def_TableBorder = 0
Const m_def_FontTransparent = True
Const m_def_HdrFontName = "Times New Roman"
Const m_def_HdrFontBold = False
Const m_def_HdrFontItalic = True
Const m_def_HdrFontSize = 10
Const m_def_PageBorder = 0
Const m_def_PrinterColorMode = 2
Const m_def_Version = 0
Const m_def_Header = "VB Print Preview"
Const m_def_HdrColor = 0
Const m_def_Footer = "Page p$"
Const m_def_DrawStyle = 0
Const m_def_ZoomMode = 0
Const m_def_TextAlign = 0
Const m_def_LineSpace = 0
Const m_def_MarginHeader = 2
Const m_def_MarginFooter = 2
Const m_def_MarginTop = 2.2
Const m_def_MarginBottom = 2.2
Const m_def_MarginLeft = 2
Const m_def_MarginRight = 2
Const m_def_Zoom = 2
Const m_def_ScaleMode = 7
Const m_def_CurrentX = 1
Const m_def_CurrentY = 1
Const m_def_DrawWidth = 1
Const m_def_FontBold = 0
Const m_def_FontItalic = 0
Const m_def_FontName = "Arial"
Const m_def_FontSize = 8
Const m_def_FontCharSet = 0
Const m_def_FontStrikethru = 0
Const m_def_FontUnderline = 0
Const m_def_ForeColor = 0
Const m_def_Orientation = 1
Const m_def_SendToPrinter = False

'Property Variables:
Dim m_NavBarMenu As String
Dim m_NavBarLabels As String
Dim m_HdrFont As Font
Dim m_Font As Font
Dim m_About As Variant
Dim m_DrawMode As DrawModeConstants
Dim m_PaperSize As PaperSizeConstans
Dim m_DocName As String
Dim m_HdrFontTransparent As Boolean
Dim m_HdrDrawWidth As Variant
Dim m_HdrDrawStyle As Integer
Dim m_PageBorderWidth As Integer
Dim m_PageBorderColor As OLE_COLOR
Dim m_PageWidth As Single
Dim m_PageHeight As Single
Dim m_PhysicalPage As Boolean
Dim m_X1 As Single
Dim m_X2 As Single
Dim m_Y1 As Single
Dim m_Y2 As Single
Dim m_IndentFirst As Variant
Dim m_IndentLeft As Variant
Dim m_IndentRight As Variant
Dim m_FillStyle As Integer
Dim m_FillColor As OLE_COLOR
Dim m_HdrFontUnderline As Boolean
Dim m_HdrFontStrikethru As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_BackColorPage As OLE_COLOR
Dim m_NavBar As NavBarSetting
Dim m_FromPage As Integer
Dim m_ToPage As Integer
Dim m_TableBorder As Integer
Dim m_FontTransparent As Boolean
Dim m_HdrFontName As String
Dim m_HdrFontBold As Boolean
Dim m_HdrFontItalic As Boolean
Dim m_HdrFontSize As Integer
Dim m_PageBorder As Integer
Dim m_PrinterColorMode As Integer
Dim m_Header As String
Dim m_HdrColor As OLE_COLOR
Dim m_MarginHeader As Single
Dim m_Footer As String
Dim m_MarginFooter As Single
Dim m_DrawStyle As Integer
Dim m_TextAlign As Integer
Dim m_LineSpace As Integer
Dim m_MarginTop As Single
Dim m_MarginBottom As Single
Dim m_MarginLeft As Single
Dim m_MarginRight As Single
Dim m_Zoom As ZoomModeConstants
Dim m_ScaleMode As ScaleModeConstants
Dim m_CurrentX As Single
Dim m_CurrentY As Single
Dim m_DrawWidth As Integer
Dim m_FontBold As Boolean
Dim m_FontItalic As Boolean
Dim m_FontName As String
Dim m_FontSize As Integer
Dim m_FontCharSet As Integer
Dim m_FontStrikethru As Boolean
Dim m_FontUnderline As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_Orientation As Integer
Dim m_SendToPrinter As Boolean

Event AfterUserScroll()
Event AfterFooter()
Event AfterHeader()
Event AfterTableCell(ByVal Row As Integer, ByVal Col As Integer, ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single, Text As String)
Event BeforeFooter()
Event BeforeHeader()
Event BeforeUserZoom()
Event BeforeTableCell(ByVal Row As Integer, ByVal Col As Integer, Text As String)
Event AfterTableEnd()
Event AfterEndDoc()
Event PageEndDoc()
Event PageEnd()
Event AfterUserZoom()
Event Error(ByVal id As Long, ByVal ErrorDescription)
Event PageView()
Event PagePrint()
Event PageNew()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Type nTableCell
                tbcCol As Integer
                tbcRow As Integer
               tbcText As String
           tbcColAlign As TextAlignConstants
          tbcBackColor As Long
          tbcForeColor As Long
           tbcFontName As String
           tbcFontSize As Integer
        tbcFontCharSet As Integer
           tbcFontBold As Boolean
         tbcFontItalic As Boolean
      tbcFontUnderline As Boolean
     tbcFontStrikethru As Boolean
    tbcFontTransparent As Boolean
            tbcPicture As StdPicture
       tbcPictureAlign As AlignmentConstants
            tbcColSpan As Integer
            tbcRowSpan As Integer
End Type
Private Type nTable
          tbCols As Integer
          tbRows As Integer
     tbLineWidth As Integer
     tbLineColor As Long
    tbColWidth() As Single
    tbColAlign() As TextAlignConstants
   tbRowHeight() As Single
      tbHeader() As String
        tbIndent As Single
      tbWordWrap As Boolean
   tbTableCell() As nTableCell
End Type
Private pTable As nTable

Public Sub About()
         MsgBox "VBPrintPreview v1.0.2" & vbNewLine & _
                "A PrintPreview ActiveX Control" & vbNewLine & vbNewLine & _
                "Created by: Diomidis G. Kiriakopoulos.", vbInformation + vbOKOnly, "About"
End Sub

Public Function FileExists(Path$) As Boolean
    If Len(Trim(Path$)) = 0 Then FileExists = False: Exit Function
    FileExists = Dir(Trim(Path$), vbNormal) <> ""
End Function

Private Sub PicRefresh(PicDest As PictureBox, PicSrc As PictureBox)
  
    If Zoom = zmRation100 Then
       PicDest.Width = PicSrc.Width
       PicDest.Height = PicSrc.Height
    End If

    If m_Zoom = zmRation100 Then
       PicDest.PaintPicture PicSrc.Picture, _
                              0, 0, PicDest.ScaleWidth, PicDest.ScaleHeight, _
                              , , , , vbSrcCopy
    Else
       Call SetStretchBltMode(PicDest.hdc, vbPaletteModeNone)
       Call StretchBlt(PicDest.hdc, 0, 0, PicDest.Width, PicDest.Height, PicSrc.hdc, 0, 0, PicSrc.Width, PicSrc.Height, vbSrcCopy)
    End If
    PicDest.Refresh

End Sub

Private Sub CmdNav_Click(Index As Integer)
  
    Select Case Index
    Case 0
        PageFirst
    Case 1
        PagePreview
    Case 2
        PageNext
    Case 3
        PageLast
    Case 4
        If m_Zoom = zmRation100 Then
           Zoom = zmWholePage
        Else
           Zoom = zmRation100
        End If
    Case 5
        PopupMenu mnuzoom, vbPopupMenuRightButton
        Exit Sub
    Case 6
        RaiseEvent PagePrint
    End Select
   ' PagePicture.SetFocus
End Sub

Private Sub HScroll1_Change()
 If HScroll1.Visible = False Then Exit Sub
    On Local Error Resume Next
    PagePicture.Left = -(HScroll1.Value)
    PicBack.Left = PagePicture.Left + 15
    HScroll1.SetFocus
    On Local Error GoTo 0
End Sub

Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyLeft 'Arrow left
        If HScroll1.Value - HScroll1.SmallChange > HScroll1.Min Then
           HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        Else
           HScroll1.Value = HScroll1.Min
        End If
    Case vbKeyRight 'Arrow right
        If HScroll1.Value + HScroll1.SmallChange < HScroll1.Max Then
           HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        Else
           HScroll1.Value = HScroll1.Max
        End If
    Case 33 'PageUp
        PagePreview
    Case 34 'PageDown
        PageNext
    Case vbKeyG  'G
        PageGoTo
     Case vbKeyUp, vbKeyDown
          VScroll1_KeyUp KeyCode, Shift
          VScroll1.SetFocus
    End Select
End Sub

Private Sub mnzoom_Click(Index As Integer)
    Select Case Index
    Case 0
       Zoom = zmRation50 '0
    Case 1
       Zoom = zmRation75 '1
    Case 2
       Zoom = zmRation100 '2
    Case 3
       Zoom = zmRation150 '3
    Case 4
       Zoom = zmRation200 '4
    Case 5 'sep
    Case 6
       Zoom = zmWholePage '5
    Case 7
       Zoom = zmPageWidth '6
    Case 8
       Zoom = zmThumbnail '7
    Case 9
       'Zoom = zmTwoPages '= 8
    End Select
End Sub

Private Sub PagePicture_DblClick()
     If m_Zoom = 4 Then Exit Sub
     m_Zoom = m_Zoom + 1
     If m_Zoom > 4 Then
       Zoom = zmRation200
     Else
       Zoom = m_Zoom
     End If
End Sub

Private Sub PicViewPort_DblClick()
     If m_Zoom = 4 Then Exit Sub
     m_Zoom = m_Zoom + 1
     If m_Zoom > 4 Then
       Zoom = zmRation200
     Else
     Zoom = m_Zoom
     End If
End Sub

Private Sub PicViewPort_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      If m_Zoom = 0 Then Exit Sub
     m_Zoom = m_Zoom - 1
     If m_Zoom < 0 Then
        Zoom = zmRation50
     Else
       Zoom = m_Zoom
     End If
    End If
End Sub

Private Sub ThumbNail_Click(Index As Integer)
         Dim uPic As PictureBox
         ViewPage = Index
         For Each uPic In ThumbNail
            If uPic.Index > 0 Then
              Unload uPic
            End If
         Next
         ThumbNail(0).Visible = False
        Zoom = zmRation100
End Sub

Private Sub UserControl_Initialize()

    TempDir = Environ("TEMP") & "\"
    
    If FileExists(TempDir & "PPView0.bmp") Then Kill TempDir & "PPView0.bmp"
    
    Orientation = Printer.Orientation
   
   If m_BackColorPage = 0 Then m_BackColorPage = m_def_BackColorPage
   FontCharSet = CharSetWin
    ScaleMode = smCentimeters
    SendToPrinter = False
'    If PageWidth = 0 Or PageHeight = 0 Then
'       PhysicalPage = False
'    End If
    If PaperSize = 0 Then PaperSize = vbPRPSA4
     
    StartDoc
    UserControl_Resize
    EndDoc
    m_Zoom = zmWholePage
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
      VScroll1_KeyUp KeyCode, Shift
End Sub

Private Sub VScroll1_Change()
    On Local Error Resume Next
    
    If Zoom <> zmThumbnail Then
        PagePicture.Top = -(VScroll1.Value)
        PicBack.Top = PagePicture.Top + 15
        If VScroll1.Visible = True Then VScroll1.SetFocus
    Else
        Dim Col As Single, Row As Single, mTop As Integer, wStep As Long, mLeft As Long, l As Integer
        Dim X As Integer, Y As Integer, AA As Integer, nRow As Integer, vsmax As Integer, i As Integer
        If ThumbNail(0).Width > ThumbNail(0).Height Then
            wStep = ThumbNail(0).Width
        Else
            wStep = ThumbNail(0).Height
        End If
    
        AA = 0
        mLeft = 100
        mTop = -VScroll1.Value * wStep
        l = 0
    
        For i = 0 To MaxPageNumber
            If mLeft + (wStep * 2) + 50 > PicViewPort.ScaleWidth Then
                mTop = mTop + wStep + 50
                mLeft = 150
                l = 1
                AA = AA + 1
            Else
                mLeft = mLeft + wStep * l + 50
                If l = 0 Then l = l + 1
            End If
            Col = (wStep - ThumbNail(i).Width) / 2
            Row = (wStep - ThumbNail(i).Height) / 2
            ThumbNail(i).Move mLeft + Col, mTop + Row
            ThumbNail(i).Enabled = True
            ThumbNail(i).Visible = True
        Next
    End If
    
    On Local Error GoTo 0
    
    
    RaiseEvent AfterUserScroll
    
End Sub

Private Sub VScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyUp ', vbKeyPageUp  'Arrow left, PageUp
        If VScroll1.Visible = False Then
            PagePreview
        Else
            If VScroll1.Value - VScroll1.SmallChange > VScroll1.Min Then
               VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
            Else
               VScroll1.Value = VScroll1.Min
            End If
        End If
    Case vbKeyDown ', vbKeyPageDown   'Arrow down, PageDown
        If VScroll1.Visible = False Then
            PageNext
        Else
           If VScroll1.Value + VScroll1.SmallChange < VScroll1.Max Then
              VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
           Else
              VScroll1.Value = VScroll1.Max
           End If
        End If
    Case vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown, vbKeyG
       HScroll1_KeyUp KeyCode, Shift
       'HScroll1.SetFocus
    End Select
    
End Sub

Private Sub PagePicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Local Error Resume Next
 
   If Button = vbLeftButton And Shift = 0 Then
      PanSet.X = X
      PanSet.Y = Y
      PagePicture.MousePointer = vbCustom
   End If
 
   Measuring X, Y
   RaiseEvent MouseDown(Button, Shift, X, Y)
     
End Sub

Private Sub PagePicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim nTop As Integer, nLeft As Integer
      
   On Local Error Resume Next

   If Button = vbLeftButton And Shift = 0 Then

      'new coordinates
      With PagePicture
         nTop = -(.Top + (Y - PanSet.Y))
         nLeft = -(.Left + (X - PanSet.X))
      End With

      'Check limits
      With VScroll1
         If .Visible Then
            If nTop < .Min Then
               nTop = .Min
            ElseIf nTop > .Max Then
               nTop = .Max
            End If
         Else
            nTop = -PagePicture.Top
         End If
      End With

      With HScroll1
         If .Visible Then
            If nLeft < .Min Then
               nLeft = .Min
            ElseIf nLeft > .Max Then
               nLeft = .Max
            End If
         Else
            nLeft = -PagePicture.Left
         End If
      End With
      
      PagePicture.Move -nLeft, -nTop
      PicBack.Move -nLeft + 30, -nTop + 15
   End If
   
   Measuring X, Y
   RaiseEvent MouseMove(Button, Shift, X, Y)
 
End Sub

Private Sub PagePicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Local Error Resume Next
   
   If Button = vbLeftButton And Shift = 0 Then
      If VScroll1.Visible Then VScroll1.Value = -(PagePicture.Top)
      If HScroll1.Visible Then HScroll1.Value = -(PagePicture.Left)
   End If
   PagePicture.MousePointer = vbDefault
   
   Measuring X, Y
   RaiseEvent MouseUp(Button, Shift, X, Y)
   
   If Button = 2 Then
     If m_Zoom = 0 Then Exit Sub
      m_Zoom = m_Zoom - 1
     If m_Zoom < 0 Then
       Zoom = zmRation50
     Else
       Zoom = m_Zoom
     End If
   End If
    
End Sub

Private Sub UserControl_Resize()
       
    If UserControl.Width < VScroll1.Width Then Exit Sub
    If m_NavBar > 0 Then
       If UserControl.Width < 6000 Then UserControl.Width = 6000
    ElseIf m_NavBar = nbNone Then
       If UserControl.Width < 4000 Then UserControl.Width = 4000
    End If
    
    If UserControl.Height < 4000 Then UserControl.Height = 4000
    
    If m_NavBar = nbNone Then
       PicViewPort.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
       HScroll1.Move 20, PicViewPort.Top + PicViewPort.Height - HScroll1.Height, UserControl.Width - VScroll1.Width
       VScroll1.Move PicViewPort.Left + PicViewPort.Width - VScroll1.Width, PicViewPort.Top
       PicNaV.Visible = False
    Else
      
      If m_NavBar = nbTop Or m_NavBar = nbTopPrint Then
         If m_NavBar = nbTop Then CmdNav(6).Visible = False Else CmdNav(6).Visible = True
         PicNaV.Align = 1
         PicViewPort.Move 0, PicNaV.Height, UserControl.ScaleWidth, UserControl.ScaleHeight - PicNaV.Height
         HScroll1.Move 20, PicViewPort.Top + PicViewPort.Height + -HScroll1.Height + PicNaV.Height, UserControl.Width - VScroll1.Width
         VScroll1.Move PicViewPort.Left + PicViewPort.Width - VScroll1.Width + PicNaV.Height, PicViewPort.Top
      ElseIf m_NavBar = nbBottom Or m_NavBar = nbBottomPrint Then
         If m_NavBar = nbBottom Then CmdNav(6).Visible = False Else CmdNav(6).Visible = True
         PicNaV.Align = 2
         PicViewPort.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - PicNaV.Height
         HScroll1.Move 20, PicViewPort.Top + PicViewPort.Height + -HScroll1.Height + PicNaV.Height, UserControl.Width - VScroll1.Width
         VScroll1.Move PicViewPort.Left + PicViewPort.Width - VScroll1.Width + PicNaV.Height, PicViewPort.Top
      End If
      
       'tDrawNav
       Dim Rct As RECT
       With PicNaV
         .Cls
         .ScaleMode = vbPixels
         Rct.Left = .ScaleLeft
         Rct.Top = .ScaleTop
         Rct.Right = .ScaleWidth
         Rct.Bottom = .ScaleHeight
      End With
      DrawEdge PicNaV.hdc, Rct, CLng(&H1 Or &H4), &H1 Or &H2 Or &H4 Or &H8
      PicNaV.Visible = True
    End If
    
    
    If HScroll1.Visible = True Then
       VScroll1.Height = PicViewPort.Height - HScroll1.Height
    Else
       VScroll1.Height = PicViewPort.Height
    End If
   
    PagePicture.Visible = False
    PicBack.Visible = False
    
    If PaperSize = 0 Then PaperSize = vbPRPSA4
    
    Call DisplayPages
    
    position.Move PicNaV.ScaleWidth - position.Width - PicNaV.TextWidth("W"), (PicNaV.ScaleHeight - position.Height) / 2
    If position.Left < CmdNav(5).Left + CmdNav(5).Width Then position.Visible = False Else position.Visible = True
    Debug.Print "Resize"
    
End Sub

Private Sub Measuring(X As Single, Y As Single)

   Dim tX As Single, ty As Single, sm As ScaleModeConstants, tsm As String
   
   sm = ObjPrint.ScaleMode
   
   Select Case m_Zoom
   Case zmRation50 '0
        tX = Round(ScaleX(X / 0.5, sm, ScaleMode), 2)
        ty = Round(ScaleY(Y / 0.5, sm, ScaleMode), 2)
   Case zmRation75 ' 1
        tX = Round(ScaleX(X / 0.75, sm, ScaleMode), 2)
        ty = Round(ScaleY(Y / 0.75, sm, ScaleMode), 2)
   Case zmRation100 ' 2
        tX = Round(ScaleX(X, sm, ScaleMode), 2)
        ty = Round(ScaleY(Y, sm, ScaleMode), 2)
   Case zmRation150 ' 3
        tX = Round(ScaleX(X / 1.5, sm, ScaleMode), 2)
        ty = Round(ScaleY(Y / 1.5, sm, ScaleMode), 2)
   Case zmRation200 ' 4
        tX = Round(ScaleX(X / 2, sm, ScaleMode), 2)
        ty = Round(ScaleY(Y / 2, sm, ScaleMode), 2)
   Case zmWholePage, zmPageWidth ' 5, 6
        tX = PagePicture.ScaleWidth
        ty = PagePicture.ScaleHeight
        tX = Round(ScaleX(tX, sm, m_ScaleMode), 2)
        ty = Round(ScaleY(ty, sm, m_ScaleMode), 2)
        tX = Round(ScaleX(X / (tX / Printer.ScaleWidth), sm, m_ScaleMode), 2)
        ty = Round(ScaleY(Y / (ty / Printer.ScaleHeight), sm, m_ScaleMode), 2)
   Case zmThumbnail
        position.Caption = ""
   End Select
   
   X = tX
   Y = ty
   If ScaleMode = smCentimeters Then tsm = "cm" Else tsm = "in"
   position.Caption = "X:" + Format(X, "0.00") + tsm + " Y:" + Format(Y, "0.00") + tsm + " "
End Sub

Private Sub DisplayPages()
        

    Dim Xmin As Single
    Dim Ymin As Single
    Dim wid As Single
    Dim hgt As Single
    Dim Aspect As Single
  
    Dim Have_wid As Single
    Dim Have_hgt As Single
    Dim Need_wid As Single
    Dim Need_hgt As Single
    Dim Need_HScroll1 As Boolean
    Dim Need_VScroll1 As Boolean
  
   If FileExists(TempDir & "PPView" & CStr(ViewPage) & ".bmp") = True Then
     ' picFullPage.Width = 100
     ' picFullPage.Height = 100
      picFullPage.Picture = LoadPicture()
      picFullPage.Picture = LoadPicture(TempDir & "PPView" & CStr(ViewPage) & ".bmp")
      picFullPage.AutoSize = True
      If picFullPage.Width > picFullPage.Height Then
         Orientation = PageLandscape
      Else
         Orientation = PagePortrait
      End If
   Else
      Exit Sub
   End If
   
   PagePicture.Visible = False
   PicBack.Visible = False
   
   If PicViewPort.Height < 180 Then Exit Sub
   
   Select Case m_Zoom
   Case 0:
        PagePicture.Width = ScaleX(PageWidth, ScaleMode, vbTwips) * 0.5 ' Printer.Width * 0.5
        PagePicture.Height = ScaleY(PageHeight, ScaleMode, vbTwips) * 0.5
   Case 1:
        PagePicture.Width = ScaleX(PageWidth, ScaleMode, vbTwips) * 0.75
        PagePicture.Height = ScaleY(PageHeight, ScaleMode, vbTwips) * 0.75
   Case 2:
        PagePicture.Width = ScaleX(PageWidth, ScaleMode, vbTwips)
        PagePicture.Height = ScaleY(PageHeight, ScaleMode, vbTwips)
   Case 3:
        PagePicture.Width = ScaleX(PageWidth, ScaleMode, vbTwips) * 1.5
        PagePicture.Height = ScaleY(PageHeight, ScaleMode, vbTwips) * 1.5
   Case 4:
        PagePicture.Width = ScaleX(PageWidth, ScaleMode, vbTwips) * 2
        PagePicture.Height = ScaleY(PageHeight, ScaleMode, vbTwips) * 2
   Case 5
        'Clear any picture and set the size and loaction
        PageWidth = picFullPage.ScaleWidth
        PageHeight = picFullPage.ScaleHeight
        PagePicture.Width = ScaleX(PageWidth, ScaleMode, vbTwips)
        PagePicture.Height = ScaleY(PageHeight, ScaleMode, vbTwips) 'Printer.Height
        
        If (PageWidth / PageHeight) < 1 Then
            picFullPage.Height = PicViewPort.Height
            picFullPage.Width = picFullPage.Height * (PageWidth / PageHeight)
        Else
            picFullPage.Width = PicViewPort.Width
            picFullPage.Height = picFullPage.Width / (PageWidth / PageHeight)
        End If
        
        picFullPage.Cls
        picFullPage.BackColor = BackColorPage
        picFullPage.Move (PicViewPort.Width - picFullPage.Width) \ 2, (PicViewPort.Height - picFullPage.Height) \ 2
     
        'Get the scale values
        Aspect = PagePicture.Height / PagePicture.Width
        wid = picFullPage.Width
        hgt = picFullPage.Height

        'MaintainRatio
        If (hgt / wid) > Aspect Then
            hgt = Aspect * wid
            Xmin = picFullPage.Left
            Ymin = (picFullPage.Height - hgt) / 2
        Else
            Xmin = (picFullPage.Width - wid) / 2
            Ymin = picFullPage.Top
        End If
        PagePicture.Move Xmin, 0, wid, hgt
        If FileExists(TempDir & "PPView" & CStr(ViewPage) & ".bmp") = True Then
              picFullPage.Picture = LoadPicture(TempDir & "PPView" & CStr(ViewPage) & ".bmp")
        End If
        
   Case 6 'zmPageWidth
       ' PageWidth = Printer.ScaleWidth
       ' PageHeight = Printer.ScaleHeight
        picFullPage.Width = PicViewPort.ScaleWidth '- IIf(VScroll1.Visible = True, VScroll1.Width + 60, 0) '- 60
        picFullPage.Height = PicViewPort.Width * (PageHeight / PageWidth)
        picFullPage.Cls
        picFullPage.BackColor = BackColorPage
        If picFullPage.Height < PicViewPort.Height Then
           PagePicture.Move (PicViewPort.Width - picFullPage.Width) / 2, (PicViewPort.Height - picFullPage.Height) / 2, picFullPage.Width, picFullPage.Height
        Else
           PagePicture.Move 0, 0, picFullPage.Width, picFullPage.Height
        End If
      
        If FileExists(TempDir & "PPView" & CStr(ViewPage) & ".bmp") = True Then
            picFullPage.Picture = LoadPicture(TempDir & "PPView" & CStr(ViewPage) & ".bmp")
        End If
        
   Case 7
        position.Caption = ""
        CreateThumbNail
        Exit Sub
   End Select
   
   picFullPage.ScaleMode = PagePicture.ScaleMode
   PicRefresh PagePicture, picFullPage
   
    Need_wid = PagePicture.Width
    Need_hgt = PagePicture.Height
    Have_wid = PicViewPort.Width
    Have_hgt = PicViewPort.Height

    ' See which scroll bars we need.
    Need_HScroll1 = (Need_wid > Have_wid)
    If Need_HScroll1 Then Have_hgt = Have_hgt - HScroll1.Height

    Need_VScroll1 = (Need_hgt > Have_hgt)
    If Need_VScroll1 Then
        ' This takes away a little width so we might need the horizontal scroll bar now.
        Have_wid = Have_wid - VScroll1.Width
        If Not Need_HScroll1 Then
            Need_HScroll1 = (Need_wid > Have_wid)
            If Need_HScroll1 Then Have_hgt = Have_hgt - HScroll1.Height
        End If
    End If
    
    If m_Zoom < 6 Then
        If PagePicture.Width > PicViewPort.Width - VScroll1.Width And PagePicture.Height > PicViewPort.Height - HScroll1.Height Then
            VScroll1.Value = VScroll1.Min '0
            HScroll1.Value = HScroll1.Min '0
            PagePicture.Top = -(VScroll1.Value)
            PagePicture.Left = -(HScroll1.Value)
        Else
            If PagePicture.Height > PicViewPort.Height Then
                VScroll1.Value = VScroll1.Min  '0
                PagePicture.Top = -(VScroll1.Value)
                PagePicture.Left = (PicViewPort.Width - PagePicture.Width) / 2
            Else
                PagePicture.Move (PicViewPort.Width - PagePicture.Width) / 2, (PicViewPort.Height - PagePicture.Height) / 2
            End If
        End If
    End If
    
    PicBack.Move PagePicture.Left + 15, PagePicture.Top + 15, PagePicture.Width + 15, PagePicture.Height + 15
        
    ' Position or hide the scroll bars.
    If Need_HScroll1 Then
       HScroll1.Move 0, Have_hgt - 60, Have_wid - 60
       HScroll1.Min = -250
       HScroll1.Max = Abs(PagePicture.Width - PicViewPort.Width + 700)
       HScroll1.SmallChange = HScroll1.Max * 0.1
       HScroll1.LargeChange = HScroll1.Max * 0.2
       HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If

    If Need_VScroll1 Then
        VScroll1.Move Have_wid - 60, 0, VScroll1.Width, Have_hgt - 60
        VScroll1.Min = -250
        VScroll1.Max = Abs(PagePicture.Height - PicViewPort.Height + 500)
        VScroll1.SmallChange = VScroll1.Max * 0.1
        VScroll1.LargeChange = VScroll1.Max * 0.2
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    
    PicBoxCorner.Move VScroll1.Left, VScroll1.Top + VScroll1.Height, VScroll1.Width, HScroll1.Height
    PicBoxCorner.ZOrder
    
    If Need_HScroll1 = False Or Need_VScroll1 = False Then
       PicBoxCorner.Visible = False
    ElseIf Need_HScroll1 = True Or Need_VScroll1 = True Then
       PicBoxCorner.Visible = True
       PicBoxCorner.ZOrder
    End If
    
    PicBack.Visible = True
    PagePicture.Visible = True
    
    WritePages
    Debug.Print "   DisplayPages"
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CurrentX() As Variant
       CurrentX = ObjPrint.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Variant)
        
    New_CurrentX = ScaleX(New_CurrentX, ScaleMode, ScaleMode)
    m_CurrentX = New_CurrentX
    ObjPrint.CurrentX = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CurrentY() As Variant
        CurrentY = ObjPrint.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Variant)
    New_CurrentY = ScaleY(New_CurrentY, ScaleMode, ScaleMode)
    m_CurrentY = New_CurrentY
    ObjPrint.CurrentY = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get DrawWidth() As Integer
    DrawWidth = m_DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
   
    m_DrawWidth = New_DrawWidth
    If m_DrawWidth < 1 Then m_DrawWidth = 1
  
    If m_SendToPrinter Then
        Printer.DrawWidth = (m_DrawWidth * 15) / 2.5
    Else
        ObjPrint.DrawWidth = m_DrawWidth
    End If
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontBold() As Boolean
       FontBold = ObjPrint.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
    ObjPrint.FontBold = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontItalic() As Boolean
       FontItalic = ObjPrint.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_FontItalic = New_FontItalic
    ObjPrint.FontItalic = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,"Arial"
Public Property Get FontName() As String
       FontName = ObjPrint.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    ObjPrint.Font.Name = New_FontName
    ObjPrint.Print "";
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,8
Public Property Get FontSize() As Integer
       FontSize = ObjPrint.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Integer)
    m_FontSize = New_FontSize
    ObjPrint.FontSize = m_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FontCharSet() As Integer
       FontCharSet = m_FontCharSet
End Property

Public Property Let FontCharSet(ByVal New_FontCharSet As Integer)
    On Error Resume Next
    If ObjPrint Is Nothing Then Exit Property
    m_FontCharSet = New_FontCharSet
    ObjPrint.Font.Charset = New_FontCharSet
    On Error GoTo 0
    PropertyChanged "FontCharSet"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontStrikethru() As Boolean
        FontStrikethru = ObjPrint.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    m_FontStrikethru = New_FontStrikethru
    ObjPrint.FontStrikethru = m_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Header() As String
    Header = m_Header
End Property

Public Property Let Header(ByVal New_Header As String)
    m_Header = New_Header
    PropertyChanged "Header"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Times New Romans
Public Property Get HdrFontName() As String
    HdrFontName = m_HdrFontName
End Property

Public Property Let HdrFontName(ByVal New_HdrFontName As String)
    m_HdrFontName = New_HdrFontName
    PropertyChanged "HdrFontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get HdrFontBold() As Boolean
    HdrFontBold = m_HdrFontBold
End Property

Public Property Let HdrFontBold(ByVal New_HdrFontBold As Boolean)
    m_HdrFontBold = New_HdrFontBold
    PropertyChanged "HdrFontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HdrFontItalic() As Boolean
    HdrFontItalic = m_HdrFontItalic
End Property

Public Property Let HdrFontItalic(ByVal New_HdrFontItalic As Boolean)
    m_HdrFontItalic = New_HdrFontItalic
    PropertyChanged "HdrFontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,10
Public Property Get HdrFontSize() As Integer
    HdrFontSize = m_HdrFontSize
End Property

Public Property Let HdrFontSize(ByVal New_HdrFontSize As Integer)
    m_HdrFontSize = New_HdrFontSize
    PropertyChanged "HdrFontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HdrColor() As OLE_COLOR
    HdrColor = m_HdrColor
End Property

Public Property Let HdrColor(ByVal New_HdrColor As OLE_COLOR)
    m_HdrColor = New_HdrColor
    PropertyChanged "HdrColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HdrFontTransparent() As Boolean
    HdrFontTransparent = m_HdrFontTransparent
End Property

Public Property Let HdrFontTransparent(ByVal New_HdrFontTransparent As Boolean)
    m_HdrFontTransparent = New_HdrFontTransparent
    PropertyChanged "HdrFontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Footer() As String
    Footer = m_Footer
End Property

Public Property Let Footer(ByVal New_Footer As String)
    m_Footer = New_Footer
    PropertyChanged "Footer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get PrinterColorMode() As PrinterColorModeTypes
    PrinterColorMode = m_PrinterColorMode
End Property

Public Property Let PrinterColorMode(ByVal New_PrinterColorMode As PrinterColorModeTypes)
    m_PrinterColorMode = New_PrinterColorMode
    PropertyChanged "PrinterColorMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get PageBorder() As PageBorderConstants
    PageBorder = m_PageBorder
End Property

Public Property Let PageBorder(ByVal New_PageBorder As PageBorderConstants)
    m_PageBorder = New_PageBorder
    PropertyChanged "PageBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontUnderline() As Boolean
        FontUnderline = ObjPrint.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    m_FontUnderline = New_FontUnderline
    ObjPrint.FontUnderline = m_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontTransparent() As Boolean
       FontTransparent = ObjPrint.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    m_FontTransparent = New_FontTransparent
    ObjPrint.FontTransparent = m_FontTransparent
    PropertyChanged "FontTransparent"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
       ForeColor = ObjPrint.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    ObjPrint.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get NavBar() As NavBarSetting
    NavBar = m_NavBar
End Property

Public Property Let NavBar(ByVal New_NavBar As NavBarSetting)
    m_NavBar = New_NavBar
    UserControl_Resize
    PropertyChanged "NavBar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicViewPort,PicViewPort,-1,BackColor
Public Property Get BackColorViewPort() As OLE_COLOR
    BackColorViewPort = PicViewPort.BackColor
End Property

Public Property Let BackColorViewPort(ByVal New_BackColorViewPort As OLE_COLOR)
    PicViewPort.BackColor() = New_BackColorViewPort
    PropertyChanged "BackColorViewPort"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPrintPic,picPrintPic,-1,BackColor
Public Property Get BackColorPage() As OLE_COLOR
    BackColorPage = m_BackColorPage
End Property

Public Property Let BackColorPage(ByVal New_BackColorPage As OLE_COLOR)
    m_BackColorPage = New_BackColorPage
    PropertyChanged "BackColorPage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Orientation() As PageOrientationConstants
    Orientation = m_Orientation 'Printer.Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As PageOrientationConstants)
    If New_Orientation = 0 Then New_Orientation = PagePortrait
    m_Orientation = New_Orientation
   MakeNewPage
   PropertyChanged "Orientation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get SendToPrinter() As Boolean
    SendToPrinter = m_SendToPrinter
End Property

Public Property Let SendToPrinter(ByVal New_SendToPrinter As Boolean)
    m_SendToPrinter = New_SendToPrinter
    If m_SendToPrinter = False Then
       Set ObjPrint = PagePicture
    Else
       Set ObjPrint = Printer
    End If
    PropertyChanged "SendToPrinter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ScaleMode() As PrinterScaleMode
    ScaleMode = m_ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As PrinterScaleMode)
    m_ScaleMode = New_ScaleMode
    If Not ObjPrint Is Nothing Then
       ObjPrint.ScaleMode = m_ScaleMode
    End If
    Printer.ScaleMode = m_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Zoom() As ZoomModeConstants
    Zoom = m_Zoom
End Property

Public Property Let Zoom(ByVal New_Zoom As ZoomModeConstants)
    
    If New_Zoom < 0 Or New_Zoom > 7 Then Exit Property
    
    RaiseEvent BeforeUserZoom
    
    m_Zoom = New_Zoom
    PropertyChanged "Zoom"
    
    Dim uPic As PictureBox
    For Each uPic In ThumbNail
         If uPic.Index > 0 Then
            Unload uPic
         End If
     Next
    ThumbNail(0).Visible = False
    If New_Zoom = 7 Then
       CmdNav(0).Enabled = False
       CmdNav(1).Enabled = False
       CmdNav(2).Enabled = False
       CmdNav(3).Enabled = False
       LabelPages.Caption = ""
    Else
       CmdNav(0).Enabled = True
       CmdNav(1).Enabled = True
       CmdNav(2).Enabled = True
       CmdNav(3).Enabled = True
    End If
    NavBarLabel.Caption = ""
    DisplayPages
    
    If Zoom < zmThumbnail Then
       PagePicture.Visible = True
       PicBack.Visible = True
    End If
    RaiseEvent AfterUserZoom
    
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,False
'Public Property Get PhysicalPage() As Boolean
'    PhysicalPage = m_PhysicalPage
'End Property

'Public Property Let PhysicalPage(ByVal New_PhysicalPage As Boolean)
'    m_PhysicalPage = New_PhysicalPage
'    If m_PhysicalPage = True Then
'       PageWidth = ScaleX(Printer.Width, vbTwips, ScaleMode)
'       PageHeight = ScaleY(Printer.Height, vbTwips, ScaleMode)
'    Else
'       PageWidth = Printer.ScaleWidth
'       PageHeight = Printer.ScaleHeight
'    End If
'    PropertyChanged "PhysicalPage"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get PageWidth() As Single
    PageWidth = m_PageWidth
End Property

Public Property Let PageWidth(ByVal New_PageWidth As Single)
    m_PageWidth = New_PageWidth
    PropertyChanged "PageWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get PageHeight() As Single
    PageHeight = m_PageHeight
End Property

Public Property Let PageHeight(ByVal New_PageHeight As Single)
    m_PageHeight = New_PageHeight
    PropertyChanged "PageHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get MarginLeft() As Variant
    MarginLeft = m_MarginLeft
End Property

Public Property Let MarginLeft(ByVal New_MarginLeft As Variant)
   If InStr(1, New_MarginLeft, "%") Then
      New_MarginLeft = Replace(New_MarginLeft, "%", "")
      New_MarginLeft = ((New_MarginLeft) / 100) * (PageWidth / 1000)
   Else
      New_MarginLeft = ScaleX(New_MarginLeft, ScaleMode, ScaleMode)
   End If
    m_MarginLeft = New_MarginLeft
    PropertyChanged "MarginLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get MarginRight() As Variant
    MarginRight = m_MarginRight
End Property

Public Property Let MarginRight(ByVal New_MarginRight As Variant)
   If InStr(1, New_MarginRight, "%") Then
      New_MarginRight = Replace(New_MarginRight, "%", "")
      New_MarginRight = ((New_MarginRight) / 100) * (PageWidth / 1000)
   Else
      New_MarginRight = ScaleX(New_MarginRight, ScaleMode, ScaleMode)
    End If
    m_MarginRight = New_MarginRight
    PropertyChanged "MarginRight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get MarginTop() As Variant
    MarginTop = m_MarginTop
End Property

Public Property Let MarginTop(ByVal New_MarginTop As Variant)
   If InStr(1, New_MarginTop, "%") Then
      New_MarginTop = Replace(New_MarginTop, "%", "")
      New_MarginTop = ((New_MarginTop) / 100) * (PageHeight / 1000)
   Else
      New_MarginTop = ScaleY(New_MarginTop, ScaleMode, ScaleMode)
   End If
    m_MarginTop = New_MarginTop
    If m_MarginTop < m_MarginHeader Then MarginHeader = m_MarginTop
    PropertyChanged "MarginTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get MarginBottom() As Variant
    MarginBottom = m_MarginBottom
End Property

Public Property Let MarginBottom(ByVal New_MarginBottom As Variant)
   
   If InStr(1, New_MarginBottom, "%") Then
      New_MarginBottom = Replace(New_MarginBottom, "%", "")
      New_MarginBottom = ((New_MarginBottom) / 100) * (PageHeight / 1000)
   Else
      New_MarginBottom = ScaleY(New_MarginBottom, ScaleMode, ScaleMode)
   End If
    m_MarginBottom = New_MarginBottom
    If m_MarginBottom < m_MarginFooter Then MarginBottom = m_MarginFooter
    PropertyChanged "MarginBottom"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MarginFooter() As Variant
    If m_MarginFooter = 0 Then
       Select Case ScaleMode
       Case smCentimeters
           m_MarginFooter = 1
'       Case smMillimeters
'           m_MarginFooter = 10
       Case Else
           m_MarginFooter = 0.5 '200 * 567
       End Select
    End If
    MarginFooter = m_MarginFooter
End Property

Public Property Let MarginFooter(ByVal New_MarginFooter As Variant)
      
   If InStr(1, New_MarginFooter, "%") Then
      New_MarginFooter = Replace(New_MarginFooter, "%", "")
      New_MarginFooter = ((New_MarginFooter) / 100) * (PageHeight / 1000)
   Else
      New_MarginFooter = ScaleY(New_MarginFooter, ScaleMode, ScaleMode)
   End If
    If m_MarginFooter = 0 Then
       Select Case ScaleMode
       Case smCentimeters
           m_MarginFooter = 1 ' * 567
       'Case smMillimeters
       '    m_MarginFooter = 10
       Case Else
           m_MarginFooter = 0.5 '200 * 567
       End Select
    End If
    m_MarginFooter = New_MarginFooter
    PropertyChanged "MarginFooter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MarginHeader() As Variant
    If m_MarginHeader = 0 Then
       Select Case ScaleMode
       Case smCentimeters
           m_MarginHeader = 1
'       Case smMillimeters
'           m_MarginHeader = 10
       Case Else
           m_MarginHeader = 0.5
       End Select
    End If
    MarginHeader = m_MarginHeader
End Property

Public Property Let MarginHeader(ByVal New_MarginHeader As Variant)

   If InStr(1, New_MarginHeader, "%") Then
      New_MarginHeader = Replace(New_MarginHeader, "%", "")
      New_MarginHeader = ((New_MarginHeader) / 100) * (PageHeight / 1000)
   Else
      New_MarginHeader = ScaleY(New_MarginHeader, ScaleMode, ScaleMode)
   End If
    If m_MarginHeader = 0 Then
       Select Case ScaleMode
       Case smCentimeters
           m_MarginHeader = 1
'       Case smMillimeters
'           m_MarginHeader = 10
       Case Else
           m_MarginHeader = 0.5
       End Select
    End If
    m_MarginHeader = New_MarginHeader
    PropertyChanged "MarginHeader"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get LineSpace() As LineSpaceConstants
    LineSpace = m_LineSpace
End Property

Public Property Let LineSpace(ByVal New_LineSpace As LineSpaceConstants)
    m_LineSpace = New_LineSpace
    PropertyChanged "LineSpace"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get TextAlign() As TextAlignConstants
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As TextAlignConstants)
    m_TextAlign = New_TextAlign
    PropertyChanged "TextAlign"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DrawStyle() As DrawStyleConstants
    DrawStyle = m_DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As DrawStyleConstants)
    m_DrawStyle = New_DrawStyle
    ObjPrint.DrawStyle = m_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get TableBorder() As TableBorderConstants
    TableBorder = m_TableBorder
End Property

Public Property Let TableBorder(ByVal New_TableBorder As TableBorderConstants)
    m_TableBorder = New_TableBorder
    If m_TableBorder > 10 Then m_TableBorder = 0
    PropertyChanged "TableBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FromPage() As Integer
    FromPage = m_FromPage
End Property

Public Property Let FromPage(ByVal New_FromPage As Integer)
    m_FromPage = New_FromPage
    PropertyChanged "FromPage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ToPage() As Integer
    ToPage = m_ToPage
End Property

Public Property Let ToPage(ByVal New_ToPage As Integer)
    m_ToPage = New_ToPage
    PropertyChanged "ToPage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get HdrFontStrikethru() As Boolean
    HdrFontStrikethru = m_HdrFontStrikethru
End Property

Public Property Let HdrFontStrikethru(ByVal New_HdrFontStrikethru As Boolean)
    m_HdrFontStrikethru = New_HdrFontStrikethru
    PropertyChanged "HdrFontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get HdrFontUnderline() As Boolean
    HdrFontUnderline = m_HdrFontUnderline
End Property

Public Property Let HdrFontUnderline(ByVal New_HdrFontUnderline As Boolean)
    m_HdrFontUnderline = New_HdrFontUnderline
    PropertyChanged "HdrFontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FillStyle() As FillStyleConstants
    FillStyle = m_FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleConstants)
    m_FillStyle = New_FillStyle
    ObjPrint.FillStyle = m_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    m_FillColor = New_FillColor
    ObjPrint.FillColor = m_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get IndentFirst() As Variant
    IndentFirst = m_IndentFirst
End Property

Public Property Let IndentFirst(ByVal New_IndentFirst As Variant)
    m_IndentFirst = ScaleX(New_IndentFirst, ScaleMode, ScaleMode)
    PropertyChanged "IndentFirst"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get IndentLeft() As Variant
    IndentLeft = m_IndentLeft
End Property

Public Property Let IndentLeft(ByVal New_IndentLeft As Variant)
    m_IndentLeft = ScaleX(New_IndentLeft, ScaleMode, ScaleMode)
    PropertyChanged "IndentLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get IndentRight() As Variant
    IndentRight = m_IndentRight
End Property

Public Property Let IndentRight(ByVal New_IndentRight As Variant)
    m_IndentRight = ScaleX(New_IndentRight, ScaleMode, ScaleMode)
    PropertyChanged "IndentRight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get x1() As Single
    x1 = m_X1
End Property

Public Property Let x1(ByVal New_X1 As Single)
    m_X1 = New_X1
    PropertyChanged "X1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get x2() As Single
    x2 = m_X2
End Property

Public Property Let x2(ByVal New_X2 As Single)
    m_X2 = New_X2
    PropertyChanged "X2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get y1() As Single
    y1 = m_Y1
End Property

Public Property Let y1(ByVal New_Y1 As Single)
    m_Y1 = New_Y1
    PropertyChanged "Y1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get y2() As Single
    y2 = m_Y2
End Property

Public Property Let y2(ByVal New_Y2 As Single)
    m_Y2 = New_Y2
    PropertyChanged "Y2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get PageBorderColor() As OLE_COLOR
    PageBorderColor = m_PageBorderColor
End Property

Public Property Let PageBorderColor(ByVal New_PageBorderColor As OLE_COLOR)
    m_PageBorderColor = New_PageBorderColor
    PropertyChanged "PageBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get PageBorderWidth() As Integer
    PageBorderWidth = m_PageBorderWidth
End Property

Public Property Let PageBorderWidth(ByVal New_PageBorderWidth As Integer)
    m_PageBorderWidth = New_PageBorderWidth
    PropertyChanged "PageBorderWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PagePicture,PagePicture,-1,hDC
Public Property Get hdc() As Long
    hdc = PagePicture.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Document
Public Property Get DocName() As String
    DocName = m_DocName
End Property

Public Property Let DocName(ByVal New_DocName As String)
    m_DocName = New_DocName
    PropertyChanged "DocName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,9
Public Property Get PaperSize() As PaperSizeConstans
    PaperSize = m_PaperSize
End Property

Public Property Let PaperSize(ByVal New_PaperSize As PaperSizeConstans)
    m_PaperSize = New_PaperSize
    PropertyChanged "PaperSize"
    SelectPaperSize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=31,0,0,0
Public Property Get DrawMode() As DrawModeConstants
    DrawMode = m_DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As DrawModeConstants)
    m_DrawMode = New_DrawMode
    ObjPrint.DrawMode = m_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    FontName = m_Font.Name
    FontBold = m_Font.Bold
    FontCharSet = m_Font.Charset
    FontItalic = m_Font.Italic
    FontSize = m_Font.Size
    FontStrikethru = m_Font.Strikethrough
    FontUnderline = m_Font.Underline
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get HdrFont() As Font
    Set HdrFont = m_HdrFont
End Property

Public Property Set HdrFont(ByVal New_HdrFont As Font)
    Set m_HdrFont = New_HdrFont
    HdrFontName = m_HdrFont.Name
    HdrFontBold = m_HdrFont.Bold
    'HdrFontCharSet = m_HdrFont.Charset
    HdrFontItalic = m_HdrFont.Italic
    HdrFontSize = m_HdrFont.Size
    HdrFontStrikethru = m_HdrFont.Strikethrough
    HdrFontUnderline = m_HdrFont.Underline
    PropertyChanged "HdrFont"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get NavBarMenu() As String
    NavBarMenu = m_NavBarMenu
End Property

Public Property Let NavBarMenu(ByVal New_NavBarMenu As String)
    Dim nArr() As String
    If InStr(1, New_NavBarMenu, "|") = 0 Then New_NavBarMenu = New_NavBarMenu + "||"
    m_NavBarMenu = New_NavBarMenu
    nArr = Split(m_NavBarMenu, "|")
    mnzoom(6).Caption = nArr(0)
    mnzoom(7).Caption = nArr(1)
    mnzoom(8).Caption = nArr(2)
    PropertyChanged "NavBarMenu"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get NavBarLabels() As String
    NavBarLabels = m_NavBarLabels
End Property

Public Property Let NavBarLabels(ByVal New_NavBarLabels As String)
    m_NavBarLabels = New_NavBarLabels
    NavBarLabel.Caption = m_NavBarLabels
    NavBarLabel.Refresh
    PropertyChanged "NavBarLabels"
End Property

'**************** End Property ****************
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CurrentX = m_def_CurrentX
    m_CurrentY = m_def_CurrentY
    m_DrawWidth = m_def_DrawWidth
    m_FontBold = m_def_FontBold
    m_FontItalic = m_def_FontItalic
    m_FontName = m_def_FontName
    m_FontSize = m_def_FontSize
    m_FontCharSet = m_def_FontCharSet
    m_FontStrikethru = m_def_FontStrikethru
    m_FontTransparent = m_def_FontTransparent
    m_FontUnderline = m_def_FontUnderline
    m_ForeColor = m_def_ForeColor
    m_Orientation = m_def_Orientation
    m_SendToPrinter = m_def_SendToPrinter
    m_ScaleMode = m_def_ScaleMode
    m_Zoom = m_def_Zoom
    m_MarginLeft = m_def_MarginLeft
    m_MarginRight = m_def_MarginRight
    m_MarginTop = m_def_MarginTop
    m_MarginBottom = m_def_MarginBottom
    m_LineSpace = m_def_LineSpace
    m_TextAlign = m_def_TextAlign
    m_DrawStyle = m_def_DrawStyle
    m_Header = m_def_Header
    m_HdrColor = m_def_HdrColor
    m_MarginHeader = m_def_MarginHeader
    m_Footer = m_def_Footer
    m_MarginFooter = m_def_MarginFooter
    m_PrinterColorMode = m_def_PrinterColorMode
    m_PageBorder = m_def_PageBorder
    m_HdrFontName = m_def_HdrFontName
    m_HdrFontBold = m_def_HdrFontBold
    m_HdrFontItalic = m_def_HdrFontItalic
    m_HdrFontSize = m_def_HdrFontSize
    m_TableBorder = m_def_TableBorder
    m_FromPage = m_def_FromPage
    m_ToPage = m_def_ToPage
    m_NavBar = m_def_NavBar
    m_BackColor = m_def_BackColor
    m_BackColorPage = m_def_BackColorPage
    PicViewPort.BackColor = &H808080
    m_HdrFontStrikethru = m_def_HdrFontStrikethru
    m_HdrFontUnderline = m_def_HdrFontUnderline
    m_FillStyle = m_def_FillStyle
    m_FillColor = m_def_FillColor
    m_IndentFirst = m_def_IndentFirst
    m_IndentLeft = m_def_IndentLeft
    m_IndentRight = m_def_IndentRight
    m_X1 = m_def_X1
    m_X2 = m_def_X2
    m_Y1 = m_def_Y1
    m_Y2 = m_def_Y2
    'PhysicalPage = m_def_PhysicalPage
    m_PageWidth = m_def_PageWidth
    m_PageHeight = m_def_PageHeight
    m_PageBorderColor = m_def_PageBorderColor
    m_PageBorderWidth = m_def_PageBorderWidth
    m_HdrDrawWidth = m_def_HdrDrawWidth
    m_HdrDrawStyle = m_def_HdrDrawStyle
    m_HdrFontTransparent = m_def_HdrFontTransparent
    m_DocName = m_def_DocName
    m_PaperSize = m_def_PaperSize
    m_DrawMode = m_def_DrawMode
    m_About = m_def_About
    Set m_Font = Ambient.Font
    Set m_HdrFont = Ambient.Font
    m_NavBarMenu = m_def_NavBarMenu
    m_NavBarLabels = m_def_NavBarLabels
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     
    m_CurrentX = PropBag.ReadProperty("CurrentX", m_def_CurrentX)
    m_CurrentY = PropBag.ReadProperty("CurrentY", m_def_CurrentY)
    m_DrawWidth = PropBag.ReadProperty("DrawWidth", m_def_DrawWidth)
    m_FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    m_FontItalic = PropBag.ReadProperty("FontItalic", m_def_FontItalic)
    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    m_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
    m_FontCharSet = PropBag.ReadProperty("FontCharSet", m_def_FontCharSet)
    m_FontStrikethru = PropBag.ReadProperty("FontStrikethru", m_def_FontStrikethru)
    m_FontTransparent = PropBag.ReadProperty("FontTransparent", m_def_FontTransparent)
    m_FontUnderline = PropBag.ReadProperty("FontUnderline", m_def_FontUnderline)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_SendToPrinter = PropBag.ReadProperty("SendToPrinter", m_def_SendToPrinter)
    m_ScaleMode = PropBag.ReadProperty("ScaleMode", m_def_ScaleMode)
    m_Zoom = PropBag.ReadProperty("Zoom", m_def_Zoom)
    m_MarginLeft = PropBag.ReadProperty("MarginLeft", m_def_MarginLeft)
    m_MarginRight = PropBag.ReadProperty("MarginRight", m_def_MarginRight)
    m_MarginTop = PropBag.ReadProperty("MarginTop", m_def_MarginTop)
    m_MarginBottom = PropBag.ReadProperty("MarginBottom", m_def_MarginBottom)
    m_LineSpace = PropBag.ReadProperty("LineSpace", m_def_LineSpace)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_DrawStyle = PropBag.ReadProperty("DrawStyle", m_def_DrawStyle)
    m_Header = PropBag.ReadProperty("Header", m_def_Header)
    m_HdrColor = PropBag.ReadProperty("HdrColor", m_def_HdrColor)
    m_MarginHeader = PropBag.ReadProperty("MarginHeader", m_def_MarginHeader)
    m_Footer = PropBag.ReadProperty("Footer", m_def_Footer)
    m_MarginFooter = PropBag.ReadProperty("MarginFooter", m_def_MarginFooter)
    m_PrinterColorMode = PropBag.ReadProperty("PrinterColorMode", m_def_PrinterColorMode)
    m_PageBorder = PropBag.ReadProperty("PageBorder", m_def_PageBorder)
    m_HdrFontName = PropBag.ReadProperty("HdrFontName", m_def_HdrFontName)
    m_HdrFontBold = PropBag.ReadProperty("HdrFontBold", m_def_HdrFontBold)
    m_HdrFontItalic = PropBag.ReadProperty("HdrFontItalic", m_def_HdrFontItalic)
    m_HdrFontSize = PropBag.ReadProperty("HdrFontSize", m_def_HdrFontSize)
    m_TableBorder = PropBag.ReadProperty("TableBorder", m_def_TableBorder)
    m_FromPage = PropBag.ReadProperty("FromPage", m_def_FromPage)
    m_ToPage = PropBag.ReadProperty("ToPage", m_def_ToPage)
    m_NavBar = PropBag.ReadProperty("NavBar", m_def_NavBar)
    PicViewPort.BackColor = PropBag.ReadProperty("BackColorViewPort", &H808080)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BackColorPage = PropBag.ReadProperty("BackColorPage", m_def_BackColorPage)
    m_HdrFontStrikethru = PropBag.ReadProperty("HdrFontStrikethru", m_def_HdrFontStrikethru)
    m_HdrFontUnderline = PropBag.ReadProperty("HdrFontUnderline", m_def_HdrFontUnderline)
    m_FillStyle = PropBag.ReadProperty("FillStyle", m_def_FillStyle)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    m_IndentFirst = PropBag.ReadProperty("IndentFirst", m_def_IndentFirst)
    m_IndentLeft = PropBag.ReadProperty("IndentLeft", m_def_IndentLeft)
    m_IndentRight = PropBag.ReadProperty("IndentRight", m_def_IndentRight)
    m_X1 = PropBag.ReadProperty("X1", m_def_X1)
    m_X2 = PropBag.ReadProperty("X2", m_def_X2)
    m_Y1 = PropBag.ReadProperty("Y1", m_def_Y1)
    m_Y2 = PropBag.ReadProperty("Y2", m_def_Y2)
    m_PageBorderColor = PropBag.ReadProperty("PageBorderColor", m_def_PageBorderColor)
    m_PageBorderWidth = PropBag.ReadProperty("PageBorderWidth", m_def_PageBorderWidth)
    
    'm_PageWidth = PropBag.ReadProperty("PageWidth", m_def_PageWidth)
    'm_PageHeight = PropBag.ReadProperty("PageHeight", m_def_PageHeight)
    m_HdrDrawWidth = PropBag.ReadProperty("HdrDrawWidth", m_def_HdrDrawWidth)
    m_HdrDrawStyle = PropBag.ReadProperty("HdrDrawStyle", m_def_HdrDrawStyle)
    m_HdrFontTransparent = PropBag.ReadProperty("HdrFontTransparent", m_def_HdrFontTransparent)
    m_DocName = PropBag.ReadProperty("DocName", m_def_DocName)
    m_PaperSize = PropBag.ReadProperty("PaperSize", m_def_PaperSize)
    m_DrawMode = PropBag.ReadProperty("DrawMode", m_def_DrawMode)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_About = PropBag.ReadProperty("About", m_def_About)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_HdrFont = PropBag.ReadProperty("HdrFont", Ambient.Font)
    NavBarMenu = PropBag.ReadProperty("NavBarMenu", m_def_NavBarMenu)
    m_NavBarLabels = PropBag.ReadProperty("NavBarLabels", m_def_NavBarLabels)
    
    m_Font.Bold = m_FontBold
    m_Font.Charset = m_FontCharSet
    m_Font.Italic = m_FontItalic
    m_Font.Name = m_FontName
    m_Font.Strikethrough = m_FontStrikethru
    m_Font.Underline = m_FontUnderline
    m_Font.Size = m_FontSize
    
    m_HdrFont.Bold = m_HdrFontBold
    m_HdrFont.Charset = m_FontCharSet
    m_HdrFont.Italic = m_HdrFontItalic
    m_HdrFont.Name = m_HdrFontName
    m_HdrFont.Strikethrough = m_HdrFontStrikethru
    m_HdrFont.Underline = m_HdrFontUnderline
    m_HdrFont.Size = m_HdrFontSize
   
    UserControl_Initialize
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CurrentX", m_CurrentX, m_def_CurrentX)
    Call PropBag.WriteProperty("CurrentY", m_CurrentY, m_def_CurrentY)
    Call PropBag.WriteProperty("DrawWidth", m_DrawWidth, m_def_DrawWidth)
    Call PropBag.WriteProperty("FontBold", m_FontBold, m_def_FontBold)
    Call PropBag.WriteProperty("FontItalic", m_FontItalic, m_def_FontItalic)
    Call PropBag.WriteProperty("FontName", m_FontName, m_def_FontName)
    Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
    Call PropBag.WriteProperty("FontCharSet", m_FontCharSet, m_def_FontCharSet)
    Call PropBag.WriteProperty("FontStrikethru", m_FontStrikethru, m_def_FontStrikethru)
    Call PropBag.WriteProperty("FontUnderline", m_FontUnderline, m_def_FontUnderline)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("SendToPrinter", m_SendToPrinter, m_def_SendToPrinter)
    Call PropBag.WriteProperty("ScaleMode", m_ScaleMode, m_def_ScaleMode)
    Call PropBag.WriteProperty("Zoom", m_Zoom, m_def_Zoom)
    Call PropBag.WriteProperty("MarginLeft", m_MarginLeft, m_def_MarginLeft)
    Call PropBag.WriteProperty("MarginRight", m_MarginRight, m_def_MarginRight)
    Call PropBag.WriteProperty("MarginTop", m_MarginTop, m_def_MarginTop)
    Call PropBag.WriteProperty("MarginBottom", m_MarginBottom, m_def_MarginBottom)
    Call PropBag.WriteProperty("LineSpace", m_LineSpace, m_def_LineSpace)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("DrawStyle", m_DrawStyle, m_def_DrawStyle)
    Call PropBag.WriteProperty("Header", m_Header, m_def_Header)
    Call PropBag.WriteProperty("HdrColor", m_HdrColor, m_def_HdrColor)
    Call PropBag.WriteProperty("MarginHeader", m_MarginHeader, m_def_MarginHeader)
    Call PropBag.WriteProperty("Footer", m_Footer, m_def_Footer)
    Call PropBag.WriteProperty("MarginFooter", m_MarginFooter, m_def_MarginFooter)
    Call PropBag.WriteProperty("PrinterColorMode", m_PrinterColorMode, m_def_PrinterColorMode)
    Call PropBag.WriteProperty("PageBorder", m_PageBorder, m_def_PageBorder)
    Call PropBag.WriteProperty("HdrFontName", m_HdrFontName, m_def_HdrFontName)
    Call PropBag.WriteProperty("HdrFontBold", m_HdrFontBold, m_def_HdrFontBold)
    Call PropBag.WriteProperty("HdrFontItalic", m_HdrFontItalic, m_def_HdrFontItalic)
    Call PropBag.WriteProperty("HdrFontSize", m_HdrFontSize, m_def_HdrFontSize)
    Call PropBag.WriteProperty("FontTransparent", m_FontTransparent, m_def_FontTransparent)
    Call PropBag.WriteProperty("TableBorder", m_TableBorder, m_def_TableBorder)
    Call PropBag.WriteProperty("FromPage", m_FromPage, m_def_FromPage)
    Call PropBag.WriteProperty("ToPage", m_ToPage, m_def_ToPage)
    Call PropBag.WriteProperty("NavBar", m_NavBar, m_def_NavBar)
    Call PropBag.WriteProperty("BackColorViewPort", PicViewPort.BackColor, &H808080)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BackColorPage", m_BackColorPage, m_def_BackColorPage)
    Call PropBag.WriteProperty("HdrFontStrikethru", m_HdrFontStrikethru, m_def_HdrFontStrikethru)
    Call PropBag.WriteProperty("HdrFontUnderline", m_HdrFontUnderline, m_def_HdrFontUnderline)
    Call PropBag.WriteProperty("FillStyle", m_FillStyle, m_def_FillStyle)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("IndentFirst", m_IndentFirst, m_def_IndentFirst)
    Call PropBag.WriteProperty("IndentLeft", m_IndentLeft, m_def_IndentLeft)
    Call PropBag.WriteProperty("IndentRight", m_IndentRight, m_def_IndentRight)
    Call PropBag.WriteProperty("X1", m_X1, m_def_X1)
    Call PropBag.WriteProperty("X2", m_X2, m_def_X2)
    Call PropBag.WriteProperty("Y1", m_Y1, m_def_Y1)
    Call PropBag.WriteProperty("Y2", m_Y2, m_def_Y2)
    'Call PropBag.WriteProperty("PhysicalPage", m_PhysicalPage, m_def_PhysicalPage)
    Call PropBag.WriteProperty("PageWidth", m_PageWidth, m_def_PageWidth)
    Call PropBag.WriteProperty("PageHeight", m_PageHeight, m_def_PageHeight)
    Call PropBag.WriteProperty("PageBorderColor", m_PageBorderColor, m_def_PageBorderColor)
    Call PropBag.WriteProperty("PageBorderWidth", m_PageBorderWidth, m_def_PageBorderWidth)
    Call PropBag.WriteProperty("HdrDrawWidth", m_HdrDrawWidth, m_def_HdrDrawWidth)
    Call PropBag.WriteProperty("HdrDrawStyle", m_HdrDrawStyle, m_def_HdrDrawStyle)
    Call PropBag.WriteProperty("HdrFontTransparent", m_HdrFontTransparent, m_def_HdrFontTransparent)
    Call PropBag.WriteProperty("DocName", m_DocName, m_def_DocName)
    Call PropBag.WriteProperty("PaperSize", m_PaperSize, m_def_PaperSize)
    
    Call PropBag.WriteProperty("DrawMode", m_DrawMode, m_def_DrawMode)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("About", m_About, m_def_About)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("HdrFont", m_HdrFont, Ambient.Font)
    Call PropBag.WriteProperty("NavBarMenu", m_NavBarMenu, m_def_NavBarMenu)
    Call PropBag.WriteProperty("NavBarLabels", m_NavBarLabels, m_def_NavBarLabels)
End Sub

Private Sub UserControl_Terminate()
   Clear
   'Printer.ScaleMode = pSM
End Sub

'Set Current Page in Clipboard
Public Function SetDataClipboard(ByVal PageNumber As Integer) As Boolean
        
        If PageNumber - 1 > MaxPageNumber Then Exit Function
        
        If FileExists(TempDir & "PPView" & CStr(PageNumber - 1) & ".bmp") = True Then
           picFullPage.Picture = LoadPicture(TempDir & "PPView" & CStr(PageNumber - 1) & ".bmp")
           Clipboard.SetData picFullPage
           SetDataClipboard = True
           picFullPage.Picture = LoadPicture() 'clear picture
        Else
           SetDataClipboard = False
        End If
End Function

Private Function CharSetWin() As Integer

   Dim hInst As Long, lResult As Long, X As Long
   Dim LCID As Long, sLcid As String
   Dim resString As String * 255
   Dim sCodePage As String
   Dim locale_id As Long

   sCodePage = String$(1024, " ")
  
   LCID = GetThreadLocale() 'Get Current locale
   
   'decimal value of the LCID (Hex in Parentheses)
   sLcid = Hex$(Trim$(CStr(LCID))) 'Convert to Hex
   sCodePage = LocaleInfo(locale_id, LOCALE_IDEFAULTANSICODEPAGE)
   CharSetWin = GetCharSet(sCodePage) 'Convert code page to charset

End Function

Private Function LocaleInfo(ByVal Locale As Long, ByVal lc_type As Long) As String
    Dim Length As Long
    Dim Buf As String * 1024
    Length = GetLocaleInfo(Locale, lc_type, Buf, Len(Buf))
    LocaleInfo = Left$(Buf, Length - 1)
End Function

Private Function GetCharSet(sCdpg As String) As Integer
   
   Select Case sCdpg
      Case "932" ' Japanese
         GetCharSet = 128
      Case "936" ' Simplified Chinese
         GetCharSet = 134
      Case "949" ' Korean
         GetCharSet = 129
      Case "950" ' Traditional Chinese
         GetCharSet = 136
      Case "1250" ' Eastern Europe
         GetCharSet = 238
      Case "1251" ' Russian
         GetCharSet = 204
      Case "1252" ' Western European Languages
         GetCharSet = 0
      Case "1253" ' Greek
         GetCharSet = 161
      Case "1254" ' Turkish
         GetCharSet = 162
      Case "1255" ' Hebrew
         GetCharSet = 177
      Case "1256" ' Arabic
         GetCharSet = 178
      Case "1257" ' Baltic
         GetCharSet = 186
      Case Else
         GetCharSet = 0
   End Select
End Function

Public Sub StartDoc()
    Dim mLeft As Single, mTop As Single, mRight As Single, mBottom As Single
    
    MaxPageNumber = 0
    
    On Local Error Resume Next
    
    'Set the Printer's scale mode
    pSM = Printer.ScaleMode
    Printer.ScaleMode = m_ScaleMode
    
    FontCharSet = CharSetWin
    
    'Get the physical printable area
    SelectPaperSize
    'PageWidth = Printer.ScaleWidth
    'PageHeight = Printer.ScaleHeight
    
    'Printer.Orientation = Orientation
    
    If m_SendToPrinter Then
        
        ChangePaperSize
        Printer.Orientation = m_Orientation
        Text " "
        CurrentX = 0
        CurrentY = 0
        If FontTransparent Then
           iBKMode = SetBkMode(Printer.hdc, TRANSPARENT)
        Else
           iBKMode = SetBkMode(Printer.hdc, OPAQUE)
        End If
    Else
        Set ObjPrint = PagePicture
        PagePicture.Visible = False
        PicBack.Visible = False

        'Scale Object to Printer's printable area
        oSM = ObjPrint.ScaleMode
        ObjPrint.ScaleMode = m_ScaleMode
        
        'Full Page size (1440 twips = 1 inch or 567 twips = 1 centimeter)
        Select Case m_ScaleMode
        Case smCentimeters
            ObjPrint.Width = PageWidth * 567
            ObjPrint.Height = PageHeight * 567
'            If MarginLeft = 0 Then MarginLeft = 2
'            If MarginRight = 0 Then MarginRight = 2
'            If MarginTop = 0 Then MarginTop = 2
'            If MarginBottom = 0 Then MarginBottom = 2
        Case Else 'inches
            ObjPrint.Width = (PageWidth + 0.2) * 1440
            ObjPrint.Height = (PageHeight + 0.2) * 1440
'            If MarginLeft = 0 Then MarginLeft = 0.78
'            If MarginRight = 0 Then MarginRight = 0.78
'            If MarginTop = 0 Then MarginTop = 0.78
'            If MarginBottom = 0 Then MarginBottom = 0.78
        End Select
        
        'Set default properties of the scroll bars
        VScroll1.Max = Abs(PagePicture.Height - PicViewPort.Height + 500)
        VScroll1.Min = -200
        VScroll1.SmallChange = VScroll1.Max * 0.1
        VScroll1.LargeChange = VScroll1.Max * 0.2
        
        HScroll1.Max = Abs(PagePicture.Width - PicViewPort.Width + 700)
        HScroll1.Min = -250
        HScroll1.SmallChange = HScroll1.Max * 0.1
        HScroll1.LargeChange = HScroll1.Max * 0.2
         
        'Set default properties of the object to match printer
        FontName = Printer.FontName
        FontSize = Printer.FontSize
        FontBold = Printer.FontBold
        ForeColor = Printer.ForeColor
        Printer.FontTransparent = True
        FontTransparent = Printer.FontTransparent
        
        ObjPrint.Picture = Nothing
        FontCharSet = CharSetWin
    End If
        
        'Page Back color
     DrawRectangle 0, 0, Printer.Width, Printer.Height, , , BackColorPage, BackColorPage, vbFSSolid
     DrawBorder
     PrintHeader
     'PrintFooter
     CurrentY = MarginTop
     
     NavBarLabels = "Page" + Str(MaxPageNumber + 1)
     
    RaiseEvent PageNew
End Sub

Public Sub EndDoc(Optional ByVal oModal As Byte = 1)
  Dim i As Integer
    RaiseEvent PageEnd
    
    If m_SendToPrinter Then
        
        If m_ToPage = MaxPageNumber + 1 Then
           PrintFooter
           DrawBorder
           Printer.EndDoc
        Else
           Printer.Print "";
           Printer.KillDoc
           Printer.EndDoc
        End If
        
        Printer.ScaleMode = pSM
        SendToPrinter = False
        RestorePrinterDefaults
        
        Debug.Print "Enddoc", MaxPageNumber + 1
    Else
       On Local Error Resume Next
       SelectPaperSize
       PrintFooter
       DrawBorder
       SavePicture ObjPrint.Image, TempDir & "PPView" & CStr(MaxPageNumber) & ".bmp"
       picFullPage.Picture = LoadPicture(TempDir & "PPView" & CStr(0) & ".bmp")
       PagePicture.Width = picFullPage.Width
       PagePicture.Height = picFullPage.Height
       'FromPage = 1
       ToPage = PageCount
       ViewPage = 0
       DisplayPages
    End If
    
      NavBarLabel.Caption = ""
    Screen.MousePointer = 0
    
    RaiseEvent PageEndDoc
End Sub

Public Function EndOfPage() As Boolean
  Dim n As Single, OldCy As Single
  Dim fTextHeight As Single
  Dim eFontsize As Integer
  
   OldCy = CurrentY
  
  CurrentY = CurrentY + ChangeLineSpace
  fTextHeight = ObjPrint.TextHeight("TextString")
  n = ObjPrint.CurrentY + ObjPrint.TextHeight("TextString") + fTextHeight
    
  If n >= PageHeight Or n >= PageHeight - MarginBottom Then
       EndOfPage = True
  Else
      EndOfPage = False
      CurrentY = OldCy
  End If
  CurrentY = OldCy
End Function

Private Sub PrintFooter()

   Dim eFontS As Integer
   Dim eFontN As String
   Dim eFontB As Boolean
   Dim eFontI As Boolean
   Dim eFontU As Boolean
   Dim eFontK As Boolean
   Dim eFontT As Boolean
   Dim eFontC As Long
   Dim tFooter As String
   Dim ArrFooter() As String, i As Integer, fTextAlign As TextAlignConstants
   Dim X As Single, Y As Single, w As Single, H As Single, wFooter As Single
   Dim oIR As Single
   Dim oIL As Single
   Dim oIF As Single
   Dim oTA As Integer
   
   If m_Footer = "" Then Exit Sub
    
   RaiseEvent BeforeFooter
  
    'Save current setting
    eFontN = FontName
    eFontS = FontSize
    eFontB = FontBold
    eFontI = FontItalic
    eFontU = FontUnderline
    eFontK = FontStrikethru
    eFontT = FontTransparent
    eFontC = ForeColor
    oIR = IndentRight
    oIL = IndentLeft
    oIF = IndentFirst
    oTA = TextAlign
    
    'Change settings
    FontName = HdrFontName
    FontSize = HdrFontSize
    FontBold = HdrFontBold
    FontItalic = HdrFontItalic
    ForeColor = HdrColor
    FontUnderline = HdrFontUnderline
    FontStrikethru = HdrFontStrikethru
    FontTransparent = HdrFontTransparent
   
    IndentRight = 0
    IndentLeft = 0
    IndentFirst = 0
    CurrentY = PageHeight - MarginFooter
    CurrentX = MarginLeft
    
    tFooter = m_Footer
    tFooter = Replace(tFooter, "p$", Trim(Str(MaxPageNumber + 1)))
    Debug.Print tFooter
    
    If InStr(1, tFooter, "|") > 0 Then
       ArrFooter = Split(tFooter, "|")
       wFooter = (m_PageWidth - m_MarginLeft - m_MarginRight) / (UBound(ArrFooter) + 1)
       For i = 0 To UBound(ArrFooter)
           If i = 0 Then
              fTextAlign = taLeftTop
              X = m_MarginLeft
              Y = m_PageHeight - m_MarginBottom
              w = wFooter
           ElseIf i = UBound(ArrFooter) Then
              fTextAlign = taRightTop
              X = X + w
              w = wFooter
           Else
              fTextAlign = taCenterTop
              X = X + w
              w = wFooter
           End If
           H = m_MarginBottom
           TextBox ArrFooter(i), X, Y, w, H, fTextAlign, False
       Next
    Else
        PrintMultiLine tFooter, MarginLeft, PageWidth - MarginRight, , CurrentY, False
    End If
    
    CurrentY = PageHeight - MarginFooter
    CurrentX = MarginLeft
    
    'Restore setting
    FontName = eFontN
    FontSize = eFontS
    FontBold = eFontB
    FontItalic = eFontI
    FontUnderline = eFontU
    FontStrikethru = eFontK
    FontTransparent = eFontT
    ForeColor = eFontC
    IndentRight = oIR
    IndentLeft = oIL
    IndentFirst = oIF
     TextAlign = oTA
   RaiseEvent AfterFooter
   
End Sub

Private Sub PrintHeader()

    Dim eFontS As Integer
    Dim eFontN As String
    Dim eFontB As Boolean
    Dim eFontI As Boolean
    Dim eFontU As Boolean
    Dim eFontK As Boolean
    Dim eFontC As Long
    Dim OldCy As Single
    Dim OldSP As Integer
    Dim ArrHeader() As String, ArrHeaderCr() As String, i As Integer, C As Integer, hTextAlign As TextAlignConstants
    Dim X As Single, Y As Single, w As Single, H As Single, wHeader As Single
    Dim oIR As Single
    Dim oIL As Single
    Dim oIF As Single
    
    'Remember Header Information
    If m_Header = "" Then Exit Sub
    RaiseEvent BeforeHeader
    
    'Save current setting
    eFontN = FontName
    eFontS = FontSize
    eFontB = FontBold
    eFontI = FontItalic
    eFontU = FontUnderline
    eFontK = FontStrikethru
    eFontC = ForeColor
    oIR = IndentRight
    oIL = IndentLeft
    oIF = IndentFirst
    OldCy = CurrentY
    OldSP = m_LineSpace
    
    'Change settings
    FontName = HdrFontName
    FontSize = HdrFontSize
    FontBold = HdrFontBold
    FontItalic = HdrFontItalic
    ForeColor = HdrColor
    FontUnderline = HdrFontUnderline
    FontStrikethru = HdrFontStrikethru
    LineSpace = lsSpaceSingle
    IndentRight = 0
    IndentLeft = 0
    IndentFirst = 0
    CurrentY = MarginHeader - TextHeight("W")
    CurrentX = MarginLeft
        
    m_Header = Replace(m_Header, "p$", Trim(Str(MaxPageNumber + 1)))
        
    If InStr(1, m_Header, "|") > 0 Then
       ArrHeader = Split(m_Header, "|")
       wHeader = (m_PageWidth - m_MarginLeft - m_MarginRight) / (UBound(ArrHeader) + 1)
       For i = 0 To UBound(ArrHeader)
           ArrHeaderCr = Split(ArrHeader(i), vbCr)
            For C = 0 To UBound(ArrHeaderCr)
                If i = 0 Then
                    hTextAlign = taLeftTop
                    X = m_MarginLeft
                    'Y = 0
                    w = wHeader
                ElseIf i = UBound(ArrHeader) Then
                    hTextAlign = taRightTop
                    X = X + w
                    w = wHeader
                Else
                    hTextAlign = taCenterTop
                    X = X + w
                    w = wHeader
                End If
                If (UBound(ArrHeaderCr) + 1) * TextHeight > H Then
                   H = (UBound(ArrHeaderCr) + 1) * TextHeight
                   If H > MarginFooter Then H = MarginFooter
                End If
                TextBox ArrHeader(i), X, MarginFooter - H, w, H, hTextAlign, False
           Next
       Next
    Else
      PrintMultiLine m_Header, MarginLeft, PageWidth - MarginRight, CurrentY, , False
    End If
    CurrentX = MarginLeft
    
    'Restore setting
    FontName = eFontN
    FontSize = eFontS
    FontBold = eFontB
    FontItalic = eFontI
    FontUnderline = eFontU
    FontStrikethru = eFontK
    ForeColor = eFontC
    CurrentY = OldCy
    LineSpace = OldSP
    IndentRight = oIR
    IndentLeft = oIL
    IndentFirst = oIF
    RaiseEvent AfterHeader
    
End Sub

Public Function GetPage() As Integer
    If m_SendToPrinter Then
       GetPage = Printer.Page
    Else
       GetPage = MaxPageNumber + 1
    End If
End Function

Private Sub PrintMultiLine(ByVal Value As Variant, _
                           Optional ByVal MrgLeft As Single = -1, _
                           Optional ByVal MrgRight As Single = -1, _
                           Optional ByVal MrgTop As Single = -1, _
                           Optional ByVal MrgBottom As Single = -1, _
                           Optional ByVal UsePageBreaks As Boolean = True, _
                           Optional ByVal SameText As Boolean = False)
 
  Dim StartChar As Integer
  Dim CharLength As Single
  Dim CurrentPos As Single
  Dim TxtLen As Single
  Dim tString As String
  Dim X As Integer, Y As Integer
  Dim ColWidth As Single
  Dim FirstIndent As Boolean
  
    If MrgLeft = -1 Then MarginLeft = CurrentX: MrgLeft = MarginLeft
    If MrgLeft > m_PageWidth - 0.1 Then MarginLeft = m_PageWidth - 0.5
    
    If MrgRight = -1 Then MrgRight = m_MarginRight
    MrgRight = MrgRight - m_IndentRight
    
    ColWidth = MrgRight - MrgLeft - (m_IndentLeft + m_IndentFirst)
             
    TxtLen = Len(Value)
    StartChar = 1
    CurrentPos = 0
    CharLength = TxtLen
    
    TmpUnit.Left = MrgLeft
    TmpUnit.Right = MrgRight
           
   If UsePageBreaks Then
      If EndOfPage Then
         PrintFooter
         NewPage
         PrintHeader
         If SameText = False Then ChangeTextAlign tString
      End If
   End If
    
    For X = 1 To TxtLen
        Y = X - CurrentPos
        If Mid(Value, X, 1) < Chr(33) Then CharLength = Y
        If TextWidth(Mid(Value, StartChar, Y)) >= ColWidth Then
           
            'If there are no spaces then break line here
            If CharLength > Y Then CharLength = Y - 1
            tString = Mid(Value, StartChar, CharLength)
                       
            If SameText = False Then
              CurrentY = CurrentY + ChangeLineSpace
              If FirstIndent = False Then
                 ChangeTextAlign tString, IndentFirst
              Else
                 ChangeTextAlign tString
              End If
            End If

            If UsePageBreaks Then
                If EndOfPage Then
                    PrintFooter
                    NewPage
                    PrintHeader
                    If SameText = False Then
                       ChangeTextAlign tString
                    End If
                End If
            End If
            
          
           'ObjPrint.Font.Charset = FontCharSet
           
           If SameText Then
                  ObjPrint.Print tString;
                  If CurrentX + TextWidth(tString) > ColWidth And _
                     CurrentX + TextWidth(tString) > PageWidth - MarginRight - MarginLeft Then
                     CurrentX = MarginLeft + IndentLeft
                     CurrentY = CurrentY + ChangeLineSpace
                     ColWidth = PageWidth - MarginRight - MarginLeft
                  End If
            Else
                ObjPrint.Print tString
            End If
            
            If CurrentX = 0 Then
              CurrentX = MarginLeft + IndentLeft
               ColWidth = MrgRight - MrgLeft - (m_IndentLeft + m_IndentRight)
            End If
            
            CurrentPos = CharLength + CurrentPos
            StartChar = CurrentPos + 1
            CharLength = TxtLen
            
            If FirstIndent = False Then
               ColWidth = ColWidth - IndentFirst
               FirstIndent = True
            End If
            
        End If
    Next X
   
    If SameText = False Then
       CurrentY = CurrentY + ChangeLineSpace
    End If
    tString = Mid(Value, StartChar)
      
   If m_TextAlign = 3 Then
      If MrgLeft = -1 Then MrgLeft = CurrentX
      If CurrentX = 0 Then
          CurrentX = MarginLeft + IndentLeft
      End If
   Else
      If SameText = False Then
         If FirstIndent = False Then
            ChangeTextAlign tString, IndentFirst
         Else
            ChangeTextAlign tString
         End If
      End If
   End If
   
'   TmpUnit.Left = m_CurrentX
'   TmpUnit.Right = MrgLeft + TextWidth(tString)
     
   'ObjPrint.Font.Charset = FontCharSet
   If SameText Then
      ObjPrint.Print tString;
    Else
      ObjPrint.Print tString
   End If
   
   If CurrentX = 0 Then CurrentX = MarginLeft + IndentLeft + IndentFirst
    
End Sub

Public Sub Paragraph(Optional ByVal Value As String = vbNullString)

     Dim TextLine() As String, i As Integer
     
    'If MarginLeft >= 0 And CurrentX < MarginLeft Then CurrentX = m_MarginLeft + m_IndentLeft
    If MarginLeft >= 0 Then CurrentX = m_MarginLeft + m_IndentLeft
    If CurrentY = 0 Then CurrentY = m_MarginTop
    If CurrentY < MarginTop Then CurrentY = m_MarginTop
            
   SetUnitParagraph.Top = CurrentY
   SetUnitParagraph.Left = CurrentX
    If Value = vbNullString Then
         CurrentY = CurrentY + ChangeLineSpace
         ObjPrint.Print Value
    Else
        Value = ReplaceChar(Value)
        If InStr(1, Value, vbCr) > 0 Then
           TextLine = Split(Value, vbCr)
        Else
           ReDim TextLine(0)
           TextLine(0) = Value
        End If
       
        For i = 0 To UBound(TextLine)
            PrintMultiLine TextLine(i), MarginLeft, PageWidth - MarginRight
        Next
    End If
    
   ' SetUnitParagraph.Left = TmpUnit.Left
    SetUnitParagraph.Right = TmpUnit.Right
    SetUnitParagraph.Bottom = CurrentY
    
    If CurrentX = 0 Then CurrentX = MarginLeft
     
End Sub

Public Sub DrawPicture(ByVal NewPic As StdPicture, _
                       ByVal Left As Variant, _
                       ByVal Top As Variant, _
                       Optional Width As Variant = 0, _
                       Optional Height As Variant = 0, _
                       Optional Opcode As RasterOpConstants = vbSrcAnd)
 
  Dim Xmin As Single
  Dim Ymin As Single
  Dim wid As Single
  Dim hgt As Single
  Dim PicBox As PictureBox
   
   Left = ScaleX(Left, ScaleMode, ScaleMode)
   Top = ScaleY(Top, ScaleMode, ScaleMode)
   
   If InStr(1, Width, "%") Then
      Width = Replace(Width, "%", "")
      Width = (Val(Width) / 100) * (NewPic.Width / 1000)
   Else
      Width = ScaleX(Width, ScaleMode, ScaleMode)
   End If
   If Width = 0 Then Width = NewPic.Width / 1000
   
   If InStr(1, Height, "%") Then
      Height = Replace(Height, "%", "")
      Height = (Val(Height) / 100) * (NewPic.Height / 1000)
   Else
     Height = ScaleY(Height, ScaleMode, ScaleMode)
   End If
   
   If Height = 0 Then Height = NewPic.Height / 1000
   
     Left = ScaleX(CSng(Left), ScaleMode, ScaleMode)
     Top = ScaleY(CSng(Top), ScaleMode, ScaleMode)
        
     picPrintPic.Cls
     Set PicBox = picPrintPic
     PicBox.Picture = NewPic
     PicBox.ScaleMode = ScaleMode
     wid = Width
     hgt = Height
     Xmin = Left
     Ymin = Top
     
     If PicBox.Picture.Type = 1 Then 'only bitmap
        ObjPrint.PaintPicture PicBox.Picture, Xmin, Ymin, wid, hgt, , , , , Opcode
     Else
        ObjPrint.PaintPicture PicBox.Picture, Xmin, Ymin, wid, hgt
     End If
     
     SetUnitPicture.Left = Xmin
     SetUnitPicture.Top = Ymin
     SetUnitPicture.Right = Xmin + wid
     SetUnitPicture.Bottom = Ymin + hgt
     Set PicBox = Nothing
End Sub

Public Sub PagePreview()
 
  On Local Error Resume Next
    
    ViewPage = ViewPage - 1
    If ViewPage < 0 Then ViewPage = 0
    Call DisplayPages
    RaiseEvent PageView
    
End Sub

Public Sub PageNext()
 
      On Local Error Resume Next
    
    ViewPage = ViewPage + 1
    If ViewPage > MaxPageNumber Then ViewPage = MaxPageNumber
    Call DisplayPages
    RaiseEvent PageView
    
End Sub

Public Sub PageFirst()
    
    On Local Error Resume Next
    ViewPage = 0
    Call DisplayPages
    RaiseEvent PageView
    
End Sub

Public Sub PageLast()

    On Local Error Resume Next
    ViewPage = MaxPageNumber
    Call DisplayPages
    RaiseEvent PageView
    
End Sub

Public Sub PageGoTo()

    Dim NewPageNo As Variant
    
    On Local Error Resume Next
    
    If ViewPage = 0 Then ViewPage = 1
    NewPageNo = InputBox("Page No")
    NewPageNo = Val(NewPageNo)
    
    If NewPageNo = 0 Then Exit Sub
    
    NewPageNo = NewPageNo - 1
    If NewPageNo > MaxPageNumber Then NewPageNo = MaxPageNumber
    ViewPage = NewPageNo
    Call DisplayPages

End Sub

Public Function TextWidth(TextString As Variant) As Single
       TextWidth = ObjPrint.TextWidth(TextString)
End Function

Public Function TextHeight(Optional TextString As String = "Hg") As Single
        TextHeight = ObjPrint.TextHeight(TextString)
End Function

Public Sub NewPage()

   On Local Error Resume Next
   
    RaiseEvent PageEnd
     
   ' PhysicalPage = PhysicalPage
    SelectPaperSize
    
    If m_SendToPrinter Then
     '
        If MaxPageNumber + 1 >= m_FromPage And MaxPageNumber + 1 <= m_ToPage Then
            PrintFooter
           If MaxPageNumber + 1 < m_ToPage Then
            Printer.NewPage
             Printer.Orientation = m_Orientation
             Printer.Print "";
             DrawRectangle 0, 0, Printer.Width, Printer.Height, , , BackColorPage, BackColorPage, vbFSSolid
             DrawBorder
             PrintHeader
           Else
           
            Printer.EndDoc

           End If
        Else
          If MaxPageNumber = m_ToPage Then
             Printer.Print "";
             Printer.KillDoc
             Printer.EndDoc
          End If
          If MaxPageNumber + 1 < m_FromPage Or MaxPageNumber > m_ToPage Then
             Printer.Print "";
             Printer.KillDoc
          'Else
             Printer.EndDoc
          End If
        End If
       Debug.Print "NewPage", MaxPageNumber + 1
    Else
       
      PrintFooter
      
      SavePicture ObjPrint.Image, TempDir & "PPView" & CStr(MaxPageNumber) & ".bmp"
      
      MakeNewPage
    End If
    
    MaxPageNumber = MaxPageNumber + 1
           
    NavBarLabels = "Pages" + Str(MaxPageNumber + 1)
      
    CurrentX = MarginLeft
    CurrentY = MarginTop
    
   RaiseEvent PageNew
End Sub

Private Sub MakeNewPage()
      If ObjPrint Is Nothing Then Exit Sub
      If SendToPrinter = True Then 'Exit Sub
      '    Printer.NewPage
      '    Printer.Orientation = m_Orientation
      '    Printer.Print "";
      '   '    DrawRectangle 0, 0, Printer.Width, Printer.Height, , , BackColorPage, BackColorPage, vbFSSolid
      '       DrawBorder
      '       PrintHeader
      Else
      
      ObjPrint.Cls
      
      SelectPaperSize
      'Printer.Orientation = m_Orientation
      'PageWidth = Printer.ScaleWidth
      'PageHeight = Printer.ScaleHeight
   
       PagePicture.Visible = False
        PicBack.Visible = False

        'Scale Object to Printer's printable area
        oSM = ObjPrint.ScaleMode
        ObjPrint.ScaleMode = m_ScaleMode
      
      If SendToPrinter = False Then
        'Full Page size (1440 twips = 1 inch or 567 twips = 1 centimeter)
        Select Case m_ScaleMode
        Case smCentimeters
            ObjPrint.Width = PageWidth * 567
            ObjPrint.Height = PageHeight * 567
           ' If MarginLeft = 0 Then MarginLeft = 2.5
           ' If MarginRight = 0 Then MarginRight = 2.5
           ' If MarginTop = 0 Then MarginTop = 2.5
           ' If MarginBottom = 0 Then MarginBottom = 2.5
        Case Else 'inches
            ObjPrint.Width = (PageWidth + 0.25) * 1440
            ObjPrint.Height = (PageHeight + 0.25) * 1440
           ' If MarginLeft = 0 Then MarginLeft = 0.25
           ' If MarginRight = 0 Then MarginRight = 0.25
           ' If MarginTop = 0 Then MarginTop = 0.25
           ' If MarginBottom = 0 Then MarginBottom = 0.25
        End Select
        
        'Set default properties of the scroll bars
        VScroll1.Max = Abs(PagePicture.Height - PicViewPort.Height + 500)
        VScroll1.Min = -200
        VScroll1.SmallChange = VScroll1.Max * 0.1
        VScroll1.LargeChange = VScroll1.Max * 0.2
        
        HScroll1.Max = Abs(PagePicture.Width - PicViewPort.Width + 700)
        HScroll1.Min = -250
        HScroll1.SmallChange = HScroll1.Max * 0.1
        HScroll1.LargeChange = HScroll1.Max * 0.2
    End If
        FontCharSet = CharSetWin
        ObjPrint.Picture = Nothing
        
        'Page Back color
        DrawRectangle 0, 0, PagePicture.Width, PagePicture.Height, , , BackColorPage, BackColorPage, vbFSSolid
        DrawBorder
        PrintHeader
        CurrentY = MarginTop
     End If
End Sub

Public Sub Clear()
     Dim i As Integer
     For i = 0 To MaxPageNumber
         If FileExists(TempDir & "PPView" & CStr(i) & ".bmp") Then
            Kill TempDir & "PPView" & CStr(i) & ".bmp"
         End If
     Next
     LineSpace = lsSpaceSingle
      
End Sub

Public Sub GetMargins()
         x1 = MarginLeft
         x2 = MarginRight
         y1 = MarginTop
         y2 = MarginBottom
End Sub

Private Function SelectPaperSize()

   Dim SW As Single, SH As Single
   
    Select Case PaperSize
    Case 1, 2, 3, 4, 5, 6, 7, 14, 16, 17, 18, 19, 20, 21, 22, 23, 37, 38, 39, 40, 41 'Inches
          ScaleMode = sminches
          'Printer.ScaleMode = sminches
    Case 8, 9, 10, 11, 12, 13, 15, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36 'Centimeters
           ScaleMode = smCentimeters
          'Printer.ScaleMode = smCentimeters
    End Select

    Select Case PaperSize
    Case 1, 2 'DMPAPER_LETTER = 1               ' Letter 8 1/2 x 11 in
    'Case 2 'DMPAPER_LETTERSMALL = 2          ' Letter Small 8 1/2 x 11 in
            SW = 8.5
            SH = 11
    Case 3 ' DMPAPER_TABLOID = 3              ' Tabloid 11 x 17 in
            SW = 11
            SH = 17
    Case 4 ' DMPAPER_LEDGER = 4               ' Ledger 17 x 11 in
            SW = 17
            SH = 11
    Case 5 ' DMPAPER_LEGAL = 5                ' Legal 8 1/2 x 14 in
            SW = 8.5
            SH = 14
    Case 6 ' DMPAPER_STATEMENT = 6            ' Statement 5 1/2 x 8 1/2 in
            SW = 5.5
            SH = 8.5
    Case 7 ' DMPAPER_EXECUTIVE = 7            ' Executive 7 1/4 x 10 1/2 in
            SW = 7.25
            SH = 10.5
    Case 8 ' DMPAPER_A3 = 8                   ' A3 297 x 420 mm
            SW = 29.7
            SH = 42
    Case 9, 10 ' DMPAPER_A4 = 9                   ' A4 210 x 297 mm
    'Case 10 ' DMPAPER_A4SMALL = 10             ' A4 Small 210 x 297 mm
            SW = 21
            SH = 29.7
    Case 11 ' DMPAPER_A5 = 11                  ' A5 148 x 210 mm
            SW = 14.8
            SH = 21
    Case 12 ' DMPAPER_B4 = 12                  ' B4 250 x 354
            SW = 25
            SH = 35.4
    Case 13 ' DMPAPER_B5 = 13                  ' B5 182 x 257 mm
            SW = 18.2
            SH = 25.7
    Case 14 ' DMPAPER_FOLIO = 14               ' Folio 8 1/2 x 13 in
            SW = 8.5
            SH = 13
    Case 15 ' DMPAPER_QUARTO = 15              ' Quarto 215 x 275 mm
            SW = 21.5
            SH = 27.5
    Case 16 ' DMPAPER_10X14 = 16               ' 10x14 in
            SW = 10
            SH = 14
    Case 17 ' DMPAPER_11X17 = 17               ' 11x17 in
            SW = 11
            SH = 17
    Case 18 ' DMPAPER_NOTE = 18                ' Note 8 1/2 x 11 in
            SW = 8.5
            SH = 11
    Case 19 ' DMPAPER_ENV_9 = 19               ' Envelope #9 3 7/8 x 8 7/8
            SW = 3.875
            SH = 8.875
    Case 20 ' DMPAPER_ENV_10 = 20              ' Envelope #10 4 1/8 x 9 1/2
            SW = 4.125
            SH = 9.5
    Case 21 ' DMPAPER_ENV_11 = 21              ' Envelope #11 4 1/2 x 10 3/8
            SW = 4.5
            SH = 10.375
    Case 22 ' DMPAPER_ENV_12 = 22              ' Envelope #12 4 3/4 x 11
            SW = 4.75
            SH = 11
    Case 23 ' DMPAPER_ENV_14 = 23              ' Envelope #14 5 x 11 1/2
            SW = 5
            SH = 11.5
    'Case 24 ' DMPAPER_CSHEET = 24              ' C size sheet
    'Case 25 'DMPAPER_DSHEET = 25              ' D size sheet
    'Case 26 'DMPAPER_ESHEET = 26              ' E size sheet
    Case 27 'DMPAPER_ENV_DL = 27              ' Envelope DL 110 x 220mm
            SW = 11
            SH = 22
    Case 28 'DMPAPER_ENV_C5 = 28              ' Envelope C5 162 x 229 mm
            SW = 16.2
            SH = 22.9
    Case 29 'DMPAPER_ENV_C3 = 29              ' Envelope C3  324 x 458 mm
            SW = 32.4
            SH = 45.8
    Case 30 'DMPAPER_ENV_C4 = 30              ' Envelope C4  229 x 324 mm
            SW = 22.9
            SH = 32.4
    Case 31 'DMPAPER_ENV_C6 = 31              ' Envelope C6  114 x 162 mm
            SW = 11.4
            SH = 16.2
    Case 32 'DMPAPER_ENV_C65 = 32             ' Envelope C65 114 x 229 mm
            SW = 11.4
            SH = 22.9
    Case 33 'DMPAPER_ENV_B4 = 33              ' Envelope B4  250 x 353 mm
            SW = 25
            SH = 36.3
    Case 34 'DMPAPER_ENV_B5 = 34              ' Envelope B5  176 x 250 mm
            SW = 17.6
            SH = 25
    Case 35 'DMPAPER_ENV_B6 = 35              ' Envelope B6  176 x 125 mm
            SW = 17.6
            SH = 12.5
    Case 36 'DMPAPER_ENV_ITALY = 36           ' Envelope 110 x 230 mm
            SW = 11
            SH = 23
    Case 37 'DMPAPER_ENV_MONARCH = 37         ' Envelope Monarch 3.875 x 7.5 in
            SW = 3.875
            SH = 7.5
    Case 38 'DMPAPER_ENV_PERSONAL = 38        ' 6 3/4 Envelope 3 5/8 x 6 1/2 in
            SW = 3.625
            SH = 6.5
    Case 39 'DMPAPER_FANFOLD_US = 39          ' US Std Fanfold 14 7/8 x 11 in
            SW = 14.875
            SH = 11
    Case 40 'DMPAPER_FANFOLD_STD_GERMAN = 40  ' German Std Fanfold 8 1/2 x 12 in
            SW = 8.5
            SH = 12
    Case 41 'DMPAPER_FANFOLD_LGL_GERMAN = 41  ' German Legal Fanfold 8 1/2 x 13 in
            SW = 8.5
            SH = 13
    'Case 42 'DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN
    Case 256 'DMPAPER_USER = 256
            SW = 21
            SH = 29.7
    End Select
    If m_Orientation = 1 Then
        PageWidth = SW
        PageHeight = SH
    Else
        PageWidth = SH
        PageHeight = SW
    End If
End Function

Public Sub TextBox(ByVal Text As String, _
                   ByVal X As Variant, ByVal Y As Variant, _
                   ByVal Width As Variant, ByVal Height As Variant, _
                   Optional ByVal tAling As TextAlignConstants = taLeftTop, _
                   Optional ByVal BackShade As Boolean = True, _
                   Optional ByVal BoxShade As Boolean = False)
  
  Dim x1 As Single, y1 As Single
  Dim oCX As Single, oCy As Single
  Dim oWL As Single, oWR As Single
  Dim oWT As Single, oWB As Single
  Dim len1 As Integer, P_Fix As Single
  Dim oIL As Single, oIR As Single
  Dim oAl As TextAlignConstants, oFT As Boolean
  Dim oDW As Integer, tmpTxt As String
  Dim TextLine() As String, id As Integer, sdpix As Single
  Dim tHeight As Single, sX As Single, sY As Single
  
  X = ScaleX(X, ScaleMode, ScaleMode)
  Y = ScaleY(Y, ScaleMode, ScaleMode)
  Width = ScaleX(Width, ScaleMode, ScaleMode)
  Height = ScaleY(Height, ScaleMode, ScaleMode)
  
  Text = ReplaceChar(Text)
  'TextLine = Split(Text, vbCr)
  TextLine = Split(BreakItemText(Text, Width - (m_IndentLeft + m_IndentRight)), vbCr)
     
  If Height = 0 Then
     Height = ((UBound(TextLine) + 1) * TextHeight("H")) + ChangeLineSpace * 2
 ' Else
     'TextLine = Split(Text, vbCr)
  End If
  
  SetUnitTextBox.Left = X
  SetUnitTextBox.Right = X + Width
  SetUnitTextBox.Top = Y
  SetUnitTextBox.Bottom = Y + Height
  
  'Read Default setting
  oCX = CurrentX
  oWL = MarginLeft
  oWR = MarginRight
  oWT = MarginTop
  oWB = MarginBottom
  oAl = m_TextAlign
  oDW = DrawWidth
  oFT = FontTransparent
  oFS = FillStyle
  oIL = m_IndentLeft
  oIR = m_IndentRight
  
  'New Setting
  m_MarginLeft = X
  m_MarginRight = m_PageWidth - (Width) - X
  m_MarginTop = Y
  m_MarginBottom = Height
  If BackShade = False Then FontTransparent = True
  tmpTxt = Text
  x1 = X
  y1 = Y

  If BoxShade Then
     FillStyle = vbFSSolid
     If SendToPrinter = False Then sdpix = 3 Else sdpix = 15
     sX = ScaleX(sdpix, vbPixels, m_ScaleMode)
     sY = ScaleY(sdpix, vbPixels, m_ScaleMode)
     DrawRectangle x1 + sX, y1 + sY, Width + sX, Height + sY, , , &H808080, &H808080, vbFSSolid
  End If
  
  If BackShade Then
     'FillStyle = vbFSSolid
     DrawRectangle x1, y1, Width, Height, , , , FillColor, FillStyle
  End If
  
  If tAling >= taLeftMiddle Then
    Select Case tAling
    Case taLeftMiddle, taRightMiddle, taCenterMiddle, taJustifyMiddle
        CurrentY = Y + ((Height - ((UBound(TextLine) + 1) * TextHeight("Hj")) + ChangeLineSpace * 2) / 2)
    Case taLeftBottom, taRightBottom, taCenterBottom, taJustifyBottom
        CurrentY = Y + Height - ((UBound(TextLine) + 1) * TextHeight("Hj")) + ChangeLineSpace * 2
    End Select
    If CurrentY > Y + Height Then
    Stop
    End If
  Else
    CurrentY = y1
  End If
  CurrentX = x1

  TextAlign = tAling
  tHeight = CurrentY
  
   For id = 0 To UBound(TextLine)
     tHeight = tHeight + TextHeight
     'ChangeTextAlign TextLine(id)
     PrintMultiLine TextLine(id), x1, Width + x1, Y, Height, False
  Next

  'Return Default setting
  FillStyle = oFS
  TextAlign = oAl
  CurrentX = oCX
  CurrentY = Y + Height
  MarginLeft = oWL
  MarginRight = oWR
  MarginTop = oWT
  MarginBottom = oWB
  FontTransparent = oFT
  IndentLeft = oIL
  IndentRight = oIR
End Sub

'   H E L P E R   F U N C T I O N S
Private Function BreakItemText(ByVal Value As String, ByVal Width As Single, Optional WrapText As String = "") As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''     Breaks a string into multiple lines that will fit in the
'''     specified width when rendered on the output device (screen or printer).
'''     The Value string is broken at work boundaries.
'''         *** YOU MUST PROVIDE AN ADEQUATE WIDTH TO FIT THE LONGEST WORD!
'''         *** THE FUNCTION WON'T BREAK A WORD THAT EXCEEDS THE SPECIFIED
'''         *** WIDTH AND, AS A RESULT, THE STRING WILL LEAK INTO ADJACENT
'''         *** CELLS WHEN PRINTED !!!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TextWidth(Value) < Width Then
        If Value = "" Then Value = " "
        BreakItemText = Value '+ vbCrLf
    Else
        Dim iChar As Integer: iChar = 1
        Dim NewValue As String: NewValue = ""
        Dim moreWords As Boolean: moreWords = True
        Dim nextWord As String
        While moreWords
            nextWord = GetNextWord(Value, iChar)
            iChar = iChar + Len(nextWord)
            If WrapText = "" Then
                If TextWidth(NewValue & nextWord) < Width Then
                    NewValue = NewValue & nextWord
                Else
                   If NewValue <> "" Then
                    NewValue = NewValue & vbCr & nextWord
                   Else
                     NewValue = NewValue & nextWord
                   End If
                End If
                BreakItemText = NewValue
                If iChar > Len(Value) Then
                    moreWords = False
                End If
            Else
               NewValue = ""
               For iChar = 1 To Len(Value)
                   NewValue = NewValue & Mid(Value, iChar, 1)
                   If TextWidth(NewValue & WrapText) < Width Then
                   Else
                      NewValue = Mid(Value, 1, iChar - 1) & WrapText
                      BreakItemText = NewValue
                      Exit Function
                   End If
               Next
            End If
        Wend
    End If
End Function
'----
Private Function GetNextWord(ByVal Value As String, ByVal pos As Integer)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''     Retrieves the following word in the specified Valueing,
'''     starting at character pos in the Valueing
'''     Spaces are added appended to the selected word
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim nextWord As String
    While pos <= Len(Value) And Mid(Value, pos, 1) <> " "
        nextWord = nextWord & Mid(Value, pos, 1)
        pos = pos + 1
    Wend
    While pos <= Len(Value) And Mid(Value, pos, 1) = " "
        nextWord = nextWord & Mid(Value, pos, 1)
        pos = pos + 1
    Wend
    GetNextWord = nextWord
End Function

Public Sub Text(Optional ByVal Value As String = vbNullString)

    Dim TextLine() As String, i As Integer, l As Integer, E As Integer
    
     
    TextLine = Split(Value, vbCr)
    
    For E = 0 To UBound(TextLine)
        
        If TextLine(E) = "" Then TextLine(E) = vbCr
        
        Value = TextLine(E)
        
        If CurrentX + TextWidth(Value) > PageWidth - MarginRight - CurrentX Or _
           CurrentY > PageHeight - MarginBottom Then
           
         If CurrentY + TextHeight(Value) + ChangeLineSpace * 2 > PageHeight - MarginBottom Then
    
          If EndOfPage Then
            PrintFooter
            ';DrawBorder
            NewPage
            PrintHeader
            'if last line is Cr
            If TextLine(E) = vbCr And E = UBound(TextLine) Then
               Exit Sub
            End If
    
          End If
          End If
        End If
    
        If CurrentX + TextWidth(Value) > PageWidth - MarginRight Then
             TextLine = Split(BreakItemText(Value, PageWidth - MarginRight - CurrentX), vbCr)
1:
             For i = 0 To UBound(TextLine)
                 If PageWidth - MarginRight - CurrentX < TextWidth(TextLine(i)) Then
                    ObjPrint.Print vbCr;
                    CurrentY = CurrentY + ChangeLineSpace
                    CurrentX = MarginLeft
                    Value = ""
                    For l = i To UBound(TextLine)
                       If Value = vbNullString Then
                          Value = TextLine(l)
                       Else
                          Value = Value + " " + TextLine(l)
                       End If
                    Next
                    TextLine = Split(BreakItemText(Value, PageWidth - MarginRight - CurrentX), vbCr)
                    GoTo 1
                 End If
                 PrintMultiLine TextLine(i), CurrentX, PageWidth - MarginRight, , MarginBottom, , True
             Next
             Exit Sub
        Else
            ObjPrint.Print Value;
            If CurrentX = 0 Then CurrentX = MarginLeft
        End If
    Next
    
End Sub

Public Sub DrawRectangle(ByVal X As Variant, _
                         ByVal Y As Variant, _
                         ByVal Width As Variant, _
                         ByVal Height As Variant, _
                         Optional ByVal Radius1 As Variant = 0, _
                         Optional ByVal Radius2 As Variant = 0, _
                         Optional ByVal ColorLine As Long = -1, _
                         Optional ByVal ColorFill As Long = -1, _
                         Optional FilledBox As FillStyleConstants = vbFSTransparent)
                         
    Dim RC As Integer, oFcl As Long
    Dim lPen As Long, hOldPen As Long, oldBrush As Long
    Dim bR As LogBrush, hBrush As Long
    
    If ColorLine = -1 Then ColorLine = ForeColor
    If ColorFill = -1 Then ColorFill = ColorLine

     oFC = FillColor
     oFS = FillStyle
    oFcl = ForeColor
    
    FillColor = ColorFill
    FillStyle = FilledBox
    ForeColor = ColorLine
    
    X = ScaleX(X, ScaleMode, vbPixels)
    Y = ScaleY(Y, ScaleMode, vbPixels)
    Width = ScaleX(Width, ScaleMode, vbPixels)
    Height = ScaleY(Height, ScaleMode, vbPixels)
    If Radius1 <> 0 Then Radius1 = ScaleX(Radius1, ScaleMode, vbPixels)
    If Radius2 <> 0 Then Radius2 = ScaleY(Radius2, ScaleMode, vbPixels)
    
'    bR.lbColor = m_FillColor
'    bR.lbHatch = m_DrawStyle
'    bR.lbStyle = 2
'    hBrush = CreateBrushIndirect(bR)
    lPen = CreatePen(m_DrawStyle, DrawWidth, ColorLine)
'    lPen = PenCreate(ColorLine)
    
    'oldBrush = SelectObject(ObjPrint.hDC, hBrush)
    hOldPen = SelectObject(ObjPrint.hdc, lPen)
      
    RC = RoundRect(ObjPrint.hdc, X, Y, X + Width, Y + Height, Radius1, Radius2)
    
    If RC = 0 Then
        RaiseEvent Error(1002, "Error DrawRectangle")
    End If
    
    Call SelectObject(ObjPrint.hdc, hOldPen)
    'Call SelectObject(ObjPrint.hDC, oldBrush)
    DeleteObject lPen
    'DeleteObject oldBrush
    
    ForeColor = oFcl
    FillColor = oFC
    FillStyle = oFS
        
End Sub


'Create pen Style
Private Function PenCreate(ColorLine As Long) As Long
Dim BrushInf As LogBrush
Dim StyleArr() As Long
Dim wLine As Long
Dim PenStyle As Long
    
    wLine = DrawWidth ' mWidthLine
        
    Select Case m_DrawStyle
    Case 0 'vbSolid
       ReDim StyleArr(1)
       StyleArr(0) = 10
       StyleArr(1) = 0
    Case 1 'vbDash
       ReDim StyleArr(1)
        StyleArr(0) = 18 * (wLine / 2)
        StyleArr(1) = 6 * (wLine / 2)
    Case 2 'vbDot
       ReDim StyleArr(3)
        StyleArr(0) = 3 * (wLine / 2)
        StyleArr(1) = 3 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 3 * (wLine / 2)
    Case 3 'vbDashDot
       ReDim StyleArr(3)
        StyleArr(0) = 9 * (wLine / 2)
        StyleArr(1) = 6 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 6 * (wLine / 2)
    Case 4 'vbDashDotDot
        ReDim StyleArr(5)
        StyleArr(0) = 9 * (wLine / 2)
        StyleArr(1) = 3 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 3 * (wLine / 2)
        StyleArr(4) = 3 * (wLine / 2)
        StyleArr(5) = 3 * (wLine / 2)
    Case 5 'vbInvisible
        PenCreate = CreatePen(PS_NULL, wLine, ColorLine)
        Exit Function
    End Select
    
    BrushInf.lbColor = ColorLine
    PenCreate = ExtCreatePen(PS_GEOMETRIC Or PS_USERSTYLE Or PS_ENDCAP_ROUND, wLine, BrushInf, UBound(StyleArr()) + 1, StyleArr(0))
    
    Erase StyleArr
    
End Function
Public Sub DrawLine(ByVal x1 As Variant, _
                    ByVal y1 As Variant, _
                    Optional ByVal x2 As Variant, _
                    Optional ByVal y2 As Variant, _
                    Optional ByVal LineWidth As Integer = 1, _
                    Optional ByVal ColorLine As Long = -1)
     
    If LineWidth > 0 Then
        oDW = DrawWidth
        DrawWidth = LineWidth
    End If
    
    If ColorLine <> -1 Then
       oFC = ForeColor
       ForeColor = ColorLine
    End If
    
    x1 = ScaleX(x1, ScaleMode, ScaleMode)
    y1 = ScaleX(y1, ScaleMode, ScaleMode)
    
    If IsMissing(x2) And IsMissing(y2) Then
        ObjPrint.Line -(x1, y1), ForeColor
    Else
        x2 = ScaleX(x2, ScaleMode, ScaleMode)
        y2 = ScaleX(y2, ScaleMode, ScaleMode)
        ObjPrint.Line (x1, y1)-(x2, y2), ForeColor
    End If
    
    If LineWidth > 0 Then DrawWidth = oDW
    If ColorLine <> -1 Then ForeColor = oFC
    
End Sub

Public Sub DrawPolyline(ByVal PointsLine As String, _
                        Optional ByVal LineWidth As Integer = 1, _
                        Optional ByVal ColorLine As Long = -1)

      Dim RC As Integer, nowCount As Integer
      Dim sArr() As String, nArr() As String, a As Integer
      Dim oFS As Integer
      Dim Point() As POINTAPI
      ReDim Point(0)
      
      If IsNull(PointsLine) Then Exit Sub
      sArr = Split(PointsLine, ",")
      If IsArray(sArr) = False Then Exit Sub
      If UBound(sArr) + 1 < 2 Then Exit Sub
      
      If LineWidth > 0 Then
         oDW = DrawWidth
         DrawWidth = LineWidth
      End If
    
      nowCount = 0
      For a = 0 To UBound(sArr)
         nArr = Split(sArr(a), " ")
         nowCount = nowCount + 1
         ReDim Preserve Point(nowCount)
         Point(nowCount).X = ScaleX(nArr(0), ScaleMode, vbPixels)
         Point(nowCount).Y = ScaleY(nArr(1), ScaleMode, vbPixels)
      Next
     
      RC = Polyline(ObjPrint.hdc, Point(1), nowCount)
      
      If RC = 0 Then
         RaiseEvent Error(1001, "Error DrawPolyline")
      End If
      
      If LineWidth > 1 Then DrawWidth = oDW
      If ColorLine <> -1 Then ForeColor = oFC
      
End Sub

Public Sub DrawPolygon(ByVal Points As String, _
                       Optional ByVal ColorLine As Long = -1, _
                       Optional ByVal ColorFill As Long = -1, _
                       Optional FilledStylePolygon As FillStyleConstants = vbSolid, _
                       Optional FillPolyMode As PolyFillMode = WINDING)
      
      Dim RC As Integer, nowCount As Integer, oFC1 As Long
      Dim sArr() As String, nArr() As String, a As Integer
      Dim Point() As POINTAPI
      ReDim Point(0)
      
      If IsNull(Points) Then Exit Sub
      
      sArr = Split(Points, ",")
      If IsArray(sArr) = False Then Exit Sub
      If UBound(sArr) + 1 < 2 Then Exit Sub
      
      If ColorLine = -1 Then ColorLine = ForeColor
      If ColorFill = -1 Then ColorFill = ColorLine
    
      oFS = FillStyle
      oFC = ForeColor
      oFC1 = FillColor
      FillStyle = FilledStylePolygon
      ForeColor = ColorLine
      FillColor = ColorFill
      nowCount = 0
      For a = 0 To UBound(sArr)
         sArr(a) = Replace(sArr(a), "  ", " ")
         nArr = Split(Trim(sArr(a)), " ")
         nowCount = nowCount + 1
         ReDim Preserve Point(nowCount)
         Point(nowCount).X = ScaleX(nArr(0), ScaleMode, vbPixels)
         Point(nowCount).Y = ScaleY(nArr(1), ScaleMode, vbPixels)
      Next
      
      RC = SetPolyFillMode(ObjPrint.hdc, FillPolyMode)
      RC = Polygon(ObjPrint.hdc, Point(1), nowCount)
      
      FillStyle = oFS
      ForeColor = oFC
      FillColor = oFC1
      If RC = 0 Then
         RaiseEvent Error(1002, "Error DrawPolygon")
      End If
End Sub

Public Sub DrawCircle(ByVal CenterX As Variant, _
                      ByVal CenterY As Variant, _
                      ByVal Radius As Variant, _
                      Optional cStart As Single = 0, _
                      Optional cEnd As Single = 0, _
                      Optional ByVal ColorLine As Long = -1, _
                      Optional ByVal ColorFill As Long = -1, _
                      Optional cDrawStyle As DrawStyleConstants = vbSolid, _
                      Optional cFillStyle As FillStyleConstants = vbFSTransparent)
                      
        
        If ColorLine = -1 Then ColorLine = ForeColor
        If ColorFill = -1 Then ColorFill = ColorLine
    
        oFS = FillStyle
        oFC = FillColor
        oDS = DrawStyle
        
        FillColor = ColorFill
        DrawStyle = cDrawStyle
        FillStyle = cFillStyle
        
        CenterX = ScaleX(CenterX, ScaleMode, ScaleMode)
        CenterY = ScaleX(CenterY, ScaleMode, ScaleMode)
        Radius = ScaleX(Radius, ScaleMode, ScaleMode)
        If cStart = 0 And cEnd = 0 Then
            ObjPrint.Circle (CenterX, CenterY), Radius, ColorLine
        Else
            ObjPrint.Circle (CenterX, CenterY), Radius, ColorLine, cStart, cEnd
        End If
        
        FillStyle = oFS
        DrawStyle = oDS
        FillColor = oFC
End Sub

Public Sub DrawEllipse(ByVal Left As Variant, ByVal Top As Variant, _
                       ByVal Width As Variant, ByVal Height As Variant, _
                       Optional cStart As Single = 0, Optional cEnd As Single = 0, _
                       Optional ByVal ColorLine As Long = -1, _
                       Optional ByVal ColorFill As Long = -1, _
                       Optional cDrawStyle As DrawStyleConstants = vbSolid, _
                       Optional cFillStyle As FillStyleConstants = vbFSTransparent)
                   
    Dim AspectRatio As Single, Radius As Single
    
    If ColorLine = -1 Then ColorLine = ForeColor
    If ColorFill = -1 Then ColorFill = ColorLine
    
    AspectRatio = Height / Width
    Radius = IIf(Width > Height, Width / 2, Height / 2)
    
    oFS = FillStyle
    oFC = FillColor
    oDS = DrawStyle
        
    FillColor = ColorFill
    DrawStyle = cDrawStyle
    FillStyle = cFillStyle
    
    Left = ScaleX(Left, ScaleMode, ScaleMode)
    Top = ScaleY(Top, ScaleMode, ScaleMode)
    Width = ScaleX(Width, ScaleMode, ScaleMode)
    Height = ScaleY(Width, ScaleMode, ScaleMode)
     If cStart = 0 And cEnd = 0 Then
        ObjPrint.Circle ((Left / 2) + (Left + Width) / 2, (Top / 2) + (Top + Height) / 2), Radius, ColorLine, , , AspectRatio
     Else
       ObjPrint.Circle ((Left / 2) + (Left + Width) / 2, (Top / 2) + (Top + Height) / 2), Radius, ColorLine, cStart, cEnd, AspectRatio
    End If
    FillStyle = oFS
    DrawStyle = oDS
    FillColor = oFC
End Sub

'H E L P  F U N C T I O N
Private Sub ChangeTextAlign(tString As String, Optional FisrtIndent As Single = 0)

        Dim tArr() As String, IdArr As Integer, nString As String, i As Integer
        
'   taLeftTop = 0
'   taRightTop = 1
'   taCenterTop = 2
'   taJustifyTop = 3
'   taLeftMiddle = 4
'   taRightMiddle = 5
'   taCenterMiddle = 6
'   taJustifyMiddle = 7
'   taLeftBottom = 8
'   taRightBottom = 9
'   taCenterBottom = 10
'   taJustifyBottom = 11
        
        Select Case m_TextAlign
        Case taLeftTop, taLeftMiddle, taLeftBottom 'Left
             tString = RTrim(tString)
             CurrentX = m_MarginLeft + m_IndentLeft + FisrtIndent
        Case taRightTop, taRightMiddle, taRightBottom 'Right
             tString = RTrim(tString)
             CurrentX = (m_PageWidth - m_MarginRight) - TextWidth(tString) - m_IndentRight - FisrtIndent
        Case taCenterTop, taCenterMiddle, taCenterBottom 'Center
             tString = Trim(tString)
             CurrentX = MarginLeft + ((m_PageWidth - m_MarginRight - m_MarginLeft - TextWidth(tString)) / 2)
        Case taJustifyTop, taJustifyMiddle, taJustifyBottom 'Justify
          If Right(tString, 1) <> "." Then
          If tString <> "" Then
             tString = RTrim(tString)
             tArr = Split(tString, " ")
             IdArr = 0
             'Do Until TextWidth(nString) > m_PageWidth - m_MarginLeft - m_MarginRight - (m_IndentLeft + m_IndentRight) - FisrtIndent
             Do Until (TextWidth(nString) + TextWidth(" ")) >= (m_PageWidth - (m_MarginRight + m_MarginLeft) - (m_IndentLeft + m_IndentRight) - FisrtIndent)
                If IdArr > UBound(tArr) - 1 Then
                  IdArr = 0
                End If
                tArr(IdArr) = tArr(IdArr) + " "
                IdArr = IdArr + 1
                nString = ""
                For i = 0 To UBound(tArr)
                If nString = "" Then
                    nString = tArr(i)
                Else
                    nString = nString + " " + tArr(i)
                End If
                Next
             Loop
             If (TextWidth(nString) + TextWidth(" ")) > (m_PageWidth - (m_MarginRight + m_MarginLeft) - (m_IndentLeft + m_IndentRight) - FisrtIndent) Then
               'nString = mid ( nString, instr(1,nString,"  ") , " ")
             End If
          End If
             tString = nString
          Else
          
          End If
             
             CurrentX = m_MarginLeft + m_IndentLeft + FisrtIndent
        End Select
End Sub

'Change Line Space
Private Function ChangeLineSpace() As Single
    Dim OldTH As Single, NewCY As Single
    
    OldTH = TextHeight("W")
    
    Select Case m_LineSpace
    Case lsSpaceSingle '= 0
       Exit Function
       
    Case lsSpaceLine15 '1
       NewCY = OldTH * 0.33
       
    Case lsSpaceDoubleline '2
      NewCY = OldTH * 0.75
      
    Case lsSpaceHalfline '3
      NewCY = -OldTH / 4
    End Select
    ChangeLineSpace = NewCY
    
End Function

'no
Public Sub CreateThumbNail()

     Dim nPic As PictureBox, i As Integer
     Dim uPic As PictureBox, wdt As Long, hgt As Long, vsmax As Integer
     Dim Col As Single, Row As Single, mTop As Integer, wStep As Long, mLeft As Long, l As Integer
     Dim X As Integer, Y As Integer, AA As Integer, nRow As Integer
     On Error Resume Next
    
     For Each uPic In ThumbNail
         If uPic.Index > 0 Then
            Unload uPic
         End If
     Next
    
     For i = 0 To MaxPageNumber
        If i > 0 Then
            Load ThumbNail(i)
        End If
        picFullPage.Picture = LoadPicture(TempDir & "PPView" & CStr(i) & ".bmp")
        
        ThumbNail(i).ScaleMode = ScaleMode
        ThumbNail(i).Width = ScaleX(picFullPage.Width, vbTwips, ScaleMode) * 100
        ThumbNail(i).Height = ScaleX(picFullPage.Height, vbTwips, ScaleMode) * 100
        ThumbNail(i).Enabled = False
        ThumbNail(i).Visible = False
        ThumbNail(i).ToolTipText = "Page" + Str(i + 1)
        PicRefresh ThumbNail(i), picFullPage
     Next
     
    If ThumbNail(0).Width > ThumbNail(0).Height Then
        wStep = ThumbNail(0).Width
    Else
        wStep = ThumbNail(0).Height
    End If
     
    AA = 0
    mLeft = 100
    mTop = 0
    l = 0
    vsmax = 0
    
    For i = 0 To MaxPageNumber
        If mLeft + (wStep * 2) + 50 > PicViewPort.ScaleWidth Then
           mTop = mTop + wStep + 50
           mLeft = 150
           l = 1
           AA = AA + 1
        Else
           mLeft = mLeft + wStep * l + 50
           If l = 0 Then l = l + 1
        End If
        Col = (wStep - ThumbNail(i).Width) / 2
        Row = (wStep - ThumbNail(i).Height) / 2
        ThumbNail(i).Move mLeft + Col, mTop + Row
        ThumbNail(i).Enabled = True
        ThumbNail(i).Visible = True
    Next
    vsmax = PicViewPort.ScaleHeight / ThumbNail(0).Height 'ScaleY(PicViewPort.ScaleHeight, PicViewPort.ScaleMode, ThumbNail(0).ScaleMode)
    VScroll1.Max = AA \ vsmax + IIf(AA / vsmax > AA \ vsmax, 1, 0) + 1
    If AA >= vsmax Then
       VScroll1.Move PicViewPort.ScaleWidth - VScroll1.Width, 0
       VScroll1.Height = PicViewPort.ScaleHeight
       VScroll1.Min = 0
       VScroll1.SmallChange = 1
       VScroll1.LargeChange = 1
       VScroll1.Value = 0
       VScroll1.Visible = True
    End If
End Sub
'
Public Function ShowSaveFile(Optional Title As String = "Export Image", _
                             Optional strFilter As String = "", _
                             Optional PathName As String = "", _
                             Optional FileName As String = "") As String
    
    Dim OFName As OPENFILENAME
    
    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hWndOwner = UserControl.hWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    
    'Set the filet
    If strFilter = "" Then
       OFName.lpstrFilter = "BMP File (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
        OFName.lpstrDefExt = "bmp"
    Else
       OFName.lpstrFilter = strFilter
    End If
    
    'Create a buffer
    If FileName = "" Then
        OFName.lpstrFile = Space$(254) + vbNullChar
    Else
        OFName.lpstrFile = FileName + "." + OFName.lpstrDefExt + Space$(254 - Len(FileName + "." + OFName.lpstrDefExt)) + vbNullChar
    End If
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
   ' If FileName = "" Then
        OFName.lpstrFileTitle = Space$(512) + vbNullChar
   ' Else
      '  OFName.lpstrFileTitle = FileName + Space$(254)
   ' End If
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    
    'OFName.lpstrInitialDir = "C:\"
    If PathName = "" Then
       OFName.lpstrInitialDir = App.Path
    Else
      OFName.lpstrInitialDir = PathName
    End If

    'Set the dialog title
    If Title <> "" Then
        OFName.lpstrTitle = Title
    Else
        OFName.lpstrTitle = "Save File"
    End If
    'no extra flags
    OFName.flags = 0

    'Show the 'Save File'-dialog
    If GetSaveFileName(OFName) Then
        ShowSaveFile = Trim$(OFName.lpstrFile)
        ShowSaveFile = Trim(Replace(ShowSaveFile, vbNullChar, ""))
    Else
        ShowSaveFile = ""
    End If
End Function

Public Function DialogPrint(DialogType As PrintDialogSettings) As Boolean
        Select Case DialogType
        Case 0
             DialogPrint = ShowPrinter(PD_PRINTSETUP)
        Case 1
             DialogPrint = ShowPrinterSetup 'ShowPageSetupDlg
        Case 2
             DialogPrint = ShowPrinter
        End Select
End Function

'Private Function ShowPageSetupDlg() As Boolean
'    Dim PSD As PageSetupDialog
'
'    'Set the structure size
'    PSD.lStructSize = Len(PSD)
'
'    'Set the owner window
'    PSD.hwndOwner = UserControl.hWnd
'
'    'Set the application instance
'    PSD.hInstance = App.hInstance
'    PSD.rtMargin.Left = MarginLeft * 1000
'    PSD.rtMargin.Top = MarginTop * 1000
'    PSD.rtMargin.Bottom = MarginBottom * 1000
'    PSD.rtMargin.Right = MarginRight * 1000
'    PSD.ptPaperSize.X = PageWidth * 1000
'    PSD.ptPaperSize.Y = PageHeight * 1000
'
'    'extra flags set MARGINS
'    PSD.flags = PSD_MARGINS Or PSD_DISABLEPAPER Or IIf(ScaleMode = smCentimeters, PSD_INHUNDREDTHSOFMILLIMETERS, PSD_INTHOUSANDTHSOFINCHES) Or PSD_DISABLEPRINTER
'
'    'Show the pagesetup dialog
'    If PAGESETUPDLG(PSD) Then
'        MarginLeft = PSD.rtMargin.Left / 1000
'        MarginTop = PSD.rtMargin.Top / 1000
'        MarginBottom = PSD.rtMargin.Bottom / 1000
'        MarginRight = PSD.rtMargin.Right / 1000
'        PageWidth = PSD.ptPaperSize.X / 1000
'        PageHeight = PSD.ptPaperSize.Y / 1000
'
'        If PageWidth > PageHeight Then Orientation = PageLandscape Else Orientation = PagePortrait
'        ShowPageSetupDlg = True
'    Else
'        ShowPageSetupDlg = False
'    End If
'
'End Function

Private Function ShowPrinterSetup() As Boolean
    
    Dim PRINTSETUPDLG As PageSetupDialog
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE
          
    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String
    Dim strSetting As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PRINTSETUPDLG.lStructSize = Len(PRINTSETUPDLG)
    PRINTSETUPDLG.hWndOwner = UserControl.hWnd

    ' Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX Or DM_COPIES
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmCopies = Printer.Copies
    On Error Resume Next
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    ' Allocate memory for the initialization hDevMode structure
    ' and copy the settings gathered above into this memory
    PRINTSETUPDLG.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PRINTSETUPDLG.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PRINTSETUPDLG.hDevMode)
    End If

    ' Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    
    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    
    'allocate appropriate values to rtMargin Structure
    'PRINTSETUPDLG.hInstance = App.hInstance
    PRINTSETUPDLG.rtMargin.Left = MarginLeft * 1000
    PRINTSETUPDLG.rtMargin.Top = MarginTop * 1000
    PRINTSETUPDLG.rtMargin.Bottom = MarginBottom * 1000
    PRINTSETUPDLG.rtMargin.Right = MarginRight * 1000
    PRINTSETUPDLG.ptPaperSize.X = PageWidth * 1000
    PRINTSETUPDLG.ptPaperSize.Y = PageHeight * 1000
    
    'instruct page setup dialog to use the values in the structure
    PRINTSETUPDLG.flags = PSD_MARGINS Or IIf(ScaleMode = smCentimeters, PSD_INHUNDREDTHSOFMILLIMETERS, PSD_INTHOUSANDTHSOFINCHES) Or PSD_DISABLEPRINTER
    
    ' Allocate memory for the initial hDevName structure
    ' and copy the settings gathered above into this memory
    PRINTSETUPDLG.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PRINTSETUPDLG.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If
    
    ' Call the print dialog up and let the user make changes
    If PAGESETUPDLG(PRINTSETUPDLG) Then

        ' First get the DevName structure.
        lpDevName = GlobalLock(PRINTSETUPDLG.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PRINTSETUPDLG.hDevNames

        ' Next get the DevMode structure and set the printer
        ' properties appropriately
        lpDevMode = GlobalLock(PRINTSETUPDLG.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PRINTSETUPDLG.hDevMode)
        GlobalFree PRINTSETUPDLG.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
               If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
               End If
            Next
        End If
        On Error Resume Next

        ' Set global page margin values
        MarginLeft = PRINTSETUPDLG.rtMargin.Left / 1000
        MarginTop = PRINTSETUPDLG.rtMargin.Top / 1000
        MarginBottom = PRINTSETUPDLG.rtMargin.Bottom / 1000
        MarginRight = PRINTSETUPDLG.rtMargin.Right / 1000
        PageWidth = PRINTSETUPDLG.ptPaperSize.X / 1000
        PageHeight = PRINTSETUPDLG.ptPaperSize.Y / 1000
    
        If PageWidth > PageHeight Then Orientation = PageLandscape Else Orientation = PagePortrait
        ShowPrinterSetup = True
       
        ' Set printer object properties according to selections made by user
        DoEvents
        With Printer
            .Copies = DevMode.dmCopies
            .Duplex = DevMode.dmDuplex
            .Orientation = DevMode.dmOrientation
        End With
        On Error GoTo 0
    Else
       ShowPrinterSetup = False
    End If

End Function

Private Function ShowPrinter(Optional PrintFlags As Long) As Boolean

Dim PrintDlg As PRINTDLG_TYPE
Dim DevMode As DEVMODE_TYPE
Dim DevName As DEVNAMES_TYPE
Dim lpDevMode As Long, lpDevName As Long
Dim bReturn As Integer
Dim objPrinter As Printer, NewPrinterName As String
Dim strSetting As String
    
    ' Use PrintSetupDialog to get the handle to a memory
    ' block with a DevMode and DevName structures
    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hWndOwner = UserControl.hWnd  'frmOwner.hWnd
    PrintDlg.flags = PrintFlags + PD_HIDEPRINTTOFILE + PD_PAGENUMS
    PrintDlg.nFromPage = 1
    PrintDlg.nToPage = PageCount
    PrintDlg.nMinPage = 1
    PrintDlg.nMaxPage = PageCount
     
    ' Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX Or DM_COPIES
    DevMode.dmOrientation = Orientation
    DevMode.dmCopies = 1
    DevMode.dmPaperSize = m_PaperSize
    
    On Error Resume Next
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0
    
    ' Allocate memory for the initialization hDevMode structure
    ' and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If
    
    ' Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    
    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    
    ' Allocate memory for the initial hDevName structure
    ' and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If
    
    ' Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) Then
        
        ' First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames
    
        ' Next get the DevMode structure and set the printer
        ' properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                End If
            Next
        End If
        FromPage = PrintDlg.nFromPage
        ToPage = PrintDlg.nToPage
        On Error Resume Next
        DoEvents
        'Set the printer properties modified by the user
            With Printer
                .ColorMode = DevMode.dmColor
                .Copies = DevMode.dmCopies
                If .PaperBin <> DevMode.dmDefaultSource Then .PaperBin = DevMode.dmDefaultSource
                .Duplex = DevMode.dmDuplex
                Orientation = DevMode.dmOrientation
                .PaperSize = DevMode.dmPaperSize
                PaperSize = DevMode.dmPaperSize
                .PrintQuality = DevMode.dmPrintQuality
                .Zoom = DevMode.dmScale
                'frompage=devmode.f
            End With
        
        On Error GoTo 0
        ShowPrinter = True
    Else
        ShowPrinter = False
    End If

End Function

Public Function SetPrinterOrientation(ByVal eOrientation As PageOrientationConstants) As Boolean

    Dim bDevMode()      As Byte
    Dim bPrinterInfo2() As Byte
    Dim hPrinter        As Long
    Dim lResult         As Long
    Dim nSize           As Long
    Dim sPrnName        As String
    Dim dm              As DEVMODE_TYPE
    Dim pd              As PRINTER_DEFAULTS
    Dim pi2             As PRINTER_INFO_2

    ' Get device name of default printer
    sPrnName = Printer.DeviceName
    ' PRINTER_ALL_ACCESS required under NT, because we're going to call SetPrinter
    pd.DesiredAccess = PRINTER_ALL_ACCESS

    ' Get a handle to the printer.
    If OpenPrinter(sPrnName, hPrinter, pd) Then
        ' Get number of bytes requires for PRINTER_INFO_2 structure
        Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
        ' Create a buffer of the required size
        ReDim bPrinterInfo2(1 To nSize) As Byte
        ' Fill buffer with structure
        lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
        ' Copy fixed portion of structure into VB Type variable
        Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))

        ' Get number of bytes requires for DEVMODE structure
        nSize = DocumentProperties(0&, hPrinter, sPrnName, 0&, 0&, 0)
        ' Create a buffer of the required size
        ReDim bDevMode(1 To nSize)

        ' If PRINTER_INFO_2 points to a DEVMODE structure, copy it into our buffer
        If pi2.pDevMode Then
           Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
        Else
           ' Otherwise, call DocumentProperties to get a DEVMODE structure
           Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), 0&, DM_OUT_BUFFER)
        End If

        ' Copy fixed portion of structure into VB Type variable
        Call CopyMemory(dm, bDevMode(1), Len(dm))
        With dm
            ' Set new orientation
            .dmOrientation = eOrientation
            .dmFields = DM_ORIENTATION
        End With
        ' Copy our Type back into buffer
        Call CopyMemory(bDevMode(1), dm, Len(dm))
        ' Set new orientation
        Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), bDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)

        ' Point PRINTER_INFO_2 at our modified DEVMODE
        pi2.pDevMode = VarPtr(bDevMode(1))
        ' Set new orientation system-wide
        lResult = SetPrinter(hPrinter, 2, pi2, 0&)

        ' Clean up and exit
        Call ClosePrinter(hPrinter)
        SetPrinterOrientation = True
    Else
        SetPrinterOrientation = False
    End If
End Function

Public Function GetPrinterOrientation(DeviceName As String, hdc As Long) As PageOrientationConstants
    
    Dim hPrinter    As Long
    Dim nSize       As Long
    Dim pDevMode    As DEVMODE_TYPE
    Dim aDevMode()  As Byte
   
    If OpenPrinter(DeviceName, hPrinter, NULLPTR) Then
       nSize = DocumentProperties(NULLPTR, hPrinter, DeviceName, NULLPTR, NULLPTR, 0)
       ReDim aDevMode(1 To nSize)
       nSize = DocumentProperties(NULLPTR, hPrinter, DeviceName, aDevMode(1), NULLPTR, DM_OUT_BUFFER)
       Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
       GetPrinterOrientation = pDevMode.dmOrientation
       Call ClosePrinter(hPrinter)
    Else
       GetPrinterOrientation = PagePortrait
    End If
End Function


'Max Pages
Public Function PageCount() As Integer
       PageCount = MaxPageNumber + 1
End Function

'Number Current Page
Public Function CurrentPage() As Integer
    If m_SendToPrinter Then
       CurrentPage = Printer.Page
    Else
      If FileExists(TempDir & "PPView" & CStr(ViewPage) & ".bmp") = True Then
         CurrentPage = ViewPage + 1
      Else
         CurrentPage = 0
      End If
    End If
End Function

'Draw PageBorder
Private Sub DrawBorder()
      Dim x1 As Single, x2 As Single, y1 As Single, y2 As Single
      
      If PageBorder = pbNone Then Exit Sub
      
      oCy = CurrentY
      oCX = CurrentX
      
      Select Case PageBorder
      Case pbNone '= 0
      
      Case pbBottom '= 1
           x1 = MarginLeft
           x2 = PageWidth - MarginLeft
           y1 = PageHeight - MarginBottom
           y2 = y1
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
      Case pbTop '= 2
           x1 = MarginLeft
           x2 = PageWidth - MarginLeft
           y1 = MarginTop
           y2 = y1
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
      Case pbTopBottom '= 3
           x1 = MarginLeft
           x2 = PageWidth - MarginLeft
           y1 = PageHeight - MarginBottom
           y2 = y1
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
           x1 = MarginLeft
           x2 = PageWidth - MarginLeft
           y1 = MarginTop
           y2 = y1
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
      Case pbBox ' = 4
           x1 = MarginLeft
           x2 = PageWidth - MarginLeft
           y1 = PageHeight - MarginBottom
           y2 = y1
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
           x1 = MarginLeft
           x2 = PageWidth - MarginLeft
           y1 = MarginTop
           y2 = y1
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
           x1 = MarginLeft
           x2 = x1
           y1 = MarginTop
           y2 = PageHeight - MarginBottom
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
           x1 = PageWidth - MarginLeft
           x2 = x1
           y1 = MarginTop
           y2 = PageHeight - MarginBottom
           DrawLine x1, y1, x2, y2, m_PageBorderWidth, m_PageBorderColor
           
      End Select
      CurrentY = oCy
      CurrentX = oCX
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub StartTable()
        pTable.tbCols = 0
        pTable.tbRows = 0
        pTable.tbIndent = 0
        ReDim pTable.tbColWidth(pTable.tbCols)
        ReDim pTable.tbColAlign(pTable.tbCols)
        ReDim pTable.tbRowHeight(pTable.tbRows)
        ReDim pTable.tbHeader(pTable.tbCols)
        Erase pTable.tbTableCell
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub EndTable()

      Dim Col As Integer, Row As Integer
      Dim newLeft As Single, newTop As Single, newRight As Single, newBottom As Single
      Dim id As Integer, OldTop As Single, mTop As Single, mWidthTable As Single, mLeft As Single
      Dim percentage As Single, RowTableHeight As Single, cSP As Integer, mRowStart As Long, mRowEnd As Long
      Dim TextLine() As String, tRowTableHeight As Single
      
      Dim OldSet1 As Variant, OldSet2 As Variant, OldSet3 As Variant, OldSet4 As Variant, OldSet5 As Variant
      Dim OldSet6 As Variant, OldSet7 As Variant, OldSet8 As Variant, OldSet9 As Variant, OldSet10 As Variant
      Dim OldSet11 As Variant, OldSet12 As Variant, OldSet13 As Variant, OldSet14 As Variant, oDW As Integer
      Dim oIL As Variant, oIR As Variant
      
      oDW = DrawWidth
      DrawWidth = pTable.tbLineWidth
      
      'Measuring Width Table
      mWidthTable = 0
      For id = 0 To pTable.tbCols
          mWidthTable = mWidthTable + pTable.tbColWidth(id)
      Next
            
      'Position Table
      Select Case m_TextAlign
      Case 0 'Left
           'mLeft = IIf(CurrentX > MarginLeft, CurrentX, MarginLeft)
           mLeft = MarginLeft
      Case 1 'Right
           mLeft = (PageWidth - MarginRight - mWidthTable)
      Case 2 'Center
           mLeft = (PageWidth - MarginRight - mWidthTable) / 2
      Case 3 'Justify
          percentage = (PageWidth - MarginRight - MarginLeft) / mWidthTable
          For id = 0 To pTable.tbCols
              pTable.tbColWidth(id) = pTable.tbColWidth(id) * percentage
          Next
          mLeft = MarginLeft
      End Select
      
      'Read start position table
      SetUnitTable.Left = mLeft
      SetUnitTable.Top = CurrentY
      'Read position Right
      newRight = mWidthTable + mLeft
      
      oIL = m_IndentLeft
      oIR = m_IndentRight
      
      'Print TABLE
      If pTable.tbRows > 0 Then
         mTop = CurrentY
         OldTop = CurrentY
            
         'Make Row Height
         If pTable.tbWordWrap Then
            For Row = 0 To pTable.tbRows
                tRowTableHeight = 0
                For Col = 0 To pTable.tbCols
                   If pTable.tbTableCell(Row, Col).tbcText <> "" Then
                    If pTable.tbTableCell(Row, Col).tbcFontSize <> -1 Then
                       FontSize = pTable.tbTableCell(Row, Col).tbcFontSize
                    End If
                    TextLine = Split(BreakItemText(pTable.tbTableCell(Row, Col).tbcText, pTable.tbColWidth(Col) - pTable.tbIndent * 2), vbCr)
                    tRowTableHeight = FindRowTableHeight(Row, pTable.tbTableCell(Row, Col).tbcFontName, TextLine(0)) * (UBound(TextLine) + 1) + (ChangeLineSpace * 2)
                    If pTable.tbRowHeight(Row) < tRowTableHeight Then
                       pTable.tbRowHeight(Row) = tRowTableHeight
                    End If
                   End If
                Next
            Next
         Else
            For Row = 0 To pTable.tbRows
               For Col = 0 To pTable.tbCols
                If pTable.tbTableCell(Row, Col).tbcText = "" Then pTable.tbTableCell(Row, Col).tbcText = " "
                TextLine = Split(BreakItemText(pTable.tbTableCell(Row, Col).tbcText, pTable.tbColWidth(Col) - pTable.tbIndent * 2, "..."), vbCr)
                pTable.tbTableCell(Row, Col).tbcText = TextLine(0)
                tRowTableHeight = FindRowTableHeight(Row, pTable.tbTableCell(Row, Col).tbcFontName)
                If tRowTableHeight > pTable.tbRowHeight(Row) Then
                pTable.tbRowHeight(Row) = tRowTableHeight 'FindRowTableHeight(Row, pTable.tbTableCell(Row, Col).tbcFontName)
                End If
               Next
            Next
         End If
         
         m_IndentLeft = pTable.tbIndent
         m_IndentRight = pTable.tbIndent
         mRowStart = 0
         
         PrintTableHeader mLeft
         
         For Row = 1 To pTable.tbRows
            
            If EndOfPage Or CurrentY + pTable.tbRowHeight(Row) > PageHeight - MarginBottom Then
               RaiseEvent AfterTableEnd
               DrawWidth = oDW
               mRowEnd = Row - 1
               BorderDraw mTop, mRowStart, mRowEnd
               PrintFooter
               NewPage
               DrawBorder
               SetUnitTable.Top = CurrentY
               PrintHeader
               Paragraph 'Null Line
               PrintTableHeader mLeft
               mTop = CurrentY
               mRowStart = Row
               oDW = DrawWidth
               DrawWidth = pTable.tbLineWidth
            End If
            
            RowTableHeight = pTable.tbRowHeight(Row)
            
            OldTop = CurrentY
            
            For Col = 0 To pTable.tbCols
                
                RaiseEvent BeforeTableCell(Row, Col, pTable.tbTableCell(Row, Col).tbcText)
                newLeft = mLeft
                
                For id = 0 To Col
                    newLeft = newLeft + pTable.tbColWidth(id)
                Next
                RowTableHeight = pTable.tbRowHeight(Row)
                newLeft = newLeft - pTable.tbColWidth(id - 1)
                newTop = OldTop
                
                If pTable.tbTableCell(Row, Col).tbcBackColor <> -1 And pTable.tbTableCell(Row, Col).tbcBackColor <> 0 Then
                    OldSet1 = FillColor
                    FillColor = pTable.tbTableCell(Row, Col).tbcBackColor
                Else
                    OldSet1 = FillColor
                    FillColor = vbWhite
                End If
                If pTable.tbTableCell(Row, Col).tbcForeColor <> -1 And pTable.tbTableCell(Row, Col).tbcForeColor <> 0 Then
                    OldSet2 = ForeColor
                    ForeColor = pTable.tbTableCell(Row, Col).tbcForeColor
                End If
                If pTable.tbTableCell(Row, Col).tbcFontSize > 0 Then
                    OldSet4 = FontSize
                    FontSize = pTable.tbTableCell(Row, Col).tbcFontSize
                End If
                If pTable.tbTableCell(Row, Col).tbcFontCharSet > 0 Then
                    OldSet5 = FontCharSet
                    FontCharSet = pTable.tbTableCell(Row, Col).tbcFontCharSet
                End If
                If pTable.tbTableCell(Row, Col).tbcFontName <> "" Then
                    OldSet3 = FontName
                    FontName = pTable.tbTableCell(Row, Col).tbcFontName
                End If
                OldSet6 = FontBold
                FontBold = pTable.tbTableCell(Row, Col).tbcFontBold
             
                OldSet7 = FontItalic
                FontItalic = pTable.tbTableCell(Row, Col).tbcFontItalic
             
                OldSet8 = FontUnderline
                FontUnderline = pTable.tbTableCell(Row, Col).tbcFontUnderline
            
                OldSet9 = FontStrikethru
                FontStrikethru = pTable.tbTableCell(Row, Col).tbcFontStrikethru
             
                OldSet10 = FontTransparent
                FontTransparent = pTable.tbTableCell(Row, Col).tbcFontTransparent
                'tcPicture '= 11
                'tcColSpan '= 12
                OldSet12 = pTable.tbColWidth(Col)
                If pTable.tbTableCell(Row, Col).tbcColSpan > 1 Then
                   For cSP = Col + 1 To Col + pTable.tbTableCell(Row, Col).tbcColSpan - 1
                       pTable.tbColWidth(Col) = pTable.tbColWidth(Col) + pTable.tbColWidth(cSP)
                   Next
                End If
                'tcRowSpan '= 13
                OldSet13 = RowTableHeight
                If pTable.tbTableCell(Row, Col).tbcRowSpan > 1 Then
                   pTable.tbRowHeight(Row) = RowTableHeight
                   For cSP = Row + 1 To Row + pTable.tbTableCell(Row, Col).tbcRowSpan - 1
                       RowTableHeight = RowTableHeight + pTable.tbRowHeight(cSP)
                   Next
                End If
                
                OldSet14 = TextAlign
                TextAlign = pTable.tbTableCell(Row, Col).tbcColAlign
                
                If pTable.tbTableCell(Row, Col).tbcPicture Is Nothing Then
                   If pTable.tbTableCell(Row, Col).tbcRowSpan > 0 Then
                    If OldSet1 <> FillColor Then
                       If TableBorder = tbAll Then
                         DrawRectangle newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight, , , pTable.tbLineColor, FillColor, vbFSSolid
                       Else
                         DrawRectangle newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight, , , FillColor, FillColor, vbFSSolid
                       End If
                    End If
                    TextBox pTable.tbTableCell(Row, Col).tbcText, _
                            newLeft, newTop, _
                            pTable.tbColWidth(Col), _
                            RowTableHeight, _
                            IIf(pTable.tbColAlign(Col) <> pTable.tbTableCell(Row, Col).tbcColAlign, pTable.tbTableCell(Row, Col).tbcColAlign, pTable.tbColAlign(Col)), _
                            False
                   End If
                Else
                    DrawPicture pTable.tbTableCell(Row, Col).tbcPicture, newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight
                    DrawRectangle newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight, , , pTable.tbLineColor
                End If
                  
                If IsEmpty(OldSet1) = False Then FillColor = OldSet1
                If IsEmpty(OldSet2) = False Then ForeColor = OldSet2
                If OldSet3 <> "" Then FontName = OldSet3
                If IsEmpty(OldSet4) = False Then FontSize = OldSet4
                If IsEmpty(OldSet5) = False Then FontCharSet = OldSet5
                FontBold = OldSet6
                FontItalic = OldSet7
                FontUnderline = OldSet8
                FontStrikethru = OldSet9
                FontTransparent = OldSet10
                ' tcPicture '= 11
                ' tcColSpan '= 12
                If pTable.tbTableCell(Row, Col).tbcColSpan > 1 Then
                   pTable.tbColWidth(Col) = OldSet12
                   Col = Col + pTable.tbTableCell(Row, Col).tbcColSpan - 1
                End If
                ' tcRowSpan '= 13
                RowTableHeight = OldSet13
                TextAlign = OldSet14
                RaiseEvent AfterTableCell(Row, Col, newLeft, newTop, newLeft + pTable.tbColWidth(Col), newTop + RowTableHeight, pTable.tbTableCell(Row, Col).tbcText)
           Next Col
        Next Row
        RaiseEvent AfterTableEnd
        
        mRowEnd = Row - 1
        BorderDraw mTop, mRowStart, mRowEnd

        m_IndentLeft = oIL
        m_IndentRight = oIR
        DrawWidth = oDW
              
        'set position
        SetUnitTable.Right = newRight
        SetUnitTable.Bottom = CurrentY
      End If
      
End Sub

Private Sub PrintTableHeader(mLeft As Single)

      Dim Col As Integer, Row As Integer
      Dim newLeft As Single, newTop As Single, newRight As Single, newBottom As Single
      Dim id As Integer, OldTop As Single, mTop As Single, mWidthTable As Single, RowTableHeight As Single, cSP As Integer
      Dim TextLine() As String, tRowTableHeight As Single, oDW As Integer
      
      Dim percentage As Single
      Dim OldSet1 As Variant
      Dim OldSet2 As Variant
      Dim OldSet3 As Variant
      Dim OldSet4 As Variant
      Dim OldSet5 As Variant
      Dim OldSet6 As Variant
      Dim OldSet7 As Variant
      Dim OldSet8 As Variant
      Dim OldSet9 As Variant
      Dim OldSet10 As Variant
      Dim OldSet11 As Variant
      Dim OldSet12 As Variant
      Dim OldSet13 As Variant
      Dim OldSet14 As Variant
      
      oDW = DrawWidth
      DrawWidth = pTable.tbLineWidth
      
         'Make Row Height
         If pTable.tbWordWrap Then
            For Row = 0 To pTable.tbRows
                tRowTableHeight = 0
                For Col = 0 To pTable.tbCols
                   If pTable.tbTableCell(Row, Col).tbcText <> "" Then
                    If pTable.tbTableCell(Row, Col).tbcFontSize <> -1 Then
                       FontSize = pTable.tbTableCell(Row, Col).tbcFontSize
                    End If
                    TextLine = Split(BreakItemText(pTable.tbTableCell(Row, Col).tbcText, pTable.tbColWidth(Col) - pTable.tbIndent * 2), vbCr)
                    tRowTableHeight = FindRowTableHeight(Row, pTable.tbTableCell(Row, Col).tbcFontName, TextLine(0)) * (UBound(TextLine) + 1) + ChangeLineSpace * 2
                    If pTable.tbRowHeight(Row) < tRowTableHeight Then
                       pTable.tbRowHeight(Row) = tRowTableHeight
                    End If
                  End If
                Next
            Next
         Else
            'For Row = 0 To pTable.tbRows
               Row = 0
               For Col = 0 To pTable.tbCols
                If pTable.tbTableCell(Row, Col).tbcText = "" Then pTable.tbTableCell(Row, Col).tbcText = " "
                TextLine = Split(BreakItemText(pTable.tbTableCell(Row, Col).tbcText, pTable.tbColWidth(Col) - pTable.tbIndent * 2, "..."), vbCr)
                pTable.tbTableCell(Row, Col).tbcText = TextLine(0)
                'pTable.tbRowHeight(Row) = FindRowTableHeight(Row, pTable.tbTableCell(Row, Col).tbcFontName)
                tRowTableHeight = FindRowTableHeight(Row, pTable.tbTableCell(Row, Col).tbcFontName)
                If tRowTableHeight > pTable.tbRowHeight(Row) Then
                    pTable.tbRowHeight(Row) = tRowTableHeight 'FindRowTableHeight(Row, pTable.tbTableCell(Row, Col).tbcFontName)
                End If
               Next
            'Next
         End If
         
          OldTop = CurrentY
          Row = 0
            For Col = 0 To pTable.tbCols
                newLeft = mLeft
                For id = 0 To Col
                    newLeft = newLeft + pTable.tbColWidth(id)
                Next
                RowTableHeight = pTable.tbRowHeight(Row)
                newLeft = newLeft - pTable.tbColWidth(id - 1)
                newLeft = newLeft
                newTop = OldTop
                
             If pTable.tbTableCell(Row, Col).tbcBackColor <> -1 And pTable.tbTableCell(Row, Col).tbcBackColor <> 0 Then
                OldSet1 = FillColor
                FillColor = pTable.tbTableCell(Row, Col).tbcBackColor
             Else
                OldSet1 = FillColor
                FillColor = vbWhite
             End If
             If pTable.tbTableCell(Row, Col).tbcForeColor <> -1 And pTable.tbTableCell(Row, Col).tbcForeColor <> 0 Then
                 OldSet2 = ForeColor
                 ForeColor = pTable.tbTableCell(Row, Col).tbcForeColor
             End If
             If pTable.tbTableCell(Row, Col).tbcFontName <> "" Then
                 OldSet3 = FontName
                 FontName = pTable.tbTableCell(Row, Col).tbcFontName
             End If
             If pTable.tbTableCell(Row, Col).tbcFontSize > 0 Then
                 OldSet4 = FontSize
                 FontSize = pTable.tbTableCell(Row, Col).tbcFontSize
             End If
             If pTable.tbTableCell(Row, Col).tbcFontCharSet > 0 Then
                 OldSet5 = FontCharSet
                 FontCharSet = pTable.tbTableCell(Row, Col).tbcFontCharSet
             End If
             OldSet6 = FontBold
             FontBold = pTable.tbTableCell(Row, Col).tbcFontBold
             
             OldSet7 = FontItalic
             FontItalic = pTable.tbTableCell(Row, Col).tbcFontItalic
             
             OldSet8 = FontUnderline
             FontUnderline = pTable.tbTableCell(Row, Col).tbcFontUnderline
            
             OldSet9 = FontStrikethru
             FontStrikethru = pTable.tbTableCell(Row, Col).tbcFontStrikethru
             
             OldSet10 = FontTransparent
             FontTransparent = pTable.tbTableCell(Row, Col).tbcFontTransparent
            'tcPicture '= 11
            'tcColSpan '= 12
             OldSet11 = pTable.tbColWidth(Col)
             If pTable.tbTableCell(Row, Col).tbcColSpan > 1 Then
                For cSP = Col + 1 To Col + pTable.tbTableCell(Row, Col).tbcColSpan - 1
                    pTable.tbColWidth(Col) = pTable.tbColWidth(Col) + pTable.tbColWidth(cSP)
                Next
             End If
            'tcRowSpan '= 13
             OldSet13 = RowTableHeight
             If pTable.tbTableCell(Row, Col).tbcRowSpan > 1 Then
                   pTable.tbRowHeight(Row) = RowTableHeight
                   For cSP = Row + 1 To Row + pTable.tbTableCell(Row, Col).tbcRowSpan - 1
                       RowTableHeight = RowTableHeight + pTable.tbRowHeight(cSP)
                   Next
             End If
             
            OldSet14 = TextAlign
            TextAlign = pTable.tbTableCell(Row, Col).tbcColAlign
                            
            If pTable.tbTableCell(Row, Col).tbcPicture Is Nothing Then
                If OldSet1 <> FillColor Then
                   If TableBorder = tbAll Then
                      DrawRectangle newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight, , , pTable.tbLineColor, FillColor, vbFSSolid
                   Else
                      DrawRectangle newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight, , , FillColor, FillColor, vbFSSolid
                   End If
               End If
                TextBox pTable.tbTableCell(Row, Col).tbcText, _
                             newLeft, newTop, _
                             pTable.tbColWidth(Col), _
                             RowTableHeight, _
                             IIf(pTable.tbColAlign(Col) <> pTable.tbTableCell(0, Col).tbcColAlign, pTable.tbTableCell(0, Col).tbcColAlign, pTable.tbColAlign(Col)), _
                             False
             Else
                DrawPicture pTable.tbTableCell(Row, Col).tbcPicture, newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight
                DrawRectangle newLeft, newTop, pTable.tbColWidth(Col), RowTableHeight, , , pTable.tbLineColor
             End If
                  
              If IsEmpty(OldSet1) = False Then FillColor = OldSet1
              If IsEmpty(OldSet2) = False Then ForeColor = OldSet2
              If OldSet3 <> "" Then FontName = OldSet3
              If IsEmpty(OldSet4) = False Then FontSize = OldSet4
              If IsEmpty(OldSet5) = False Then FontCharSet = OldSet5
              FontBold = OldSet6
              FontItalic = OldSet7
              FontUnderline = OldSet8
              FontStrikethru = OldSet9
              FontTransparent = OldSet10
              
            ' tcPicture '= 11
            ' tcColSpan '= 12
              If pTable.tbTableCell(Row, Col).tbcColSpan > 1 Then
                 pTable.tbColWidth(Col) = OldSet11
                 Col = Col + pTable.tbTableCell(Row, Col).tbcColSpan - 1
               End If
            ' tcRowSpan '= 13
            ' tcTextAling '= 14
              TextAlign = OldSet14
           Next
      DrawWidth = oDW
End Sub

Private Sub BorderDraw(ByVal mTop As Single, ByVal RowStart As Long, ByVal RowEnd As Long)
        
        Dim newLeft As Single, newTop As Single, newRight As Single, newBottom As Single
        Dim id As Integer, OldTop As Single, mTRh As Single
        
        OldTop = CurrentY
        oDW = DrawWidth
        DrawWidth = pTable.tbLineWidth
             
        Select Case TableBorder
        Case tbNone   '= 0
        
        Case tbBottom '= 1
            'header down
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop + pTable.tbRowHeight(RowStart)
             Else
                newTop = mTop
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             
             'tablebottom
             newLeft = MarginLeft
             newTop = mTop
             For id = RowStart To RowEnd
                 newTop = newTop + pTable.tbRowHeight(id)
             Next
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
                
        Case tbTop '= 2
             'header up
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop
             Else
                newTop = mTop - pTable.tbRowHeight(0)
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             'header down
             If RowStart = 0 Then
                newTop = mTop + pTable.tbRowHeight(0)
             Else
                newTop = mTop
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             
        Case tbTopBottom '= 3
             'header up
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop
             Else
                newTop = mTop - pTable.tbRowHeight(0)
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             'header down
             If RowStart = 0 Then
                newTop = mTop + pTable.tbRowHeight(0)
             Else
                newTop = mTop
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             
            'tablebottom
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols '- 1
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             newTop = OldTop
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             
        Case tbBox '= 4
             'header up
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop
             Else
                newTop = mTop - pTable.tbRowHeight(0)
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             'header down
             If RowStart = 0 Then
                newTop = mTop + pTable.tbRowHeight(0)
             Else
                newTop = mTop
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
            
             'Box
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols '- 1
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop
             Else
                newTop = mTop - pTable.tbRowHeight(0)
             End If
            
             newBottom = OldTop
             DrawRectangle newLeft, newTop, newRight - newLeft, newBottom - newTop, , , pTable.tbLineColor
             
        Case tbColums '= 5
             'Colums
             If RowStart = 0 Then
                newTop = mTop
             Else
                newTop = mTop - pTable.tbRowHeight(0)
             End If
            
             newBottom = OldTop
             newLeft = MarginLeft
             For id = 0 To pTable.tbCols '- 1
                 DrawLine newLeft, newTop, newLeft, newBottom, pTable.tbLineWidth, pTable.tbLineColor
                 newLeft = newLeft + pTable.tbColWidth(id)
             Next
             DrawLine newLeft, newTop, newLeft, newBottom, pTable.tbLineWidth, pTable.tbLineColor
             
        Case tbColTopBottom '= 6
        
              'header up
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop
             Else
                newTop = mTop - pTable.tbRowHeight(0)
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             'header down
             If RowStart = 0 Then
                newTop = mTop + pTable.tbRowHeight(0)
             Else
                newTop = mTop
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
              
              'tablebottom
              newLeft = MarginLeft
              newRight = 0
              For id = 0 To pTable.tbCols '- 1
                  newRight = newRight + pTable.tbColWidth(id)
              Next
              newRight = newRight + newLeft
              newTop = OldTop
              DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
              
              'Colums
              newTop = mTop
              newBottom = OldTop
              newLeft = MarginLeft
              For id = 0 To pTable.tbCols - 1
                  newLeft = newLeft + pTable.tbColWidth(id)
                  DrawLine newLeft, newTop, newLeft, newBottom, pTable.tbLineWidth, pTable.tbLineColor
               Next
               
        Case tbAll '= 7
            
        Case tbBoxRows '= 8
             'rows
              newLeft = MarginLeft
              newRight = 0
              For id = 0 To pTable.tbCols
                  newRight = newRight + pTable.tbColWidth(id)
              Next
              'newRight = newRight
              newTop = mTop
              If RowStart > 0 Then 'HEADER
                DrawRectangle newLeft, newTop - pTable.tbRowHeight(0), newRight, pTable.tbRowHeight(0), , , pTable.tbLineColor
              End If
              
              For id = RowStart To RowEnd
                  DrawRectangle newLeft, newTop, newRight, pTable.tbRowHeight(id), , , pTable.tbLineColor
                  newTop = newTop + pTable.tbRowHeight(id)
              Next
            
        Case tbBoxColumns ' = 9
                'header up
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop
             Else
                newTop = mTop - pTable.tbRowHeight(0)
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
             'header down
             If RowStart = 0 Then
                newTop = mTop + pTable.tbRowHeight(0)
             Else
                newTop = mTop
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
                          
              'Bottom
              newLeft = MarginLeft
              newRight = 0
              For id = 0 To pTable.tbCols
                  newRight = newRight + pTable.tbColWidth(id)
              Next
              newRight = newRight + newLeft
              newTop = OldTop
              DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
    
              'Colums
              If RowStart = 0 Then
                newTop = mTop
              Else
                newTop = mTop - pTable.tbRowHeight(0)
              End If
              newBottom = OldTop
              newLeft = MarginLeft
              For id = 0 To pTable.tbCols
                  DrawLine newLeft, newTop, newLeft, newBottom, pTable.tbLineWidth, pTable.tbLineColor
                  newLeft = newLeft + pTable.tbColWidth(id)
              Next
              DrawLine newLeft, newTop, newLeft, newBottom, pTable.tbLineWidth, pTable.tbLineColor
              
      Case tbBelowHeader '= 10
            'header down
             newLeft = MarginLeft
             newRight = 0
             For id = 0 To pTable.tbCols
                 newRight = newRight + pTable.tbColWidth(id)
             Next
             newRight = newRight + newLeft
             If RowStart = 0 Then
                newTop = mTop + pTable.tbRowHeight(RowStart)
             Else
                newTop = mTop
             End If
             DrawLine newLeft, newTop, newRight, newTop, pTable.tbLineWidth, pTable.tbLineColor
       End Select
       CurrentY = OldTop
       DrawWidth = oDW
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
'Read Row Table Height
Private Function FindRowTableHeight(Row As Integer, FontName As String, Optional txt As String = "W") As Single
         
         Dim tFSize As Integer, Cl As Integer, oFS As Integer, oFN As String
         
         picPrintPic.Cls
         picPrintPic.ScaleMode = ScaleMode
         For Cl = 0 To pTable.tbCols
             If tFSize < pTable.tbTableCell(Row, Cl).tbcFontSize Then
                tFSize = pTable.tbTableCell(Row, Cl).tbcFontSize
             End If
         Next
         oFS = ObjPrint.FontSize
         oFN = ObjPrint.FontName
         If FontName <> "" Then ObjPrint.FontName = FontName
        If tFSize > 0 Then
            ObjPrint.FontSize = tFSize
        Else
            ObjPrint.FontSize = FontSize
        End If
        ObjPrint.Print "";
        FindRowTableHeight = ObjPrint.TextHeight(txt)
        'End If
       ObjPrint.FontSize = oFS
       ObjPrint.FontName = oFN
       ObjPrint.Print "";
       
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub Table(ByVal FormatCol As String, ByVal Header As String, ByVal Body As String, _
                 Optional HeaderShade As Long = -1, Optional BodyShade As Long = -1, _
                 Optional LineColor As Long = 0, Optional LineWidth As Integer = 1, _
                 Optional Wrap As Boolean = True, Optional Indent As Variant = "50tw")
                      
     Dim tArr() As String, tCAling As TextAlignConstants
     Dim tCols() As String, tHdr() As String, mBody() As Variant
     Dim tRows() As String, C As Integer, R As Integer, Numr As Variant
     Dim i As Long
     On Error GoTo err1
     Body = Replace(Body, "|;", ";")
     Body = Replace(Body, ";|", ";")
     If Body = "" Then
        Numr = Split(FormatCol, "|")
        Body = String(UBound(Numr), "|") + ";"
     End If
     FormatCol = Replace(FormatCol, ";", "")
     tRows = Split(Body, ";")
     tCols = Split(Header, "|")
     
     ReDim mBody(UBound(tCols), UBound(tRows) - 1)
     
     For R = 0 To UBound(tRows) - 1
         tCols = Split(tRows(R), "|")
         For C = 0 To UBound(tCols)
            mBody(C, R) = tCols(C)
         Next
     Next
     
     TableArray FormatCol, Header, mBody, HeaderShade, BodyShade, LineColor, LineWidth, Wrap, Indent
     
     Exit Sub
err1:
    Stop
End Sub
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
'Write-Read Property Table
Public Function TableCell(ByVal Settings As TableSettingConstants, _
                          Optional ByVal Row As Variant, Optional ByVal Col As Variant, _
                          Optional ByVal Value As Variant = Null) As Variant
                          
        Dim C As Integer, R As Integer
        
        'Write Property
        If Not IsNull(Value) Then
            If Settings = tcText Then
               If IsMissing(Col) Or IsMissing(Row) Then Exit Function
               pTable.tbTableCell(Row - 1, Col - 1).tbcText = Value
            
            ElseIf Settings = tcCols Then
               pTable.tbCols = Value - 1
               If pTable.tbRows > 0 And pTable.tbCols > 0 Then
                  ReDim Preserve pTable.tbTableCell(pTable.tbRows, pTable.tbCols)
                  For R = 0 To pTable.tbRows
                     For C = 0 To pTable.tbCols
                        pTable.tbTableCell(R, C).tbcColSpan = 1
                        pTable.tbTableCell(R, C).tbcRowSpan = 1
                     Next
                  Next
                  
               End If
               ReDim Preserve pTable.tbColWidth(pTable.tbCols)
               ReDim Preserve pTable.tbColAlign(pTable.tbCols)
               ReDim Preserve pTable.tbHeader(pTable.tbCols)
               
            ElseIf Settings = tcRows Then
               Dim oRs As Integer
               Dim tTableCell() As nTableCell
               oRs = pTable.tbRows
               If oRs > 0 Then
                    pTable.tbRows = Value '- 1
                    If pTable.tbRows > 0 And pTable.tbCols > 0 Then
                       tTableCell = pTable.tbTableCell
                       ReDim pTable.tbTableCell(0 To pTable.tbRows, 0 To pTable.tbCols)
                       For R = 0 To oRs
                          For C = 0 To pTable.tbCols
                             pTable.tbTableCell(R, C) = tTableCell(R, C)
                          Next
                       Next
                    End If
                    For R = oRs + 1 To pTable.tbRows
                        For C = 0 To pTable.tbCols
                            pTable.tbTableCell(R, C) = pTable.tbTableCell(oRs, C)
                            pTable.tbTableCell(R, C).tbcText = ""
                            If pTable.tbTableCell(R, C).tbcRowSpan = 0 Then pTable.tbTableCell(R, C).tbcRowSpan = 1
                        Next
                    Next
                    Erase tTableCell
                    ReDim Preserve pTable.tbRowHeight(pTable.tbRows)
               Else
                    pTable.tbRows = Value - 1
                    If pTable.tbRows > 0 And pTable.tbCols > 0 Then
                       ReDim pTable.tbTableCell(0 To pTable.tbRows, 0 To pTable.tbCols)
                       For R = 0 To pTable.tbRows - 1
                          For C = 0 To pTable.tbCols
                             If pTable.tbTableCell(R, C).tbcColSpan = 0 Then pTable.tbTableCell(R, C).tbcColSpan = 1
                             If pTable.tbTableCell(R, C).tbcRowSpan = 0 Then pTable.tbTableCell(R, C).tbcRowSpan = 1
                          Next
                       Next
                    End If
                    ReDim Preserve pTable.tbRowHeight(pTable.tbRows)
                    pTable.tbWordWrap = True
               End If
               
            ElseIf Settings = tcColWidth Then
               If IsMissing(Col) Then Exit Function
               Value = ScaleX(Value, ScaleMode, ScaleMode)
               pTable.tbColWidth(Col - 1) = Value
               If Col = 1 Then
                  For C = 0 To pTable.tbCols
                    pTable.tbColWidth(C) = Value
                  Next
               End If
               
            ElseIf Settings = tcIndent Then
                  Value = ScaleX(Value, ScaleMode, ScaleMode)
                  pTable.tbIndent = Value

            ElseIf Settings = tcRowHeight Then
               Value = ScaleX(Value, ScaleMode, ScaleMode)
               If IsMissing(Row) Then
                  For R = 0 To pTable.tbRows
                      pTable.tbRowHeight(R) = Value
                  Next
               Else
                  pTable.tbRowHeight(Row - 1) = Value
                  If Row = 1 Then
                     For R = 0 To pTable.tbRows
                        pTable.tbRowHeight(R) = Value
                     Next
                  End If
               End If
            ElseIf Settings = tcTextAling Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcColAlign = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcColAlign = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcColAlign = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcColAlign = Value
               End If
            
            ElseIf Settings = tcBackColor Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(Row - 1, C - 1).tbcBackColor = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcBackColor = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcBackColor = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcBackColor = Value
               End If
               
            ElseIf Settings = tcColSpan Then
               If Col - 1 + Value > pTable.tbCols Then Value = pTable.tbCols + 1 - (Col) + 1
               pTable.tbTableCell(Row - 1, Col - 1).tbcColSpan = Value
               
            ElseIf Settings = tcFontBold Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontBold = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcFontBold = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontBold = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontBold = Value
               End If
               
            ElseIf Settings = tcFontCharSet Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontCharSet = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcFontCharSet = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontCharSet = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontCharSet = Value
               End If
               
            ElseIf Settings = tcFontItalic Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontItalic = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcFontItalic = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontItalic = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontItalic = Value
               End If
               
            ElseIf Settings = tcFontName Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontName = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcFontName = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontName = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontName = Value
               End If
              
            ElseIf Settings = tcFontSize Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontSize = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(R - 1, C - 1).tbcFontSize = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontSize = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontSize = Value
               End If
               
            ElseIf Settings = tcFontStrikethru Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontStrikethru = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcFontStrikethru = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontStrikethru = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontStrikethru = Value
               End If
               
            ElseIf Settings = tcFontTransparent Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontTransparent = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcFontTransparent = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontTransparent = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontTransparent = Value
               End If
               
            ElseIf Settings = tcFontUnderline Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcFontUnderline = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcFontUnderline = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcFontUnderline = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcFontUnderline = Value
               End If
               
            ElseIf Settings = tcForeColor Then
               If IsMissing(Col) And IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      For C = 1 To pTable.tbCols + 1
                          pTable.tbTableCell(R - 1, C - 1).tbcForeColor = Value
                      Next
                  Next
               ElseIf IsMissing(Col) Then
                  For C = 1 To pTable.tbCols + 1
                      pTable.tbTableCell(Row - 1, C - 1).tbcForeColor = Value
                  Next
               ElseIf IsMissing(Row) Then
                  For R = 1 To pTable.tbRows + 1
                      pTable.tbTableCell(R - 1, Col - 1).tbcForeColor = Value
                  Next
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                  pTable.tbTableCell(Row - 1, Col - 1).tbcForeColor = Value
               End If
               
            ElseIf Settings = tcPicture Then
               If IsMissing(Col) Or IsMissing(Row) Then Exit Function
               Set pTable.tbTableCell(Row - 1, Col - 1).tbcPicture = Value
               
            ElseIf Settings = tcRowSpan Then
               If IsMissing(Col) Or IsMissing(Row) Then Exit Function
               Dim Rw As Long
               pTable.tbTableCell(Row - 1, Col - 1).tbcRowSpan = Value
               For Rw = Row To Row + Value - 2
                   pTable.tbTableCell(Rw, Col - 1).tbcRowSpan = 0
               Next
               
            End If
            
        'Read Property
        Else
            If Settings = tcText Then
               If IsMissing(Col) Or IsMissing(Row) Then Exit Function
               TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcText
               
            ElseIf Settings = tcCols Then
               TableCell = pTable.tbCols + 1
               
            ElseIf Settings = tcRows Then
               TableCell = pTable.tbRows + 1
               
            ElseIf Settings = tcColWidth Then
               If IsMissing(Col) Then Exit Function
               TableCell = pTable.tbColWidth(Col - 1)
               
            ElseIf Settings = tcIndent Then
               TableCell = pTable.tbIndent
                  
            ElseIf Settings = tcRowHeight Then
               If IsMissing(Row) Then Exit Function
               TableCell = pTable.tbRowHeight(Row - 1)
                  
            ElseIf Settings = tcTextAling Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcColAlign
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcColAlign
               End If
               
            ElseIf Settings = tcBackColor Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcBackColor
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcBackColor
               End If
               
            ElseIf Settings = tcColSpan Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcColSpan
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcColSpan
               End If
               
            ElseIf Settings = tcFontBold Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontBold
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontBold
               End If
               
            ElseIf Settings = tcFontCharSet Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontCharSet
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontCharSet
               End If
               
            ElseIf Settings = tcFontItalic Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontItalic
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontItalic
               End If
               
            ElseIf Settings = tcFontName Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontName
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontName
               End If
               
            ElseIf Settings = tcFontSize Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontSize
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontSize
               End If
               
            ElseIf Settings = tcFontStrikethru Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontStrikethru
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontStrikethru
               End If
               
            ElseIf Settings = tcFontTransparent Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontTransparent
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontTransparent
               End If
               
            ElseIf Settings = tcFontUnderline Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcFontUnderline
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcFontUnderline
               End If
               
            ElseIf Settings = tcForeColor Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   TableCell = pTable.tbTableCell(R - 1, C - 1).tbcForeColor
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcForeColor
               End If
               
            ElseIf Settings = tcPicture Then
               If IsMissing(Col) Or IsMissing(Row) Then
                   If IsMissing(Col) Then C = 1 Else C = Col
                   If IsMissing(Row) Then R = 1 Else R = Row
                   Set TableCell = pTable.tbTableCell(R - 1, C - 1).tbcPicture
               ElseIf IsMissing(Col) = False And IsMissing(Row) = False Then
                   Set TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcPicture
               End If
               
            ElseIf Settings = tcRowSpan Then
               If IsMissing(Col) Or IsMissing(Row) Then Exit Function
               TableCell = pTable.tbTableCell(Row - 1, Col - 1).tbcRowSpan
            End If
        End If
        
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub TableArray(ByVal FormatCol As String, ByVal Header As String, Body As Variant, _
                      Optional HeaderShade As Long = -1, Optional BodyShade As Long = -1, _
                      Optional LineColor As Long = 0, Optional LineWidth As Integer = 1, _
                      Optional Wrap As Boolean = True, Optional Indent As Variant = "50tw")

     Dim tArr() As Variant, tCAling As TextAlignConstants
     Dim tCols() As String, nCols() As Single, tHdr() As String
     Dim tRows() As String, C As Integer, R As Integer
     Dim i As Long
                    
     
     tArr = Body
     pTable.tbRows = UBound(tArr, 2) + 1
     pTable.tbCols = UBound(tArr, 1) '+ 1
     
     If LineWidth < 1 Then LineWidth = 1
     pTable.tbLineColor = LineColor
     pTable.tbLineWidth = LineWidth
     
     If Indent <> -1 Then
        Indent = ScaleX(Indent, ScaleMode, ScaleMode)
        pTable.tbIndent = Indent
     End If
     pTable.tbWordWrap = Wrap
     
     ReDim pTable.tbColWidth(pTable.tbCols)
     ReDim pTable.tbHeader(pTable.tbCols)
     ReDim pTable.tbColAlign(pTable.tbCols)
     ReDim pTable.tbTableCell(pTable.tbRows, pTable.tbCols)
     ReDim pTable.tbRowHeight(pTable.tbRows)
     
     FormatCol = Replace(FormatCol, ";", "")
     Header = Replace(Header, ";", "")
     tCols = Split(FormatCol, "|")
     ReDim nCols(UBound(tCols))
      tHdr = Split(Header, "|")

     'Align
     For i = 0 To pTable.tbCols
        If i <= UBound(tCols) Then
         If InStr(1, tCols(i), "<+") > 0 Then
            pTable.tbColAlign(i) = taLeftMiddle
            tCols(i) = Replace(tCols(i), "<+", "")
         ElseIf InStr(1, tCols(i), ">+") > 0 Then
            pTable.tbColAlign(i) = taRightMiddle
            tCols(i) = Replace(tCols(i), ">+", "")
         ElseIf InStr(1, tCols(i), "^+") > 0 Then
            pTable.tbColAlign(i) = taCenterMiddle
            tCols(i) = Replace(tCols(i), "^+", "")
         ElseIf InStr(1, tCols(i), "=+") > 0 Then
            pTable.tbColAlign(i) = taJustifyMiddle
            tCols(i) = Replace(tCols(i), "=+", "")
               
         ElseIf InStr(1, tCols(i), "<_") > 0 Then
            pTable.tbColAlign(i) = taLeftBottom
            tCols(i) = Replace(tCols(i), "<_", "")
         ElseIf InStr(1, tCols(i), ">_") > 0 Then
            pTable.tbColAlign(i) = taRightBottom
            tCols(i) = Replace(tCols(i), ">_", "")
         ElseIf InStr(1, tCols(i), "^_") > 0 Then
            pTable.tbColAlign(i) = taCenterBottom
            tCols(i) = Replace(tCols(i), "^_", "")
         ElseIf InStr(1, tCols(i), "=_") > 0 Then
            pTable.tbColAlign(i) = taJustifyBottom
            tCols(i) = Replace(tCols(i), "=_", "")
            
         ElseIf InStr(1, tCols(i), "<") > 0 Then
            pTable.tbColAlign(i) = taLeftTop
            tCols(i) = Replace(tCols(i), "<", "")
         ElseIf InStr(1, tCols(i), ">") > 0 Then
            pTable.tbColAlign(i) = taRightTop
            tCols(i) = Replace(tCols(i), ">", "")
         ElseIf InStr(1, tCols(i), "^") > 0 Then
            pTable.tbColAlign(i) = taCenterTop
            tCols(i) = Replace(tCols(i), "^", "")
         ElseIf InStr(1, tCols(i), "=") > 0 Then
            pTable.tbColAlign(i) = taJustifyTop
            tCols(i) = Replace(tCols(i), "=", "")
            
         Else
            pTable.tbColAlign(i) = taLeftTop
            tCols(i) = Replace(tCols(i), "<", "")
         End If
         nCols(i) = ScaleX(tCols(i), ScaleMode, ScaleMode)
         pTable.tbColWidth(i) = nCols(i)
        End If
     Next
     
     If HeaderShade = -1 Then HeaderShade = vbWhite
     If BodyShade = -1 Then BodyShade = vbWhite
     
     For C = 0 To pTable.tbCols
       If C <= UBound(tCols) Then
         pTable.tbTableCell(0, C).tbcCol = 0
         pTable.tbTableCell(0, C).tbcRow = C
         pTable.tbTableCell(0, C).tbcText = tHdr(C)
         pTable.tbTableCell(0, C).tbcColAlign = pTable.tbColAlign(C)
         pTable.tbTableCell(0, C).tbcColSpan = 1
         pTable.tbTableCell(0, C).tbcBackColor = HeaderShade
         pTable.tbTableCell(0, C).tbcFontBold = m_FontBold
         pTable.tbTableCell(0, C).tbcFontCharSet = m_FontCharSet
         pTable.tbTableCell(0, C).tbcFontItalic = m_FontItalic
         pTable.tbTableCell(0, C).tbcFontName = m_FontName
         pTable.tbTableCell(0, C).tbcFontSize = m_FontSize
         pTable.tbTableCell(0, C).tbcFontStrikethru = m_FontStrikethru
         pTable.tbTableCell(0, C).tbcFontTransparent = True
         pTable.tbTableCell(0, C).tbcFontUnderline = m_FontUnderline
         pTable.tbTableCell(0, C).tbcForeColor = m_ForeColor
         Set pTable.tbTableCell(0, C).tbcPicture = Nothing
         pTable.tbTableCell(0, C).tbcRowSpan = 1
         pTable.tbHeader(C) = tHdr(C)
       End If
     Next
     
     For R = 0 To UBound(tArr, 2)
         For C = 0 To UBound(tArr, 1)
            pTable.tbTableCell(R + 1, C).tbcCol = R
            pTable.tbTableCell(R + 1, C).tbcRow = C
            pTable.tbTableCell(R + 1, C).tbcText = tArr(C, R)
            pTable.tbTableCell(R + 1, C).tbcColAlign = pTable.tbColAlign(C)
            pTable.tbTableCell(R + 1, C).tbcColSpan = 1
            pTable.tbTableCell(R + 1, C).tbcBackColor = BodyShade
            pTable.tbTableCell(R + 1, C).tbcFontBold = m_FontBold
            pTable.tbTableCell(R + 1, C).tbcFontCharSet = m_FontCharSet
            pTable.tbTableCell(R + 1, C).tbcFontItalic = m_FontItalic
            pTable.tbTableCell(R + 1, C).tbcFontName = m_FontName
            pTable.tbTableCell(R + 1, C).tbcFontSize = m_FontSize
            pTable.tbTableCell(R + 1, C).tbcFontStrikethru = m_FontStrikethru
            pTable.tbTableCell(R + 1, C).tbcFontTransparent = True
            pTable.tbTableCell(R + 1, C).tbcFontUnderline = m_FontUnderline
            pTable.tbTableCell(R + 1, C).tbcForeColor = m_ForeColor
            Set pTable.tbTableCell(R + 1, C).tbcPicture = Nothing
            pTable.tbTableCell(R + 1, C).tbcRowSpan = 1
         Next
     Next

End Sub

Private Sub WritePages()
      LabelPages.Caption = Str(CurrentPage) + "/" + Str(PageCount)
End Sub


Public Function ScaleX(Unt As Variant, _
                            mFromScaleMode As ScaleModeConstants, _
                            mToScaleMode As ScaleModeConstants) As Single
        
        If ObjPrint Is Nothing Then Exit Function
        
        If InStr(1, Unt, "cm") Then
            Unt = (Replace(Unt, "cm", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleX = Round(ObjPrint.ScaleX((Unt), vbCentimeters, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "mm") Then
            Unt = (Replace(Unt, "mm", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
             ScaleX = Round(ObjPrint.ScaleX((Unt), vbMillimeters, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "in") Then
            Unt = (Replace(Unt, "in", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleX = Round(ObjPrint.ScaleX((Unt), vbInches, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "px") Then
            Unt = (Replace(Unt, "px", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleX = Round(ObjPrint.ScaleX((Unt), vbPixels, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "tw") Then
            Unt = (Replace(Unt, "tw", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleX = Round(ObjPrint.ScaleX((Unt), vbTwips, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "pt") Then
            Unt = (Replace(Unt, "pt", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleX = Round(ObjPrint.ScaleX((Unt), vbPoints, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "ch") Then
            Unt = (Replace(Unt, "ch", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleX = Round(ObjPrint.ScaleX((Unt), vbCharacters, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "%") Then
            Unt = Val(Replace(Unt, "%", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            Unt = Unt * PageWidth * (Unt / 100)
            ScaleX = Round(ObjPrint.ScaleX(Unt, mFromScaleMode, mToScaleMode), 3)
        Else
           If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
           ScaleX = Round(ObjPrint.ScaleX(Unt, mFromScaleMode, mToScaleMode), 3)
        End If
        
End Function

Public Function ScaleY(ByVal Unt As Variant, _
                            Optional mFromScaleMode As ScaleModeConstants, _
                            Optional mToScaleMode As ScaleModeConstants) As Single
        If ObjPrint Is Nothing Then Exit Function
        If InStr(1, Unt, "cm") Then
            Unt = Replace(Unt, "cm", "")
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY((Unt), vbCentimeters, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "mm") Then
            Unt = Replace(Unt, "mm", "")
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY((Unt), vbMillimeters, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "in") Then
            Unt = Replace(Unt, "in", "")
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY((Unt), vbInches, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "px") Then
            Unt = Replace(Unt, "px", "")
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY((Unt), vbPixels, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "tw") Then
            Unt = Replace(Unt, "tw", "")
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY((Unt), vbTwips, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "pt") Then
            Unt = (Replace(Unt, "pt", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY((Unt), vbPoints, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "ch") Then
            Unt = (Replace(Unt, "ch", ""))
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY((Unt), vbCharacters, mToScaleMode), 3)
        ElseIf InStr(1, Unt, "%") Then
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            Unt = Replace(Unt, "%", "")
            Unt = CSng(Unt) * PageHeight * Val(Unt) / 100
            ScaleY = Round(ObjPrint.ScaleY((Unt), mFromScaleMode, mToScaleMode), 3)
        Else
            If InStr(1, Unt, ".") > 0 Then Unt = Val(Unt)
            ScaleY = Round(ObjPrint.ScaleY(Unt, mFromScaleMode, mToScaleMode), 3)
        End If
        
End Function

Private Function ReplaceChar(ByVal txt As String) As String
                ReplaceChar = Replace(txt, vbCrLf, vbCr)
                ReplaceChar = Replace(ReplaceChar, vbLf, vbCr)
                ReplaceChar = Replace(ReplaceChar, Chr(0), "")
                ReplaceChar = Replace(ReplaceChar, vbTab, String(5, " "))
End Function

Public Sub CalcPicture()
      m_X1 = SetUnitPicture.Left
      m_Y1 = SetUnitPicture.Top
      m_X2 = SetUnitPicture.Right
      m_Y2 = SetUnitPicture.Bottom
End Sub

Public Sub CalcParagraph()
      m_X1 = SetUnitParagraph.Left
      m_Y1 = SetUnitParagraph.Top
      m_X2 = SetUnitParagraph.Right
      m_Y2 = SetUnitParagraph.Bottom
End Sub

Public Sub CalcTable()
      m_X1 = SetUnitTable.Left
      m_Y1 = SetUnitTable.Top
      m_X2 = SetUnitTable.Right
      m_Y2 = SetUnitTable.Bottom
End Sub

Public Sub CalcTextBox()
    m_X1 = SetUnitTextBox.Left
    m_Y1 = SetUnitTextBox.Top
    m_X2 = SetUnitTextBox.Right
    m_Y2 = SetUnitTextBox.Bottom
End Sub

Private Sub ChangePaperSize()
    Dim PrinterHandle As Long
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim MyDevMode As DEVMODE_TYPE
    Dim Result As Long
    Dim Needed As Long
    Dim pFullDevMode As Long
    Dim pi2_buffer() As Long
    CallingHwnd = hWnd
    PrinterName = Printer.DeviceName
    
    If PrinterName = "" Then
        Exit Sub
    End If
    
    pd.pDataType = vbNullString
    pd.pDevMode = 0&
    'Printer_Access_All is required for NT security
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    
    Result = OpenPrinter(PrinterName, PrinterHandle, pd)
    
    'The first call to GetPrinter gets the size, in bytes, of the buffer needed.
    'This value is divided by 4 since each element of pi2_buffer is a long.
    Result = GetPrinter(PrinterHandle, 2, ByVal 0&, 0, Needed)
    ReDim pi2_buffer((Needed \ 4))
    Result = GetPrinter(PrinterHandle, 2, pi2_buffer(0), Needed, Needed)
    
    'The seventh element of pi2_buffer is a Pointer to a block of memory
    ' which contains the full DevMode (including the PRIVATE portion).
    pFullDevMode = pi2_buffer(7)
    
    'Copy the Public portion of FullDevMode into our DevMode structure
    Call CopyMemory(MyDevMode, ByVal pFullDevMode, Len(MyDevMode))
    
    'Make desired changes
    InitialPaperSize = MyDevMode.dmPaperSize
    InitialPaperOrit = MyDevMode.dmOrientation
    MyDevMode.dmPaperSize = PaperSize
    If Orientation <> 0 Then MyDevMode.dmOrientation = Orientation
    
    'Copy our DevMode structure back into FullDevMode
    Call CopyMemory(ByVal pFullDevMode, MyDevMode, Len(MyDevMode))
    
    'Copy our changes to "the PUBLIC portion of the DevMode" into "the PRIVATE portion of the DevMode"
    Result = DocumentProperties(hWnd, PrinterHandle, PrinterName, ByVal pFullDevMode, ByVal pFullDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
    
    'Update the printer's default properties (to verify, go to the Printer folder
    ' and check the properties for the printer)
    Result = SetPrinter(PrinterHandle, 2, pi2_buffer(0), 0&)
    
    Call ClosePrinter(PrinterHandle)
End Sub

Private Sub RestorePrinterDefaults()
    Dim PrinterHandle As Long
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim MyDevMode As DEVMODE_TYPE
    Dim Result As Long
    Dim Needed As Long
    Dim pFullDevMode As Long
    Dim pi2_buffer() As Long
    
    PrinterName = Printer.DeviceName
    If PrinterName = "" Then
        Exit Sub
    End If
    
    pd.pDataType = vbNullString
    pd.pDevMode = 0&
    'Printer_Access_All is required for NT security
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    
    Result = OpenPrinter(PrinterName, PrinterHandle, pd)
    
    'The first call to GetPrinter gets the size, in bytes, of the buffer needed.
    'This value is divided by 4 since each element of pi2_buffer is a long.
    Result = GetPrinter(PrinterHandle, 2, ByVal 0&, 0, Needed)
    ReDim pi2_buffer((Needed \ 4))
    Result = GetPrinter(PrinterHandle, 2, pi2_buffer(0), Needed, Needed)
    
    'The seventh element of pi2_buffer is a Pointer to a block of memory
    ' which contains the full DevMode (including the PRIVATE portion).
    pFullDevMode = pi2_buffer(7)
    
    'Copy the Public portion of FullDevMode into our DevMode structure
    Call CopyMemory(MyDevMode, ByVal pFullDevMode, Len(MyDevMode))
    
    'Make desired changes
    MyDevMode.dmPaperSize = InitialPaperSize
    MyDevMode.dmOrientation = InitialPaperOrit
    
    'Copy our DevMode structure back into FullDevMode
    Call CopyMemory(ByVal pFullDevMode, MyDevMode, Len(MyDevMode))
    
    'Copy our changes to "the PUBLIC portion of the DevMode" into "the PRIVATE portion of the DevMode"
    Result = DocumentProperties(CallingHwnd, PrinterHandle, PrinterName, ByVal pFullDevMode, ByVal pFullDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
    
    'Update the printer's default properties (to verify, go to the Printer folder
    ' and check the properties for the printer)
    Result = SetPrinter(PrinterHandle, 2, pi2_buffer(0), 0&)
    
    Call ClosePrinter(PrinterHandle)
End Sub
