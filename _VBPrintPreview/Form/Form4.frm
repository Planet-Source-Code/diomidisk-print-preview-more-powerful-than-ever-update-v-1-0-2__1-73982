VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Fax Form"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12510
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicViewPort 
      Height          =   8625
      Left            =   420
      ScaleHeight     =   8565
      ScaleWidth      =   12345
      TabIndex        =   2
      Top             =   1050
      Width           =   12405
      Begin VB.PictureBox PicFrame 
         Height          =   8295
         Left            =   15
         ScaleHeight     =   8235
         ScaleWidth      =   11925
         TabIndex        =   5
         Top             =   -15
         Width           =   11985
         Begin VB.PictureBox PicConatiner 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   15840
            Left            =   120
            Picture         =   "Form4.frx":038A
            ScaleHeight     =   15810
            ScaleWidth      =   12210
            TabIndex        =   6
            Top             =   90
            Width           =   12240
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7140
               Index           =   8
               Left            =   1500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Text            =   "Form4.frx":275234
               Top             =   6435
               Width           =   9270
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   6435
               TabIndex        =   14
               Top             =   5580
               Width           =   3930
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   6600
               TabIndex        =   13
               Top             =   5085
               Width           =   3930
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   5
               Left            =   6735
               TabIndex        =   12
               Text            =   "1"
               Top             =   4560
               Width           =   3930
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   6675
               TabIndex        =   11
               Text            =   "My name"
               Top             =   4050
               Width           =   3930
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   1845
               TabIndex        =   10
               Top             =   5580
               Width           =   4005
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   2145
               TabIndex        =   9
               Text            =   "009876543210"
               Top             =   5070
               Width           =   3690
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1875
               TabIndex        =   8
               Text            =   "001234567890"
               Top             =   4560
               Width           =   3990
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   1800
               TabIndex        =   7
               Text            =   "Company "
               Top             =   4050
               Width           =   4035
            End
         End
      End
      Begin VB.VScrollBar VScroll 
         Height          =   3195
         Left            =   12030
         TabIndex        =   4
         Top             =   45
         Width           =   255
      End
      Begin VB.HScrollBar HScroll 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   8265
         Width           =   3435
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   450
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   90
      Width           =   1395
   End
   Begin PrintPreviewVB.VBPrintPreview VBPrintPreview1 
      Height          =   6615
      Left            =   465
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   11668
      CurrentX        =   2
      CurrentY        =   27,7
      FontName        =   "Lucida Sans Unicode"
      FontCharSet     =   161
      Zoom            =   5
      Header          =   ""
      Footer          =   "||Page p$"
      FromPage        =   1
      ToPage          =   1
      NavBar          =   3
      PageWidth       =   21
      PageHeight      =   29,7
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   450
      Index           =   1
      Left            =   105
      TabIndex        =   16
      Top             =   90
      Width           =   1395
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
  Select Case Index
  Case 0
     VBPrintPreview1.Visible = True
     PicViewPort.Visible = False
     Command1(0).Visible = False
     Command1(1).Visible = True
     PrintControls
  Case 1
     VBPrintPreview1.Visible = False
     PicViewPort.Visible = True
     Command1(1).Visible = False
     Command1(0).Visible = True
  End Select
End Sub

Sub PrintControls()
    Dim i As Integer, X As Single, Y As Single, w As Single, txt As String

    With VBPrintPreview1
        .Orientation = PagePortrait
        .Zoom = zmRation100
        .ScaleMode = smCentimeters
        ' prepare printer
        .StartDoc
        
        ' draw form
        .GetMargins
        .DrawPicture PicConatiner.Picture, 0, 0, "100%", "100%", vbSrcAnd
        
        ' draw each control
        On Error Resume Next
        For i = 0 To Controls.Count - 1
            If TypeName(Controls(i)) = "TextBox" Then
                ' adjust font
                .FontName = Controls(i).FontName
                .FontBold = Controls(i).FontBold
                .FontItalic = Controls(i).FontItalic
                .FontSize = Controls(i).FontSize
    
                ' adjust position
                X = .ScaleX(Controls(i).Left, PicConatiner.ScaleMode, .ScaleMode)
                Y = .ScaleY(Controls(i).Top, PicConatiner.ScaleMode, .ScaleMode)
                w = .ScaleX(Controls(i).Width, PicConatiner.ScaleMode, .ScaleMode)
                txt = Controls(i).Text
                .FillColor = RGB(255, 255, 255)
                .TextBox txt, X, Y, w, .ScaleY(Controls(i).Height, PicConatiner.ScaleMode, .ScaleMode), taLeftTop, False
           End If
        Next i
        On Error GoTo 0
        
        ' all done
        .EndDoc
    
    End With
    
End Sub

Private Sub Form_Load()
             
    SetScrolls VScroll
    SetScrolls HScroll
    Text1(6).Text = Format(Now, "dd/mm/yyyy Hh:Nn")
End Sub

Private Sub Form_Resize()
      
   PicViewPort.Move 0, 600, ScaleWidth, ScaleHeight - 600
   VBPrintPreview1.Move 0, 600, ScaleWidth, ScaleHeight - 600
   VBPrintPreview1.ZOrder 1
End Sub

Private Sub HScroll_Change()
    With PicConatiner
        .Left = -((.Width - PicFrame.Width) * HScroll.Value) / 100
    End With
End Sub

Private Sub VBPrintPreview1_PagePrint()

      If VBPrintPreview1.DialogPrint(pdPrint) Then
           VBPrintPreview1.SendToPrinter = True
           PrintControls
           VBPrintPreview1.SendToPrinter = False
           Command1_Click 1
      End If
        
End Sub

Private Sub VScroll_Change()
    With PicConatiner
        .Top = -((.Height - PicFrame.Height) * VScroll.Value) / 100
    End With
End Sub

Private Sub PicViewPort_Resize()
      PicFrame.Move 0, 0, PicViewPort.ScaleWidth - VScroll.Width, PicViewPort.ScaleHeight - HScroll.Height
      VScroll.Move PicViewPort.ScaleWidth - VScroll.Width, 0, VScroll.Width, PicViewPort.ScaleHeight - HScroll.Height
      HScroll.Move 0, PicViewPort.ScaleHeight - HScroll.Height, PicViewPort.ScaleWidth - VScroll.Width
End Sub

Private Sub SetScrolls(ByVal ObjS As Object)
    
    With ObjS
        .Min = 1
        .Max = 100
        .LargeChange = 10
    End With
End Sub

