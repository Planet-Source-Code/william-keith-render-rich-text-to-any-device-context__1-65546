VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTransparentRTB 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picContainer 
      Height          =   2895
      Left            =   3660
      ScaleHeight     =   2835
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   3180
      Width           =   3915
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2835
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5001
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmTransparentRTB.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditLeft 
         Caption         =   "&Left"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEditCenter 
         Caption         =   "C&enter"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditRight 
         Caption         =   "&Right"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditItalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEditUnderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditBlack 
         Caption         =   "Blac&k"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuEditRed 
         Caption         =   "Re&d"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditGreen 
         Caption         =   "&Green"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditBlue 
         Caption         =   "Blue"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit8pt 
         Caption         =   "8pt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEdit10pt 
         Caption         =   "10pt"
      End
      Begin VB.Menu mnuEdit14pt 
         Caption         =   "14pt"
      End
   End
End
Attribute VB_Name = "frmTransparentRTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_BTNFACE As Long = 15

Private Const COLORONCOLOR As Long = 3
Private Const HALFTONE As Long = 4

Private Const WM_ERASEBKGND As Long = &H14

Private Const WM_USER As Long = &H400

Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72

Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16

Private Const WS_EX_TRANSPARENT As Long = &H20&

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CHARRANGE
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FORMATRANGE
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As RECT        ' Region of the DC to draw to (in twips)
  rcPage As RECT    ' Region of the entire DC (page size) (in twips)
  chrg As CHARRANGE ' Range of text to draw (see above declaration)
End Type

Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, lp As Any) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function MaskBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim p As StdPicture

Private Sub Form_Load()
    '-- Load our background image
    Set p = LoadPicture(App.Path & "\DSCF0849.JPG")
    '-- Make the RTB transparent
    SetWindowLong rtb.hwnd, GWL_EXSTYLE, GetWindowLong(rtb.hwnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '-- Clean up
    Set p = Nothing
End Sub

Private Sub Form_Resize()
    '-- Make sure we paint when the form's size is reduced as well!
    Form_Paint
    '-- Move the RichTextBox's container PictureBox
    picContainer.Move Me.ScaleWidth - picContainer.Width - 180, Me.ScaleHeight - picContainer.Height - 180
End Sub

Private Sub rtb_Change()
    '-- Paint the form when the user types in the RichTextBox
    Form_Paint
End Sub

Private Sub Form_Paint()
    Dim mDC As Long
    Dim mBmp As Long
    Dim rc As RECT
    Dim fmt As FORMATRANGE
    Dim hBrush As Long
    
    '-- Get the width and height of the form. Actually shorter than
    '   converting from the twip values
    GetClientRect Me.hwnd, rc
    
    '-- Create a backbuffer
    mDC = CreateCompatibleDC(Me.hdc)
    mBmp = CreateCompatibleBitmap(Me.hdc, rc.Right, rc.Bottom)
    DeleteObject SelectObject(mDC, mBmp)
    
    '-- Make sure we won't get any of the horrid drawing artifacts that
    '   we would in the default BLACKONWHITE mode.  One could also set this
    '   to HALFTONE mode to get really high quality stretching, but that would
    '   be slow
    SetStretchBltMode mDC, COLORONCOLOR
    '-- Stretch the background image onto the backbuffer
    StretchImage _
        mDC, _
        0, 0, _
        HimetricToPixelsX(p.Width), HimetricToPixelsY(p.Height), _
        rc.Right, rc.Bottom, p.Handle, True
    
    '-- Set up our FORMATRANGE structure to draw the text to the backbuffer
    With fmt
        .hdc = mDC
        .chrg.cpMin = 0
        .chrg.cpMax = Len(rtb.Text)
        .hdcTarget = mDC
        SetRect .rc, 0, 0, Me.ScaleWidth, Me.ScaleHeight
        SetRect .rcPage, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End With
    
    '-- Get the RichTextBox to draw the text.  Since it has the transparent style,
    '   it will draw it transparently on to the image background
    SendMessage rtb.hwnd, EM_FORMATRANGE, 1, fmt
    
    '-- Blit the backbuffer to the form's DC
    BitBlt Me.hdc, 0, 0, rc.Right, rc.Bottom, mDC, 0, 0, vbSrcCopy
    
    '-- Clean up
    DeleteDC mDC
    DeleteObject mBmp
End Sub

Private Sub mnuEdit8pt_Click()
    rtb.SelFontSize = 8
    mnuEdit8pt.Checked = True
    mnuEdit10pt.Checked = False
    mnuEdit14pt.Checked = False
End Sub

Private Sub mnuEdit10pt_Click()
    '-- Set the font size
    rtb.SelFontSize = 10
    '-- Check/uncheck other items
    mnuEdit8pt.Checked = False
    mnuEdit10pt.Checked = True
    mnuEdit14pt.Checked = False
End Sub

Private Sub mnuEdit14pt_Click()
    '-- Set the font size
    rtb.SelFontSize = 14
    '-- Check/uncheck other items
    mnuEdit8pt.Checked = False
    mnuEdit10pt.Checked = False
    mnuEdit14pt.Checked = False
End Sub

Private Sub mnuEditBlack_Click()
    '-- Set the text colour
    rtb.SelColor = vbBlack
    '-- Check/uncheck other items
    mnuEditBlack.Checked = True
    mnuEditRed.Checked = False
    mnuEditGreen.Checked = False
    mnuEditBlue.Checked = False
End Sub

Private Sub mnuEditRed_Click()
    '-- Set the text colour
    rtb.SelColor = vbRed
    '-- Check/uncheck other items
    mnuEditBlack.Checked = False
    mnuEditRed.Checked = True
    mnuEditGreen.Checked = False
    mnuEditBlue.Checked = False
End Sub

Private Sub mnuEditGreen_Click()
    '-- Set the text colour
    rtb.SelColor = vbGreen
    '-- Check/uncheck other items
    mnuEditBlack.Checked = False
    mnuEditRed.Checked = False
    mnuEditGreen.Checked = True
    mnuEditBlue.Checked = False
End Sub

Private Sub mnuEditBlue_Click()
    '-- Set the text colour
    rtb.SelColor = vbBlue
    '-- Check/uncheck other items
    mnuEditBlack.Checked = False
    mnuEditRed.Checked = False
    mnuEditGreen.Checked = False
    mnuEditBlue.Checked = True
End Sub

Private Sub mnuEditBold_Click()
    '-- Make the text bold (or not)
    rtb.SelBold = Not mnuEditBold.Checked
    '-- Check/uncheck the item
    mnuEditBold.Checked = Not mnuEditBold.Checked
End Sub

Private Sub mnuEditItalic_Click()
    '-- Italicise the text (or not)
    rtb.SelItalic = Not mnuEditItalic.Checked
    '-- Check/uncheck the item
    mnuEditItalic.Checked = Not mnuEditItalic.Checked
End Sub

Private Sub mnuEditUnderline_Click()
    '-- Underline the text (or not)
    rtb.SelUnderline = Not mnuEditUnderline.Checked
    '-- Check/uncheck the item
    mnuEditUnderline.Checked = Not mnuEditUnderline.Checked
End Sub

Private Sub mnuEditLeft_Click()
    '-- Set the text colour
    rtb.SelAlignment = rtfLeft
    '-- Check/uncheck other items
    mnuEditLeft.Checked = True
    mnuEditCenter.Checked = False
    mnuEditRight.Checked = False
End Sub

Private Sub mnuEditCenter_Click()
    '-- Set the text colour
    rtb.SelAlignment = rtfCenter
    '-- Check/uncheck other items
    mnuEditLeft.Checked = False
    mnuEditCenter.Checked = True
    mnuEditRight.Checked = False
End Sub

Private Sub mnuEditRight_Click()
    '-- Set the text colour
    rtb.SelAlignment = rtfRight
    '-- Check/uncheck other items
    mnuEditLeft.Checked = False
    mnuEditCenter.Checked = False
    mnuEditRight.Checked = True
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Public Sub DrawImage(hdc As Long, x As Long, y As Long, Width As Long, Height As Long, hBmp As Long)
    Dim hTmpDC As Long
    
    hTmpDC = CreateCompatibleDC(hdc)
    DeleteObject SelectObject(hTmpDC, hBmp)
    BitBlt hdc, x, y, Width, Height, hTmpDC, 0, 0, vbSrcCopy
    DeleteDC hTmpDC
End Sub

Public Sub StretchImage(hdc As Long, x As Long, y As Long, Width As Long, Height As Long, DestWidth As Long, DestHeight As Long, hBmp As Long, Optional FixAspectRatio As Boolean = False, Optional ScaleToWidth As Boolean = False)
    Dim hTmpDC As Long
    
    hTmpDC = CreateCompatibleDC(hdc)
    DeleteObject SelectObject(hTmpDC, hBmp)
    
    If FixAspectRatio Then
        If ScaleToWidth Then
            DestHeight = DestWidth * (Height / Width)
        Else
            DestWidth = DestHeight * (Width / Height)
        End If
    End If
    
    StretchBlt hdc, x, y, DestWidth, DestHeight, hTmpDC, 0, 0, Width, Height, vbSrcCopy
    DeleteDC hTmpDC
End Sub

Public Function HimetricToPixelsX(ByVal Value As Single) As Single
    HimetricToPixelsX = (Value / 1000) * 567 / Screen.TwipsPerPixelX
End Function

Public Function HimetricToPixelsY(ByVal Value As Single) As Single
    HimetricToPixelsY = (Value / 1000) * 567 / Screen.TwipsPerPixelY
End Function

