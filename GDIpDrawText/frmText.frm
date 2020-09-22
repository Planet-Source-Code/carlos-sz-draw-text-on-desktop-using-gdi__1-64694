VERSION 5.00
Begin VB.Form frmText 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmText"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tokenText As Long
Private myFontName As String
Private mySize As Single
Private formSW As Long
Private formSH As Long

Private tempBlend As BLENDFUNCTION
Private tempBI As BITMAPINFO
Private graphics As Long, brush As Long, pen As Long
Attribute brush.VB_VarUserMemId = 1073938439
Attribute pen.VB_VarUserMemId = 1073938439
Private brush_sha As Long
Attribute brush_sha.VB_VarUserMemId = 1073938442
Private pen_sha As Long
Attribute pen_sha.VB_VarUserMemId = 1073938443
Private fontFam As Long, curFont As Long, strFormat As Long
Attribute fontFam.VB_VarUserMemId = 1073938437
Attribute curFont.VB_VarUserMemId = 1073938437
Attribute strFormat.VB_VarUserMemId = 1073938437
Private rcLayout As RECTF   ' Designates the string drawing bounds
Attribute rcLayout.VB_VarUserMemId = 1073938440
Private path As Long
Attribute path.VB_VarUserMemId = 1073938441
Private path_sha As Long
Attribute path_sha.VB_VarUserMemId = 1073938446
Private winSize As size
Attribute winSize.VB_VarUserMemId = 1073938443
Private srcPoint As POINTAPI
Attribute srcPoint.VB_VarUserMemId = 1073938444

Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hDcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private mDC As Long  ' Memory hDC
Attribute mDC.VB_VarUserMemId = 1073938451
Private mainBitmap As Long    ' Memory Bitmap
Attribute mainBitmap.VB_VarUserMemId = 1073938452
Private blendFunc32bpp As BLENDFUNCTION
Attribute blendFunc32bpp.VB_VarUserMemId = 1073938453
Private oldBitmap As Long
Attribute oldBitmap.VB_VarUserMemId = 1073938454

Private Const PixelFormatBPPMask As Long = &HFF00&
Private dcMemory As Long
Attribute dcMemory.VB_VarUserMemId = 1073938457

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Private Sub Form_Initialize()

    Dim GpInput As GdiplusStartupInput
    GpInput.GdiplusVersion = 1
    If GdiplusStartup(tokenText, GpInput) <> Ok Then
        MsgBox "Error loading GDI+!", vbCritical
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    Call pvInitialize
    Call pvLoadTools
    Call pvDrawText
    Call pvFinalize

End Sub

Private Sub pvInitialize()

'I use a label to have an idea of the final form's size
'There's probably a better approach... Please tell me!
'you can load here e.g. all GDI+ settings that you need (from an INI file)
    Label1.Caption = "Draw Nice Text With GDI+"
    'Define the GDI+ font
    myFontName = "Tahoma"
    'Define the GDI+ font size
    mySize = 18
    Label1.Font = myFontName
    Label1.FontSize = mySize

    'define the initial form size
    Me.ScaleMode = 1
    Me.Width = Label1.Width
    Me.Height = 500
    Me.ScaleMode = 3

    'record for bitmap use
    'it's interesting to get this if you need to refresh the text every n secs (like time function)
    formSW = Me.ScaleWidth
    formSH = Me.ScaleHeight

End Sub

Private Sub pvFinalize()

'placed at top so you can see how nice the text is!
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

    Me.ScaleMode = 1
    Me.Width = Label1.Width
    Me.Height = 500
    Me.ScaleMode = 3

    Me.Top = frmMain.Top + frmMain.Height + 80
    Me.Left = (frmMain.Width * 0.5) + frmMain.Left - (Me.Width * 0.5)

End Sub

Private Sub pvLoadTools()

'the idea here is to create all tools that you'll need to draw the text just once.
'interesting, again, if you have to update the text all the time (this can slow
'things down if you create every e.g. second)

    Dim curWinLong As Long

    curWinLong = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED

    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32
        .biHeight = formSH    'Me.ScaleHeight
        .biWidth = formSW    'Me.ScaleWidth
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
    End With

    'drawing bounds (form size)
    rcLayout.Right = Label1.Width
    rcLayout.Bottom = 500

    'create the outline font pen
    Call GdipCreatePen1(GetRGB_VB2GDIP(vbBlack, 155), 2, UnitPixel, pen)
    'create brush to fill the font area
    Call GdipCreateSolidFill(GetRGB_VB2GDIP(vbWhite, 225), brush)

    'create the font
    Call GdipCreateFontFamilyFromName(StrConv(myFontName, vbUnicode), 0, fontFam)
    Call GdipCreateFont(fontFam, mySize, FontStyleBold, UnitPoint, curFont)

    'create the string
    Call GdipCreateStringFormat(0, 0, strFormat)
    'NOTE: Center was selected because the text will be placed at the bottom center of frmMain
    Call GdipSetStringFormatAlign(strFormat, StringAlignmentCenter)

    'create the font path (where it will be made)
    Call GdipCreatePath(FillModeWinding, path)

    'create a simple shadow effect
    Call GdipCreatePen1(GetRGB_VB2GDIP(vbGrayed, 50), 2, UnitPixel, pen_sha)
    Call GdipCreateSolidFill(GetRGB_VB2GDIP(vbGrayed, 50), brush_sha)
    Call GdipCreatePath(FillModeWinding, path_sha)

    'update layer window stuff (that will blend and show the GDI+ text)
    srcPoint.X = 0
    srcPoint.Y = 0
    winSize.cx = Me.ScaleWidth
    winSize.cy = Me.ScaleHeight

    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With

End Sub

Private Sub pvDrawText()

'initialize graphics
    mDC = CreateCompatibleDC(Me.hdc)
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    oldBitmap = SelectObject(mDC, mainBitmap)

    Call GdipCreateFromHDC(mDC, graphics)

    'pad shadow left/top
    rcLayout.Left = 2
    rcLayout.Top = 2

    'add the shadow text to its path
    Call GdipAddPathString(path_sha, StrConv(Label1, vbUnicode), -1, fontFam, 1, mySize, rcLayout, strFormat)

    'main text position
    rcLayout.Left = 0
    rcLayout.Top = 0

    'add it to its path
    Call GdipAddPathString(path, StrConv(Label1, vbUnicode), -1, fontFam, 1, mySize, rcLayout, strFormat)

    'it'll make the final draw (graphics) real better!
    Call GdipSetSmoothingMode(graphics, SmoothingModeAntiAlias)

    'draw and fill the paths
    'shadow
    Call GdipDrawPath(graphics, pen_sha, path_sha)
    Call GdipFillPath(graphics, brush_sha, path_sha)
    'main text
    Call GdipDrawPath(graphics, pen, path)
    Call GdipFillPath(graphics, brush, path)

    'delete things
    'NOTE: you can delete the font family, pens, brushes etc.
    'but this code aimed a text that needs to be updated every second

    Call GdipDeletePath(path_sha)
    Call GdipDeletePath(path)
    Call GdipDeleteGraphics(graphics)

    'now the text will be shown
    Call UpdateLayeredWindow(Me.hwnd, Me.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)

    'cleaning
    SelectObject mDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC mDC

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call GdiplusShutdown(tokenText)
    
End Sub

