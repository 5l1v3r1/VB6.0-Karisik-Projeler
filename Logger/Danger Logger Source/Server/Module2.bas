Attribute VB_Name = "Module2"
Option Explicit
Option Base 0

Private Type PALETTEENTRY
    peRed   As Byte
    peGreen As Byte
    peBlue  As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY
End Type
         
Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
    ByVal hDC As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long) As Long

Private Declare Function GetDeviceCaps Lib "GDI32" ( _
    ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long

Private Declare Function GetSystemPaletteEntries Lib "GDI32" ( _
    ByVal hDC As Long, ByVal wStartIndex As Long, _
    ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
    As Long

Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long

Private Declare Function CreatePalette Lib "GDI32" ( _
    lpLogPalette As LOGPALETTE) As Long


Private Declare Function SelectPalette Lib "GDI32" ( _
    ByVal hDC As Long, ByVal hPalette As Long, _
    ByVal bForceBackground As Long) As Long

Private Declare Function RealizePalette Lib "GDI32" ( _
    ByVal hDC As Long) As Long


Private Declare Function SelectObject Lib "GDI32" ( _
    ByVal hDC As Long, ByVal hObject As Long) As Long


Private Declare Function BitBlt Lib "GDI32" ( _
    ByVal hDCDest As Long, ByVal XDest As Long, _
    ByVal YDest As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hDCSrc As Long, _
    ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
    As Long


Private Declare Function GetWindowDC Lib "USER32" ( _
    ByVal hWnd As Long) As Long


Private Declare Function GetDC Lib "USER32" ( _
    ByVal hWnd As Long) As Long


Private Declare Function ReleaseDC Lib "USER32" ( _
    ByVal hWnd As Long, ByVal hDC As Long) As Long


Private Declare Function DeleteDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long


Private Declare Function GetWindowRect Lib "USER32" ( _
    ByVal hWnd As Long, lpRect As RECT) As Long


Private Declare Function GetDesktopWindow Lib "USER32" () As Long


Private Declare Function GetForegroundWindow Lib "USER32" () As Long


Private Declare Function OleCreatePictureIndirect _
    Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Function CreateBitmapPicture(ByVal hBmp As Long, _
        ByVal hPal As Long) As Picture

Dim r   As Long
Dim Pic As PicBmp

Dim IPic          As IPicture
Dim IID_IDispatch As GUID

With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
End With

With Pic
    .Size = Len(Pic)
    .Type = vbPicTypeBitmap
    .hBmp = hBmp
    .hPal = hPal
End With

r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
Set CreateBitmapPicture = IPic
End Function


Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal bClient As Boolean, ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture

Dim hDCMemory       As Long
Dim hBmp            As Long
Dim hBmpPrev        As Long
Dim r               As Long
Dim hDCSrc          As Long
Dim hPal            As Long
Dim hPalPrev        As Long
Dim RasterCapsScrn  As Long
Dim HasPaletteScrn  As Long
Dim PaletteSizeScrn As Long
Dim LogPal          As LOGPALETTE

If bClient Then
    hDCSrc = GetDC(hWndSrc)
Else
    hDCSrc = GetWindowDC(hWndSrc)
End If

hDCMemory = CreateCompatibleDC(hDCSrc)

hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)

RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
HasPaletteScrn = RasterCapsScrn And RC_PALETTE
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)

If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    
    LogPal.palVersion = &H300
    LogPal.palNumEntries = 256
    r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
    hPal = CreatePalette(LogPal)
    
    hPalPrev = SelectPalette(hDCMemory, hPal, 0)
    r = RealizePalette(hDCMemory)
End If

r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
    LeftSrc, TopSrc, vbSrcCopy)

hBmp = SelectObject(hDCMemory, hBmpPrev)

If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
End If

r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hDCSrc)

Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CaptureScreen() As Picture
Dim hWndScreen As Long

hWndScreen = GetDesktopWindow()

With Screen
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
            .Width \ .TwipsPerPixelX, .Height \ .TwipsPerPixelY)
End With
End Function

Public Function CaptureForm(frm As Form) As Picture
With frm
    Set CaptureForm = CaptureWindow(.hWnd, False, 0, 0, _
            .ScaleX(.Width, vbTwips, vbPixels), _
            .ScaleY(.Height, vbTwips, vbPixels))
End With
End Function

Public Function CaptureClient(frm As Form) As Picture

With frm
    Set CaptureClient = CaptureWindow(.hWnd, True, 0, 0, _
            .ScaleX(.ScaleWidth, .ScaleMode, vbPixels), _
            .ScaleY(.ScaleHeight, .ScaleMode, vbPixels))
End With
End Function

Public Function CaptureActiveWindow() As Picture
Dim hWndActive As Long
Dim RectActive As RECT

hWndActive = GetForegroundWindow()
Call GetWindowRect(hWndActive, RectActive)

With RectActive
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
            .Right - .Left, .Bottom - .Top)
End With
End Function

Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)

Dim PicRatio     As Double
Dim PrnWidth     As Double
Dim PrnHeight    As Double
Dim PrnRatio     As Double
Dim PrnPicWidth  As Double
Dim PrnPicHeight As Double
Const vbHiMetric As Integer = 8

If Pic.Height >= Pic.Width Then
    Prn.Orientation = vbPRORPortrait
Else
    Prn.Orientation = vbPRORLandscape
End If

PicRatio = Pic.Width / Pic.Height

With Prn
    PrnWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbHiMetric)
    PrnHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbHiMetric)
End With

PrnRatio = PrnWidth / PrnHeight

If PicRatio >= PrnRatio Then
  
    PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
    PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
Else
   
    PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
    PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
End If

Call Prn.PaintPicture(Pic, 0, 0, PrnPicWidth, PrnPicHeight)
End Sub

