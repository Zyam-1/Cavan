VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Logical Font
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_MASK = &HF
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_TT_ALWAYS = 32
Private Const CLIP_EMBEDDED = 128

Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255

' Font Families
'
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Private Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Private Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Private Const FF_SCRIPT = 64     '  Cursive, etc.
Private Const FF_DECORATIVE = 80 '  Old English, etc.

' Font Weights
Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_HEAVY = 900

Private Const FW_ULTRALIGHT = FW_EXTRALIGHT
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_BLACK = FW_HEAVY

Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_PRECIS = 4
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_OUTLINE_PRECIS = 8

Private Type LogFont
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private m_Font As StdFont
Private m_hFont As Long
Private m_Rotation As Single

Private Sub Class_Terminate()
          '
          ' Clean-up created objects!!!
          '
6040      If m_hFont Then
6050          Call DeleteObject(m_hFont)
6060          Set m_Font = Nothing
6070      End If
End Sub

Public Property Set LogFont(ByVal NewFont As Font)
6080      If m_hFont Then
6090          Call DeleteObject(m_hFont)
6100          m_hFont = 0
6110      End If
         
6120      If NewFont Is Nothing Then
6130          Set m_Font = Nothing
6140      Else
              '
              ' Stash a copy of the passed object,
              ' to avoid a new reference to it.
              '
6150          Set m_Font = New StdFont
6160          With m_Font
6170              .Bold = NewFont.Bold
6180              .Charset = NewFont.Charset
6190              .Italic = NewFont.Italic
6200              .Name = NewFont.Name
6210              .size = NewFont.size
6220              .Strikethrough = NewFont.Strikethrough
6230              .Underline = NewFont.Underline
6240              .Weight = NewFont.Weight
6250          End With
6260          m_hFont = CreateLogFont
6270      End If
End Property

Public Property Get LogFont() As Font
6280      Set LogFont = m_Font
End Property

Public Property Let Rotation(ByVal NewVal As Single)
6290      If NewVal <> m_Rotation Then
6300          m_Rotation = NewVal
6310          If m_hFont Then
6320              Call DeleteObject(m_hFont)
6330              m_hFont = 0
6340          End If
6350          If Not (m_Font Is Nothing) Then
6360              m_hFont = CreateLogFont
6370          End If
6380      End If
End Property

Public Property Get Rotation() As Single
6390      Rotation = m_Rotation
End Property

Public Property Get Handle() As Long
6400      Handle = m_hFont
End Property

Private Function CreateLogFont() As Long
          Dim lf As LogFont
          Dim hWnd As Long
          Dim hdc As Long
         
6410      hWnd = GetDesktopWindow
6420      hdc = GetDC(hWnd)
         
6430      With lf
              '
              ' All but two properties are very straight-forward,
              ' even with rotation, and map directly.
              '
6440          .lfHeight = -(m_Font.size * GetDeviceCaps(hdc, LOGPIXELSY)) / 72
6450          .lfWidth = 0
6460          .lfEscapement = m_Rotation * 10
6470          .lfOrientation = .lfEscapement
6480          .lfWeight = m_Font.Weight
6490          .lfItalic = m_Font.Italic
6500          .lfUnderline = m_Font.Underline
6510          .lfStrikeOut = m_Font.Strikethrough
6520          .lfClipPrecision = CLIP_DEFAULT_PRECIS
6530          .lfQuality = PROOF_QUALITY
6540          .lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
6550          .lfFaceName = m_Font.Name & vbNullChar
              '
              ' OEM fonts can't rotate, and we must force
              ' substitution with something ANSI.
              '
6560          .lfCharSet = m_Font.Charset
6570          If .lfCharSet = OEM_CHARSET Then
6580              If (m_Rotation Mod 360) <> 0 Then
6590                  .lfCharSet = ANSI_CHARSET
6600              End If
6610          End If
              '
              ' Only TrueType fonts can rotate, so we must
              ' specify TT-only if angle is not zero.
              '
6620          If (m_Rotation Mod 360) <> 0 Then
6630              .lfOutPrecision = OUT_TT_ONLY_PRECIS
6640          Else
6650              .lfOutPrecision = OUT_DEFAULT_PRECIS
6660          End If
6670      End With
         
6680      CreateLogFont = CreateFontIndirect(lf)
6690      Call ReleaseDC(hWnd, hdc)
End Function
