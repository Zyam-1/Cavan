VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmPrintBarCodeUnit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Barcode Printing ..."
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timBarcode 
      Interval        =   2000
      Left            =   3105
      Top             =   195
   End
   Begin VB.PictureBox picBarCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3720
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   1
      Top             =   1176
      Visible         =   0   'False
      Width           =   2685
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   228
      Left            =   0
      TabIndex        =   2
      Top             =   1008
      Visible         =   0   'False
      Width           =   24
      _ExtentX        =   53
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   75
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3105
   End
   Begin VB.Label lblSID 
      BackColor       =   &H8000000E&
      Height          =   180
      Left            =   630
      TabIndex        =   0
      Top             =   570
      Width           =   1830
   End
End
Attribute VB_Name = "frmPrintBarCodeUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pUnit As String

Public Property Let Unit(ByVal strUnit As String)

10    pUnit = strUnit

End Property

Private Function GenerateCode128(Str As String, xPos As Single, yPos As Single, Optional BarWidth As Integer = 1) As Single
    Dim Code128 As New clsCode128
    Dim BarCodeWidth As Long

10    Me.picBarCode.Cls
20    Me.picBarCode.Width = 1
30    BarCodeWidth = Code128.Code128_Print(Str, Me.picBarCode, BarWidth, True)
40    Me.PaintPicture Me.picBarCode.Image, xPos, yPos, Me.picBarCode.ScaleWidth, Me.picBarCode.ScaleHeight, 0, 0, Me.picBarCode.ScaleWidth, Me.picBarCode.ScaleHeight
50    Me.CurrentX = xPos + BarCodeWidth / 2 - Me.TextWidth(Str) / 2
60    Me.CurrentY = yPos + Me.picBarCode.ScaleHeight + 2
70    Me.Print Str

80    GenerateCode128 = Me.CurrentY



End Function




Private Sub Form_Load()
    Dim yPos As Single
    Dim Px As Printer
    Dim blnPrinterSet As Boolean

10  Me.FontBold = True

20  lblSID = pUnit
30  lblSID.FontSize = 8

40  blnPrinterSet = False
50  For Each Px In Printers
60      If UCase$(Px.DeviceName) = UCase(TransfusionLabel) Then
70          Set Printer = Px
80          blnPrinterSet = True
90          Printer.Orientation = vbPRORPortrait    'vbPRORLandscape
100         Exit For
110     End If
120 Next

130 If blnPrinterSet Then

    '130     LoadSIDdetails
    ' Using CodeB
    'YPos = GenerateCode128("Testing Code 128", 10, 10, 1)
140     yPos = GenerateCode128(pUnit, 150, 1, 2)    'yPos = GenerateCode128(pSampleId, 150, 1, 2)


150     picBarCode.Picture = picBarCode.Image
160     Image1.Picture = picBarCode.Picture

170     PrintForm
180     Printer.EndDoc
190 End If


End Sub

Private Sub timBarcode_Timer()
10    Unload Me
End Sub


