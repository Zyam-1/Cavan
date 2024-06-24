VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmPrintBarCodeDemo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Barcode Printing ..."
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBarCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3792
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   2
      Top             =   912
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Timer timBarcode 
      Interval        =   1000
      Left            =   2424
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   228
      Left            =   2328
      TabIndex        =   4
      Top             =   1656
      Visible         =   0   'False
      Width           =   24
      _ExtentX        =   53
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblChartDOB 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   24
      TabIndex        =   3
      Top             =   1224
      Width           =   2724
   End
   Begin VB.Label lblSID 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   24
      TabIndex        =   1
      Top             =   720
      Width           =   2160
   End
   Begin VB.Label lblName 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   24
      TabIndex        =   0
      Top             =   960
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   648
      Left            =   0
      Stretch         =   -1  'True
      Top             =   48
      Width           =   2160
   End
End
Attribute VB_Name = "frmPrintBarCodeDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String
Private pPatName As String
Private pChart As String
Private pPatientDOB As String


Public Property Let SampleID(ByVal strSampleID As String)

10    pSampleID = strSampleID

End Property

Public Property Let PatientName(ByVal strPatName As String)

10    pPatName = strPatName

End Property


Public Property Let Chart(ByVal strChart As String)

10    pChart = strChart

End Property

Public Property Let PatientDOB(ByVal strPatientDOB As String)

10    pPatientDOB = strPatientDOB

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

10    Me.FontBold = True

20    lblSID = pSampleID
30    lblSID.FontSize = 10 '8
40    lblName.FontSize = 10 '6

50    lblName = pPatName
60    lblChartDOB = "Ch: " & pChart & " DoB: " & pPatientDOB

70    blnPrinterSet = False
80    For Each Px In Printers
90        If UCase$(Px.DeviceName) = UCase(TransfusionLabel) Then
100     Set Printer = Px
110     blnPrinterSet = True
120     Printer.Orientation = vbPRORPortrait    'vbPRORLandscape
130     Exit For
140       End If
150   Next

160   If blnPrinterSet Then

    '130     LoadSIDdetails
    ' Using CodeB
    'YPos = GenerateCode128("Testing Code 128", 10, 10, 1)
170     yPos = GenerateCode128(pSampleID, 150, 1, 2)    'yPos = GenerateCode128(pSampleId, 150, 1, 2)


180     picBarCode.Picture = picBarCode.Image
190     Image1.Picture = picBarCode.Picture

200     PrintForm
210     Printer.EndDoc
220   End If


End Sub

Private Sub timBarcode_Timer()
10    Unload Me
End Sub

