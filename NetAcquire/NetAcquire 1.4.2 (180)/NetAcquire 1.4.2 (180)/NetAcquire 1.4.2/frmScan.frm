VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Scan"
   ClientHeight    =   8355
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "From Scanner"
      Height          =   1100
      Left            =   45
      TabIndex        =   9
      Top             =   510
      Width           =   7380
      Begin VB.OptionButton optPixelType 
         Caption         =   "B/W"
         Height          =   195
         Index           =   1
         Left            =   3750
         TabIndex        =   13
         Top             =   750
         Width           =   675
      End
      Begin VB.OptionButton optPixelType 
         Caption         =   "Grey"
         Height          =   195
         Index           =   2
         Left            =   4965
         TabIndex        =   12
         Top             =   750
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optPixelType 
         Caption         =   "Colour"
         Height          =   195
         Index           =   4
         Left            =   6120
         TabIndex        =   11
         Top             =   750
         Width           =   795
      End
      Begin VB.ComboBox cmbResolution 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   630
         Width           =   2835
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pixel Type"
         Height          =   195
         Left            =   3720
         TabIndex        =   15
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Resolution"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "From a File"
      Height          =   6285
      Left            =   45
      TabIndex        =   7
      Top             =   1740
      Width           =   7380
      Begin VB.TextBox txtFilePath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   5940
         Width           =   7260
      End
      Begin VB.Image Picture1 
         Height          =   5640
         Left            =   90
         Stretch         =   -1  'True
         Top             =   180
         Width           =   7170
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   1100
      Left            =   7560
      Picture         =   "frmScan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1860
      Width           =   1200
   End
   Begin VB.CommandButton cmdSaveFile 
      Caption         =   "Save"
      Height          =   1100
      Left            =   7560
      Picture         =   "frmScan.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3030
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   5460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   1100
      Left            =   7560
      Picture         =   "frmScan.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6885
      Width           =   1200
   End
   Begin VB.TextBox txtSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   1875
   End
   Begin VB.CommandButton cmdscan 
      Caption         =   "&Scan"
      Height          =   1100
      Left            =   7560
      Picture         =   "frmScan.frx":1DFE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   510
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   45
      TabIndex        =   4
      Top             =   8085
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BLOCK_SIZE = 10000

Private Sub cmdBrowse_Click()



48710 On Error GoTo cmdBrowse_Click_Error

48720 With CommonDialog1
48730   .CancelError = False
48740   .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
48750   .Filter = "(*.jpg)|*.jpg|(*.bmp)|*.bmp"
48760   .ShowOpen
48770   If .FileName = "" Then
48780       Exit Sub
48790   Else
48800       txtFilePath = .FileName
48810       Call LoadAndScalePic(Picture1, .FileName)
            '100               Picture1.Picture = LoadPicture(CommonDialog1.FileName)
48820   End If

48830 End With


48840 Exit Sub

cmdBrowse_Click_Error:
          Dim strES As String
          Dim intEL As Integer

48850 intEL = Erl
48860 strES = Err.Description
48870 LogError "frmScan", "cmdBrowse_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

48880 Unload Me

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSaveFile_Click
' Author    : Masood
' Date      : 03/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveFile_Click()
48890 On Error GoTo cmdSaveFile_Click_Error
      Dim PictBmp As String
      Dim ByteData() As Byte   'Byte array for Blob data.
      Dim SourceFile As Integer
      Dim filelength As Long
      Dim Numblocks As Integer
      Const BlockSize = 100000
      Dim LeftOver As Long
      Dim i As Long

      Dim tb As Recordset
      Dim sql As String
      Dim FilePath As String
48900 If Len(txtFilePath) = 0 Then Exit Sub

48910 sql = "SELECT * FROM ScannedImages WHERE 0 = 1"
48920 Set tb = New Recordset
48930 RecOpenClient 0, tb, sql
48940 tb.AddNew



48950 PictBmp = txtFilePath
48960 Me.MousePointer = vbHourglass
48970 SourceFile = FreeFile
48980 Open PictBmp For Binary Access Read As SourceFile
48990 filelength = LOF(SourceFile)  ' Get the length of the file.
49000 If filelength = 0 Then
49010     Close SourceFile
          'MsgBox PictBmp & " empty or not found."
49020     Me.MousePointer = vbNormal
49030     Exit Sub
49040 Else
49050     Numblocks = filelength / BlockSize
49060     LeftOver = filelength Mod BlockSize
49070     ReDim ByteData(LeftOver)
49080     Get SourceFile, , ByteData()
49090     tb.Fields("ScannedImage").AppendChunk ByteData()
49100     ReDim ByteData(BlockSize)
49110     For i = 1 To Numblocks
49120         Get SourceFile, , ByteData()
49130         tb.Fields("ScannedImage").AppendChunk ByteData()
49140     Next i
          'tb!filelength = filelength
49150     Close SourceFile
49160 End If
49170 tb!SampleID = txtSampleID
49180 tb!ScannedName = Right(txtFilePath, Len(txtFilePath) - InStrRev(txtFilePath, "\"))    'txtFilePath 'ScannedName
49190 tb!RemoveFromLisDisplay = 0

49200 Me.MousePointer = vbNormal

49210 tb.Update

49220 Picture1.Picture = Nothing
49230 txtFilePath = ""

49240 Exit Sub


cmdSaveFile_Click_Error:

      Dim strES As String
      Dim intEL As Integer

49250 intEL = Erl
49260 strES = Err.Description
49270 LogError "frmScan", "cmdSaveFile_Click", intEL, strES, sql

End Sub
Private Sub LoadAndScalePic(ByRef imgTarget As Image, ByVal strPicName As String)
          Dim Left As Long, Top As Long, width As Long, height As Long
          Dim AspectRate As Double
49280     On Error GoTo LoadAndScalePic_Error

49290 If Len(strPicName) > 0 Then
49300   imgTarget.Visible = False
49310   imgTarget.Stretch = False
49320   imgTarget.Picture = LoadPicture(strPicName)
49330   AspectRate = imgTarget.width / imgTarget.height
49340   If AspectRate > imgTarget.Container.width / imgTarget.Container.height Then
            'Wide
49350       width = imgTarget.Container.width - 500
49360       height = width / AspectRate - 500
49370   Else
            'High
49380       height = imgTarget.Container.height - 1000
49390       width = height * AspectRate
49400   End If
49410   Left = (imgTarget.Container.width - width) / 4
49420   Top = (imgTarget.Container.height - height) / 4
49430   imgTarget.Stretch = True
49440   imgTarget.Move Left, Top, width, height
49450   imgTarget.Visible = True
49460 End If

49470     Exit Sub

LoadAndScalePic_Error:
          Dim strES As String
          Dim intEL As Integer

49480     intEL = Erl
49490     strES = Err.Description
49500     LogError "frmScan", "LoadAndScalePic", intEL, strES

End Sub
'---------------------------------------------------------------------------------------
' Procedure : SavePic
' Author    : Masood
' Date      : 03/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : ChunkPic
' Author    : Masood
' Date      : 03/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ChunkPic(ByVal rs As Recordset, FileName As String)

          Dim file_num As String
          Dim file_length As String
          Dim bytes() As Byte
          Dim num_blocks As Long
          Dim left_over As Long
          Dim block_num As Long

          Dim Fragment As Integer, Chunk() As Byte, i As Integer

49510 On Error GoTo ChunkPic_Error


          Dim nFile As Integer



          Dim src() As Byte
          '    Dim nFile As Integer
          Dim file As Integer
          Dim PictureFile As Integer

49520 nFile = FreeFile
          ' Open FileName For Input Access Read As #nFile
49530 Open FileName For Binary Access Read As #nFile
49540 ReDim src(0 To LOF(1) - 1)
49550 Get #nFile, , src
49560 Close


49570 file_num = FreeFile
49580 Open FileName For Binary Access Read As #file_num

49590 file_length = LOF(file_num)
49600 If file_length > 0 Then
49610   num_blocks = file_length / BLOCK_SIZE
49620   left_over = file_length Mod BLOCK_SIZE
49630   ReDim bytes(BLOCK_SIZE)
49640   For block_num = 1 To num_blocks
49650       Get #file_num, , bytes()
49660       rs.Fields("ScannedImage").AppendChunk Chunk()
49670   Next block_num

49680   If left_over > 0 Then
49690       ReDim bytes(left_over)
49700       Get #file_num, , bytes()
49710       rs.Fields("ScannedImage").AppendChunk Chunk()
49720   End If
49730   Close #file_num
49740 End If


49750 Exit Sub


ChunkPic_Error:

          Dim strES As String
          Dim intEL As Integer

49760 intEL = Erl
49770 strES = Err.Description
49780 LogError "frmScan", "ChunkPic", intEL, strES
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdscan_Click
' Author    : Masood
' Date      : 24/Feb/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdScan_Click()

          Dim PixelType As Integer
          Dim Resolution As Integer
49790 On Error GoTo cmdScan_Error
49800 If optPixelType(1).Value = True Then
49810   PixelType = 1
49820 ElseIf optPixelType(2).Value = True Then
49830   PixelType = 2
49840 ElseIf optPixelType(4).Value = True Then
49850   PixelType = 4
49860 End If

49870 Resolution = cmbResolution.Text

49880 Scan txtSampleID, PixelType, Resolution
49890 Exit Sub

cmdScan_Error:

          Dim strES As String
          Dim intEL As Integer

49900 intEL = Erl
49910 strES = Err.Description
49920 LogError "frmScan", "cmdScan_Click", intEL, strES
End Sub







'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Babar Shahzad
' Date      : 01/10/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Babar Shahzad
' Date      : 01/10/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

          Dim TwainOK As Long
          Dim n As Integer


49930     On Error GoTo Form_Load_Error

49940     TwainOK = TWAIN_IsAvailable()

49950     If TwainOK = 1 Then
49960         cmdscan.Enabled = True
49970     Else
49980         cmdscan.Enabled = False
49990     End If

50000     cmbResolution.Clear
50010     For n = 1 To 13
50020         cmbResolution.AddItem Choose(n, 75, 100, 150, 200, 300, _
                                           600, 1200, 2400, 3600, 4800, _
                                           7200, 9600, 19200)
50030     Next
50040     cmbResolution.ListIndex = 0

50050     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50060     intEL = Erl
50070     strES = Err.Description
50080     LogError "frmScan", "Form_Load", intEL, strES

End Sub

Private Sub mnuFile_Click()

          Dim Folder As String
          Dim FolderFound As Boolean
          Dim Temp As String

50090 On Error GoTo mnuFile_Click_Error

50100 Temp = GetOptionSetting("ScanPath", "")
50110 Folder = iBOX("Path to save", , Temp)
50120 If Right$(Folder, 1) <> "\" Then
50130   Folder = Folder & "\"
50140 End If
50150 If Trim$(Folder) = "\" Then
50160   Folder = ""
50170   FolderFound = True
50180 Else
50190   If Dir(Folder, vbDirectory) = "" Then
50200       FolderFound = False
50210       If iMsg("Folder does not exist" & vbCrLf & _
                    "Do you want to create it?", vbQuestion + vbYesNo) = vbYes Then
50220           MkDir Folder
50230           FolderFound = True
50240       End If
50250   Else
50260       FolderFound = True
50270   End If
50280 End If

50290 If FolderFound Then
50300   SaveOptionSetting "ScanPath", Folder
50310 End If

50320 Exit Sub

mnuFile_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50330 intEL = Erl
50340 strES = Err.Description
50350 iMsg "Error - Line " & intEL & ". " & strES

End Sub

