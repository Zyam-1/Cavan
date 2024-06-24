VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - View Scanned Reports"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   15585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDefaultZoom 
      Alignment       =   2  'Center
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
      Left            =   12330
      TabIndex        =   22
      Text            =   "70"
      Top             =   2760
      Width           =   435
   End
   Begin VB.CommandButton cmdGrowShrink 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   12045
      TabIndex        =   21
      Top             =   105
      Width           =   255
   End
   Begin VB.CommandButton cmdZoom 
      Height          =   585
      Index           =   1
      Left            =   12330
      Picture         =   "frmViewScan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6390
      Width           =   690
   End
   Begin VB.CommandButton cmdZoom 
      Height          =   585
      Index           =   0
      Left            =   12330
      Picture         =   "frmViewScan.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3210
      Width           =   690
   End
   Begin MSComctlLib.Slider sldZoom 
      Height          =   2505
      Left            =   12375
      TabIndex        =   17
      Top             =   3825
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   4419
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   18
      SmallChange     =   10
      Min             =   10
      Max             =   300
      SelStart        =   10
      TickStyle       =   1
      TickFrequency   =   10
      Value           =   10
   End
   Begin VB.OptionButton optZoom 
      Caption         =   "25%"
      Height          =   225
      Index           =   2
      Left            =   13620
      TabIndex        =   16
      Top             =   9570
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.OptionButton optZoom 
      Caption         =   "50%"
      Height          =   225
      Index           =   1
      Left            =   13620
      TabIndex        =   15
      Top             =   9300
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.OptionButton optZoom 
      Caption         =   "100%"
      Height          =   225
      Index           =   0
      Left            =   13620
      TabIndex        =   14
      Top             =   9060
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox PB 
      Height          =   9450
      Left            =   0
      ScaleHeight     =   165.629
      ScaleMode       =   0  'User
      ScaleWidth      =   252.717
      TabIndex        =   13
      Top             =   0
      Width           =   11685
      Begin VB.Image ImgOrig 
         Height          =   3555
         Left            =   0
         Top             =   -30
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Image Img 
         Height          =   3555
         Left            =   0
         Top             =   0
         Width           =   6045
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   500
      Left            =   90
      Max             =   1000
      SmallChange     =   100
      TabIndex        =   12
      Top             =   9480
      Width           =   11625
   End
   Begin VB.CommandButton cmdRecover 
      Caption         =   "Recover"
      Height          =   405
      Left            =   13110
      TabIndex        =   11
      Top             =   6420
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox lstDeletions 
      Height          =   1035
      Left            =   13080
      TabIndex        =   9
      Top             =   5370
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   1035
      Left            =   13110
      Picture         =   "frmViewScan.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Remove Selected Scan"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9645
      LargeChange     =   500
      Left            =   11760
      Max             =   5000
      SmallChange     =   100
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1035
      Left            =   13110
      Picture         =   "frmViewScan.frx":116E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7395
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   1065
      Left            =   13080
      Picture         =   "frmViewScan.frx":2038
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8550
      Width           =   1395
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   13080
      TabIndex        =   2
      Top             =   1650
      Width           =   1395
   End
   Begin VB.TextBox txtSampleID 
      Enabled         =   0   'False
      Height          =   315
      Left            =   13080
      TabIndex        =   0
      Top             =   480
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   7
      Top             =   9780
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Default"
      Height          =   195
      Index           =   1
      Left            =   12330
      TabIndex        =   24
      Top             =   2580
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   12780
      TabIndex        =   23
      Top             =   2880
      Width           =   150
   End
   Begin VB.Label lblQuality 
      Alignment       =   2  'Center
      Caption         =   "50"
      Height          =   315
      Left            =   13080
      TabIndex        =   18
      Top             =   2730
      Width           =   1440
   End
   Begin VB.Label lblViewDeletions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on list to view deletions"
      Height          =   525
      Left            =   13080
      TabIndex        =   10
      Top             =   4590
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image imgViewDeletions 
      Height          =   240
      Index           =   1
      Left            =   13650
      Picture         =   "frmViewScan.frx":2F02
      Top             =   5130
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   13650
      Picture         =   "frmViewScan.frx":348C
      Top             =   1410
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on list to view report"
      Height          =   525
      Index           =   0
      Left            =   13080
      TabIndex        =   4
      Top             =   870
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Index           =   0
      Left            =   13410
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmViewScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_sCallerDepartment As String
Private m_SampleID As String

Private Sub cmdCancel_Click()

47480     Unload Me

End Sub

Private Sub cmdGrowShrink_Click()

47490 On Error GoTo cmdGrowShrink_Click_Error

47500 Select Case cmdGrowShrink.Caption
          Case ">>"
47510         frmViewScan.width = 14685
47520         cmdGrowShrink.Caption = "<<"
47530     Case "<<"
47540         frmViewScan.width = 12420
47550         cmdGrowShrink.Caption = ">>"
47560 End Select
47570 Exit Sub

cmdGrowShrink_Click_Error:

       Dim strES As String
       Dim intEL As Integer

47580  intEL = Erl
47590  strES = Err.Description
47600  LogError "frmViewScan", "cmdGrowShrink_Click", intEL, strES
          
End Sub

Private Sub cmdPrint_Click()
47610  AutoRedraw = -1
47620 Call PrintPictureToFitPage(ImgOrig.Picture)
47630 Exit Sub
      '40        Printer.Print ;
      '50        Printer.PaintPicture PB.Picture, 0, 0
      '60        Printer.EndDoc

End Sub



Private Sub cmdRecover_Click()

          Dim sql As String

47640     On Error GoTo cmdRecover_Click_Error

47650     txtSampleID = Replace(txtSampleID, "/", "-")

47660     sql = "UPDATE ScannedImages " & _
                "SET RemoveFromLisDisplay = 0 WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND ScannedName = '" & lstDeletions & "'"
47670     Cnxn(0).Execute sql

47680     ListScannedNameForSampleId

47690     Exit Sub

cmdRecover_Click_Error:

          Dim strES As String
          Dim intEL As Integer

47700     intEL = Erl
47710     strES = Err.Description
47720     LogError "frmViewScan", "cmdRecover_Click", intEL, strES, sql

End Sub

Private Sub cmdRemove_Click()

          Dim sql As String

47730     On Error GoTo cmdRemove_Click_Error

47740     txtSampleID = Replace(txtSampleID, "/", "-")

47750     sql = "UPDATE ScannedImages " & _
                "SET RemoveFromLisDisplay = 1 WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND ScannedName = '" & List1 & "'"
47760     Cnxn(0).Execute sql

47770     pb.Picture = Nothing

47780     ListScannedNameForSampleId

47790     Exit Sub

cmdRemove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

47800     intEL = Erl
47810     strES = Err.Description
47820     LogError "frmViewScan", "cmdRemove_Click", intEL, strES, sql

End Sub

Private Sub cmdZoom_Click(Index As Integer)
47830     If Index = 0 Then
47840         sldZoom.Value = sldZoom.Value - 10
47850     Else
47860         sldZoom.Value = sldZoom.Value + 10
47870     End If
47880     ZoomPicture
End Sub

Private Sub Form_Activate()

47890 On Error GoTo Form_Activate_Error


47900 If Trim$(txtSampleID) = "" Then
47910     iMsg "Sample ID?", vbExclamation
47920     Exit Sub
47930 End If
      '    frmViewScan.Width = 14685
47940 cmdGrowShrink.Caption = "<<"
47950 cmdGrowShrink_Click
      'sldZoom.Value = 25
47960 lblQuality.Caption = "Zoom : " & sldZoom.Value & " %"
47970 ListScannedNameForSampleId

47980 cmdPrint.Enabled = False
47990 List1.Selected(0) = True
48000 LoadScannedImage
48010 DoEvents
48020 DoEvents
48030 Me.Caption = "NetAcquire - View Scanned Reports (" & List1.ListCount & ")"
48040 UpdateScanViewLog
48050 Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

48060 intEL = Erl
48070 strES = Err.Description
48080 LogError "frmViewScan", "Form_Activate", intEL, strES

End Sub

Private Sub ListScannedNameForSampleId()

          Dim sql As String
          Dim tb As Recordset

48090     On Error GoTo ListScannedNameForSampleId_Error

48100     txtSampleID = Replace(txtSampleID, "/", "-")
48110     List1.Clear

48120     sql = "SELECT ScannedName FROM ScannedImages WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND RemoveFromLisDisplay = 0"
48130     Set tb = New Recordset
48140     RecOpenServer 0, tb, sql

48150     Do While Not tb.EOF
48160         List1.AddItem tb!ScannedName & ""
48170         tb.MoveNext
48180     Loop
48190     DoEvents
48200     DoEvents
48210     lblViewDeletions.Visible = False
48220     imgViewDeletions(1).Visible = False
48230     lstDeletions.Visible = False
48240     cmdRecover.Visible = False
48250     If UserMemberOf = "Managers" Then
48260         lstDeletions.Clear
48270         sql = "SELECT ScannedName FROM ScannedImages WHERE " & _
                    "SampleID = '" & txtSampleID & "' " & _
                    "AND RemoveFromLisDisplay = 1"
48280         Set tb = New Recordset
48290         RecOpenServer 0, tb, sql
48300         If Not tb.EOF Then
48310             lblViewDeletions.Visible = True
48320             imgViewDeletions(1).Visible = True
48330             lstDeletions.Visible = True
48340             cmdRecover.Visible = True
48350             Do While Not tb.EOF
48360                 lstDeletions.AddItem tb!ScannedName & ""
48370                 tb.MoveNext
48380             Loop
48390         End If
48400     End If

48410     Exit Sub

ListScannedNameForSampleId_Error:

          Dim strES As String
          Dim intEL As Integer

48420     intEL = Erl
48430     strES = Err.Description
48440     LogError "frmViewScan", "ListScannedNameForSampleId", intEL, strES, sql

End Sub




Private Sub Form_Load()

48450 On Error GoTo Form_Load_Error
48460 Me.Move frmEditAll.Left + (frmEditAll.width / 3), frmEditAll.Top + (frmEditAll.height / 4)
48470 sldZoom.Value = GetOptionSetting("DocumentViewerDefaultZoom", "40")
48480 txtDefaultZoom.Text = sldZoom.Value
48490 Exit Sub

Form_Load_Error:

       Dim strES As String
       Dim intEL As Integer

48500  intEL = Erl
48510  strES = Err.Description
48520  LogError "frmViewScan", "Form_Load", intEL, strES
          
End Sub

Private Sub UpdateScanViewLog()

      Dim sql As String

48530 On Error GoTo UpdateScanViewLog_Error

48540 sql = "IF NOT EXISTS(SELECT 1 FROM ScanViewLog WHERE SampleID = '" & SampleID & "' AND Department = '" & CallerDepartment & "' ) " & _
          "INSERT INTO ScanViewLog (SampleID, ScanName, Department, Viewed, Username, DateTimeOfRecord) " & _
          "VALUES ('" & SampleID & "', '" & List1 & "', '" & CallerDepartment & "', 1, '" & UserName & "', '" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "') "
48550 Cnxn(0).Execute sql

48560 Exit Sub

UpdateScanViewLog_Error:

       Dim strES As String
       Dim intEL As Integer

48570  intEL = Erl
48580  strES = Err.Description
48590  LogError "frmViewScan", "UpdateScanViewLog", intEL, strES, sql
          
End Sub

Private Sub HScroll1_Change()

48600     Img.Left = -HScroll1

End Sub

Private Sub HScroll1_Scroll()

48610     Img.Left = -HScroll1

End Sub

Private Function LoadPictureFromDB(tb As ADODB.Recordset)

48620     On Error GoTo procNoPicture

          'If Recordset is Empty, Then Exit
          Dim strStream As ADODB.Stream

48630     Set strStream = New ADODB.Stream
48640     strStream.Type = adTypeBinary
48650     strStream.Open

48660     strStream.Write tb!ScannedImage.Value


48670     strStream.SaveToFile List1, adSaveCreateOverWrite
48680     strStream.Close
48690     Set strStream = Nothing
          'Image1.Picture = LoadPicture("C:\Temp.bmp")
          'Kill ("C:\Temp.bmp")
          'PB.Picture = LoadPicture(List1)
48700     LoadPictureFromDB = True

procExitFunction:
48710     Exit Function
procNoPicture:
48720     LoadPictureFromDB = False
48730     GoTo procExitFunction
End Function

Private Sub Img2_Click()

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

         
48740     On Error GoTo List1_MouseUp_Error
          
48750     LoadScannedImage

48760     Exit Sub

List1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

48770     intEL = Erl
48780     strES = Err.Description
48790     LogError "frmViewScan", "List1_MouseUp", intEL, strES

End Sub


Private Sub optZoom_Click(Index As Integer)
48800     Select Case Index
          Case 0:
48810         Img.Stretch = False
48820     Case 1:
48830         Img.Stretch = True
48840         Img.width = ImgOrig.width * 0.5
48850         Img.height = ImgOrig.height * 0.5
48860     Case 2:
48870         Img.Stretch = True
48880         Img.width = ImgOrig.width * 0.25
48890         Img.height = ImgOrig.height * 0.25

48900     End Select
48910     Img.Picture = ImgOrig.Picture
End Sub

Private Sub sldZoom_Click()
48920     ZoomPicture
End Sub
Private Sub ZoomPicture()
48930     On Error Resume Next
48940     If sldZoom.Value = 100 Then
48950         Img.Stretch = False
48960     Else
48970         Img.Stretch = True
48980         Img.width = ImgOrig.width * sldZoom.Value / 100
48990         Img.height = ImgOrig.height * sldZoom.Value / 100
49000     End If
49010     Img.Picture = ImgOrig.Picture

49020     lblQuality.Caption = "Zoom : " & sldZoom.Value & " %"
49030     If Img.width <= pb.ScaleWidth And Img.height <= pb.ScaleHeight Then
49040         VScroll1.Visible = False
49050         HScroll1.Visible = False
49060     Else
49070         With VScroll1
49080             .Min = 1
49090             .max = Round(Img.height - pb.ScaleHeight)
49100             .SmallChange = IIf(Round(.max / 5) < 0, 1, Round(.max / 5))
49110             .LargeChange = IIf(Round(.max / 5) < 0, 1, Round(.max / 5))
                  
                  
49120             .Visible = True
                      
49130         End With
49140         With HScroll1
49150             .Min = 1
49160             .max = Round(Img.width - pb.ScaleWidth)
49170             .SmallChange = IIf(Round(.max / 5) < 0, 1, Round(.max / 5))
49180             .LargeChange = IIf(Round(.max / 5) < 0, 1, Round(.max / 5))
49190             .Visible = True
49200         End With
              
49210     End If
End Sub



Private Sub txtDefaultZoom_KeyPress(KeyAscii As Integer)

49220 On Error GoTo txtDefaultZoom_KeyPress_Error

49230 KeyAscii = VI(KeyAscii, Numeric_Only)

49240 Exit Sub

txtDefaultZoom_KeyPress_Error:

       Dim strES As String
       Dim intEL As Integer

49250  intEL = Erl
49260  strES = Err.Description
49270  LogError "frmViewScan", "txtDefaultZoom_KeyPress", intEL, strES
          
End Sub

Private Sub txtDefaultZoom_LostFocus()

49280 On Error GoTo txtDefaultZoom_LostFocus_Error

49290 If Val(txtDefaultZoom) > sldZoom.max Then
49300     txtDefaultZoom = sldZoom.max
49310 ElseIf Val(txtDefaultZoom) < sldZoom.Min Then
49320     txtDefaultZoom = sldZoom.Min
49330 End If
49340 SaveOptionSetting "DocumentViewerDefaultZoom", Val(txtDefaultZoom)

49350 Exit Sub

txtDefaultZoom_LostFocus_Error:

       Dim strES As String
       Dim intEL As Integer

49360  intEL = Erl
49370  strES = Err.Description
49380  LogError "frmViewScan", "txtDefaultZoom_LostFocus", intEL, strES
          
End Sub

Private Sub VScroll1_Change()
49390     Img.Top = -VScroll1
End Sub


Private Sub VScroll1_Scroll()
49400     Img.Top = -VScroll1
End Sub

Private Sub LoadScannedImage()

      Dim FilePath As String
      Dim sql As String
      Dim tb As Recordset
      Dim src() As Byte
      Dim a() As Byte
      Dim f As Integer


49410 On Error GoTo LoadScannedImage_Error

49420 If List1.SelCount = 0 Then Exit Sub

49430 Img.Picture = LoadPicture("")
49440 ImgOrig.Picture = LoadPicture("")
49450 FilePath = GetOptionSetting("ScanPath", "")

49460 If UCase(FilePath) = "APPPATH" Then
49470       FilePath = App.Path & "\"
49480 End If

49490 sql = "SELECT ScannedImage FROM ScannedImages WHERE " & _
            "ScannedName = '" & List1 & "' AND SampleID = " & txtSampleID
49500 Set tb = New Recordset
49510 RecOpenServer 0, tb, sql
49520 If Not tb.EOF Then
49530     LoadPictureFromDB tb
49540     src = tb!ScannedImage.GetChunk(100000000#)
49550 End If

      'a = Decompress(src)
49560 f = FreeFile
49570 Open FilePath & List1 For Binary Access Write As f
49580 Put f, , src
49590 Close f
49600 Img.Picture = LoadPicture(FilePath & List1)
49610 ImgOrig.Picture = LoadPicture(FilePath & List1)
49620 If UserMemberOf = "MANAGERS" Then
49630     cmdRemove.Visible = True
49640 End If

49650 Kill FilePath & List1

49660 ZoomPicture
49670 cmdPrint.Enabled = True

49680 Exit Sub

LoadScannedImage_Error:

      Dim strES As String
      Dim intEL As Integer

49690 intEL = Erl
49700 strES = Err.Description
49710 LogError "frmViewScan", "LoadScannedImage", intEL, strES, sql

End Sub


Public Property Get CallerDepartment() As String

49720     CallerDepartment = m_sCallerDepartment

End Property

Public Property Let CallerDepartment(ByVal sCallerDepartment As String)

49730     m_sCallerDepartment = sCallerDepartment

End Property

Public Property Get SampleID() As String

49740     SampleID = m_SampleID

End Property

Public Property Let SampleID(ByVal sSampleID As String)

49750     m_SampleID = sSampleID

End Property
