VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
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

10        Unload Me

End Sub

Private Sub cmdGrowShrink_Click()

10    On Error GoTo cmdGrowShrink_Click_Error

20    Select Case cmdGrowShrink.Caption
          Case ">>"
30            frmViewScan.Width = 14685
40            cmdGrowShrink.Caption = "<<"
50        Case "<<"
60            frmViewScan.Width = 12420
70            cmdGrowShrink.Caption = ">>"
80    End Select
90    Exit Sub

cmdGrowShrink_Click_Error:

       Dim strES As String
       Dim intEL As Integer

100    intEL = Erl
110    strES = Err.Description
120    LogError "frmViewScan", "cmdGrowShrink_Click", intEL, strES
          
End Sub

Private Sub cmdPrint_Click()
'10     AutoRedraw = -1
'20    Call PrintPictureToFitPage(ImgOrig.Picture)
'30    Exit Sub


End Sub



Private Sub cmdRecover_Click()

          Dim sql As String

10        On Error GoTo cmdRecover_Click_Error

20        txtSampleID = Replace(txtSampleID, "/", "-")

30        sql = "UPDATE ScannedImages " & _
                "SET RemoveFromLisDisplay = 0 WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND ScannedName = '" & lstDeletions & "'"
40        Cnxn(0).Execute sql

50        ListScannedNameForSampleId

60        Exit Sub

cmdRecover_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmViewScan", "cmdRecover_Click", intEL, strES, sql

End Sub

Private Sub cmdRemove_Click()

          Dim sql As String

10        On Error GoTo cmdRemove_Click_Error

20        txtSampleID = Replace(txtSampleID, "/", "-")

30        sql = "UPDATE ScannedImages " & _
                "SET RemoveFromLisDisplay = 1 WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND ScannedName = '" & List1 & "'"
40        Cnxn(0).Execute sql

50        PB.Picture = Nothing

60        ListScannedNameForSampleId

70        Exit Sub

cmdRemove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmViewScan", "cmdRemove_Click", intEL, strES, sql

End Sub

Private Sub cmdZoom_Click(Index As Integer)
10        If Index = 0 Then
20            sldZoom.Value = sldZoom.Value - 10
30        Else
40            sldZoom.Value = sldZoom.Value + 10
50        End If
60        ZoomPicture
End Sub

Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error


20    If Trim$(txtSampleID) = "" Then
30        iMsg "Sample ID?", vbExclamation
40        Exit Sub
50    End If
      '    frmViewScan.Width = 14685
60    cmdGrowShrink.Caption = "<<"
70    cmdGrowShrink_Click
      'sldZoom.Value = 25
80    lblQuality.Caption = "Zoom : " & sldZoom.Value & " %"
90    ListScannedNameForSampleId

100   cmdPrint.Enabled = False
110   List1.Selected(0) = True
120   LoadScannedImage
130   Me.Caption = "NetAcquire - View Scanned Reports (" & List1.ListCount & ")"
140   UpdateScanViewLog
150   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmViewScan", "Form_Activate", intEL, strES

End Sub

Private Sub ListScannedNameForSampleId()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo ListScannedNameForSampleId_Error

20        txtSampleID = Replace(txtSampleID, "/", "-")
30        List1.Clear

40        sql = "SELECT ScannedName FROM ScannedImages WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND RemoveFromLisDisplay = 0"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        Do While Not tb.EOF
80            List1.AddItem tb!ScannedName & ""
90            tb.MoveNext
100       Loop

110       lblViewDeletions.Visible = False
120       imgViewDeletions(1).Visible = False
130       lstDeletions.Visible = False
140       cmdRecover.Visible = False
150       If UserMemberOf = "Managers" Then
160           lstDeletions.Clear
170           sql = "SELECT ScannedName FROM ScannedImages WHERE " & _
                    "SampleID = '" & txtSampleID & "' " & _
                    "AND RemoveFromLisDisplay = 1"
180           Set tb = New Recordset
190           RecOpenServer 0, tb, sql
200           If Not tb.EOF Then
210               lblViewDeletions.Visible = True
220               imgViewDeletions(1).Visible = True
230               lstDeletions.Visible = True
240               cmdRecover.Visible = True
250               Do While Not tb.EOF
260                   lstDeletions.AddItem tb!ScannedName & ""
270                   tb.MoveNext
280               Loop
290           End If
300       End If

310       Exit Sub

ListScannedNameForSampleId_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmViewScan", "ListScannedNameForSampleId", intEL, strES, sql

End Sub




Private Sub Form_Load()

10    On Error GoTo Form_Load_Error
20    Me.Move frmViewResultsWE.Left + (frmViewResultsWE.Width / 3), frmViewResultsWE.Top + (frmViewResultsWE.Height / 4)
30    sldZoom.Value = GetOptionSetting("DocumentViewerDefaultZoom", "40", "")
40    txtDefaultZoom.Text = sldZoom.Value
50    Exit Sub

Form_Load_Error:

       Dim strES As String
       Dim intEL As Integer

60     intEL = Erl
70     strES = Err.Description
80     LogError "frmViewScan", "Form_Load", intEL, strES
          
End Sub

Private Sub UpdateScanViewLog()

      Dim sql As String

10    On Error GoTo UpdateScanViewLog_Error

20    sql = "IF NOT EXISTS(SELECT 1 FROM ScanViewLog WHERE SampleID = '" & SampleID & "' AND Department = '" & CallerDepartment & "' ) " & _
          "INSERT INTO ScanViewLog (SampleID, ScanName, Department, Viewed, Username, DateTimeOfRecord) " & _
          "VALUES ('" & SampleID & "', '" & List1 & "', '" & CallerDepartment & "', 1, '" & UserName & "', '" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "') "
30    Cnxn(0).Execute sql

40    Exit Sub

UpdateScanViewLog_Error:

       Dim strES As String
       Dim intEL As Integer

50     intEL = Erl
60     strES = Err.Description
70     LogError "frmViewScan", "UpdateScanViewLog", intEL, strES, sql
          
End Sub

Private Sub HScroll1_Change()

10        Img.Left = -HScroll1

End Sub

Private Sub HScroll1_Scroll()

10        Img.Left = -HScroll1

End Sub

Private Function LoadPictureFromDB(tb As ADODB.Recordset)

10        On Error GoTo procNoPicture

          'If Recordset is Empty, Then Exit
          Dim strStream As ADODB.Stream

20        Set strStream = New ADODB.Stream
30        strStream.Type = adTypeBinary
40        strStream.Open

50        strStream.Write tb!ScannedImage.Value


60        strStream.SaveToFile List1, adSaveCreateOverWrite
70        strStream.Close
80        Set strStream = Nothing
          'Image1.Picture = LoadPicture("C:\Temp.bmp")
          'Kill ("C:\Temp.bmp")
          'PB.Picture = LoadPicture(List1)
90        LoadPictureFromDB = True

procExitFunction:
100       Exit Function
procNoPicture:
110       LoadPictureFromDB = False
120       GoTo procExitFunction
End Function

Private Sub Img2_Click()

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

         
10        On Error GoTo List1_MouseUp_Error
          
20        LoadScannedImage

30        Exit Sub

List1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewScan", "List1_MouseUp", intEL, strES

End Sub


Private Sub optZoom_Click(Index As Integer)
10        Select Case Index
          Case 0:
20            Img.Stretch = False
30        Case 1:
40            Img.Stretch = True
50            Img.Width = ImgOrig.Width * 0.5
60            Img.Height = ImgOrig.Height * 0.5
70        Case 2:
80            Img.Stretch = True
90            Img.Width = ImgOrig.Width * 0.25
100           Img.Height = ImgOrig.Height * 0.25

110       End Select
120       Img.Picture = ImgOrig.Picture
End Sub

Private Sub sldZoom_Click()
10        ZoomPicture
End Sub
Private Sub ZoomPicture()
10        If sldZoom.Value = 100 Then
20            Img.Stretch = False
30        Else
40            Img.Stretch = True
50            Img.Width = ImgOrig.Width * sldZoom.Value / 100
60            Img.Height = ImgOrig.Height * sldZoom.Value / 100
70        End If
80        Img.Picture = ImgOrig.Picture

90        lblQuality.Caption = "Zoom : " & sldZoom.Value & " %"
100       If Img.Width <= PB.ScaleWidth And Img.Height <= PB.ScaleHeight Then
110           VScroll1.Visible = False
120           HScroll1.Visible = False
130       Else
140           With VScroll1
150               .Min = 1
160               .Max = Round(Img.Height - PB.ScaleHeight)
170               .SmallChange = IIf(Round(.Max / 5) < 0, 1, Round(.Max / 5))
180               .LargeChange = IIf(Round(.Max / 5) < 0, 1, Round(.Max / 5))
                  
                  
190               .Visible = True
                      
200           End With
210           With HScroll1
220               .Min = 1
230               .Max = Round(Img.Width - PB.ScaleWidth)
240               .SmallChange = IIf(Round(.Max / 5) < 0, 1, Round(.Max / 5))
250               .LargeChange = IIf(Round(.Max / 5) < 0, 1, Round(.Max / 5))
260               .Visible = True
270           End With
              
280       End If
End Sub



Private Sub txtDefaultZoom_KeyPress(KeyAscii As Integer)

10    On Error GoTo txtDefaultZoom_KeyPress_Error

20    KeyAscii = VI(KeyAscii, Numeric_Only)

30    Exit Sub

txtDefaultZoom_KeyPress_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmViewScan", "txtDefaultZoom_KeyPress", intEL, strES
          
End Sub

Private Sub txtDefaultZoom_LostFocus()

10    On Error GoTo txtDefaultZoom_LostFocus_Error

20    If Val(txtDefaultZoom) > sldZoom.Max Then
30        txtDefaultZoom = sldZoom.Max
40    ElseIf Val(txtDefaultZoom) < sldZoom.Min Then
50        txtDefaultZoom = sldZoom.Min
60    End If
70    SaveOptionSetting "DocumentViewerDefaultZoom", Val(txtDefaultZoom), ""

80    Exit Sub

txtDefaultZoom_LostFocus_Error:

       Dim strES As String
       Dim intEL As Integer

90     intEL = Erl
100    strES = Err.Description
110    LogError "frmViewScan", "txtDefaultZoom_LostFocus", intEL, strES
          
End Sub

Private Sub VScroll1_Change()
10        Img.Top = -VScroll1
End Sub


Private Sub VScroll1_Scroll()
10        Img.Top = -VScroll1
End Sub

Private Sub LoadScannedImage()

      Dim FilePath As String
      Dim sql As String
      Dim tb As Recordset
      Dim src() As Byte
      Dim a() As Byte
      Dim f As Integer


10    On Error GoTo LoadScannedImage_Error

20    If List1.SelCount = 0 Then Exit Sub

30    Img.Picture = LoadPicture("")
40    ImgOrig.Picture = LoadPicture("")
50    FilePath = GetOptionSetting("ScanPath", "", "")

52    If UCase(FilePath) = "APPPATH" Then
54          FilePath = App.Path & "\"
56    End If
      
60    sql = "SELECT ScannedImage FROM ScannedImages WHERE " & _
            "ScannedName = '" & List1 & "' AND SampleID = " & txtSampleID
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If Not tb.EOF Then
100       LoadPictureFromDB tb
110       src = tb!ScannedImage.GetChunk(100000000#)
120   End If

      'a = Decompress(src)
130   f = FreeFile
140   Open FilePath & List1 For Binary Access Write As f
150   Put f, , src
160   Close f
170   Img.Picture = LoadPicture(FilePath & List1)
180   ImgOrig.Picture = LoadPicture(FilePath & List1)
190   If UserMemberOf = "MANAGERS" Then
200       cmdRemove.Visible = True
210   End If

220   Kill FilePath & List1

230   ZoomPicture
240   cmdPrint.Enabled = True

250   Exit Sub

LoadScannedImage_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmViewScan", "LoadScannedImage", intEL, strES, sql

End Sub


Public Property Get CallerDepartment() As String

10        CallerDepartment = m_sCallerDepartment

End Property

Public Property Let CallerDepartment(ByVal sCallerDepartment As String)

10        m_sCallerDepartment = sCallerDepartment

End Property

Public Property Get SampleID() As String

10        SampleID = m_SampleID

End Property

Public Property Let SampleID(ByVal sSampleID As String)

10        m_SampleID = sSampleID

End Property
