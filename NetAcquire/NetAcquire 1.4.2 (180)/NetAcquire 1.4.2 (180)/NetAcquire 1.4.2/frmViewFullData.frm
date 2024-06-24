VERSION 5.00
Begin VB.Form frmViewFullData 
   Caption         =   "NetAcquire - View Full HbA1c Data"
   ClientHeight    =   3195
   ClientLeft      =   3540
   ClientTop       =   4005
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4935
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete this Sample"
      Height          =   495
      Left            =   3810
      TabIndex        =   2
      Top             =   2610
      Width           =   975
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   1830
      Picture         =   "frmViewFullData.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2610
      Width           =   975
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      DrawWidth       =   2
      Height          =   2490
      Left            =   90
      ScaleHeight     =   2430
      ScaleWidth      =   4650
      TabIndex        =   0
      Top             =   60
      Width           =   4710
   End
End
Attribute VB_Name = "frmViewFullData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String

Private Activated As Boolean

'Private Sub DisplayGraph()
''
''Dim Block1 As String
''Dim Block2 As String
''Dim Block3 As String
''
''Block1 = STX & _
''         "NO.0001 BC.---------0001 92-07-29" & vbCr & vbLf & _
''         "16:57 HbA1c05.2 HbA108.3 HbF00.0 *" & vbCr & vbLf & ETB
''
''Block2 = STX & _
''         "0'35#000465 01.6 0'57#000141 00.5" & vbCr & vbLf & _
''         "1'03#000292 01.0 2'03S001460 05.2" & vbCr & vbLf & _
''         "3'06#025979 91.7 -'--#------ --.-" & vbCr & vbLf & _
''         "-'--#------ --.- -'--#------ --.-" & vbCr & vbLf & _
''         "-'--#------ --.- -'--#------ --.-" & vbCr & vbLf & _
''         "-'--#------ --.-" & vbCr & vbLf & ETB
''
''Block3 = STX & "020,020,019,020,019,019,019,019,019,019,019,019,019,019,019,019,019," & _
''               "019,020,021,026,043,063,069,067,066,064,062,060,058,055,053,053,053," & _
''               "051,049,053,057,058,056" & vbCr & vbLf & ETB & _
''         STX & "054,055,056,065,053,051,048,045,043,042,040,039,039,038,037,037,037," & _
''               "037,037,037,037,038,039,040,041,043,045,046,047,047,048,048,048,050," & _
''               "052,055,060,067,074,081" & vbCr & vbLf & ETB & _
''         STX & "086,089,090,089,085,081,075,070,064,059,055,050,047,044,041,039,038," & _
''               "036,035,034,033,033,032,032,032,031,031,031,031,032,032,032,033,033," & _
''               "034,035,037,043,053,104" & vbCr & vbLf & ETB & _
''         STX & "188,188,188,188,188,188,188,188,188,188,188,188,188,165,150,133,122," & _
''               "112,112,112,112,097,097,097,097,097,084,084,072,072,065,065,065,056," & _
''               "040,040,040,030,025,023,022,021,021,020,020,020,020,019,019,019,019," & _
''               "019,019,019,020,020,020,020,020,020" & vbCr & vbLf & ETB & _
''         STX & "0001,0046,018,035,040,054,107,---,---,---,---,---,---,018,107,---,---," & _
''               "---,---,---,---,---,---,054,107" & vbCr & vbLf & ETX
'
'FillGraphData Block3
'
'End Sub
Private Sub DrawGraph()

      Dim n As Integer
      Dim Position As Integer
      Dim gdArray(0 To 179) As Integer
      Dim HbStart As Integer
      Dim HbEnd As Integer
      Dim EachX As Single
      Dim EachY As Single
      Dim max As Integer
      Dim HbPeakX As Integer
      Dim HbPeakY As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim Block3 As String

33100 On Error GoTo DrawGraph_Error

33110 sql = "Select * from HbA1c where " & _
            "SampleID = '" & mSampleID & "'"
33120 Set tb = New Recordset
33130 RecOpenClient 0, tb, sql

33140 If tb.EOF Then Exit Sub

33150 Block3 = Mid$(tb!block & "", 269)

33160 Position = 0
33170 For n = 2 To 158 Step 4
33180   gdArray(Position) = Val(Mid$(Block3, n, 3))
33190   Position = Position + 1
33200 Next
33210 For n = 165 To 321 Step 4
33220   gdArray(Position) = Val(Mid$(Block3, n, 3))
33230   Position = Position + 1
33240 Next
33250 For n = 328 To 477 Step 4
33260   gdArray(Position) = Val(Mid$(Block3, n, 3))
33270   Position = Position + 1
33280 Next
33290 For n = 491 To 727 Step 4
33300   gdArray(Position) = Val(Mid$(Block3, n, 3))
33310   Position = Position + 1
33320 Next

33330 max = 0
33340 For n = 0 To 179
33350   If gdArray(n) > max Then
33360     max = gdArray(n)
33370   End If
33380 Next
33390 If max < 2 Then
33400   picGraph.Print "No Graph available."
33410   Exit Sub
33420 End If

33430 EachY = picGraph.height / max
33440 EachX = picGraph.width / 179

33450 HbStart = Val(Mid$(Block3, 828, 3))
33460 HbEnd = Val(Mid$(Block3, 832, 3))
33470 HbPeakX = ((HbStart + HbEnd) / 2) * EachX
33480 HbPeakY = picGraph.height - (IIf(gdArray(HbStart) > gdArray(HbEnd), gdArray(HbStart), gdArray(HbEnd)) * EachY)

33490 For n = 0 To 179
33500   If n = 0 Then
33510     picGraph.PSet (n * EachX, picGraph.height - gdArray(n) * EachY), vbBlack
33520   Else
33530     If n > HbStart And n < HbEnd Then
33540       picGraph.Line -(n * EachX, picGraph.height - gdArray(n) * EachY), vbYellow
33550     Else
33560       picGraph.Line -(n * EachX, picGraph.height - gdArray(n) * EachY), vbBlack
33570     End If
33580   End If
33590 Next
33600 picGraph.Line (HbStart * EachX, picGraph.height - gdArray(HbStart) * EachY)-(HbEnd * EachX, picGraph.height - gdArray(HbEnd) * EachY), vbYellow

33610 Exit Sub

DrawGraph_Error:

      Dim strES As String
      Dim intEL As Integer

33620 intEL = Erl
33630 strES = Err.Description
33640 LogError "fViewFullData", "DrawGraph", intEL, strES, sql


End Sub


Private Sub bcancel_Click()

33650 Unload Me

End Sub


Private Sub cmdDelete_Click()

      Dim s As String
      Dim sql As String
      Dim HbA1cCode As String

33660 On Error GoTo cmdDelete_Click_Error

33670 HbA1cCode = GetOptionSetting("BioCodeForHbA1c", "")
33680 If HbA1cCode = "" Then Exit Sub

33690 s = "Delete Sample " & mSampleID & " ?"
33700 If iMsg(s, vbQuestion + vbYesNo) = vbYes Then
33710   s = "You wont be able to undo this action." & vbCrLf & _
            "Are you sure?"
33720   If iMsg(s, vbQuestion + vbYesNo, "Confirm Deletion", vbRed, 14) = vbYes Then
          
33730     sql = "Delete from HbA1c where " & _
                "SampleID = '" & mSampleID & "'"
33740     Cnxn(0).Execute sql
          
33750     sql = "Delete from BioResults where " & _
                "SampleID = '" & mSampleID & "' " & _
                "and Code = '" & HbA1cCode & "'"
33760     Cnxn(0).Execute sql
          
33770     sql = "Delete from BioRepeats where " & _
                "SampleID = '" & mSampleID & "' " & _
                "and Code = '" & HbA1cCode & "'"
33780     Cnxn(0).Execute sql
          
33790     Unload Me
          
33800   End If
33810 End If

33820 Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

33830 intEL = Erl
33840 strES = Err.Description
33850 LogError "fViewFullData", "cmdDelete_Click", intEL, strES, sql


End Sub

Private Sub Form_Activate()

33860 If Activated Then
33870   Exit Sub
33880 End If
33890 Activated = True

33900 DrawGraph

End Sub


Public Property Let SampleID(ByVal sNewValue As String)

33910 mSampleID = sNewValue

End Property

Private Sub Form_Load()

33920 Activated = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

33930 Activated = False

End Sub

