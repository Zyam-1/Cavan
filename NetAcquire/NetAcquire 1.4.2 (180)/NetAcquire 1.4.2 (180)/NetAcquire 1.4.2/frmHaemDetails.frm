VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHaemDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Haematology"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnDo 
      Caption         =   "&Undo"
      Height          =   1155
      Left            =   6750
      Picture         =   "frmHaemDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1155
      Left            =   5520
      Picture         =   "frmHaemDetails.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame fraSelectPrint 
      BorderStyle     =   0  'None
      Caption         =   "Valid"
      Height          =   285
      Index           =   1
      Left            =   1710
      TabIndex        =   5
      Top             =   600
      Width           =   2805
      Begin VB.CommandButton cmdRedCross 
         Height          =   285
         Index           =   0
         Left            =   120
         Picture         =   "frmHaemDetails.frx":284C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdGreenTick 
         Height          =   285
         Index           =   0
         Left            =   420
         Picture         =   "frmHaemDetails.frx":2B22
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdGreenTick 
         Height          =   285
         Index           =   1
         Left            =   1440
         Picture         =   "frmHaemDetails.frx":2DF8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdRedCross 
         Height          =   285
         Index           =   1
         Left            =   1120
         Picture         =   "frmHaemDetails.frx":30CE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdGreenTick 
         Height          =   285
         Index           =   2
         Left            =   2430
         Picture         =   "frmHaemDetails.frx":33A4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdRedCross 
         Height          =   285
         Index           =   2
         Left            =   2120
         Picture         =   "frmHaemDetails.frx":367A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   315
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   8055
      Left            =   150
      TabIndex        =   4
      Top             =   180
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   14208
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedRows       =   2
      RowHeightMin    =   330
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      ScrollBars      =   2
      FormatString    =   "<Test   |<Result   |^Validate   |^Authorise |^Release    |<Code"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2265
      Left            =   5130
      TabIndex        =   3
      Top             =   180
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3995
      _Version        =   393217
      TextRTF         =   $"frmHaemDetails.frx":3950
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1155
      Left            =   8790
      Picture         =   "frmHaemDetails.frx":39DB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6420
      Width           =   1035
   End
   Begin VB.Label lblName 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   24
      Top             =   5040
      Width           =   3315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   6060
      TabIndex        =   23
      Top             =   5040
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   5910
      TabIndex        =   22
      Top             =   2640
      Width           =   570
   End
   Begin VB.Label lblHospital 
      BackColor       =   &H80000016&
      Caption         =   "CAVAN GENERAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      Top             =   2640
      Width           =   3315
   End
   Begin VB.Label lblDoB 
      BackColor       =   &H80000016&
      Caption         =   "88/88/8888"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   20
      Top             =   3030
      Width           =   3315
   End
   Begin VB.Label lblSex 
      BackColor       =   &H80000016&
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   19
      Top             =   4200
      Width           =   3315
   End
   Begin VB.Label lblChart 
      BackColor       =   &H80000016&
      Caption         =   "123456789"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   18
      Top             =   3420
      Width           =   3315
   End
   Begin VB.Label lblSampleDateTime 
      BackColor       =   &H80000016&
      Caption         =   "88/88/88 88:88"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      Top             =   3810
      Width           =   3315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sample Date/Time"
      Height          =   195
      Left            =   5145
      TabIndex        =   16
      Top             =   3810
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Index           =   0
      Left            =   6210
      TabIndex        =   14
      Top             =   4200
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   6105
      TabIndex        =   13
      Top             =   3420
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   6165
      TabIndex        =   12
      Top             =   3030
      Width           =   315
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmHaemDetails.frx":48A5
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmHaemDetails.frx":4B7B
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000016&
      Caption         =   "123456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   2
      Top             =   4590
      Width           =   3315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   5745
      TabIndex        =   1
      Top             =   4590
      Width           =   735
   End
End
Attribute VB_Name = "frmHaemDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCode As String
Private pSampleID As String
Private Sub FillG()
          
          Dim sql As String
          Dim tb As Recordset
          Dim SelectColour As String
          Dim s As String
          Dim DoB As String
          Dim DaysOld As Long

2590      On Error GoTo FillG_Error

2600      If Trim$(lblSampleID) = "" Then Exit Sub

2610      grd.Visible = False
2620      grd.Rows = 3
2630      grd.AddItem ""
2640      grd.RemoveItem 2

2650      Select Case Left$(UCase$(Trim$(lblSex)), 1)
              Case "M":
2660              SelectColour = "CASE WHEN (H.Result < MaleLow) THEN 16711680 " & _
                      "     WHEN (H.Result > MaleHigh) THEN 255 " & _
                      "     ELSE -2147483624 END BackColour, " & _
                      "CASE WHEN (H.Result < MaleLow) THEN 16777215  " & _
                      "     WHEN (H.Result > MaleHigh) THEN  65535 " & _
                      "     ELSE 0 END ForeColour "
2670          Case "F":
2680              SelectColour = "CASE WHEN (H.Result < FemaleLow) THEN 16711680 " & _
                      "     WHEN (H.Result > FemaleHigh) THEN 255 " & _
                      "     ELSE -2147483624 END BackColour, " & _
                      "CASE WHEN (H.Result < FemaleLow) THEN 16777215  " & _
                      "     WHEN (H.Result > FemaleHigh) THEN  65535 " & _
                      "     ELSE 0 END ForeColour "
2690          Case Else:
2700              SelectColour = "CASE WHEN (H.Result < FemaleLow) THEN 16711680 " & _
                      "     WHEN (H.Result > MaleHigh) THEN 255 " & _
                      "     ELSE -2147483624 END BackColour, " & _
                      "CASE WHEN (H.Result < FemaleLow) THEN 16777215  " & _
                      "     WHEN (H.Result > MaleHigh) THEN  65535 " & _
                      "     ELSE 0 END ForeColour "
2710      End Select

2720      DoB = lblDoB
2730      If IsDate(DoB) Then
2740          DaysOld = DateDiff("d", DoB, lblSampleDateTime)
2750      Else
2760          DaysOld = 365 * 20
2770      End If

2780      sql = "SELECT DISTINCT H.SampleID, UPPER(H.Code) Code, D.ShortName, H.Result, " & _
              "H.Analyser, COALESCE(D.PlausibleLow, 0) PlausibleLow, " & _
              "COALESCE(D.PlausibleHigh, 9999) PlausibleHigh, " & _
              SelectColour & ", " & _
              "D.PrintPriority, " & _
              "CASE Valid WHEN 1 THEN 'V' ELSE '' END V, " & _
              "CASE Authorised WHEN 1 THEN 'A' ELSE '' END A, " & _
              "CASE Released WHEN 1 THEN 'R' ELSE '' END R " & _
              "FROM Haem50Results H INNER JOIN HaemTestDefinitions D " & _
              "ON H.Code = D.AnalyteName " & _
              "WHERE H.SampleID = '" & lblSampleID & "' " & _
              "AND D.AgeFromDays <= " & DaysOld & " " & _
              "AND D.AgeToDays >= " & DaysOld & " " & _
              "ORDER BY PrintPriority"

2790      Set tb = New Recordset
2800      RecOpenServer 0, tb, sql
2810      Do While Not tb.EOF
2820          s = tb!ShortName & vbTab & _
                  tb!Result & vbTab & vbTab & vbTab & vbTab & tb!Code & ""
2830          grd.AddItem s
2840          grd.row = grd.Rows - 1
2850          grd.Col = 1
2860          grd.CellBackColor = tb!BackColour
2870          grd.CellForeColor = tb!ForeColour
        
2880          SetgPicture 2, "V", tb!v
2890          SetgPicture 3, "A", tb!a
2900          SetgPicture 4, "R", tb!R
        
2910          tb.MoveNext
2920      Loop

2930      If grd.Rows > 3 Then
2940          grd.RemoveItem 2
2950      End If
2960      grd.Visible = True

2970      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

2980      intEL = Erl
2990      strES = Err.Description
3000      LogError "frmHaemDetails", "FillG", intEL, strES, sql
3010      grd.Visible = True

End Sub
Private Sub FillVPARFE(ByVal VPARFE As String, _
          ByVal Base As String, _
          ByVal BaseBy As String, _
          ByVal BaseTime As String)
3020      With rtb

3030          .SelFontName = "Courier New"
3040          .SelFontSize = 10
3050          .SelColor = vbBlack
        
3060          If Base = 1 Then
3070              .SelText = VPARFE & " by "
          
3080              If BaseBy <> "" Then
3090                  .SelText = BaseBy & " "
3100              Else
3110                  .SelText = "Unknown "
3120              End If
          
3130              If IsDate(BaseTime) Then
3140                  .SelText = "on " & Format$(BaseTime, "dd/MM/yyyy at HH:mm") & vbCrLf
3150              Else
3160                  .SelText = " (Time unknown)" & vbCrLf
3170              End If
          
3180          Else
3190              .SelText = "Not " & VPARFE & vbCrLf
3200          End If

3210      End With

End Sub

Private Sub SetgPicture(ByVal gCol As Integer, _
          ByVal colTitle As String, _
          ByVal colValue As String)
        
3220      grd.Col = gCol
3230      If colValue = colTitle Then
3240          Set grd.CellPicture = imgGreenTick.Picture
3250      Else
3260          Set grd.CellPicture = imgRedCross.Picture
3270      End If
3280      grd.CellPictureAlignment = flexAlignCenterCenter

End Sub

Private Sub LoadCode()

          Dim tb As Recordset
          Dim sql As String

3290      On Error GoTo LoadCode_Error

3300      rtb.TextRTF = ""

3310      sql = "SELECT DISTINCT H.SampleID, UPPER(H.Code) Code, D.ShortName, H.Result, D.DoDelta, D.DeltaValue, H.Units, " & _
              "H.Analyser, D.PrintPriority, " & _
              "Valid, ValidBy, ValidTime, " & _
              "Printed, PrintedBy, PrintedTime, " & _
              "Authorised, AuthorisedBy, AuthorisedTime, " & _
              "Released, ReleasedBy, ReleasedTime, " & _
              "Faxed, FaxedBy, FaxedTime, " & _
              "Emailed, EmailedBy, EmailedTime " & _
              "FROM Haem50Results H INNER JOIN HaemTestDefinitions D " & _
              "ON H.Code = D.AnalyteName " & _
              "WHERE H.SampleID = '" & pSampleID & "' " & _
              "AND Code = '" & pCode & "'"
3320      Set tb = New Recordset
3330      RecOpenServer 0, tb, sql
3340      If Not tb.EOF Then
         
3350          rtb.SelFontName = "Courier New"
3360          rtb.SelFontSize = 16
3370          rtb.SelBold = True
3380          rtb.SelColor = vbBlue
        
3390          rtb.SelText = tb!ShortName & " "
3400          rtb.SelText = tb!Result & vbCrLf

3410          FillVPARFE "Validated", tb!Valid, tb!ValidBy & "", tb!ValidTime & ""
3420          FillVPARFE "Printed", tb!Printed, tb!PrintedBy & "", tb!PrintedTime & ""
3430          FillVPARFE "Authorised", tb!Authorised, tb!AuthorisedBy & "", tb!AuthorisedTime & ""
3440          FillVPARFE "Released", tb!Released, tb!ReleasedBy & "", tb!ReleasedTime & ""
3450          FillVPARFE "Faxed", tb!FAXed, tb!FaxedBy & "", tb!FaxedTime & ""
3460          FillVPARFE "eMailed", tb!Emailed, tb!EmailedBy & "", tb!EmailedTime & ""

3470      End If

3480      Exit Sub

LoadCode_Error:

          Dim strES As String
          Dim intEL As Integer

3490      intEL = Erl
3500      strES = Err.Description
3510      LogError "frmHaemDetails", "LoadCode", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

3520      Unload Me

End Sub


Public Property Let SampleID(ByVal sNewValue As String)

3530      pSampleID = sNewValue

3540      lblSampleID = pSampleID

End Property
Public Property Let Code(ByVal sNewValue As String)

3550      pCode = sNewValue

End Property

Private Sub cmdGreenTick_Click(Index As Integer)

          Dim Y As Integer

3560      With grd
3570          .Col = Index + 2
3580          For Y = 2 To .Rows - 1
3590              .row = Y
3600              If .CellPicture = imgRedCross.Picture Then
3610                  Set .CellPicture = imgGreenTick.Picture
3620                  If grd.TextMatrix(grd.row, grd.Col) = "_" Then 'Flag to mark it as changed
3630                      grd.TextMatrix(grd.row, grd.Col) = ""
3640                  Else
3650                      grd.TextMatrix(grd.row, grd.Col) = "_"
3660                  End If
3670                  cmdSave.Visible = True
3680                  cmdUnDo.Visible = True
3690                  cmdCancel.Visible = False
3700              End If
3710          Next
3720      End With

End Sub

Private Sub cmdRedCross_Click(Index As Integer)

          Dim Y As Integer

3730      With grd
3740          .Col = Index + 2
3750          For Y = 2 To .Rows - 1
3760              .row = Y
3770              If .CellPicture = imgGreenTick.Picture Then
3780                  Set .CellPicture = imgRedCross.Picture
3790                  If grd.TextMatrix(grd.row, grd.Col) = "_" Then 'Flag to mark it as changed
3800                      grd.TextMatrix(grd.row, grd.Col) = ""
3810                  Else
3820                      grd.TextMatrix(grd.row, grd.Col) = "_"
3830                  End If
3840                  cmdSave.Visible = True
3850                  cmdUnDo.Visible = True
3860                  cmdCancel.Visible = False
3870              End If
3880          Next
3890      End With

End Sub


Private Sub cmdSave_Click()

          Dim Y As Integer
          Dim sql As String

          'Validate
3900      On Error GoTo cmdSave_Click_Error

3910      grd.Col = 2
3920      For Y = 2 To grd.Rows - 1
3930          If grd.TextMatrix(Y, 2) = "_" Then
3940              grd.row = Y
3950              If grd.CellPicture = imgGreenTick.Picture Then
3960                  sql = "UPDATE Haem50Results " & _
                          "SET Valid = 1, " & _
                          "ValidBy = '" & UserName & "', " & _
                          "ValidTime = getdate() " & _
                          "WHERE SampleID = '" & lblSampleID & "' " & _
                          "AND Code = '" & grd.TextMatrix(Y, 5) & "'"
3970              Else
3980                  sql = "UPDATE Haem50Results " & _
                          "SET Valid = 0, " & _
                          "ValidBy = '', " & _
                          "ValidTime = NULL " & _
                          "WHERE SampleID = '" & lblSampleID & "' " & _
                          "AND Code = '" & grd.TextMatrix(Y, 5) & "'"
3990              End If
4000              Cnxn(0).Execute sql
4010          End If
4020      Next

          'Authorise
4030      grd.Col = 3
4040      For Y = 2 To grd.Rows - 1
4050          If grd.TextMatrix(Y, 3) = "_" Then
4060              grd.row = Y
4070              If grd.CellPicture = imgGreenTick.Picture Then
4080                  sql = "UPDATE Haem50Results " & _
                          "SET Authorised = 1, " & _
                          "AuthorisedBy = '" & UserName & "', " & _
                          "AuthorisedTime = getdate() " & _
                          "WHERE SampleID = '" & lblSampleID & "' " & _
                          "AND Code = '" & grd.TextMatrix(Y, 5) & "'"
4090              Else
4100                  sql = "UPDATE Haem50Results " & _
                          "SET Authorised = 0, " & _
                          "AuthorisedBy = '', " & _
                          "AuthorisedTime = NULL " & _
                          "WHERE SampleID = '" & lblSampleID & "' " & _
                          "AND Code = '" & grd.TextMatrix(Y, 5) & "'"
4110              End If
4120              Cnxn(0).Execute sql
4130          End If
4140      Next

          'Released
4150      grd.Col = 4
4160      For Y = 2 To grd.Rows - 1
4170          If grd.TextMatrix(Y, 4) = "_" Then
4180              grd.row = Y
4190              If grd.CellPicture = imgGreenTick.Picture Then
4200                  sql = "UPDATE Haem50Results " & _
                          "SET Released = 1, " & _
                          "ReleasedBy = '" & UserName & "', " & _
                          "ReleasedTime = getdate() " & _
                          "WHERE SampleID = '" & lblSampleID & "' " & _
                          "AND Code = '" & grd.TextMatrix(Y, 5) & "'"
4210              Else
4220                  sql = "UPDATE Haem50Results " & _
                          "SET Released = 0, " & _
                          "ReleasedBy = '', " & _
                          "ReleasedTime = NULL " & _
                          "WHERE SampleID = '" & lblSampleID & "' " & _
                          "AND Code = '" & grd.TextMatrix(Y, 5) & "'"
4230              End If
4240              Cnxn(0).Execute sql
4250          End If
4260      Next

4270      cmdSave.Visible = False
4280      cmdUnDo.Visible = False
4290      cmdCancel.Visible = True

4300      Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

4310      intEL = Erl
4320      strES = Err.Description
4330      LogError "frmHaemDetails", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub cmdUnDo_Click()

          Dim X As Integer
          Dim Y As Integer

4340      For X = 2 To 4
4350          For Y = 2 To grd.Rows - 1
4360              If grd.TextMatrix(Y, X) = "_" Then
4370                  grd.Col = X
4380                  grd.row = Y
4390                  If grd.CellPicture = imgRedCross.Picture Then
4400                      Set grd.CellPicture = imgGreenTick.Picture
4410                  Else
4420                      Set grd.CellPicture = imgRedCross.Picture
4430                  End If
4440                  grd.TextMatrix(Y, X) = ""
4450              End If
4460          Next
4470      Next

4480      cmdSave.Visible = False
4490      cmdUnDo.Visible = False
4500      cmdCancel.Visible = True

End Sub

Private Sub Form_Activate()

4510      LoadCode
4520      FillG

End Sub

Private Sub Form_Load()

4530      grd.ColWidth(5) = 0

End Sub


Private Sub grd_Click()

4540      If grd.row < 2 Then Exit Sub

4550      Select Case grd.Col
        
              Case 0, 1:
4560              pCode = grd.TextMatrix(grd.row, 5)
4570              LoadCode
        
4580          Case Else:
4590              If grd.CellPicture = imgGreenTick.Picture Then
4600                  Set grd.CellPicture = imgRedCross.Picture
4610              Else
4620                  Set grd.CellPicture = imgGreenTick.Picture
4630              End If
          
                  'Flag to mark it as changed
4640              If grd.TextMatrix(grd.row, grd.Col) = "_" Then
4650                  grd.TextMatrix(grd.row, grd.Col) = ""
4660              Else
4670                  grd.TextMatrix(grd.row, grd.Col) = "_"
4680                  cmdSave.Visible = True
4690                  cmdUnDo.Visible = True
4700                  cmdCancel.Visible = False
4710              End If
          
4720      End Select

End Sub


