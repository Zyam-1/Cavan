VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmSuggestFromStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1125
      HelpContextID   =   10090
      Left            =   8295
      Picture         =   "frmSuggestFromStock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   1125
      HelpContextID   =   10080
      Left            =   8295
      Picture         =   "frmSuggestFromStock.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1830
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid gSuggest 
      Height          =   3495
      Left            =   150
      TabIndex        =   0
      Top             =   1830
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   3
      FormatString    =   "<Unit Number              |<ABO Rh          |<Kell                                |<Expiry                     "
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   5340
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblKell 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4020
      TabIndex        =   15
      Top             =   1320
      Width           =   2760
   End
   Begin VB.Label lblGroup 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   885
      TabIndex        =   14
      Top             =   1320
      Width           =   2760
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1995
      TabIndex        =   10
      Top             =   870
      Width           =   1470
   End
   Begin VB.Label lblSex 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5835
      TabIndex        =   9
      Top             =   510
      Width           =   960
   End
   Begin VB.Label lblAge 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3945
      TabIndex        =   8
      Top             =   510
      Width           =   1470
   End
   Begin VB.Label lblDoB 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1995
      TabIndex        =   7
      Top             =   510
      Width           =   1470
   End
   Begin VB.Label lblPatName 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1995
      TabIndex        =   6
      Top             =   150
      Width           =   3420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lab Number"
      Height          =   195
      Left            =   1005
      TabIndex        =   5
      Top             =   930
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   5505
      TabIndex        =   4
      Top             =   555
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Age"
      Height          =   195
      Left            =   3585
      TabIndex        =   3
      Top             =   555
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "D.o.B."
      Height          =   195
      Left            =   1425
      TabIndex        =   2
      Top             =   555
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Patient Name"
      Height          =   195
      Left            =   945
      TabIndex        =   1
      Top             =   210
      Width           =   960
   End
End
Attribute VB_Name = "frmSuggestFromStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pGroup As String

Private pNumberOfPacks As Integer


Private Sub FillG()

      Dim s() As String
      Dim n As Integer
      Dim sql As String
      Dim tb As Recordset

      'Female < 60 years and not K+
10    On Error GoTo FillG_Error

20    If lblSex = "Female" And lblKell <> "K+" And DateDiff("YYYY", lblDoB, Now) < 60 Then
30      GenerateSQLF60 s
40    Else
50      GenerateSQLOther s
60    End If

70    gSuggest.Rows = 2
80    gSuggest.AddItem ""
90    gSuggest.RemoveItem 1

100   For n = 0 To UBound(s)
110     sql = s(n)
120     Set tb = New Recordset
130     RecOpenServerBB 0, tb, sql
140     Do While Not tb.EOF
150       gSuggest.AddItem tb!ISBT128 & "" & vbTab & _
                           Bar2Group(tb!GroupRh) & vbTab & _
                           tb!Screen & vbTab & _
                           Format(tb!DateExpiry, "dd/mmm/yyyy hh:mm")
160       If gSuggest.Rows = pNumberOfPacks + 2 Then
170         Exit For
180       End If
190       tb.MoveNext
200     Loop
210   Next

220   If gSuggest.Rows > 2 Then
230     gSuggest.RemoveItem 1
240   End If

250   Exit Sub

FillG_Error:

Dim strES As String
Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmSuggestFromStock", "FillG", intEL, strES, sql

End Sub



Private Sub GenerateSQLF60(ByRef s() As String)

      Dim Gr As String

10    On Error GoTo GenerateSQLF60_Error

20    Gr = UCase$(Bar2Group(pGroup))

30    Select Case Gr
        Case "O POS"
40        ReDim s(0 To 1)
50        s(0) = GenerateSQLBody(pGroup, "-")
60        s(1) = GenerateSQLBody(Group2Bar("O Neg"), "-")
70      Case "O NEG"
80        ReDim s(0 To 0)
90        s(0) = GenerateSQLBody(pGroup, "-")
100     Case "A POS"
110       ReDim s(0 To 3)
120       s(0) = GenerateSQLBody(pGroup, "-")
130       s(1) = GenerateSQLBody(Group2Bar("O Pos"), "-")
140       s(2) = GenerateSQLBody(Group2Bar("A Neg"), "-")
150       s(3) = GenerateSQLBody(Group2Bar("O Neg"), "-")
160     Case "A NEG"
170       ReDim s(0 To 1)
180       s(0) = GenerateSQLBody(pGroup, "-")
190       s(1) = GenerateSQLBody(Group2Bar("O Neg"), "-")
200     Case "B POS"
210       ReDim s(0 To 3)
220       s(0) = GenerateSQLBody(pGroup, "-")
230       s(1) = GenerateSQLBody(Group2Bar("O Pos"), "-")
240       s(2) = GenerateSQLBody(Group2Bar("B Neg"), "-")
250       s(3) = GenerateSQLBody(Group2Bar("O Neg"), "-")
260     Case "B NEG"
270       ReDim s(0 To 1)
280       s(0) = GenerateSQLBody(pGroup, "-")
290       s(1) = GenerateSQLBody(Group2Bar("O Neg"), "-")
300     Case "AB POS"
310       ReDim s(0 To 7)
320       s(0) = GenerateSQLBody(pGroup, "-")
330       s(1) = GenerateSQLBody(Group2Bar("A Pos"), "-")
340       s(2) = GenerateSQLBody(Group2Bar("B Pos"), "-")
350       s(3) = GenerateSQLBody(Group2Bar("AB Neg"), "-")
360       s(4) = GenerateSQLBody(Group2Bar("A Neg"), "-")
370       s(5) = GenerateSQLBody(Group2Bar("B Neg"), "-")
380       s(6) = GenerateSQLBody(Group2Bar("O Pos"), "-")
390       s(7) = GenerateSQLBody(Group2Bar("O Neg"), "-")
400     Case "AB NEG"
410       ReDim s(0 To 2)
420       s(0) = GenerateSQLBody(pGroup, "-")
430       s(1) = GenerateSQLBody(Group2Bar("A Neg"), "-")
440       s(2) = GenerateSQLBody(Group2Bar("B Neg"), "-")
450   End Select

460   Exit Sub

GenerateSQLF60_Error:

Dim strES As String
Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "frmSuggestFromStock", "GenerateSQLF60", intEL, strES

End Sub

Private Function GenerateSQLBody(ByVal GRh As String, ByVal K As String) As String

      Dim s As String

10    On Error GoTo GenerateSQLBody_Error

20    s = "SELECT ISBT128, GroupRh, Screen, DateExpiry FROM Latest WHERE " & _
          "GroupRh = '" & GRh & "' " & _
          "AND Screen LIKE '%K" & K & "%' " & _
          "AND (Event = 'C' OR Event = 'R') " & _
          "AND DATEDIFF(minute, getdate(), DateExpiry) >= 0 " & _
          "ORDER BY DateExpiry asc"

30    GenerateSQLBody = s

40    Exit Function

GenerateSQLBody_Error:

Dim strES As String
Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSuggestFromStock", "GenerateSQLBody", intEL, strES

End Function

Private Sub GenerateSQLOther(ByRef s() As String)

      Dim Gr As String

10    On Error GoTo GenerateSQLOther_Error

20    Gr = UCase$(Bar2Group(pGroup))

30    Select Case Gr
        Case "O POS"
40        ReDim s(0 To 3)
50        s(0) = GenerateSQLBody(pGroup, "+")
60        s(1) = GenerateSQLBody(Group2Bar("O Pos"), "-")
70        s(2) = GenerateSQLBody(Group2Bar("O Neg"), "+")
80        s(3) = GenerateSQLBody(Group2Bar("O Neg"), "-")
90      Case "O NEG"
100       ReDim s(0 To 1)
110       s(0) = GenerateSQLBody(pGroup, "+")
120       s(1) = GenerateSQLBody(Group2Bar("O Neg"), "-")
130     Case "A POS"
140       ReDim s(0 To 7)
150       s(0) = GenerateSQLBody(pGroup, "+")
160       s(1) = GenerateSQLBody(Group2Bar("A Pos"), "-")
170       s(2) = GenerateSQLBody(Group2Bar("A Neg"), "+")
180       s(3) = GenerateSQLBody(Group2Bar("A Neg"), "-")
190       s(4) = GenerateSQLBody(Group2Bar("O Pos"), "+")
200       s(5) = GenerateSQLBody(Group2Bar("O Pos"), "-")
210       s(6) = GenerateSQLBody(Group2Bar("O Neg"), "+")
220       s(7) = GenerateSQLBody(Group2Bar("O Neg"), "-")
230     Case "A NEG"
240       ReDim s(0 To 3)
250       s(0) = GenerateSQLBody(pGroup, "+")
260       s(1) = GenerateSQLBody(Group2Bar("A Neg"), "-")
270       s(2) = GenerateSQLBody(Group2Bar("O Neg"), "+")
280       s(3) = GenerateSQLBody(Group2Bar("O Neg"), "-")
290     Case "B POS"
300       ReDim s(0 To 7)
310       s(0) = GenerateSQLBody(pGroup, "+")
320       s(1) = GenerateSQLBody(Group2Bar("B Pos"), "-")
330       s(2) = GenerateSQLBody(Group2Bar("B Neg"), "+")
340       s(3) = GenerateSQLBody(Group2Bar("B Neg"), "-")
350       s(4) = GenerateSQLBody(Group2Bar("O Pos"), "+")
360       s(5) = GenerateSQLBody(Group2Bar("O Pos"), "-")
370       s(6) = GenerateSQLBody(Group2Bar("O Neg"), "+")
380       s(7) = GenerateSQLBody(Group2Bar("O Neg"), "-")
390     Case "B NEG"
400       ReDim s(0 To 3)
410       s(0) = GenerateSQLBody(pGroup, "+")
420       s(1) = GenerateSQLBody(Group2Bar("B Neg"), "-")
430       s(2) = GenerateSQLBody(Group2Bar("O Neg"), "+")
440       s(3) = GenerateSQLBody(Group2Bar("O Neg"), "-")
450     Case "AB POS"
460       ReDim s(0 To 7)
470       s(0) = GenerateSQLBody(pGroup, "+")
480       s(1) = GenerateSQLBody(Group2Bar("AB Pos"), "-")
490       s(2) = GenerateSQLBody(Group2Bar("AB Neg"), "+")
500       s(3) = GenerateSQLBody(Group2Bar("AB Neg"), "-")
510       s(4) = GenerateSQLBody(Group2Bar("A Pos"), "+")
520       s(5) = GenerateSQLBody(Group2Bar("A Pos"), "-")
530       s(6) = GenerateSQLBody(Group2Bar("B Pos"), "+")
540       s(7) = GenerateSQLBody(Group2Bar("B Pos"), "-")
550     Case "AB NEG"
560       ReDim s(0 To 3)
570       s(0) = GenerateSQLBody(pGroup, "+")
580       s(1) = GenerateSQLBody(Group2Bar("AB Neg"), "-")
590       s(2) = GenerateSQLBody(Group2Bar("A Neg"), "+")
600       s(3) = GenerateSQLBody(Group2Bar("A Neg"), "-")

610   End Select

620   Exit Sub

GenerateSQLOther_Error:

Dim strES As String
Dim intEL As Integer

630   intEL = Erl
640   strES = Err.Description
650   LogError "frmSuggestFromStock", "GenerateSQLOther", intEL, strES

End Sub

Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    On Error GoTo cmdPrint_Click_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetFormPrinter() Then Exit Sub

40    PrintBlueHeading lblSampleID

50    Printer.Print

60    Printer.ForeColor = vbBlue

70    Printer.Font.Bold = True
80    Printer.Font.Size = 12
90    Printer.Print " Suggested Packs: "
100   Printer.Print
110   Printer.Print "Unit Number        Pack Group         Screen                 Expiry"
120   For Y = 1 To gSuggest.Rows - 1
130     Printer.Print gSuggest.TextMatrix(Y, 0); Tab(20);
140     Printer.Print gSuggest.TextMatrix(Y, 1); Tab(39);
150     Printer.Print gSuggest.TextMatrix(Y, 2); Tab(62);
160     Printer.Print gSuggest.TextMatrix(Y, 3)
170   Next
180   Printer.EndDoc

190   For Each Px In Printers
200     If Px.DeviceName = OriginalPrinter Then
210       Set Printer = Px
220       Exit For
230     End If
240   Next

250   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmSuggestFromStock", "cmdPrint_Click", intEL, strES

End Sub

Private Sub PrintBlueHeading(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintBlueHeading_Error

20    sql = "Select * from PatientDetails where " & _
            "LabNumber = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then Exit Sub
60    Printer.Font.Name = "Courier New"
70    Printer.Font.Size = 14
80    Printer.Font.Bold = True

90    Printer.ForeColor = vbBlue
100   Printer.Print "CAVAN GENERAL HOSPITAL : Blood Transfusion Laboratory"
110   Printer.Font.Size = 10
120   Printer.CurrentY = 100
130   Printer.Print '" Phone 38833"

140   Printer.CurrentY = 320

150   Printer.Font.Size = 4
160   Printer.Print String$(250, "-")

170   Printer.Font.Name = "Courier New"
180   Printer.Font.Size = 12
190   Printer.Font.Bold = False

200   Printer.Print " Sample ID:";
210   Printer.Print SampleID;
  
220   Printer.Print Tab(35); "Name:";
230   Printer.Font.Bold = True
240   Printer.Font.Size = 14
250   Printer.Print tb!Name
260   Printer.Font.Size = 12
270   Printer.Font.Bold = False
  
280   Printer.Print "      Ward:";
290   Printer.Print tb!Ward & "";
  
300   Printer.Print Tab(35); " DOB:";
310   Printer.Print Format(tb!DoB, "dd/mm/yyyy");
320   Printer.Print Tab(60); "Chart #:";
330   Printer.Print tb!Patnum
 
340   If Trim$(tb!Clinician & "") <> "" Then
350     Printer.Print "Consultant:";
360     Printer.Print tb!Clinician & "";
370   Else
380     Printer.Print "        GP:";
390     Printer.Print tb!GP & "";
400   End If

410   Printer.Print Tab(35); "Addr:";
420   Printer.Print tb!Addr1 & "";
430   Printer.Print Tab(60); "    Sex:";
440   Select Case Left$(UCase$(Trim$(tb!Sex & "")), 1)
        Case "M": Printer.Print "Male"
450     Case "F": Printer.Print "Female"
460     Case Else: Printer.Print
470   End Select
  
480   Printer.Font.Bold = False
490   Printer.Print Tab(35); "     ";
500   Printer.Print tb!Addr2 & ""

510   Printer.Font.Size = 4
520   Printer.Print String$(250, "-")

530   Exit Sub

PrintBlueHeading_Error:

Dim strES As String
Dim intEL As Integer

540   intEL = Erl
550   strES = Err.Description
560   LogError "frmSuggestFromStock", "PrintBlueHeading", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub Form_Activate()

10    FillG

End Sub

Public Property Let PatName(ByVal sNewValue As String)

10    lblPatName = sNewValue

End Property
Public Property Let DoB(ByVal sNewValue As String)

10    lblDoB = sNewValue

End Property

Public Property Let Age(ByVal sNewValue As String)

10    lblAge = sNewValue

End Property


Public Property Let NumberOfPacks(ByVal iNewValue As Integer)

10    pNumberOfPacks = iNewValue

End Property
Public Property Let Group(ByVal sNewValue As String)

10    Select Case UCase$(sNewValue)
        Case "O POS": lblGroup = "O Rh Positive"
20      Case "O NEG": lblGroup = "O Rh Negative"
30      Case "A POS": lblGroup = "A Rh Positive"
40      Case "A NEG": lblGroup = "A Rh Negative"
50      Case "B POS": lblGroup = "B Rh Positive"
60      Case "B NEG": lblGroup = "B Rh Negative"
70      Case "AB POS": lblGroup = "AB Rh Positive"
80      Case "AB NEG": lblGroup = "AB Rh Negative"
90    End Select

100   pGroup = Group2Bar(sNewValue)

End Property

Public Property Let Kell(ByVal sNewValue As String)

10    lblKell = sNewValue

End Property

Public Property Let Sex(ByVal sNewValue As String)

10    Select Case UCase$(Left$(sNewValue & " ", 1))
        Case "M": lblSex = "Male"
20      Case "F": lblSex = "Female"
30      Case Else: lblSex = "Unknown"
40    End Select

End Property



Public Property Let SampleID(ByVal sNewValue As String)

10    lblSampleID = sNewValue

End Property




