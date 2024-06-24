VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form freaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfusion Reaction Investigation"
   ClientHeight    =   7185
   ClientLeft      =   1215
   ClientTop       =   570
   ClientWidth     =   10710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FREACTIO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7185
   ScaleWidth      =   10710
   Begin VB.TextBox txtResultComment 
      Height          =   2775
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Top             =   4065
      Width           =   5115
   End
   Begin VB.TextBox cUnit 
      Height          =   285
      Left            =   5700
      TabIndex        =   39
      Top             =   150
      Width           =   2700
   End
   Begin VB.CommandButton bprint 
      Caption         =   "&Print"
      Height          =   1155
      Left            =   9180
      Picture         =   "FREACTIO.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   315
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   1155
      Left            =   9180
      Picture         =   "FREACTIO.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4395
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1155
      Left            =   9180
      Picture         =   "FREACTIO.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5715
      Width           =   1245
   End
   Begin VB.TextBox txtUrineHb 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   33
      Top             =   3540
      Width           =   1275
   End
   Begin VB.TextBox txtXMPost 
      Height          =   285
      Left            =   4740
      MaxLength       =   10
      TabIndex        =   31
      Top             =   3240
      Width           =   1275
   End
   Begin VB.TextBox txtXMPre 
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   30
      Top             =   3240
      Width           =   1275
   End
   Begin VB.TextBox txtDATUnit 
      Height          =   285
      Left            =   6315
      MaxLength       =   16
      TabIndex        =   29
      Top             =   2940
      Width           =   1935
   End
   Begin VB.TextBox txtDATPost 
      Height          =   285
      Left            =   4740
      MaxLength       =   10
      TabIndex        =   28
      Top             =   2940
      Width           =   1275
   End
   Begin VB.TextBox txtDATPre 
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   27
      Top             =   2940
      Width           =   1275
   End
   Begin VB.TextBox txtAntiUnit 
      Height          =   285
      Left            =   6315
      MaxLength       =   16
      TabIndex        =   26
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtAntiPost 
      Height          =   285
      Left            =   4740
      MaxLength       =   10
      TabIndex        =   25
      Top             =   2640
      Width           =   1275
   End
   Begin VB.TextBox txtAntiPre 
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   24
      Top             =   2640
      Width           =   1275
   End
   Begin VB.TextBox txtABOUnit 
      Height          =   285
      Left            =   6315
      MaxLength       =   16
      TabIndex        =   23
      Top             =   2340
      Width           =   1935
   End
   Begin VB.TextBox txtABOPost 
      Height          =   285
      Left            =   4740
      MaxLength       =   10
      TabIndex        =   22
      Top             =   2340
      Width           =   1275
   End
   Begin VB.TextBox txtABOPre 
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   21
      Top             =   2340
      Width           =   1275
   End
   Begin VB.TextBox txtLabUnit 
      Height          =   285
      Left            =   6315
      MaxLength       =   16
      TabIndex        =   20
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtLabPost 
      Height          =   285
      Left            =   4740
      MaxLength       =   10
      TabIndex        =   19
      Top             =   2040
      Width           =   1275
   End
   Begin VB.TextBox txtLabPre 
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   18
      Top             =   2040
      Width           =   1275
   End
   Begin VB.TextBox txtWard 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   315
      Left            =   900
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txtAddr 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   900
      MaxLength       =   20
      TabIndex        =   16
      Top             =   780
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   900
      MaxLength       =   20
      TabIndex        =   15
      Top             =   450
      Width           =   3615
   End
   Begin VB.TextBox txtChart 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   14
      Top             =   165
      Width           =   1635
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   45
      TabIndex        =   38
      Top             =   7005
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   315
      Left            =   915
      TabIndex        =   40
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   125829121
      CurrentDate     =   36963
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Result Comment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1665
      TabIndex        =   41
      Top             =   4095
      Width           =   1155
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Urine Uroblinogen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   32
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6720
      TabIndex        =   13
      Top             =   1800
      Width           =   285
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Post Sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   12
      Top             =   1800
      Width           =   885
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Pre Sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   11
      Top             =   1800
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   0
      X2              =   8955
      Y1              =   1620
      Y2              =   1605
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Unit No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5055
      TabIndex        =   10
      Top             =   180
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "X-Match"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2310
      TabIndex        =   9
      Top             =   3300
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "D.A.T."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2445
      TabIndex        =   8
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Antibody Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1740
      TabIndex        =   7
      Top             =   2700
      Width           =   1170
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ABO/Rh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2295
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Lab No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2340
      TabIndex        =   5
      Top             =   2100
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   4
      Top             =   1140
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2460
      TabIndex        =   3
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   840
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   540
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   240
      Width           =   345
   End
End
Attribute VB_Name = "freaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bprint_Click()

10    On Error GoTo bprint_Click_Error

20    Printer.FontName = "Courier New"
30    Printer.Orientation = vbPRORPortrait

40    PrintHeadingCavan

50    TextOut vbCrLf
60    TextOut FormatString("Date : ", 8, , Alignleft), 10
70    TextOut FormatString(dtDate, 20, , Alignleft), 10, True
80    TextOut vbCrLf

90    TextOut String$(247, "-"), 4
100   TextOut vbCrLf

110   TextOut Space(10)
120   TextOut FormatString("", 35, , AlignRight)
130   TextOut FormatString("Pre Sample", 15, , AlignCenter)
140   TextOut FormatString("Post Sample", 15, , AlignCenter)
150   TextOut FormatString("Unit", 15, , AlignCenter)
160   TextOut vbCrLf

170   TextOut String$(247, "-"), 4
180   TextOut vbCrLf

190   TextOut Space(10)
200   TextOut FormatString("Lab No. ", 35, "|", AlignRight)
210   TextOut FormatString(txtLabPre, 15, , AlignCenter), , True
220   TextOut FormatString(txtLabPost, 15, , AlignCenter), , True
230   TextOut FormatString(txtLabUnit, 15, , AlignCenter), , True
240   TextOut vbCrLf

250   TextOut Space(10)
260   TextOut FormatString("ABO/Rh ", 35, "|", AlignRight)
270   TextOut FormatString(txtABOPre, 15, , AlignCenter), , True
280   TextOut FormatString(txtABOPost, 15, , AlignCenter), , True
290   TextOut FormatString(txtABOUnit, 15, , AlignCenter), , True
300   TextOut vbCrLf

310   TextOut Space(10)
320   TextOut FormatString("Antibody Screen ", 35, "|", AlignRight)
330   TextOut FormatString(txtAntiPre, 15, , AlignCenter), , True
340   TextOut FormatString(txtAntiPost, 15, , AlignCenter), , True
350   TextOut FormatString(txtAntiUnit, 15, , AlignCenter), , True
360   TextOut vbCrLf

370   TextOut Space(10)
380   TextOut FormatString("D.A.T. ", 35, "|", AlignRight)
390   TextOut FormatString(txtDATPre, 15, , AlignCenter), , True
400   TextOut FormatString(txtDATPost, 15, , AlignCenter), , True
410   TextOut FormatString(txtDATUnit, 15, , AlignCenter), , True
420   TextOut vbCrLf

430   TextOut Space(10)
440   TextOut FormatString("X-Match ", 35, "|", AlignRight)
450   TextOut FormatString(txtXMPre, 15, , AlignCenter), , True
460   TextOut FormatString(txtXMPost, 15, , AlignCenter), , True
470   TextOut vbCrLf

480   TextOut Space(10)
490   TextOut FormatString("Urine Uroblinogen ", 35, "|", AlignRight)
500   TextOut FormatString(txtUrineHb, 15, , AlignCenter), , True
510   TextOut vbCrLf

520   TextOut FormatString("Result Comment ", 15, "|", Alignleft)
530   TextOut FormatString(txtResultComment, 200, , Alignleft), , True
540   TextOut vbCrLf

550   Do While Printer.CurrentY < 6700
560        Printer.Print
570   Loop

580   Printer.ForeColor = vbRed
590   Printer.Font.Size = 4
600   Printer.Print String(250, "-")

610   Printer.Font.Size = 10
620   Printer.Font.Bold = False
630   Printer.Print " Report Date:"; Format(Now, "dd/mm/yyyy hh:mm");
640   Printer.Print Tab(60); "Printed By "; TechnicianNameForCode(UserCode & "")

650   Printer.EndDoc

660   Exit Sub

bprint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

670   intEL = Erl
680   strES = Err.Description
690   LogError "freaction", "bprint_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    If Trim$(cUnit) = "" Then
30      iMsg "Unit No. must be specified.", vbCritical
40      If TimedOut Then Unload Me: Exit Sub
50      Exit Sub
60    End If

70    sql = "Select * from BadReact where " & _
            "UnitNo = '" & cUnit & "'"
80    Set tb = New Recordset
90    RecOpenServerBB 0, tb, sql

100   If tb.EOF Then
110     tb.AddNew
120   End If
130   With tb
140     !patno = txtChart
150     !UnitNo = cUnit
160     !Name = txtName
170     !addr = txtAddr
180     !Ward = txtWard
190     !labpre = txtLabPre
200     !labpost = txtLabPost
210     !labunit = txtLabUnit
220     !abopre = txtABOPre
230     !abopost = txtABOPost
240     !abounit = txtABOUnit
250     !antipre = txtAntiPre
260     !antipost = txtAntiPost
270     !antiunit = txtAntiUnit
280     !datpre = txtDATPre
290     !datpost = txtDATPost
300     !datunit = txtDATUnit
310     !xmpre = txtXMPre
320     !xmpost = txtXMPost
330     !urinehb = txtUrineHb
340     !ResultComment = txtResultComment
350     !DateTime = Format(dtDate, "dd/MMM/yyyy HH:mm:ss")
360     .Update
370   End With
380   cmdSave.Enabled = False
390   iMsg "Unit " & cUnit & " updated!"

400   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "freaction", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub SearchBadReact()


      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SearchBadReact_Error

20    If Trim$(cUnit) = "" Then Exit Sub

          '1.   Check if  unit logged as Bad Reaction previously
          '2.   If not THEN
                '2a.    Check if units exists and get latest patient details for unit and fill in demo details on screen.
          '3.   ELSE load previouly reaction results

        
30    sql = "Select * from BadReact where " & _
            "UnitNo = '" & cUnit & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql

60    If tb.EOF Then
        'unit NOT logged as Bad Reaction previously
        'get patient details for that unit from Product history
        
70      If fillPatientDetails Then
80        txtLabPre = ""
90        txtLabPost = ""
100       txtLabUnit = ""
110       txtABOPre = ""
120       txtABOPost = ""
130       txtABOUnit = ""
140       txtAntiPre = ""
150       txtAntiPost = ""
160       txtAntiUnit = ""
170       txtDATPre = ""
180       txtDATPost = ""
190       txtDATUnit = ""
200       txtXMPre = ""
210       txtXMPost = ""
220       txtUrineHb = ""
230       txtResultComment = ""
240     Else
                  
250       iMsg "Unit Number not found.", vbInformation
260       If TimedOut Then Unload Me: Exit Sub
270       txtChart = ""
280       cUnit = ""
290       txtName = ""
300       txtAddr = ""
310       txtWard = ""
320       txtLabPre = ""
330       txtLabPost = ""
340       txtLabUnit = ""
350       txtABOPre = ""
360       txtABOPost = ""
370       txtABOUnit = ""
380       txtAntiPre = ""
390       txtAntiPost = ""
400       txtAntiUnit = ""
410       txtDATPre = ""
420       txtDATPost = ""
430       txtDATUnit = ""
440       txtXMPre = ""
450       txtXMPost = ""
460       txtUrineHb = ""
470       txtResultComment = ""
480   End If

490   Else
500     With tb
510       txtChart = !patno & ""
520       cUnit = !UnitNo & ""
530       txtName = !Name & ""
540       txtAddr = !addr & ""
550       txtWard = !Ward & ""
560       txtLabPre = !labpre & ""
570       txtLabPost = !labpost & ""
580       txtLabUnit = !labunit & ""
590       txtABOPre = !abopre & ""
600       txtABOPost = !abopost & ""
610       txtABOUnit = !abounit & ""
620       txtAntiPre = !antipre & ""
630       txtAntiPost = !antipost & ""
640       txtAntiUnit = !antiunit & ""
650       txtDATPre = !datpre & ""
660       txtDATPost = !datpost & ""
670       txtDATUnit = !datunit & ""
680       txtXMPre = !xmpre & ""
690       txtXMPost = !xmpost & ""
700       txtUrineHb = !urinehb & ""
          
710       If IsNull(!ResultComment) Then 'If null load preious fields in txtResultComment
720           txtResultComment = !WCADate & "" & " " & !wcares & "" & vbCrLf & !BactDate & "" & " " & !bactres & "" & vbCrLf & !typeofreact & ""
730       Else
740           txtResultComment = !ResultComment & ""
750       End If
          
760       dtDate = !DateTime & ""
770     End With
780     cmdSave.Enabled = False
790   End If

800   Exit Sub

SearchBadReact_Error:

      Dim strES As String
      Dim intEL As Integer

810   intEL = Erl
820   strES = Err.Description
830   LogError "freaction", "SearchBadReact", intEL, strES, sql


End Sub

Private Function fillPatientDetails() As Boolean
      Dim tb As Recordset
      Dim pd As Recordset
      Dim sql As String
      Dim strAddr As String

10       On Error GoTo fillPatientDetails_Error

20    sql = "Select Top 1 * from Product where " & _
            "ISBT128 = '" & cUnit & "' and PatName <> '' and labnumber <> '' order by datetime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60          sql = "Select PD.patnum, PD.name, PD.Addr1, PD.addr2, PD.addr3, PD.addr4, PD.Ward from PatientDetails as PD " & _
            "where PD.LabNumber = '" & tb!LabNumber & "' "
70          Set pd = New Recordset
80          RecOpenServerBB 0, pd, sql
90          If Not pd.EOF Then
100             fillPatientDetails = True
110             txtChart = pd!patnum & ""
120             txtName = pd!Name & ""
130             If Len(pd!Addr1 & "") > 0 Then strAddr = Trim$(pd!Addr1)
140             If Len(pd!Addr2 & "") > 0 Then strAddr = strAddr & " " & Trim$(pd!Addr2 & "")
150             If Len(pd!Addr3 & "") > 0 Then strAddr = strAddr & " " & Trim$(pd!Addr3 & "")
160             If Len(pd!Addr4 & "") > 0 Then strAddr = strAddr & " " & Trim$(pd!Addr4 & "")
170             txtAddr = strAddr
180             txtWard = pd!Ward & ""
190             dtDate = Format(Now, "dd/mmm/yyyy")
200         Else
210             fillPatientDetails = False
220         End If
230   End If

240   Exit Function

fillPatientDetails_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "freaction", "fillPatientDetails", intEL, strES, sql

End Function

Private Sub cUnit_LostFocus()
10    SearchBadReact
End Sub



Private Sub dtDate_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub dtDate_Click()
10    cmdSave.Enabled = True
End Sub

Private Sub Form_Load()

10    dtDate = Format(Now, "dd/mmm/yyyy")

End Sub


Private Sub txtABOPost_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtABOPre_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtABOUnit_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtAntiPost_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtAntiPre_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtAntiUnit_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtDATPost_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtDATPre_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtDATUnit_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtLabPost_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtLabPre_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtLabUnit_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtResultComment_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtUrineHb_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtXMPost_Change()
10    cmdSave.Enabled = True
End Sub

Private Sub txtXMPre_Change()
10    cmdSave.Enabled = True
End Sub


Public Sub PrintHeadingCavan()

      Dim sql As String
      Dim Heading As String
      
10       On Error GoTo PrintHeadingCavan_Error

20    Heading = "Transfusion Reaction Investigation"
30    Printer.Font.Name = "Courier New"
40    Printer.Font.Size = 14
50    Printer.Font.Bold = True

60    Printer.ForeColor = vbRed
70    Printer.Print "CAVAN GENERAL HOSPITAL : " & Heading;
80    Printer.Font.Size = 10
90    Printer.CurrentY = 100
100   Printer.Print '" Phone 38833"

110   Printer.CurrentY = 320

120   Printer.Font.Size = 4
130   Printer.Print String$(250, "-")

140   Printer.ForeColor = vbBlack

150   Printer.Font.Name = "Courier New"
160   Printer.Font.Size = 12
170   Printer.Font.Bold = False

180   Printer.Print " Unit No:";
190   Printer.Print cUnit;
  
200   Printer.Print Tab(35); "Name:";
210   Printer.Font.Bold = True
220   Printer.Font.Size = 14
230   Printer.Print txtName
240   Printer.Font.Size = 12
250   Printer.Font.Bold = False
  
260   Printer.Print " Ward:";
270   Printer.Print txtWard & "";
  
280   Printer.Print Tab(35);
290   Printer.Print ; "Chart #:";
300   Printer.Print txtChart
 
310   Printer.Print " Addr:";
320   Printer.Print txtAddr & ""
  
330   Printer.Font.Bold = False

340   Printer.Font.Size = 4
350   Printer.Print String$(250, "-")



360   Exit Sub

PrintHeadingCavan_Error:

Dim strES As String
Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "freaction", "PrintHeadingCavan", intEL, strES, sql

End Sub

