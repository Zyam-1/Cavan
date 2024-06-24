VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMicroUnusedSampleIDs 
   Caption         =   "NetAcquire - Unused Sample ID's"
   ClientHeight    =   6570
   ClientLeft      =   540
   ClientTop       =   675
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   6090
   Begin VB.TextBox txtReport 
      Height          =   5025
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1350
      Width           =   5715
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between"
      Height          =   1125
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   3315
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1020
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optDates 
         Alignment       =   1  'Right Justify
         Caption         =   "Dates"
         Height          =   225
         Left            =   780
         TabIndex        =   4
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSIDs 
         Caption         =   "Sample Numbers"
         Height          =   225
         Left            =   1530
         TabIndex        =   3
         Top             =   0
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1020
         TabIndex        =   7
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220266497
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220266497
         CurrentDate     =   38126
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.CommandButton breCalc 
      Caption         =   "Calculate"
      Height          =   825
      Left            =   3750
      Picture         =   "frmMicroUnusedSampleIDs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   5070
      Picture         =   "frmMicroUnusedSampleIDs.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   825
   End
End
Attribute VB_Name = "frmMicroUnusedSampleIDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calculate()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim lngFrom As Long
      Dim lngTo As Long
      Dim n As Integer
      Dim ThisSID As Long

55250 On Error GoTo Calculate_Error

55260 sql = "Select SampleID from Demographics where "
55270 If optDates Then
55280   If Abs(DateDiff("d", dtFrom, dtTo)) > 60 Then
55290     iMsg "Maximum 60 Days!", vbExclamation
55300     Exit Sub
55310   End If
55320   sql = sql & "Rundate between '" & Format(dtFrom, "dd/mmm/yyyy") & _
                    "' and '" & Format(dtTo, "dd/mmm/yyyy") & "' and sampleid > '" & sysOptMicroOffset(0) & "'"
55330 Else
55340   lngFrom = Val(txtFrom)
55350   lngTo = Val(txtTo)
55360   If lngTo < lngFrom Then
55370     txtFrom = Format(lngTo)
55380     txtTo = Format(lngFrom)
55390     lngFrom = Val(txtFrom)
55400     lngTo = Val(txtTo)
55410   End If
55420   If lngFrom < 1 Or lngFrom > 9999999 Then
55430     iMsg "Number <From> is incorrect!", vbExclamation
55440     txtFrom = ""
55450     Exit Sub
55460   End If
55470   If lngTo < 1 Or lngTo > 9999999 Then
55480     iMsg "Number <To> is incorrect!", vbExclamation
55490     txtTo = ""
55500     Exit Sub
55510   End If
55520   If lngTo - lngFrom > 5000 Then
55530     iMsg "Maximum 5000 Records!", vbExclamation
55540     Exit Sub
55550   End If
55560   sql = sql & "SampleID between '" & Format$(Val(txtFrom)) & "' " & _
                    " and '" & Format$(Val(txtTo)) & "' "
55570 End If
55580 sql = sql & "and (PatName = '' or PatName is null) " & _
                  "order by SampleID"

55590 Set tb = New Recordset
55600 RecOpenClient 0, tb, sql
55610 If tb.EOF Then
55620   s = "No Sample ID's found with Blank Patient Names."
55630 Else
55640   If tb.RecordCount = 1 Then
55650     s = "Sample ID " & tb!SampleID & " has " & _
              "no Patient Name associated."
55660   Else
55670     s = "The following " & tb.RecordCount & " Sample ID's have " & _
              "no Patient Names associated."
55680     Do While Not tb.EOF
55690       s = s & tb!SampleID & ",   "
55700       tb.MoveNext
55710     Loop
55720     s = Left$(s, Len(s) - 4)
55730   End If
55740 End If
55750 s = s & vbCrLf & vbCrLf

      'end of blank pat names

55760 s = s & "Unused Sample ID's :-" & vbCrLf & vbCrLf
55770 n = 0
55780 sql = "Select top 1 SampleID from Demographics where "
55790 If optDates Then
55800   sql = sql & "Rundate between '" & Format(dtFrom, "dd/mmm/yyyy") & _
                    "' and '" & Format(dtTo, "dd/mmm/yyyy") & "' and sampleid > '" & sysOptMicroOffset(0) & "'"
55810 Else
55820   sql = sql & "SampleID between '" & Format$(Val(txtFrom)) & "' " & _
                    " and '" & Format$(Val(txtTo)) & "' "
55830 End If
55840 sql = sql & "order by SampleID"

55850 Set tb = New Recordset
55860 RecOpenClient 0, tb, sql
55870 If Not tb.EOF Then
55880   lngFrom = tb!SampleID
55890   sql = "Select top 1 SampleID from Demographics where "
55900   If optDates Then
55910     sql = sql & "Rundate between '" & Format(dtFrom, "dd/mmm/yyyy") & _
                      "' and '" & Format(dtTo, "dd/mmm/yyyy") & "' and sampleid > '" & sysOptMicroOffset(0) & "'"
55920   Else
55930     sql = sql & "SampleID between '" & Format$(Val(txtFrom)) & "' " & _
                      " and '" & Format$(Val(txtTo)) & "' "
55940   End If
55950   sql = sql & "order by SampleID desc"
55960   Set tb = New Recordset
55970   RecOpenClient 0, tb, sql
55980   lngTo = tb!SampleID
55990   If lngTo - lngFrom > 5000 Then
56000     s = s & lngTo - lngFrom & " Difference. Unable to calculate."
56010   Else
56020     For ThisSID = lngFrom To lngTo
56030       sql = "Select top 1 SampleID from Demographics where " & _
                  "SampleID = '" & ThisSID & "'"
56040       Set tb = New Recordset
56050       RecOpenServer 0, tb, sql
56060       If tb.EOF Then
56070         s = s & ThisSID & ",   "
56080         n = n + 1
56090       End If
56100     Next
56110     If n > 0 Then
56120       s = Left$(s, Len(s) - 4)
56130     Else
56140       s = s & "All Used."
56150     End If
56160   End If
56170 Else
56180   s = s & "None."
56190 End If

56200 s = s & vbCrLf & vbCrLf & "Lowest Sample ID found : " & lngFrom & vbCrLf
56210 s = s & "Highest Sample ID found : " & lngTo

56220 txtReport = s

56230 Exit Sub

Calculate_Error:

      Dim strES As String
      Dim intEL As Integer

56240 intEL = Erl
56250 strES = Err.Description
56260 LogError "frmMicroUnusedSampleIDs", "Calculate", intEL, strES, sql


End Sub

Private Sub breCalc_Click()

56270 Screen.MousePointer = vbHourglass

56280 Calculate

56290 Screen.MousePointer = 0

End Sub


Private Sub cmdCancel_Click()

56300 Unload Me

End Sub


Private Sub Form_Load()

56310 dtFrom = Format(Now, "dd/mmm/yyyy")
56320 dtTo = dtFrom

End Sub

Private Sub optDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

56330 dtFrom.Visible = True
56340 dtTo.Visible = True
56350 txtFrom.Visible = False
56360 txtTo.Visible = False

End Sub


Private Sub optSIDs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

56370 dtFrom.Visible = False
56380 dtTo.Visible = False
56390 txtFrom.Visible = True
56400 txtTo.Visible = True

End Sub


