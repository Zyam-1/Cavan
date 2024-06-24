VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm24hrUrine 
   Caption         =   "NetAcquire - 24 Hr Urine Excretion"
   ClientHeight    =   6555
   ClientLeft      =   2445
   ClientTop       =   600
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   5310
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   675
      Left            =   3990
      Picture         =   "frm24hrUrine.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   3810
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   675
      Left            =   3990
      Picture         =   "frm24hrUrine.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5670
      Width           =   885
   End
   Begin VB.CommandButton bPrint 
      Height          =   675
      Left            =   3990
      Picture         =   "frm24hrUrine.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   4500
      Width           =   885
   End
   Begin VB.Frame Frame4 
      Caption         =   "Units"
      Height          =   825
      Left            =   3780
      TabIndex        =   40
      Top             =   2700
      Width           =   1245
      Begin VB.TextBox tUnits 
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   41
         Text            =   "24"
         Top             =   180
         Width           =   300
      End
      Begin ComCtl2.UpDown udHours 
         Height          =   195
         Left            =   150
         TabIndex        =   44
         Top             =   510
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   344
         _Version        =   327681
         Value           =   24
         BuddyControl    =   "tUnits"
         BuddyDispid     =   196613
         OrigLeft        =   3180
         OrigTop         =   3600
         OrigRight       =   3420
         OrigBottom      =   4365
         Max             =   24
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hrs"
         Height          =   195
         Left            =   930
         TabIndex        =   43
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "mmol/"
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4425
      Left            =   150
      TabIndex        =   10
      Top             =   1890
      Width           =   3405
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   9
         Left            =   2190
         TabIndex        =   52
         Top             =   3930
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   9
         Left            =   1140
         TabIndex        =   51
         Top             =   3930
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   8
         Left            =   2190
         TabIndex        =   37
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   7
         Left            =   2190
         TabIndex        =   36
         Top             =   2940
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   6
         Left            =   2190
         TabIndex        =   35
         Top             =   2580
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   5
         Left            =   2190
         TabIndex        =   34
         Top             =   2220
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   4
         Left            =   2190
         TabIndex        =   33
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   3
         Left            =   2190
         TabIndex        =   32
         Top             =   1500
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   2
         Left            =   2190
         TabIndex        =   31
         Top             =   1140
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   1
         Left            =   2190
         TabIndex        =   30
         Top             =   780
         Width           =   795
      End
      Begin VB.TextBox tper24 
         Height          =   285
         Index           =   0
         Left            =   2190
         TabIndex        =   29
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   8
         Left            =   1140
         TabIndex        =   28
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   7
         Left            =   1140
         TabIndex        =   27
         Top             =   2940
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   6
         Left            =   1140
         TabIndex        =   26
         Top             =   2580
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   5
         Left            =   1140
         TabIndex        =   25
         Top             =   2220
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   4
         Left            =   1140
         TabIndex        =   24
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   3
         Left            =   1140
         TabIndex        =   23
         Top             =   1500
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   2
         Left            =   1140
         TabIndex        =   22
         Top             =   1140
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   21
         Top             =   780
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   20
         Top             =   420
         Width           =   795
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Nitrogen"
         Height          =   195
         Index           =   9
         Left            =   450
         TabIndex        =   50
         Top             =   3990
         Width           =   600
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Potassium"
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   19
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Sodium"
         Height          =   195
         Index           =   2
         Left            =   525
         TabIndex        =   18
         Top             =   1170
         Width           =   525
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Urea"
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   17
         Top             =   825
         Width           =   345
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Calcium"
         Height          =   195
         Index           =   5
         Left            =   495
         TabIndex        =   16
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Phosphorus"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   2610
         Width           =   840
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "T. Prot"
         Height          =   195
         Index           =   8
         Left            =   570
         TabIndex        =   14
         Top             =   3360
         Width           =   480
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Magnesium"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   2970
         Width           =   810
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Chloride"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   12
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Creatinine"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   11
         Top             =   480
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Volume"
      Height          =   675
      Left            =   3780
      TabIndex        =   9
      Top             =   1890
      Width           =   1245
      Begin VB.TextBox tVolume 
         Height          =   285
         Left            =   105
         TabIndex        =   38
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label16 
         Caption         =   "ml"
         Height          =   195
         Left            =   930
         TabIndex        =   39
         Top             =   330
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5085
      Begin VB.TextBox tDoB 
         Height          =   285
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1245
      End
      Begin VB.TextBox tChart 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   1245
      End
      Begin VB.TextBox tName 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   510
         Width           =   4035
      End
      Begin VB.TextBox txtSampleID 
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label lRunTime 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   49
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lRunDate 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   48
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Run Date/Time"
         Height          =   195
         Left            =   2550
         TabIndex        =   47
         Top             =   900
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   540
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         Height          =   195
         Left            =   3330
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm24hrUrine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strmMol(0 To 9) As String
Dim strPer24(0 To 9) As String
Dim strDP(0 To 9) As String

Private pPrintToPrinter As String
Private Sub Calculate()

          Dim n As Integer

34600     For n = 0 To 8
34610         tper24(n) = ""
34620     Next

34630     If Val(tvolume) = 0 Then Exit Sub
34640     If Val(tUnits) = 0 Then Exit Sub

34650     For n = 0 To 8
34660         If Val(strmMol(n)) > 0 Then
34670             strPer24(n) = (Val(tvolume) / 1000) * Val(strmMol(n))
34680         End If
34690     Next
        
34700     If Val(strPer24(1)) > 0 Then
34710         strPer24(9) = Format(Val(strPer24(1)) * 0.028, "0.0")
34720     End If

34730     For n = 0 To 8
34740         tper24(n) = Format$(strPer24(n), strDP(n))
34750         tmmol(n) = strmMol(n)
34760     Next

End Sub

Private Sub bPrint_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim Ward As String
          Dim Clin As String
          Dim GP As String

34770     On Error GoTo bPrint_Click_Error

34780     GetWardClinGP txtSampleID, Ward, Clin, GP

34790     sql = "Select * from PrintPending where " & _
              "Department = '2' " & _
              "and SampleID = '" & txtSampleID & "'"
34800     Set tb = New Recordset
34810     RecOpenClient 0, tb, sql
34820     If tb.EOF Then
34830         tb.AddNew
34840     End If
34850     tb!SampleID = txtSampleID
34860     tb!Ward = Ward
34870     tb!Clinician = Clin
34880     tb!GP = GP
34890     tb!Department = "2"
34900     tb!Initiator = UserName
34910     tb!UsePrinter = pPrintToPrinter
34920     tb.Update


34930     Exit Sub

bPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

34940     intEL = Erl
34950     strES = Err.Description
34960     LogError "frm24hrUrine", "bPrint_Click", intEL, strES, sql


End Sub

Public Property Get PrintToPrinter() As String

34970     PrintToPrinter = pPrintToPrinter

End Property


Public Property Let PrintToPrinter(ByVal strNewValue As String)

34980     pPrintToPrinter = strNewValue

End Property

Private Sub cmdCancel_Click()

34990     Unload Me

End Sub

Private Sub cmdSetPrinter_Click()

35000     frmForcePrinter.From = Me
35010     frmForcePrinter.Show 1

35020     If pPrintToPrinter = "Automatic Selection" Then
35030         pPrintToPrinter = ""
35040     End If

35050     If pPrintToPrinter <> "" Then
35060         cmdSetPrinter.BackColor = vbRed
35070         cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
35080     Else
35090         cmdSetPrinter.BackColor = vbButtonFace
35100         pPrintToPrinter = ""
35110         cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
35120     End If

End Sub

Private Sub txtsampleid_LostFocus()

          Dim sn As Recordset
          Dim sql As String
          Dim BR As BIEResult
          Dim BRs As New BIEResults
          Dim n As Integer
          Dim OffSet As Integer
          Dim CurrentPrintFormat As String

35130     On Error GoTo txtsampleid_LostFocus_Error

35140     If Trim(txtSampleID) = "" Then Exit Sub

35150     txtSampleID = Format(Val(txtSampleID))

35160     sql = "select * from demographics where " & _
              "sampleid = '" & txtSampleID & "'"
35170     Set sn = New Recordset
35180     RecOpenServer 0, sn, sql
35190     If sn.EOF Then
35200         tName = ""
35210         tChart = ""
35220         tDoB = ""
35230         lRunDate = ""
35240     Else
35250         tChart = sn!Chart & ""
35260         tName = sn!PatName & ""
35270         tDoB = sn!DoB & ""
35280         lRunDate = Format(sn!Rundate, "dd/mm/yyyy")
35290     End If

35300     lRunTime = ""

35310     For n = 0 To 9
35320         strmMol(n) = ""
35330         strPer24(n) = ""
35340     Next

35350     Set BRs = BRs.Load("Bio", txtSampleID, "Results", gDONTCARE, gDONTCARE)
35360     For Each BR In BRs
35370         Select Case UCase$(BR.ShortName)
                  Case "URVOL": tvolume = BR.Result: OffSet = -1 'Urine Volume
35380             Case 887: tUnits = BR.Result: OffSet = -1 'Urine Collection Hours
35390             Case "UCRE", "CREJC", "CREA", "CREAT": OffSet = 0
35400             Case "UREA": OffSet = 1
35410             Case "USOD", "NA": OffSet = 2
35420             Case "UPOT", "K": OffSet = 3
35430             Case "CL": OffSet = 4
35440             Case "CA": OffSet = 5
35450             Case "PHOS", "PO4": OffSet = 6
35460             Case "MG": OffSet = 7
35470             Case "UPRO", "UPROT", "PRO", "PROT": OffSet = 8
35480             Case Else:
35490                 OffSet = -1
35500         End Select
35510         If OffSet > -1 Then
35520             Select Case BR.Printformat
                      Case 0: CurrentPrintFormat = "####0   "
35530                 Case 1: CurrentPrintFormat = "####0.0 "
35540                 Case 2: CurrentPrintFormat = "####0.00"
35550                 Case 3: CurrentPrintFormat = "###0.000"
35560             End Select
35570             lRunTime = Format(BR.RunTime, "dd/mm/yy hh:mm")
35580             strmMol(OffSet) = Format(BR.Result, CurrentPrintFormat)
35590             strDP(OffSet) = CurrentPrintFormat
35600         End If
35610     Next

35620     Calculate

35630     Exit Sub

txtsampleid_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

35640     intEL = Erl
35650     strES = Err.Description
35660     LogError "frm24hrUrine", "txtsampleid_LostFocus", intEL, strES, sql


End Sub

Private Sub tVolume_LostFocus()

35670     Calculate

End Sub

Private Sub udHours_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

35680     Calculate

End Sub

