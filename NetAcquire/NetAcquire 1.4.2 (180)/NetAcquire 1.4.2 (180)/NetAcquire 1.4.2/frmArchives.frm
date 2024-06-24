VERSION 5.00
Begin VB.Form frmArchives 
   Caption         =   "NetAcquire - Archives"
   ClientHeight    =   3405
   ClientLeft      =   330
   ClientTop       =   1095
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   10725
   Begin VB.TextBox txtReport 
      Height          =   2505
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   750
      Width           =   10395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   2940
      Picture         =   "frmArchives.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox txtSampleID 
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Number"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   1125
   End
End
Attribute VB_Name = "frmArchives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckDemographics()

          Dim s As String
          Dim sql As String
          Dim tb As Recordset
          Dim tbCurrent As Recordset
          Dim RecHead(0 To 18) As String
          Dim RecFrom(0 To 18) As String
          Dim RecTo(0 To 18) As String
          Dim n As Integer

52100     On Error GoTo CheckDemographics_Error

52110     RecHead(0) = "Chart Number"
52120     RecHead(1) = "Patient Name"
52130     RecHead(2) = "Age"
52140     RecHead(3) = "Sex"
52150     RecHead(4) = "Date of Birth"
52160     RecHead(5) = "Address Line 1"
52170     RecHead(6) = "Address Line 2"
52180     RecHead(7) = "Ward"
52190     RecHead(8) = "Clinician"
52200     RecHead(9) = "GP"
52210     RecHead(10) = "Clinical Details"
52220     RecHead(11) = "Routine"
52230     RecHead(12) = "FAXed"
52240     RecHead(13) = "Fasting"
52250     RecHead(14) = "On Warfarin"
52260     RecHead(15) = "Pregnant"
52270     RecHead(16) = "A and E"
52280     RecHead(17) = "Category"
52290     RecHead(18) = "Urgent"

52300     sql = "Select * from Demographics where " & _
              "SampleID = '" & Val(txtSampleID) & "'"
52310     Set tbCurrent = New Recordset
52320     RecOpenServer 0, tbCurrent, sql
52330     If Not tbCurrent.EOF Then
52340         RecTo(0) = Trim$(tbCurrent!Chart & "")
52350         RecTo(1) = Trim$(tbCurrent!PatName & "")
52360         RecTo(2) = Trim$(tbCurrent!Age & "")
52370         RecTo(3) = Trim$(tbCurrent!Sex & "")
52380         RecTo(4) = Trim$(tbCurrent!DoB & "")
52390         RecTo(5) = Trim$(tbCurrent!Addr0 & "")
52400         RecTo(6) = Trim$(tbCurrent!Addr1 & "")
52410         RecTo(7) = Trim$(tbCurrent!Ward & "")
52420         RecTo(8) = Trim$(tbCurrent!Clinician & "")
52430         RecTo(9) = Trim$(tbCurrent!GP & "")
52440         RecTo(10) = Trim$(tbCurrent!ClDetails & "")
52450         RecTo(11) = IIf(IsNull(tbCurrent!RooH), "", tbCurrent!RooH)
52460         RecTo(12) = IIf(IsNull(tbCurrent!FAXed), "", tbCurrent!FAXed)
52470         RecTo(13) = IIf(IsNull(tbCurrent!Fasting), "", tbCurrent!Fasting)
52480         RecTo(14) = IIf(IsNull(tbCurrent!OnWarfarin), "", tbCurrent!OnWarfarin)
52490         RecTo(15) = IIf(IsNull(tbCurrent!Pregnant), "", tbCurrent!Pregnant)
52500         RecTo(16) = tbCurrent!AandE & ""
52510         RecTo(17) = tbCurrent!Category & ""
52520         RecTo(18) = IIf(IsNull(tbCurrent!Urgent), "", tbCurrent!Urgent)
52530     End If

52540     s = ""
52550     sql = "Select * from ArcDemographics where " & _
              "SampleID = '" & Val(txtSampleID) & "' " & _
              "order by DateTimeOfArchive desc"
52560     Set tb = New Recordset
52570     RecOpenClient 0, tb, sql
52580     If Not tb.EOF Then
52590         Do While Not tb.EOF
52600             If Trim$(tb!ArchiveOperator & "") <> "" Then
52610                 s = s & "Demographics changed" & _
                          " on " & Format$(tb!DateTimeOfArchive, "dd/mmm/yyyy") & _
                          " at " & Format$(tb!DateTimeOfArchive, "hh:nn:ss") & _
                          " by " & tb!ArchiveOperator & vbCrLf
52620                 RecFrom(0) = Trim$(tb!Chart & "")
52630                 RecFrom(1) = Trim$(tb!PatName & "")
52640                 RecFrom(2) = Trim$(tb!Age & "")
52650                 RecFrom(3) = Trim$(tb!Sex & "")
52660                 RecFrom(4) = Trim$(tb!DoB & "")
52670                 RecFrom(5) = Trim$(tb!Addr0 & "")
52680                 RecFrom(6) = Trim$(tb!Addr1 & "")
52690                 RecFrom(7) = Trim$(tb!Ward & "")
52700                 RecFrom(8) = Trim$(tb!Clinician & "")
52710                 RecFrom(9) = Trim$(tb!GP & "")
52720                 RecFrom(10) = Trim$(tb!ClDetails & "")
52730                 RecFrom(11) = IIf(IsNull(tb!RooH), "", tb!RooH)
52740                 RecFrom(12) = IIf(IsNull(tb!FAXed), "", tb!FAXed)
52750                 RecFrom(13) = IIf(IsNull(tb!Fasting), "", tb!Fasting)
52760                 RecFrom(14) = IIf(IsNull(tb!OnWarfarin), "", tb!OnWarfarin)
52770                 RecFrom(15) = IIf(IsNull(tb!Pregnant), "", tb!Pregnant)
52780                 RecFrom(16) = tb!AandE & ""
52790                 RecFrom(17) = tb!Category & ""
52800                 RecFrom(18) = IIf(IsNull(tb!Urgent), "", tb!Urgent)

52810                 For n = 0 To 10
52820                     If RecFrom(n) <> RecTo(n) Then
52830                         s = s & RecHead(n) & " changed from " & RecFrom(n) & _
                                  " to " & RecTo(n) & vbCrLf
52840                     End If
52850                 Next

52860                 For n = 0 To 10
52870                     RecTo(n) = RecFrom(n)
52880                 Next
          
52890             End If
          
52900             tb.MoveNext
52910         Loop
52920     Else
52930         s = "Demographics:- No changes recorded."
52940     End If

52950     CheckDemographics = s

52960     Exit Function

CheckDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

52970     intEL = Erl
52980     strES = Err.Description
52990     LogError "frmArchives", "CheckDemographics", intEL, strES, sql


End Function

Private Function CheckMasks() As String

          Dim s As String
          Dim sql As String
          Dim tb As Recordset
          Dim tbCurrent As Recordset
          Dim RecHead(0 To 10) As String
          Dim RecFrom(0 To 10) As String
          Dim RecTo(0 To 10) As String
          Dim n As Integer

53000     On Error GoTo CheckMasks_Error

53010     RecHead(0) = "Haemolysed"
53020     RecHead(1) = "Slightly Haemolysed"
53030     RecHead(2) = "Lipaemic"
53040     RecHead(3) = "Old Sample"
53050     RecHead(4) = "Grossly Haemolysed"
53060     RecHead(5) = "Icteric"
53070     RecHead(6) = "LIH Value"

53080     sql = "Select * from Masks where " & _
              "SampleID = '" & Val(txtSampleID) & "'"
53090     Set tbCurrent = New Recordset
53100     RecOpenClient 0, tbCurrent, sql
53110     If Not tbCurrent.EOF Then
53120         RecTo(0) = tbCurrent!H
53130         RecTo(1) = tbCurrent!s
53140         RecTo(2) = tbCurrent!l
53150         RecTo(3) = tbCurrent!o
53160         RecTo(4) = tbCurrent!g
53170         RecTo(5) = tbCurrent!J
53180         RecTo(6) = tbCurrent!LIH
53190     End If

53200     s = ""
53210     sql = "Select * from ArcMasks where " & _
              "SampleID = '" & Val(txtSampleID) & "' " & _
              "order by DateTimeOfArchive"
53220     Set tb = New Recordset
53230     RecOpenClient 0, tb, sql
53240     If Not tb.EOF Then
53250         Do While Not tb.EOF
53260             s = s & "Masks changed" & _
                      " on " & Format$(tb!DateTimeOfArchive, "dd/mmm/yyyy") & _
                      " at " & Format$(tb!DateTimeOfArchive, "hh:nn:ss") & _
                      " by " & tb!Operator & vbCrLf

53270             RecFrom(0) = tb!H
53280             RecFrom(1) = tb!s
53290             RecFrom(2) = tb!l
53300             RecFrom(3) = tb!o
53310             RecFrom(4) = tb!g
53320             RecFrom(5) = tb!J
53330             RecFrom(6) = tb!LIH

53340             For n = 0 To 6
53350                 If RecFrom(n) <> RecTo(n) Then
53360                     s = s & RecHead(n)
53370                     If n <> 6 Then
53380                         s = s & " Flag"
53390                     Else
53400                         s = s & " Value"
53410                     End If
53420                     s = s & " changed from " & _
                              IIf(RecFrom(n) = "", "<Blank>", RecFrom(n)) & _
                              " to " & _
                              IIf(RecTo(n) = "", "<Blank>", RecTo(n)) & vbCrLf
53430                 End If
53440             Next
53450             For n = 0 To 10
53460                 RecTo(n) = RecFrom(n)
53470             Next
          
53480             tb.MoveNext
53490         Loop
53500     Else
53510         s = "Masks:- No changes recorded."
53520     End If

53530     CheckMasks = s

53540     Exit Function

CheckMasks_Error:

          Dim strES As String
          Dim intEL As Integer

53550     intEL = Erl
53560     strES = Err.Description
53570     LogError "frmArchives", "CheckMasks", intEL, strES, sql


End Function
Private Sub cmdCancel_Click()

53580     Unload Me

End Sub


Private Sub Form_Resize()

53590     If Me.width < 4305 Then
53600         Me.width = 4305
53610     End If

53620     txtReport.width = Me.width - 510

53630     If Me.height < 3810 Then
53640         Me.height = 3810
53650     End If

53660     txtReport.height = Me.height - 1425

End Sub


Private Sub txtsampleid_LostFocus()

          Dim s As String

53670     txtSampleID = Format$(Val(txtSampleID))
53680     If txtSampleID = "0" Then
53690         Exit Sub
53700     End If

          '50    s = CheckComments()
          '60    If s <> "" Then
          '70      s = s & vbCrLf & String(40, "-") & vbCrLf
          '80    End If

53710     s = s & CheckMasks()
53720     s = s & vbCrLf & String(40, "-") & vbCrLf

53730     s = s & CheckDemographics()
53740     s = s & vbCrLf & String(40, "-") & vbCrLf

53750     txtReport = s

End Sub

