VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Print Results"
   ClientHeight    =   3840
   ClientLeft      =   2220
   ClientTop       =   2235
   ClientWidth     =   6075
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   6075
   Begin MSComCtl2.DTPicker dt 
      Height          =   345
      Left            =   900
      TabIndex        =   14
      Top             =   390
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Format          =   219348993
      CurrentDate     =   37112
   End
   Begin VB.OptionButton o 
      Caption         =   "Biochemistry"
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
      Index           =   1
      Left            =   4290
      TabIndex        =   13
      Top             =   510
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton o 
      Alignment       =   1  'Right Justify
      Caption         =   "Haematology"
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
      Index           =   0
      Left            =   2970
      TabIndex        =   12
      Top             =   510
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2760
      Picture         =   "frmPrintOptions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2850
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1515
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton v 
         Caption         =   "All"
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
         Index           =   2
         Left            =   270
         TabIndex        =   11
         Top             =   1020
         Width           =   555
      End
      Begin VB.OptionButton v 
         Caption         =   "Only Valid"
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
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   690
         Width           =   1095
      End
      Begin VB.OptionButton v 
         Caption         =   "Valid, not Printed"
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
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   450
      TabIndex        =   5
      Top             =   960
      Width           =   2655
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   1500
         TabIndex        =   7
         Top             =   900
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tto"
         BuddyDispid     =   196615
         OrigLeft        =   1650
         OrigTop         =   1140
         OrigRight       =   1890
         OrigBottom      =   1335
         Max             =   99999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1485
         TabIndex        =   6
         Top             =   450
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tfrom"
         BuddyDispid     =   196614
         OrigLeft        =   1620
         OrigTop         =   540
         OrigRight       =   1860
         OrigBottom      =   945
         Max             =   99999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox tfrom 
         Height          =   315
         Left            =   270
         MaxLength       =   12
         TabIndex        =   0
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox tto 
         Height          =   315
         Left            =   270
         MaxLength       =   12
         TabIndex        =   1
         Top             =   870
         Width           =   1215
      End
   End
   Begin VB.CommandButton bstop 
      Caption         =   "&Stop Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4170
      Picture         =   "frmPrintOptions.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2850
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1350
      Picture         =   "frmPrintOptions.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2850
      Width           =   975
   End
End
Attribute VB_Name = "frmPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tonumber As Long
Dim fromnumber As Long

Dim printing As Boolean


Private Sub bcancel_Click()

39310 If printing Then Exit Sub

39320 Unload Me

End Sub

Private Sub bPrint_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim Temp As Long
      Dim t
      Dim printit As Integer
      Dim total As Long

39330 On Error GoTo bPrint_Click_Error

39340 If printing Then Exit Sub

39350 printing = True

39360 If Val(tfrom) > Val(tto) Then
39370   Temp = tto
39380   tto = tfrom
39390   tfrom = tto
39400 End If

39410 total = Abs(Val(tfrom) - Val(tto)) + 1
39420 If total > 50 Then
39430   iMsg "Too many to print (" & Format$(total) & ") reports." & vbCrLf & "Maximum 50"
39440   printing = False
39450   Exit Sub
39460 End If
39470 If total > 20 Then
39480   If iMsg("You requested to print " & Format$(total) & " reports." & vbCrLf & _
                "Are you sure?", vbYesNo + vbQuestion) = vbNo Then
39490     printing = False
39500     Exit Sub
39510   End If
39520 End If

39530 For Temp = Val(tfrom) To Val(tto)
39540   printit = True
39550   If o(0) Then
39560     sql = "select * from HaemResults where " & _
                "SampleID = '" & Format$(Temp) & "'"
39570   Else
39580     sql = "Select * from BioResults where " & _
                "SampleID = '" & Format$(Temp) & "'"
39590   End If
        
39600   Set tb = New Recordset
39610   RecOpenClient 0, tb, sql
39620   If tb.EOF Then
39630     printit = False
39640   Else
39650     If v(0) Then
39660       If tb!Valid = 0 Then printit = False
39670       If tb!Printed = 1 Then printit = False
39680     ElseIf v(1) Then
39690       If tb!Valid = 0 Then printit = False
39700     End If
39710   End If
39720   If printit Then
39730     bPrint.Visible = False
39740     bstop.Visible = True
39750     t = Timer
39760     If o(0) Then
39770       PrintResultHaemWin Format$(Temp)
39780     Else
39790       sql = "Update BioResults " & _
                  "Set Valid = 1, Printed = 0 " & _
                  "where SampleID = '" & Format$(Temp) & "'"
39800       Cnxn(0).Execute sql
39810       PrintResultBioWin Format$(Temp)
39820     End If
39830     Do While Timer - t < 2
39840       DoEvents
39850     Loop
39860   End If
39870 Next

39880 printing = False
39890 bPrint.Visible = True
39900 bstop.Visible = False

39910 Exit Sub

bPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

39920 intEL = Erl
39930 strES = Err.Description
39940 LogError "fprintoptions", "bPrint_Click", intEL, strES, sql

End Sub

Private Sub bstop_Click()

39950 Unload Me

End Sub

Private Sub FillToAndFrom()

      Dim tb As Recordset
      Dim sql As String

39960 On Error GoTo FillToAndFrom_Error

39970 If o(0) Then
39980   sql = "select sampleid from haemresults where " & _
              "rundate = '" & Format$(dt, "dd/mmm/yyyy") & "' " & _
              "order by sampleid"
39990 Else
40000   sql = "select distinct sampleid from BioResults where " & _
              "rundate = '" & Format$(dt, "dd/mmm/yyyy") & "' " & _
              "order by sampleid"
40010 End If
40020 Set tb = New Recordset
40030 RecOpenClient 0, tb, sql

40040 If tb.EOF Then
40050   fromnumber = 0
40060   tfrom = ""
40070   tonumber = 0
40080   tto = ""
40090   Exit Sub
40100 Else
40110   tfrom = tb!SampleID
40120   fromnumber = Val(tfrom)
40130   tb.MoveLast
40140   tto = tb!SampleID
40150   tonumber = Val(tto)
40160 End If

40170 Exit Sub

FillToAndFrom_Error:

      Dim strES As String
      Dim intEL As Integer

40180 intEL = Erl
40190 strES = Err.Description
40200 LogError "fprintoptions", "FillToAndFrom", intEL, strES, sql

End Sub


Private Sub dt_CloseUp()

40210 FillToAndFrom

End Sub

Private Sub Form_Load()

40220 dt = Format$(Now, "dd/mm/yyyy")

40230 UpDown1.max = 9999999
40240 UpDown2.max = 9999999

      'FillToAndFrom

40250 printing = False

End Sub

Private Sub o_Click(Index As Integer)

40260 FillToAndFrom

End Sub


