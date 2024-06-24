VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmXMO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Units Pending"
   ClientHeight    =   7875
   ClientLeft      =   270
   ClientTop       =   525
   ClientWidth     =   8145
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
   Icon            =   "frmXMO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7875
   ScaleWidth      =   8145
   Begin VB.CommandButton cmdIssue 
      Appearance      =   0  'Flat
      Caption         =   "&Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   660
      Picture         =   "frmXMO.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5520
      Width           =   915
   End
   Begin VB.CommandButton cmdXM 
      Appearance      =   0  'Flat
      Caption         =   "X-M"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1920
      Picture         =   "frmXMO.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5520
      Width           =   915
   End
   Begin VB.Frame frKey 
      Height          =   645
      Left            =   2790
      TabIndex        =   32
      Top             =   4110
      Width           =   585
      Begin VB.Image iKey 
         Height          =   480
         Left            =   30
         Picture         =   "frmXMO.frx":0EDE
         Top             =   150
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   2010
      TabIndex        =   29
      Top             =   3090
      Width           =   4335
      Begin VB.CommandButton cmdLogAsChecked 
         Caption         =   "Log as Checked"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2610
         Picture         =   "frmXMO.frx":1320
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   210
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblChecked 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Frame frXM 
      Enabled         =   0   'False
      Height          =   2505
      Left            =   360
      TabIndex        =   16
      Top             =   4110
      Width           =   3015
      Begin VB.CheckBox cXM 
         Alignment       =   1  'Right Justify
         Caption         =   "Enzyme"
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
         Left            =   450
         TabIndex        =   28
         Top             =   840
         Width           =   915
      End
      Begin VB.CheckBox cXM 
         Alignment       =   1  'Right Justify
         Caption         =   "Coombs"
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
         Left            =   450
         TabIndex        =   27
         Top             =   540
         Width           =   915
      End
      Begin VB.CheckBox cXM 
         Alignment       =   1  'Right Justify
         Caption         =   "Room Temp"
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
         Left            =   150
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   1470
         ScaleHeight     =   285
         ScaleWidth      =   885
         TabIndex        =   19
         Top             =   480
         Width           =   915
         Begin VB.OptionButton xReaction 
            Caption         =   "+"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   23
            Top             =   60
            Width           =   375
         End
         Begin VB.OptionButton xReaction 
            Alignment       =   1  'Right Justify
            Caption         =   "O"
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   22
            Top             =   60
            Width           =   405
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1470
         ScaleHeight     =   285
         ScaleWidth      =   885
         TabIndex        =   18
         Top             =   180
         Width           =   915
         Begin VB.OptionButton xReaction 
            Caption         =   "+"
            Height          =   225
            Index           =   1
            Left            =   450
            TabIndex        =   21
            Top             =   30
            Width           =   375
         End
         Begin VB.OptionButton xReaction 
            Alignment       =   1  'Right Justify
            Caption         =   "O"
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   20
            Top             =   30
            Width           =   405
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1470
         ScaleHeight     =   285
         ScaleWidth      =   885
         TabIndex        =   17
         Top             =   780
         Width           =   915
         Begin VB.OptionButton xReaction 
            Caption         =   "+"
            Height          =   195
            Index           =   5
            Left            =   450
            TabIndex        =   25
            Top             =   60
            Width           =   405
         End
         Begin VB.OptionButton xReaction 
            Alignment       =   1  'Right Justify
            Caption         =   "O"
            Height          =   195
            Index           =   4
            Left            =   30
            TabIndex        =   24
            Top             =   60
            Width           =   405
         End
      End
      Begin VB.Label lblReaction 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Left            =   1380
         TabIndex        =   36
         Top             =   630
         Width           =   285
      End
      Begin VB.Label lblIAT 
         AutoSize        =   -1  'True
         Caption         =   "I.A.T."
         Height          =   195
         Left            =   840
         TabIndex        =   35
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.ListBox List1 
      Columns         =   1
      Height          =   2400
      Left            =   3540
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4170
      Width           =   4065
   End
   Begin VB.TextBox txtUnitNumber 
      BackColor       =   &H00FFFFFF&
      DataField       =   "number"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1140
      MaxLength       =   14
      TabIndex        =   0
      Top             =   270
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
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
      Height          =   795
      Left            =   1080
      Picture         =   "frmXMO.frx":1762
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1965
   End
   Begin VB.CommandButton cmdRemove 
      Appearance      =   0  'Flat
      Caption         =   "&Remove from pending list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5610
      Picture         =   "frmXMO.frx":1DCC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1995
   End
   Begin VB.CommandButton cmdAddPending 
      Appearance      =   0  'Flat
      Caption         =   "&Mark as Pending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3570
      Picture         =   "frmXMO.frx":220E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1905
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   40
      Top             =   7680
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ISBT128"
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
      Left            =   4230
      TabIndex        =   44
      Top             =   330
      Width           =   630
   End
   Begin VB.Label lblISBT128 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4950
      TabIndex        =   43
      Top             =   270
      Width           =   2670
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Lab Number"
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
      Left            =   120
      TabIndex        =   42
      Top             =   3330
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblLabNo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   3510
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This Unit is Tagged. Click here to view."
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1140
      TabIndex        =   39
      Top             =   630
      Visible         =   0   'False
      Width           =   6465
   End
   Begin VB.Label lblKell 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6300
      TabIndex        =   38
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Kell"
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
      Left            =   6000
      TabIndex        =   37
      Top             =   1530
      Width           =   255
   End
   Begin VB.Label lblProduct 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      TabIndex        =   15
      Top             =   1020
      Width           =   6465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Product"
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
      Left            =   540
      TabIndex        =   14
      Top             =   1110
      Width           =   555
   End
   Begin VB.Label lblGroupRh 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3990
      TabIndex        =   12
      Top             =   1470
      Width           =   915
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1140
      TabIndex        =   11
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Screen"
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
      Left            =   600
      TabIndex        =   10
      Top             =   2010
      Width           =   510
   End
   Begin VB.Label lblScreen 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1140
      TabIndex        =   9
      Top             =   1860
      Width           =   6465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
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
      Left            =   690
      TabIndex        =   8
      Top             =   1500
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Group/Rh"
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
      TabIndex        =   7
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1140
      TabIndex        =   6
      Top             =   2460
      Width           =   6465
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status"
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
      Left            =   660
      TabIndex        =   5
      Top             =   2580
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit Number"
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
      TabIndex        =   4
      Top             =   360
      Width           =   885
   End
End
Attribute VB_Name = "frmXMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AmendProduct(TimeNow As String, bbEvent As String)

      Dim tb As Recordset
      Dim sql As String
      Dim pc As String

10    On Error GoTo AmendProduct_Error

20    pc = ProductBarCodeFor(Trim$(lblProduct))

30    sql = "Select * from Latest where " & _
            "BarCode = '" & pc & "' " & _
            "and Number = '" & txtUnitNumber & "' " & _
            "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70            tb!crt = cXM(0)
80            tb!cco = cXM(1)
90            tb!cen = cXM(2)
100           tb!crtr = xReaction(1)
110           tb!ccor = xReaction(3)
120           tb!cenr = xReaction(5)
130       tb!Event = bbEvent
140       tb!Patid = frmxmatch.txtChart
150       tb!PatName = frmxmatch.txtName
160       tb!Operator = UserCode
170       tb!DateTime = TimeNow
180       tb!LabNumber = frmxmatch.tLabNum
190       tb.Update

200       sql = "Insert into Product " & _
                "  Select * from Latest where " & _
                "  BarCode = '" & pc & "' " & _
                "  and Number = '" & txtUnitNumber & "' " & _
                "  and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
210       CnxnBB(0).Execute sql

220   End If

230   Exit Sub

AmendProduct_Error:

Dim strES As String
Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmXMO", "AmendProduct", intEL, strES, sql

End Sub

Private Sub FillDetails(ByVal p As Product)

      Dim sql As String
      Dim GRh As String
      Dim tb As Recordset

10    lblTag.Visible = False

20    lblExpiry = Format(p.DateExpiry, "dd/mm/yyyy")
30    lblProduct = ProductWordingFor(p.BarCode)

40    If Trim$(frmxmatch.txtChart) <> "" Then
50        sql = "Select * from Product where " & _
                "Number = '" & Trim$(txtUnitNumber) & "' " & _
                "and PatID = '" & frmxmatch.txtChart & "' " & _
                "and crt = 1 and crtr = 1 " & _
                "AND DateExpiry = '" & Format(p.DateExpiry, "dd/MMM/yyyy") & "'"
60        Set tb = New Recordset
70        RecOpenServerBB 0, tb, sql
80        If Not tb.EOF Then
90            iMsg "Unit has been found incompatible for this patient and cannot be issued!", vbExclamation
100           If TimedOut Then Unload Me: Exit Sub
110           txtUnitNumber = ""
120           Exit Sub
130       End If
140   End If

150   lblISBT128 = p.ISBT128

160   If InStr(p.Screen, "K+") Then
170       lblKell = "Positive"
180   ElseIf InStr(p.Screen & "", "K-") Then
190       lblKell = "Negative"
200   Else
210       lblKell = ""
220   End If
230   If Trim$(lblProduct) = "" Then Exit Sub

240   If DateDiff("d", Now, lblExpiry) < 0 Then
250       iMsg "Unit expired!", vbExclamation
260       If TimedOut Then Unload Me: Exit Sub
270       RemoveDetails
280       txtUnitNumber = ""
290       Exit Sub
300   End If

310   GRh = p.GroupRh

320   If Not Allowed(GRh) Then
330       txtUnitNumber = ""
340       Exit Sub
350   End If

360   WarnAboutGroup GRh, lblKell

370   If Left$(frmxmatch.lstfg & "  ", 2) <> "O " Then
380       If InStr(UCase$(p.Screen & ""), "O RECIPIENT ONLY") Then
390           iMsg "Screen: " & tb!Screen, vbInformation
400           If TimedOut Then Unload Me: Exit Sub
410       End If
420   End If

430   lblGroupRh = Bar2Group(GRh)

440   lblScreen = p.Screen

450   Select Case p.PackEvent
        Case "Y":
460       lblStatus = "Removed Pending Transfusion."
470       cmdRemove.Enabled = True
  Case "P":
480       lblStatus = "Pending for "
490       If p.SampleID = frmxmatch.tLabNum Then
500         lblStatus = lblStatus & "this patient."
510         cmdRemove.Enabled = True
520       Else
530         lblStatus = lblStatus & "another patient."
540         cmdAddPending.Enabled = True
550       End If
560     Case "D": lblStatus = "This unit number has been destroyed."
570     Case "S": lblStatus = "This unit has been Transfused."
580     Case "T": lblStatus = "This unit has been returned."
590     Case "X": lblStatus = "Crossmatched for "
600       If p.Chart = frmxmatch.txtChart Or p.PatName & "" = frmxmatch.txtName Then
610         lblStatus = lblStatus & "this patient."
620       Else
630         lblStatus = lblStatus & "another patient."
640       End If
650       cmdAddPending.Enabled = True
660     Case "I": lblStatus = "Issued to "
670       If p.Chart = frmxmatch.txtChart Or p.PatName & "" = frmxmatch.txtName Then
680         lblStatus = lblStatus & "this patient."
690       Else
700         lblStatus = lblStatus & "another patient."
710       End If
720       cmdAddPending.Enabled = True

721     Case "V": lblStatus = "Electronic Issue to "
722       If p.Chart = frmxmatch.txtChart Or p.PatName & "" = frmxmatch.txtName Then
723         lblStatus = lblStatus & "this patient."
724       Else
725         lblStatus = lblStatus & "another patient."
726       End If
727       cmdAddPending.Enabled = True

730     Case "C", "R": lblStatus = "In free stock."
740       cmdAddPending.Enabled = True

750   End Select

760   If p.Checked Then
770       lblChecked = "Group  Checked"
780       cmdLogAsChecked.Visible = False
790   Else
800       lblChecked = "Group Not Checked"
810       cmdLogAsChecked.Visible = True
820   End If

830   If TagIsPresent(txtUnitNumber, CDate(lblExpiry)) Then
840       lblTag.Visible = True
850   End If

End Sub

Private Sub WarnAboutGroup(ByVal GRh As String, ByVal UnitKell As String)

      Dim RightGroup As Integer
      Dim s As String

10    RightGroup = True

20    Select Case Left$(frmxmatch.lstfg & "  ", 2)
      Case "O ": If GRh <> "51" And GRh <> "95" Then RightGroup = False
30    Case "A ": If GRh <> "62" And GRh <> "06" Then RightGroup = False
40    Case "B ": If GRh <> "73" And GRh <> "17" Then RightGroup = False
50    Case "AB": If GRh <> "84" And GRh <> "28" Then RightGroup = False
60    End Select

70    If InStr(UCase$(frmxmatch.lstfg), "NEG") And (GRh = "51" Or GRh = "62" Or GRh = "73" Or GRh = "84") Then RightGroup = False
80    If InStr(UCase$(frmxmatch.lstfg), "POS") And (GRh = "95" Or GRh = "06" Or GRh = "17" Or GRh = "28") Then RightGroup = False

90    If Not RightGroup Then
100       iMsg "Donor unit and Patient group differ.", vbInformation
110       If TimedOut Then Unload Me: Exit Sub
120       If Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "50y" Then
130           s = "It is recommended practice to" & vbCrLf & _
                  "transfuse fully compatible blood" & vbCrLf & _
                  "to this category of patient." & vbCrLf & _
                  "(Female of less than 50 years)" & vbCrLf & _
                  "This practice is not being followed."
140           iMsg s, vbInformation
150           If TimedOut Then Unload Me: Exit Sub
160       End If
170   End If

180   If Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "60y" And frmxmatch.cmbKell = "K-" And UnitKell = "Positive" Then
190       s = "It is recommended practice to" & vbCrLf & _
              "transfuse Kell Negative blood" & vbCrLf & _
              "to this category of patient." & vbCrLf & _
              "(Female of less than 60 years K-)" & vbCrLf & _
              "This practice is not being followed."
200       iMsg s, vbInformation
210       If TimedOut Then Unload Me: Exit Sub
220   ElseIf Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "60y" And frmxmatch.cmbKell = "K-" And UnitKell = "" Then
230       iMsg "Unit of unknown Kell Status." & vbCrLf & _
               "Kell Negative blood should be given to this Patient. ", vbInformation
240       If TimedOut Then Unload Me: Exit Sub
250   ElseIf Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "60y" And frmxmatch.cmbKell = "" And UnitKell <> "Negative" Then
260       iMsg "Patient of unknown Kell Status." & vbCrLf & _
               "Kell Negative blood should be given to this Patient. ", vbInformation
270       If TimedOut Then Unload Me: Exit Sub
280   End If

End Sub


Private Function Allowed(ByVal Unit As String) As Integer

      Dim Patient As String
      Dim s As String
      Dim A As Integer
      Dim pGroup As String
      Dim prh As String

10    image2grh pGroup, prh

20    If InStr(UCase$(frmxmatch.lstfg), "NEG") _
          Or InStr(UCase$(frmxmatch.lstfg), "DVI") _
          Or InStr(frmxmatch.lstfg, "C/E") Then
30      Patient = "-"
40      Patient = Trim$(Left$(frmxmatch.lstfg, 2)) & Patient
50    ElseIf InStr(UCase$(frmxmatch.lstfg), "POS") Then
60      Patient = "+"
70      Patient = Trim$(Left$(frmxmatch.lstfg, 2)) & Patient
80    Else
90      Patient = prh
100     s = "Confirm Patient Group/Rhesus? " & _
            vbCrLf & vbCrLf & _
            Space(10) & pGroup & _
            IIf(Patient = "-", " Negative", " Positive")
110     Answer = iMsg(s, vbYesNo + vbQuestion)
120     If TimedOut Then Unload Me: Exit Function
130     If Answer = vbNo Then
140       Allowed = False
150       Exit Function
160     Else
170       Patient = pGroup & Patient
180     End If
190   End If

200   A = False
210   If InStr(frmXM.lblProduct, "Plasma") Then
220     A = True
230   Else
240       Select Case Patient
          Case "O-": If Unit = "95" Then A = True
250           If Unit = "51" Then A = True
260       Case "O+": If Unit = "51" Then A = True
270           If Unit = "95" Then A = True
280       Case "A-": If Unit = "06" Then A = True
290           If Unit = "95" Then A = True
300           If Unit = "62" Then A = True
310           If Unit = "51" Then A = True
320       Case "A+": If Unit = "62" Then A = True
330           If Unit = "06" Then A = True
340           If Unit = "51" Then A = True
350           If Unit = "95" Then A = True
360       Case "B-": If Unit = "17" Then A = True
370           If Unit = "95" Then A = True
380           If Unit = "73" Then A = True
390           If Unit = "51" Then A = True
400       Case "B+": If Unit = "73" Then A = True
410           If Unit = "17" Then A = True
420           If Unit = "51" Then A = True
430           If Unit = "95" Then A = True
440       Case "AB-": A = True
450       Case "AB+": A = True
460       End Select
470   End If

480   If Not A Then
490       s = "This unit is not compatible " & vbCrLf & _
              "with this patient and cannot" & vbCrLf & _
              "be crossmatched and must not" & vbCrLf & _
              "be transfused."
500       iMsg s, vbCritical
510       If TimedOut Then Unload Me: Exit Function
520   End If

530   Allowed = A

End Function
Private Sub XmOrIssue(ByVal XorIorK As String)

      Dim tb As Recordset
      Dim sql As String
      Dim TimeNow As String
      Dim n As Integer
      Dim Counter As Integer
      Dim XMLine As String
      Dim pc As String

10    On Error GoTo XmOrIssue_Error

20    cmdXM.Enabled = False
30    cmdRemove.Enabled = False
40    cmdAddPending.Enabled = False

50    TimeNow = Now

60    If Trim$(txtUnitNumber) = "" Then
70        Counter = 0
80        For n = 0 To List1.ListCount - 1
90            If List1.Selected(n) Then
100               txtUnitNumber = Left$(List1.List(n), 7)
110               sql = "Select DateExpiry from Product where " & _
                        "Number = '" & txtUnitNumber & "' " & _
                        "order by DateExpiry Desc"
120               Set tb = New Recordset
130               RecOpenServerBB 0, tb, sql
140               If Not tb.EOF Then
150                   lblExpiry = Format(tb!DateExpiry, "dd/MM/yyyy")
160               End If
170               lblProduct = Trim$(Mid$(List1.List(n), 8))
180               pc = ProductBarCodeFor(lblProduct)

190               sql = "Select * from Latest where " & _
                        "BarCode = '" & pc & "' " & _
                        "and Number = '" & txtUnitNumber & "' " & _
                        "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
200               Set tb = New Recordset
210               RecOpenServerBB 0, tb, sql
220               If Not tb.EOF Then
230                   Counter = Counter + 1
240                   lblProduct = Mid$(List1.List(n), 8)
250                   pc = ProductBarCodeFor(lblProduct)
260                   lblExpiry = tb!DateExpiry
270                   lblGroupRh = Bar2Group(tb!GroupRh)
280                   lblScreen = tb!Screen
290                   AmendProduct TimeNow, XorIorK
300                   If lblReaction = "+" Then
310                       Validate "R", "", "", "", "", "", DateAdd("s", 2, Now)
320                   End If
330                   GoSub UpdateXMForm
340               End If
350           End If
360       Next
370       If Counter = 0 Then
380           iMsg "Highlight Units to be Crossmatched.", vbExclamation
390           If TimedOut Then Unload Me: Exit Sub
400           Exit Sub
410       End If
420   Else
430       pc = ProductBarCodeFor(lblProduct)
440       AmendProduct TimeNow, XorIorK
450       If lblReaction = "+" Then
460           Validate "R", "", "", "", "", "", DateAdd("s", 2, Now)
470       End If
480       GoSub UpdateXMForm
490   End If

500   cXM(0) = False
510   cXM(1) = False
520   cXM(2) = False

530   RemoveDetails
540   txtUnitNumber = ""
550   lblStatus = ""

560   fill_list

570   If txtUnitNumber.Visible Then
580       txtUnitNumber.SetFocus
590   End If

600   Exit Sub

UpdateXMForm:

610   XMLine = txtUnitNumber & vbTab & _
               lblGroupRh & vbTab & _
               lblExpiry & vbTab & _
               lblScreen & vbTab
620   If XorIorK = "X" Then
630     XMLine = XMLine & "Xmatched to this Patient"
640   ElseIf XorIorK = "I" Then
650     XMLine = XMLine & "Issued to this Patient"
653   ElseIf XorIorK = "V" Then
656     XMLine = XMLine & "Electronic Issued to this Patient"
660   ElseIf XorIorK = "K" Then
670     XMLine = XMLine & "Awaiting Release"
680   End If
690   XMLine = XMLine & vbTab
700       If cXM(0) Then
710           XMLine = XMLine & IIf(xReaction(0), "O", "+")
720       End If
730       XMLine = XMLine & vbTab
740       If cXM(1) Then
750           XMLine = XMLine & IIf(xReaction(2), "O", "+")
760       End If
770       XMLine = XMLine & vbTab
780       If cXM(2) Then
790           XMLine = XMLine & IIf(xReaction(4), "O", "+")
800       End If
810   XMLine = XMLine & vbTab & _
               lblProduct & vbTab & _
               UserCode & vbTab & _
               Format(TimeNow, "dd/mm/yyyy hh:mm:ss")
820   frmxmatch.gXmatch.AddItem XMLine
830   Return

840   Exit Sub

XmOrIssue_Error:

Dim strES As String
Dim intEL As Integer

850   intEL = Erl
860   strES = Err.Description
870   LogError "frmXMO", "XmOrIssue", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

      Dim n As Long
      Dim sql As String
      Dim Un() As String

10    On Error GoTo cmdCancel_Click_Error

20        Answer = iMsg("Print Labels?", vbQuestion + vbYesNo)
30        If TimedOut Then Unload Me: Exit Sub
40        If Answer = vbYes Then
50            For n = 0 To List1.ListCount - 1
60                If List1.List(n) <> "" Then
70                    Un = Split(List1.List(n), " ")
80                    PrintBarCodesN Un(0), 1, "", "", "", ""
90                End If
100           Next
110       End If
120       Answer = iMsg("Program Analyser?", vbQuestion + vbYesNo)
130       If TimedOut Then Unload Me: Exit Sub
140       If Answer = vbYes Then
150           For n = 0 To List1.ListCount - 1
160               If List1.List(n) <> "" Then
170                   Un = Split(List1.List(n), " ")
180                   sql = "Insert into BBOrderComms " & _
                            "(TestRequired, UnitNumber, SampleID) VALUES " & _
                            "('CrossMatch', '" & Un(0) & "', '" & frmxmatch.tLabNum & "')"
190                   CnxnBB(0).Execute sql
200               End If
210           Next
220       End If

230   Unload Me

240   Exit Sub

cmdCancel_Click_Error:

Dim strES As String
Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmXMO", "cmdcancel_Click", intEL, strES, sql

End Sub

Private Sub cmdIssue_Click()

10    XmOrIssue "I"

End Sub

Private Sub cmdRemove_Click()

      Dim te As Recordset
      Dim tb As Recordset
      Dim tl As Recordset
      Dim sql As String
      Dim pc As String

10    On Error GoTo bremove_Click_Error

20    cmdRemove.Enabled = False

30    pc = ProductBarCodeFor(lblProduct)

40    sql = "Select * from Product where " & _
            "BarCode = '" & pc & "' " & _
            "and Number = '" & txtUnitNumber & "'"
50    Set tl = New Recordset
60    RecOpenServerBB 0, tl, sql

70    sql = "Select * from Latest where " & _
            "BarCode = '" & pc & "' " & _
            "and Number = '" & txtUnitNumber & "'"
80    Set tb = New Recordset
90    RecOpenServerBB 0, tb, sql

100   sql = "Select * from Product where 0 = 1"
110   Set te = New Recordset
120   RecOpenServerBB 0, te, sql
130   te.AddNew
140   te!Number = UCase$(Replace(txtUnitNumber, "+", "X"))
150   te!Event = "R"
160   te!Operator = UserCode
170   te!DateTime = Now
180   te!GroupRh = tl!GroupRh
190   te!BarCode = tl!BarCode
200   te!Supplier = tl!Supplier
210   te!DateExpiry = tl!DateExpiry
220   te!Screen = tl!Screen
230   te!crt = 0
240   te!crtr = 0
250   te!cco = 0
260   te!ccor = 0
270   te!cenr = 0
280   te!cen = 0
290   te!Checked = 0
300   te.Update

310   If Not tb.EOF Then
320       tb!Number = UCase$(Replace(txtUnitNumber, "+", "X"))
330       tb!Event = "R"
340       tb!Operator = UserCode
350       tb!DateTime = Now
360       tb!GroupRh = tl!GroupRh
370       tb!BarCode = tl!BarCode
380       tb!Supplier = tl!Supplier
390       tb!DateExpiry = tl!DateExpiry
400       tb!Screen = tl!Screen
410       tb.Update
420   End If

430   fill_list

440   Exit Sub

bremove_Click_Error:

Dim strES As String
Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "frmXMO", "bremove_Click", intEL, strES, sql

End Sub

Private Sub cmdXM_Click()

10    XmOrIssue "K"

End Sub

Private Sub cmdAddPending_Click()

      Dim sql As String
      Dim ta As Recordset
      Dim tb As Recordset
      Dim tl As Recordset
      Dim pc As String

10    On Error GoTo cmdAddPending_Click_Error

20    pc = ProductBarCodeFor(lblProduct)

30    cmdAddPending.Enabled = False

40    sql = "Select * from Latest where " & _
            "BarCode = '" & pc & "' " & _
            "and Number = '" & txtUnitNumber & "' " & _
            "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
50    Set ta = New Recordset
60    RecOpenServerBB 0, ta, sql
70    If ta.EOF Then
80        iMsg lblProduct & "/" & txtUnitNumber & vbCrLf & "Unit not found", vbExclamation
90        If TimedOut Then Unload Me: Exit Sub
100       Exit Sub
110   End If

120   sql = "Select * from Latest where " & _
            "BarCode = '" & pc & "' " & _
            "and Number = '" & txtUnitNumber & "' " & _
            "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
130   Set tl = New Recordset
140   RecOpenServerBB 0, tl, sql
150   If tl.EOF Then
160       iMsg lblProduct & "/" & txtUnitNumber & vbCrLf & "Unit not found", vbExclamation
170       If TimedOut Then Unload Me: Exit Sub
180       Exit Sub
190   End If

200   sql = "Select * from Product where 0 = 1"
210   Set tb = New Recordset
220   RecOpenServerBB 0, tb, sql

230   tb.AddNew
240   tb!Number = UCase$(Replace(ta!Number, "+", "X"))
250   tb!Event = "P"
260   tb!Patid = frmxmatch.txtChart
270   tb!PatName = frmxmatch.txtName
280   tb!Operator = UserCode
290   tb!DateTime = Now
300   tb!GroupRh = ta!GroupRh
310   tb!BarCode = ta!BarCode
320   tb!Supplier = ta!Supplier
330   tb!DateExpiry = ta!DateExpiry
340   tb!Screen = ta!Screen
350   tb!LabNumber = frmxmatch.tLabNum
360   tb!crt = cXM(0)
370   tb!cco = cXM(1)
380   tb!cen = cXM(2)
390   tb!crtr = xReaction(1)
400   tb!ccor = xReaction(3)
410   tb!cenr = xReaction(5)
420   tb!Checked = ta!Checked
430   tb.Update

440   tl!Number = UCase$(Replace(ta!Number, "+", "X"))
450   tl!Event = "P"
460   tl!Patid = frmxmatch.txtChart
470   tl!PatName = frmxmatch.txtName
480   tl!Operator = UserCode
490   tl!DateTime = Now
500   tl!GroupRh = ta!GroupRh
510   tl!BarCode = ta!BarCode
520   tl!Supplier = ta!Supplier
530   tl!DateExpiry = ta!DateExpiry
540   tl!Screen = ta!Screen
550   tl!LabNumber = frmxmatch.tLabNum
560   tl.Update

570   fill_list

580   lblProduct = ""
590   txtUnitNumber = ""
600   lblGroupRh = ""
610   lblExpiry = ""
620   lblStatus = ""
630   lblChecked = ""
640   cmdLogAsChecked.Visible = False

650   txtUnitNumber.SetFocus

660   Exit Sub

cmdAddPending_Click_Error:

Dim strES As String
Dim intEL As Integer

670   intEL = Erl
680   strES = Err.Description
690   LogError "frmXMO", "cmdAddPending_Click", intEL, strES, sql

End Sub

Private Sub cmdLogAsChecked_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdLogAsChecked_Click_Error

20    sql = "Select * from Latest where " & _
            "BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
            "and Number = '" & txtUnitNumber & "' " & _
            "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60        tb!Checked = True
70        tb!Operator = UserCode
80        tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
90        tb.Update

100       sql = "Insert into Product " & _
                "  Select * from Latest where " & _
                "  BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
                "  and Number = '" & txtUnitNumber & "' " & _
                "  and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
110       CnxnBB(0).Execute sql
120   End If

130   cmdLogAsChecked.Visible = False
140   lblChecked = "Group Checked"

150   Exit Sub

cmdLogAsChecked_Click_Error:

Dim strES As String
Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmXMO", "cmdLogAsChecked_Click", intEL, strES, sql

End Sub

Private Sub cXM_Click(Index As Integer)

      Dim n As Integer
      Dim done As Integer

10    If cXM(Index) = 1 Then
20        xReaction(Index * 2).Enabled = True
30        xReaction(Index * 2 + 1).Enabled = True
40    Else
50        xReaction(Index * 2).Enabled = False
60        xReaction(Index * 2 + 1).Enabled = False
70    End If

80    done = False

90    For n = 0 To 5
100       If xReaction(n) Then done = True
110   Next

120   If done Then
130       cmdXM.Enabled = True
140   Else
150       cmdXM.Enabled = False
160   End If

End Sub

Private Sub fill_list()

      Dim tb As Recordset
      Dim snp As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo fill_list_Error

20    sql = "SELECT DISTINCT Number, BarCode " & _
            "FROM product " & _
            "WHERE LabNumber = '" & frmxmatch.tLabNum & "' " & _
            "AND (Event = 'P' OR Event = 'Y')"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    List1.Clear

60    Do While Not tb.EOF
70        sql = "SELECT * FROM Product WHERE " & _
                "Number = '" & tb!Number & "' " & _
                "AND BarCode = '" & tb!BarCode & "' " & _
                "ORDER BY Counter desc"
80        Set snp = New Recordset
90        RecOpenServerBB 0, snp, sql
100       If (snp!Event = "Y" Or snp!Event = "P") And snp!LabNumber = frmxmatch.tLabNum Then
110           s = snp!Number & " " & ProductWordingFor(snp!BarCode)
120           List1.AddItem s
130       End If
140       tb.MoveNext
150   Loop

160   Exit Sub

fill_list_Error:

Dim strES As String
Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmXMO", "fill_list", intEL, strES, sql

End Sub

Private Sub Form_Load()

      Dim n As Integer

10        frXM.Enabled = True
20        frKey.Visible = False
30        lblIAT.Visible = False
40        lblReaction.Visible = False

50    fill_list

60    For n = 0 To 5
70        xReaction(n).Value = False
80    Next
90    cmdXM.Enabled = False

100   lblProduct = ""

110   cmdRemove.Enabled = False
120   cmdAddPending.Enabled = False

End Sub

Private Sub iKey_Click()

      Dim Reason As String
10    Answer = iMsg("Do you want to enter the Crossmatch manually?", vbQuestion + vbYesNo)
20    If TimedOut Then Unload Me: Exit Sub
30    If Answer = vbYes Then
40        Reason = iBOX("Why?")
50        If TimedOut Then Unload Me: Exit Sub
60        Reason = Trim$(Reason)
70        If Reason = "" Then
80            Exit Sub
90        End If
100       LogReasonWhy Reason, "XM"
110       frXM.Enabled = True
120       frKey.Visible = False
130   End If

End Sub

Private Sub lblReaction_Click()

10    Select Case lblReaction
      Case "": lblReaction = "O": cmdXM.Enabled = True
20    Case "O": lblReaction = "+": cmdXM.Enabled = True
30    Case "+": lblReaction = "": cmdXM.Enabled = False
40    End Select

End Sub

Private Sub lblTag_Click()
10    frmUnitNotes.txtUnitNumber = txtUnitNumber
20    frmUnitNotes.txtExpiry = lblExpiry
30    frmUnitNotes.Show 1
      'LoadDetails

End Sub

Private Sub List1_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Counter As Integer
      Dim pc As String
      Dim X As Integer

10    On Error GoTo List1_Click_Error

20    If List1.ListCount = -1 Then Exit Sub

30    Counter = 0
40    For n = 0 To List1.ListCount - 1
50        If List1.Selected(n) Then
60            Counter = Counter + 1
70            X = InStr(List1.List(n), " ")
80            If X = 0 Then Exit Sub
90            txtUnitNumber = Left$(List1.List(n), X - 1)
100           lblProduct = Mid$(List1.List(n), X + 1)
110       End If
120   Next

130   If Counter <> 1 Then
140       cmdRemove.Enabled = False
150       RemoveDetails
160       cmdXM.Enabled = True
170       txtUnitNumber = ""
180       lblProduct = ""
190       Exit Sub
200   End If

210   pc = ProductBarCodeFor(lblProduct)

220   sql = "Select * from Product  where " & _
            "BarCode = '" & pc & "' " & _
            "and Number = '" & txtUnitNumber & "'"
230   Set tb = New Recordset
240   RecOpenServerBB 0, tb, sql

250   lblGroupRh = Bar2Group(tb!GroupRh & "")

260   lblExpiry = tb!DateExpiry & ""
270   lblScreen = tb!Screen & ""

280   cmdRemove.Enabled = True
290   lblStatus = ""

300   Exit Sub

List1_Click_Error:

Dim strES As String
Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmXMO", "List1_Click", intEL, strES, sql

End Sub

Private Sub RemoveDetails()

      Dim n As Integer

10    For n = 0 To 5
20        xReaction(n) = False
30    Next

40    For n = 0 To 2
50        cXM(n) = False
60    Next

70    lblReaction = ""
80    lblExpiry = ""
90    lblGroupRh = ""
100   lblScreen = ""
110   lblChecked = ""
120   lblProduct = ""
130   cmdLogAsChecked.Visible = False
140   lblKell = ""

End Sub


Private Sub txtUnitNumber_Change()

10    If Trim$(txtUnitNumber) <> "" Then
20        cmdAddPending.Enabled = True
30        cmdRemove.Enabled = True
40    Else
50        cmdAddPending.Enabled = False
60        cmdRemove.Enabled = False
70    End If

80    lblExpiry = ""
90    lblGroupRh = ""
100   lblScreen = ""
110   lblKell = ""
120   lblProduct = ""

End Sub

Private Sub txtUnitNumber_LostFocus()

      Dim Ps As New Products
      Dim f As Form
      Dim p As Product
      Dim Check As String
      
      txtUnitNumber = Replace(txtUnitNumber, "'", "")
      
10    If Trim$(txtUnitNumber) = "" Then Exit Sub
20    If Len(txtUnitNumber) = 9 And UCase$(Left$(txtUnitNumber, 1)) = "D" And UCase$(Right$(txtUnitNumber, 1)) = "D" Then
30        txtUnitNumber = Mid$(txtUnitNumber, 2, 7)
40    End If
50    If Len(txtUnitNumber) = 7 Then
60        txtUnitNumber = UCase$(Replace(txtUnitNumber, "+", "X"))
70        Check = ChkDig(Left$(txtUnitNumber, 6))
80        If Check <> Right$(txtUnitNumber, 1) Then
90            iMsg "Check Digit incorrect!", vbCritical
100           If TimedOut Then Unload Me
110           Exit Sub
120       End If
130   End If

140   Ps.LoadLatestByUnitNumber (txtUnitNumber)

150   If Ps.Count = 0 Then
160       iMsg "Unit Number not found."
170       If TimedOut Then Unload Me: Exit Sub
180       txtUnitNumber = ""
190       Exit Sub
200   ElseIf Ps.Count > 1 Then 'multiple products found
210     Set f = New frmSelectFromMultiple
220     f.ProductList = Ps
230     f.Show 1
240     Set p = f.SelectedProduct
250     Unload f
260     Set f = Nothing
270   Else
280     Set p = Ps.Item(1)
290   End If

300   If p.PackEvent = "Z" Then
310     iMsg "This unit cannot be crossmatched/issued. It is already transfused as Emergency ONeg."
320     If TimedOut Then Unload Me: Exit Sub
330     txtUnitNumber = ""
340     Exit Sub
350   End If

360   FillDetails p

370   If InStr("XP", p.PackEvent) > 0 Then
380       iMsg "This unit is currently crossmatched " & _
               "for a different patient. " & vbCrLf & _
               "Please restock the unit before crossmatching again.", vbCritical
390       If TimedOut Then Unload Me: Exit Sub
400       RemoveDetails
410       lblStatus.Caption = ""
420       txtUnitNumber = ""
430   End If

End Sub



Private Sub xReaction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim done As Integer

10    done = True
      'If trim$(tproduct) = "" Then done = False
      'If trim$(lproduct) = "" Then done = False
      'If trim$(tunitnum) = "" Then done = False

20    If done Then
30        cmdXM.Enabled = True
40    Else
50        cmdXM.Enabled = False
60    End If

End Sub

Private Sub Validate(ByVal EventCode As String, _
                     ByVal PatName As String, _
                     ByVal Chart As String, _
                     ByVal Reason As String, _
                     ByVal EndDateTime As String, _
                     ByVal StartDateTime As String, _
                     ByVal TimeNow As String)

      Dim tLatest As Recordset
      Dim tProduct As Recordset
      Dim sql As String
      Dim ComponentCode As String

10    On Error GoTo Validate_Error

20    ComponentCode = ProductBarCodeFor(lblProduct)

30    If TimeNow = "" Then
40        TimeNow = Format(Now, "dd/mmm/yyyy hh:mm:ss")
50    End If
60    sql = "Select * from Latest where " & _
            "BarCode = '" & ComponentCode & "' " & _
            "and Number = '" & txtUnitNumber & "' " & _
            "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
70    Set tLatest = New Recordset
80    RecOpenServerBB 0, tLatest, sql
90    If Not tLatest.EOF Then
100       tLatest!DateTime = TimeNow
110       If EndDateTime = "" Then
120           tLatest!EventEnd = Null
130       Else
140           tLatest!EventEnd = Format(EndDateTime, "dd/mmm/yyyy hh:mm:ss")
150       End If
160       If StartDateTime <> "" Then
170           tLatest!EventStart = Format(StartDateTime, "dd/mmm/yyyy hh:mm:ss")
180       End If
190       tLatest!Event = EventCode
200       tLatest!Patid = Chart
210       tLatest!PatName = PatName
220       tLatest!Operator = UserCode
230       tLatest!Reason = Reason
240       tLatest.Update
250   End If

260   sql = "Select * from Product where " & _
            "BarCode = '" & ComponentCode & "' " & _
            "and Number = '" & txtUnitNumber & "' " & _
            "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "'"
270   Set tProduct = New Recordset
280   RecOpenServerBB 0, tProduct, sql
290   tProduct.AddNew
300   With tLatest
310       tProduct!Number = UCase$(Replace(!Number, "+", "X"))
320       tProduct!Event = !Event
330       tProduct!Patid = !Patid
340       tProduct!PatName = !PatName
350       tProduct!Operator = !Operator
360       tProduct!DateTime = !DateTime
370       tProduct!GroupRh = !GroupRh
380       tProduct!Supplier = !Supplier
390       tProduct!DateExpiry = !DateExpiry
400       tProduct!Screen = !Screen
410       If EventCode = "S" Then
420           tProduct!LabNumber = !LabNumber
430       Else
440           tProduct!LabNumber = ""
450       End If
460       tProduct!crt = !crt
470       tProduct!cco = !cco
480       tProduct!cen = !cen
490       tProduct!crtr = !crtr
500       tProduct!ccor = !ccor
510       tProduct!cenr = !cenr
520       tProduct!BarCode = !BarCode
530       tProduct!Checked = !Checked
540       tProduct!Notes = !Notes
550       tProduct!Reason = !Reason
560       tProduct!EventStart = !EventStart
570       tProduct!EventEnd = !EventEnd
580       tProduct!OrderNumber = !OrderNumber
590       tProduct.Update
600   End With

610   Exit Sub

Validate_Error:

Dim strES As String
Dim intEL As Integer

620   intEL = Erl
630   strES = Err.Description
640   LogError "frmXMO", "Validate", intEL, strES, sql

End Sub


