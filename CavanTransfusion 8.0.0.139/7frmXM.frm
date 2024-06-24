VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Units Pending"
   ClientHeight    =   8685
   ClientLeft      =   270
   ClientTop       =   525
   ClientWidth     =   9735
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "7frmXM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8685
   ScaleWidth      =   9735
   Begin VB.TextBox txtSQL 
      Height          =   1335
      Left            =   4320
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame fmeProducts 
      Caption         =   "Product Names"
      Height          =   2565
      Left            =   2280
      TabIndex        =   38
      Top             =   2580
      Visible         =   0   'False
      Width           =   4125
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   735
         Left            =   2430
         TabIndex        =   40
         Top             =   1710
         Width           =   765
      End
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   735
         Left            =   3240
         TabIndex        =   39
         Top             =   1710
         Width           =   765
      End
      Begin MSFlexGridLib.MSFlexGrid flxProducts 
         Height          =   2235
         Left            =   150
         TabIndex        =   41
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   5
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2475
      Left            =   150
      TabIndex        =   36
      Top             =   4200
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   4366
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"7frmXM.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdIssue 
      Appearance      =   0  'Flat
      Caption         =   "&Issue"
      Height          =   825
      Left            =   3960
      Picture         =   "7frmXM.frx":096D
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7110
      Width           =   915
   End
   Begin VB.CommandButton cmdXM 
      Appearance      =   0  'Flat
      Caption         =   "X-M"
      Enabled         =   0   'False
      Height          =   825
      Left            =   4950
      Picture         =   "7frmXM.frx":0C77
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7080
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   2010
      TabIndex        =   27
      Top             =   3150
      Width           =   4335
      Begin VB.CommandButton cmdLogAsChecked 
         Caption         =   "Change EI status"
         Height          =   645
         Left            =   2610
         Picture         =   "7frmXM.frx":0F81
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   210
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblChecked 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Frame frXM 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      TabIndex        =   14
      Top             =   6840
      Width           =   4755
      Begin VB.CheckBox cXM 
         Alignment       =   1  'Right Justify
         Caption         =   "Enzyme"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   26
         Top             =   840
         Width           =   915
      End
      Begin VB.CheckBox cXM 
         Alignment       =   1  'Right Justify
         Caption         =   "Coombs"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   25
         Top             =   540
         Width           =   915
      End
      Begin VB.CheckBox cXM 
         Alignment       =   1  'Right Justify
         Caption         =   "Room Temp"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   315
         Index           =   1
         Left            =   1470
         ScaleHeight     =   285
         ScaleWidth      =   885
         TabIndex        =   17
         Top             =   480
         Width           =   915
         Begin VB.OptionButton xReaction 
            Caption         =   "+"
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
            Index           =   3
            Left            =   480
            TabIndex        =   21
            Top             =   60
            Width           =   375
         End
         Begin VB.OptionButton xReaction 
            Alignment       =   1  'Right Justify
            Caption         =   "O"
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
            Index           =   2
            Left            =   30
            TabIndex        =   20
            Top             =   60
            Width           =   405
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   315
         Index           =   0
         Left            =   1470
         ScaleHeight     =   285
         ScaleWidth      =   885
         TabIndex        =   16
         Top             =   180
         Width           =   915
         Begin VB.OptionButton xReaction 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   480
            TabIndex        =   19
            Top             =   30
            Width           =   375
         End
         Begin VB.OptionButton xReaction 
            Alignment       =   1  'Right Justify
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   18
            Top             =   30
            Width           =   405
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   315
         Index           =   2
         Left            =   1470
         ScaleHeight     =   285
         ScaleWidth      =   885
         TabIndex        =   15
         Top             =   780
         Width           =   915
         Begin VB.OptionButton xReaction 
            Caption         =   "+"
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
            Index           =   5
            Left            =   480
            TabIndex        =   23
            Top             =   60
            Width           =   405
         End
         Begin VB.OptionButton xReaction 
            Alignment       =   1  'Right Justify
            Caption         =   "O"
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
            Index           =   4
            Left            =   30
            TabIndex        =   22
            Top             =   60
            Width           =   405
         End
      End
   End
   Begin VB.TextBox txtUnitNumber 
      BackColor       =   &H00FFFFFF&
      DataField       =   "number"
      DataSource      =   "Data1"
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
      Left            =   1140
      MaxLength       =   16
      TabIndex        =   0
      Top             =   210
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   1095
      Left            =   8430
      Picture         =   "7frmXM.frx":13C3
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6930
      Width           =   1155
   End
   Begin VB.CommandButton cmdRemove 
      Appearance      =   0  'Flat
      Caption         =   "&Remove from pending list"
      Height          =   1005
      Left            =   8415
      Picture         =   "7frmXM.frx":228D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5715
      Width           =   1185
   End
   Begin VB.CommandButton cmdAddPending 
      Appearance      =   0  'Flat
      Caption         =   "&Mark as Pending"
      Height          =   1005
      Left            =   8430
      Picture         =   "7frmXM.frx":26CF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1185
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   135
      TabIndex        =   35
      Top             =   8310
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ISBT128"
      Height          =   195
      Left            =   465
      TabIndex        =   37
      Top             =   270
      Width           =   630
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This Unit is Tagged. Click here to view."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1140
      TabIndex        =   34
      Top             =   570
      Visible         =   0   'False
      Width           =   6465
   End
   Begin VB.Label lblKell 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6300
      TabIndex        =   33
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Kell"
      Height          =   195
      Left            =   6000
      TabIndex        =   32
      Top             =   1590
      Width           =   255
   End
   Begin VB.Label lblProduct 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1140
      TabIndex        =   13
      Top             =   1080
      Width           =   6465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   540
      TabIndex        =   12
      Top             =   1170
      Width           =   555
   End
   Begin VB.Label lblGroupRh 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3990
      TabIndex        =   11
      Top             =   1530
      Width           =   915
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1140
      TabIndex        =   10
      Top             =   1530
      Width           =   1875
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Screen"
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   2070
      Width           =   510
   End
   Begin VB.Label lblScreen 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1140
      TabIndex        =   8
      Top             =   1920
      Width           =   6465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   690
      TabIndex        =   7
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Group/Rh"
      Height          =   195
      Left            =   3240
      TabIndex        =   6
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1140
      TabIndex        =   5
      Top             =   2520
      Width           =   6465
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   660
      TabIndex        =   4
      Top             =   2640
      Width           =   450
   End
End
Attribute VB_Name = "frmXM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mElectronicIssue As Boolean
Private m_SampleID As String
Private m_Flag As Boolean
Private m_UID As String
Private m_ISBT As String

Private Const fcsLine_NO = 0
Private Const fcsRID = 1
Private Const fcsSID = 2
Private Const fcsUID = 3
Private Const fcsPrd = 4


Private Sub FormatGrid()
    On Error GoTo ERROR_FormatGrid
    
    flxProducts.Rows = 1
    flxProducts.row = 0
    
    flxProducts.ColWidth(fcsLine_NO) = 250
    
    flxProducts.TextMatrix(0, fcsRID) = ""
    flxProducts.ColWidth(fcsRID) = 0
    flxProducts.ColAlignment(fcsRID) = flexAlignLeftCenter
    
    flxProducts.TextMatrix(0, fcsSID) = ""
    flxProducts.ColWidth(fcsSID) = 0
    flxProducts.ColAlignment(fcsSID) = flexAlignLeftCenter
    
    flxProducts.TextMatrix(0, fcsUID) = ""
    flxProducts.ColWidth(fcsUID) = 0
    flxProducts.ColAlignment(fcsUID) = flexAlignLeftCenter
    
    flxProducts.TextMatrix(0, fcsPrd) = "Products"
    flxProducts.ColWidth(fcsPrd) = 1550
    flxProducts.ColAlignment(fcsPrd) = flexAlignLeftCenter
    
        
    Exit Sub
ERROR_FormatGrid:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "FormatGrid", intEL, strES
End Sub


Public Property Let ElectronicIssue(ByVal sNewValue As Boolean)

10  mElectronicIssue = sNewValue

End Property

Public Function GetElectronicIssue() As Boolean

10 GetElectronicIssue = mElectronicIssue
End Function

Private Sub AmendProduct(TimeNow As String, bbEvent As String)

    Dim tb As Recordset
    Dim sql As String
    Dim pc As String

10  On Error GoTo AmendProduct_Error

20  pc = ProductBarCodeFor(Trim$(lblProduct))

30  sql = "Select * from Latest where " & _
          "BarCode = '" & pc & "' " & _
          "and ISBT128 = '" & txtUnitNumber & "' "
40  Set tb = New Recordset
50  RecOpenServerBB 0, tb, sql
60  If Not tb.EOF Then
70      tb!crt = cXM(0)
80      tb!cco = cXM(1)
90      tb!cen = cXM(2)
100     tb!crtr = xReaction(1)
110     tb!ccor = xReaction(3)
120     tb!cenr = xReaction(5)
130     tb!Event = bbEvent
140     tb!Patid = frmxmatch.txtChart
150     tb!PatName = frmxmatch.txtName
160     tb!Operator = UserCode
170     tb!DateTime = TimeNow
180     tb!LabNumber = frmxmatch.tLabNum

190     tb.Update

200     sql = "Insert into Product " & _
              "  Select * from Latest where " & _
              "  BarCode = '" & pc & "' " & _
              "  and ISBT128 = '" & txtUnitNumber & "' "
210     CnxnBB(0).Execute sql

220 End If

230 Exit Sub

AmendProduct_Error:

    Dim strES As String
    Dim intEL As Integer

240 intEL = Erl
250 strES = Err.Description
260 LogError "frmXM", "AmendProduct", intEL, strES, sql


End Sub

Private Sub FillDetails(ByVal p As Product)

    Dim sql As String
    Dim GRh As String
    Dim tb As Recordset

10  lblTag.Visible = False

20  lblExpiry = Format(p.DateExpiry, "dd/mm/yyyy HH:mm")
30  lblProduct = ProductWordingFor(p.BarCode)

40  If Not mElectronicIssue Then
50      If Trim$(frmxmatch.txtChart) <> "" Then
60          sql = "Select * from Product where " & _
                  "ISBT128 = '" & Trim$(txtUnitNumber) & "' " & _
                  "and PatID = '" & frmxmatch.txtChart & "' "
70          sql = sql & "and crt = 1 and crtr = 1 "
80          sql = sql & "AND DateExpiry = '" & Format(p.DateExpiry, "dd/MMM/yyyy HH:mm") & "' and Barcode = '" & p.BarCode & "'"

90          Set tb = New Recordset
100         RecOpenServerBB 0, tb, sql
110         If Not tb.EOF Then
120             iMsg "Unit has been found incompatible for this patient and cannot be issued!", vbExclamation
130             If TimedOut Then Unload Me: Exit Sub
140             txtUnitNumber = ""
150             Exit Sub
160         End If
170     End If
180 End If

190 If InStr(p.Screen, "K+") Then
200     lblKell = "Positive"
210 ElseIf InStr(p.Screen & "", "K-") Then
220     lblKell = "Negative"
230 Else
240     lblKell = ""
250 End If
260 If Trim$(lblProduct) = "" Then Exit Sub

270 If Format(Now, "yyyymmddhhmm") > Format(lblExpiry, "yyyymmddhhmm") Then
    '230   If DateDiff("d", Now, lblexpiry) < 0 Then
280     iMsg "Unit expired!", vbExclamation
290     If TimedOut Then Unload Me: Exit Sub
300     RemoveDetails
310     txtUnitNumber = ""
320     Exit Sub
330 End If

340 GRh = p.GroupRh

350 If Not Allowed(GRh) Then
360     txtUnitNumber = ""
370     Exit Sub
380 End If

390 WarnAboutGroup GRh, lblKell

400 If Left$(frmxmatch.lstfg & "  ", 2) <> "O " Then
410     If InStr(UCase$(p.Screen & ""), "O RECIPIENT ONLY") Then
420         iMsg "Screen: " & tb!Screen, vbInformation
430         If TimedOut Then Unload Me: Exit Sub
440     End If
450 End If

    'IF Patient Sample and Unit Elegible
460 If mElectronicIssue Then
470     If Not getSampleEligibility4EI(frmxmatch.tLabNum) Then
480         iMsg "Sample Not Eligible for Electronic Issue!", vbExclamation
490         If TimedOut Then Unload Me: Exit Sub
500         RemoveDetails
510         txtUnitNumber = ""
520         Exit Sub
530     End If

540     If Not getUnitEligibilityEI(Trim$(txtUnitNumber)) Then
550         iMsg "Unit Not Eligible for Electronic Issue!", vbExclamation
560         If TimedOut Then Unload Me: Exit Sub
570         RemoveDetails
580         txtUnitNumber = ""
590         Exit Sub
600     End If

610 End If

620 lblGroupRh = Bar2Group(GRh)

630 lblScreen = p.Screen

640 Select Case p.PackEvent
        Case "Y":
650         lblStatus = "Removed Pending Transfusion."
660         cmdRemove.Enabled = True
670     Case "P":
680         lblStatus = "Pending for "
690         If p.SampleID = frmxmatch.tLabNum Then
700             lblStatus = lblStatus & "this patient."
710             cmdRemove.Enabled = True
720         Else
730             lblStatus = lblStatus & "another patient."
740             cmdAddPending.Enabled = True
750         End If
760     Case "D": lblStatus = "This unit number has been destroyed."
770     Case "S": lblStatus = "This unit has been Transfused."
780     Case "T": lblStatus = "This unit has been returned."
790     Case "X": lblStatus = "Crossmatched for "
800         If p.Chart = frmxmatch.txtChart Or p.PatName & "" = frmxmatch.txtName Then
810             lblStatus = lblStatus & "this patient."
820         Else
830             lblStatus = lblStatus & "another patient."
840         End If
850         cmdAddPending.Enabled = True
860     Case "I": lblStatus = "Issued to "
870         If p.Chart = frmxmatch.txtChart Or p.PatName & "" = frmxmatch.txtName Then
880             lblStatus = lblStatus & "this patient."
890         Else
900             lblStatus = lblStatus & "another patient."
910         End If
920         cmdAddPending.Enabled = True

930     Case "V": lblStatus = "Electronic Issue to "
940         If p.Chart = frmxmatch.txtChart Or p.PatName & "" = frmxmatch.txtName Then
950             lblStatus = lblStatus & "this patient."
960         Else
970             lblStatus = lblStatus & "another patient."
980         End If
990         cmdAddPending.Enabled = True

1000    Case "C", "R": lblStatus = "In free stock."
1010        cmdAddPending.Enabled = True

1020 End Select

1030 If p.Checked Then
1040    lblChecked = "EI suitable"    '"Group  Checked"
1050    cmdLogAsChecked.Visible = False
1060 Else
1070    lblChecked = "Not EI suitable"    ' "Group Not Checked"
1080    cmdLogAsChecked.Visible = True
1090 End If

1100 If TagIsPresent(txtUnitNumber, CDate(lblExpiry)) Then
1110    lblTag.Visible = True
1120 End If

End Sub

Private Function getSampleEligibility4EI(ByVal strSID As String) As Boolean
    Dim tb As Recordset
    Dim sql As String

    On Error GoTo getSampleEligibility4EI_Error

10  getSampleEligibility4EI = False

    'Is sample Eligible for EI  (either Automatically Eligible or Forced Eligible)
20  sql = "Select Eligible4EI from PatientDetails where " & _
          "labnumber = '" & strSID & "' and (Eligible4EI = 'A_E' or Eligible4EI = 'F_E')"
30  Set tb = New Recordset
40  RecOpenServerBB 0, tb, sql

50  If Not tb.EOF Then
90      getSampleEligibility4EI = True
110 Else
120     getSampleEligibility4EI = False
130 End If

    Exit Function

getSampleEligibility4EI_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmXM", "getSampleEligibility4EI", intEL, strES, sql

End Function


Private Function getUnitEligibilityEI(ByVal strUnit As String) As Boolean
    Dim tb As Recordset
    Dim sql As String

    On Error GoTo getUnitEligibilityEI_Error

10  getUnitEligibilityEI = False

20  sql = "Select Checked from Latest where " & _
          "ISBT128 = '" & strUnit & "' and Checked = 1"
30  Set tb = New Recordset
40  RecOpenServerBB 0, tb, sql

50  If Not tb.EOF Then
60      getUnitEligibilityEI = True
70  Else
80      getUnitEligibilityEI = False
90  End If

    Exit Function

getUnitEligibilityEI_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmXM", "getUnitEligibilityEI", intEL, strES, sql

End Function


Private Function Allowed(ByVal Unit As String) As Integer

    Dim Patient As String
    Dim s As String
    Dim A As Integer
    Dim pGroup As String
    Dim prh As String
    Dim PatientABO_Group As String

10  On Error GoTo Allowed_Error

20  image2grh pGroup, prh

30  If InStr(UCase$(frmxmatch.lstfg), "NEG") _
       Or InStr(UCase$(frmxmatch.lstfg), "DVI") _
       Or InStr(frmxmatch.lstfg, "C/E") Then
40      Patient = "-"
50      Patient = Trim$(Left$(frmxmatch.lstfg, 2)) & Patient
60  ElseIf InStr(UCase$(frmxmatch.lstfg), "POS") Then
70      Patient = "+"
80      Patient = Trim$(Left$(frmxmatch.lstfg, 2)) & Patient
90  Else
100     Patient = prh
110     s = "Confirm Patient Group/Rhesus? " & _
            vbCrLf & vbCrLf & _
            Space(10) & pGroup
120     If Patient = "-" Then
130         s = s & "Negative"
140     ElseIf Patient = "+" Then
150         s = s & "Positive"
160     Else
170         s = s & "Unknown"
180     End If
190     Answer = iMsg(s, vbYesNo + vbQuestion)
200     If TimedOut Then Unload Me: Exit Function
210     If Answer = vbNo Then
220         Allowed = False
230         Exit Function
240     Else
250         Patient = pGroup & Patient
260     End If
270 End If

280 PatientABO_Group = ABOGroup(Patient)

290 A = False
300 If InStr(frmXM.lblProduct, "Plasma") Then
310     A = True
320 ElseIf InStr(UCase$(ProductGenericFor(ProductBarCodeFor(frmXM.lblProduct))), "PLATELET") Then
330     A = True
340 ElseIf InStr(UCase$(ProductGenericFor(ProductBarCodeFor(frmXM.lblProduct))), "LG OCTAPLAS") Then
350     Select Case PatientABO_Group
            Case "O": If Unit = "55" Then A = True    'O
380                   If Unit = "88" Then A = True    'AB
390         Case "A": If Unit = "66" Then A = True    'A
400                   If Unit = "88" Then A = True    'AB
410         Case "B": If Unit = "77" Then A = True    'B
420                   If Unit = "88" Then A = True    'AB
430         Case "AB": If Unit = "88" Then A = True    'AB
440         Case "UN": If Unit = "88" Then A = True    'AB
450     End Select
        'LG OCTAPLAS Rules
        'Patient    Product
        'O          O AB
        'A          A AB
        'B          B AB
        'AB         AB

460 Else
470     Select Case Patient
            Case "", "O-": If Unit = "95" Then A = True
480             If Unit = "51" Then A = True
490         Case "O+": If Unit = "51" Then A = True
500             If Unit = "95" Then A = True
510         Case "A-": If Unit = "06" Then A = True
520             If Unit = "95" Then A = True
530             If Unit = "62" Then A = True
540             If Unit = "51" Then A = True
550         Case "A+": If Unit = "62" Then A = True
560             If Unit = "06" Then A = True
570             If Unit = "51" Then A = True
580             If Unit = "95" Then A = True
590         Case "B-": If Unit = "17" Then A = True
600             If Unit = "95" Then A = True
610             If Unit = "73" Then A = True
620             If Unit = "51" Then A = True
630         Case "B+": If Unit = "73" Then A = True
640             If Unit = "17" Then A = True
650             If Unit = "51" Then A = True
660             If Unit = "95" Then A = True
670         Case "AB-": A = True
680         Case "AB+": A = True
690     End Select
700 End If

710 If Not A Then
720     s = "This unit is not compatible " & vbCrLf & _
            "with this patient and cannot" & vbCrLf & _
            "be crossmatched and must not" & vbCrLf & _
            "be transfused."
730     iMsg s, vbCritical
740     If TimedOut Then Unload Me: Exit Function
750 End If

760 Allowed = A

770 Exit Function

Allowed_Error:

    Dim strES As String
    Dim intEL As Integer

780 intEL = Erl
790 strES = Err.Description
800 LogError "frmXM", "Allowed", intEL, strES

End Function

Public Function ABOGroup(ByVal strBloodGroup) As String

10  Select Case UCase(strBloodGroup)
        Case "O-": ABOGroup = "O"
20      Case "O+": ABOGroup = "O"
30      Case "A-": ABOGroup = "A"
40      Case "A+": ABOGroup = "A"
50      Case "B-": ABOGroup = "B"
60      Case "B+": ABOGroup = "B"
70      Case "AB-": ABOGroup = "AB"
80      Case "AB+": ABOGroup = "AB"
90      Case "UN": ABOGroup = "UN"
100 End Select

End Function

Private Sub UpdateXMForm(TimeNow As String, ByVal XorIorK As String)

    Dim XMLine As String

10  XMLine = txtUnitNumber & vbTab & _
             lblGroupRh & vbTab & _
             lblExpiry & vbTab & _
             lblScreen & vbTab
20  If XorIorK = "X" Then
30      XMLine = XMLine & "Xmatched to this Patient"
40  ElseIf XorIorK = "I" Then
50      XMLine = XMLine & "Issued to this Patient"
60  ElseIf XorIorK = "K" Then
70      XMLine = XMLine & "Awaiting Release"
80  ElseIf XorIorK = "V" Then
90      XMLine = XMLine & "EI to this Patient"
100 End If
110 XMLine = XMLine & vbTab
120 If cXM(0) Then
130     XMLine = XMLine & IIf(xReaction(0), "O", "+")
140 End If
150 XMLine = XMLine & vbTab
160 If cXM(1) Then
170     XMLine = XMLine & IIf(xReaction(2), "O", "+")
180 End If
190 XMLine = XMLine & vbTab
200 If cXM(2) Then
210     XMLine = XMLine & IIf(xReaction(4), "O", "+")
220 End If
230 XMLine = XMLine & vbTab & _
             lblProduct & vbTab & _
             UserCode & vbTab & _
             Format(TimeNow, "dd/mm/yyyy hh:mm:ss")

240 frmxmatch.gXmatch.AddItem XMLine

End Sub

Private Sub WarnAboutGroup(ByVal GRh As String, ByVal UnitKell As String)

    Dim RightGroup As Integer
    Dim s As String

10  RightGroup = True

20  Select Case Left$(frmxmatch.lstfg & "  ", 2)
        Case "O ": If GRh <> "51" And GRh <> "95" Then RightGroup = False
30      Case "A ": If GRh <> "62" And GRh <> "06" Then RightGroup = False
40      Case "B ": If GRh <> "73" And GRh <> "17" Then RightGroup = False
50      Case "AB": If GRh <> "84" And GRh <> "28" Then RightGroup = False
60  End Select

70  If InStr(UCase$(frmxmatch.lstfg), "NEG") And (GRh = "51" Or GRh = "62" Or GRh = "73" Or GRh = "84") Then RightGroup = False
80  If InStr(UCase$(frmxmatch.lstfg), "POS") And (GRh = "95" Or GRh = "06" Or GRh = "17" Or GRh = "28") Then RightGroup = False

90  If Not RightGroup Then
100     iMsg "Donor unit and Patient group differ.", vbInformation
110     If TimedOut Then Unload Me: Exit Sub
120     If Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "50y" Then
130         s = "It is recommended practice to" & vbCrLf & _
                "transfuse fully compatible blood" & vbCrLf & _
                "to this category of patient." & vbCrLf & _
                "(Female of less than 50 years)" & vbCrLf & _
                "This practice is not being followed."
140         iMsg s, vbInformation
150         If TimedOut Then Unload Me: Exit Sub
160     End If
170 End If

180 If Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "60y" And frmxmatch.cmbKell = "K-" And UnitKell = "Positive" Then
190     s = "It is recommended practice to" & vbCrLf & _
            "transfuse Kell Negative blood" & vbCrLf & _
            "to this category of patient." & vbCrLf & _
            "(Female of less than 60 years K-)" & vbCrLf & _
            "This practice is not being followed."
200     iMsg s, vbInformation
210     If TimedOut Then Unload Me: Exit Sub
220 ElseIf Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "60y" And frmxmatch.cmbKell = "K-" And UnitKell = "" Then
230     iMsg "Unit of unknown Kell Status." & vbCrLf & _
             "Kell Negative blood should be given to this Patient. ", vbInformation
240     If TimedOut Then Unload Me: Exit Sub
250 ElseIf Left$(frmxmatch.lSex, 1) = "F" And frmxmatch.tAge < "60y" And frmxmatch.cmbKell = "" And UnitKell <> "Negative" Then
260     iMsg "Patient of unknown Kell Status." & vbCrLf & _
             "Kell Negative blood should be given to this Patient. ", vbInformation
270     If TimedOut Then Unload Me: Exit Sub
280 End If

End Sub

Private Sub XmOrIssue(ByVal XorIorK As String)

    Dim sql As String
    Dim TimeNow As String
    Dim Counter As Integer
    Dim pc As String
    Dim Y As Integer

10  On Error GoTo XmOrIssue_Error

20  cmdXM.Enabled = False
30  cmdRemove.Enabled = False
40  cmdAddPending.Enabled = False

50  TimeNow = Now

60  If Trim$(txtUnitNumber) = "" Then
70      Counter = 0
80      g.col = 0
90      For Y = 1 To g.Rows - 1
100         g.row = Y
110         If g.CellBackColor = vbYellow Then
120             txtUnitNumber = g.TextMatrix(Y, 0)
                m_ISBT = g.TextMatrix(Y, 0)
130             lblExpiry = g.TextMatrix(Y, 1)
140             lblGroupRh = g.TextMatrix(Y, 2)
150             lblProduct = g.TextMatrix(Y, 3)
160             txtUnitNumber = g.TextMatrix(Y, 0)
170             lblScreen = g.TextMatrix(Y, 4)
180             pc = ProductBarCodeFor(lblProduct)
190             Counter = Counter + 1
200             AmendProduct TimeNow, XorIorK
210             UpdateXMForm TimeNow, XorIorK
                
220         End If
230     Next
240     If Counter = 0 Then
250         iMsg "Highlight Units to be Crossmatched.", vbExclamation
260         If TimedOut Then Unload Me: Exit Sub
270         Exit Sub
280     End If
290 Else
300     pc = ProductBarCodeFor(lblProduct)
310     AmendProduct TimeNow, XorIorK
320     UpdateXMForm TimeNow, XorIorK
330 End If
    m_ISBT = txtUnitNumber.Text
    If m_Flag Then
        If m_UID <> "" Then
'            Call UpdateIdentifier(m_UID, m_ISBT)
'            Call CountAndUpdate(Trim(frmxmatch.tLabNum))
            m_Flag = False
            m_UID = ""
        End If
    End If
    
340 cXM(0) = False
350 cXM(1) = False
360 cXM(2) = False

370 RemoveDetails
380 txtUnitNumber = ""
390 lblStatus = ""

400 FillG

410 If txtUnitNumber.Visible Then
420     txtUnitNumber.SetFocus
430 End If

440 Exit Sub

XmOrIssue_Error:

    Dim strES As String
    Dim intEL As Integer

450 intEL = Erl
460 strES = Err.Description
MsgBox Err.Description & " " & "XmOrIssue"
470 LogError "frmXM", "XmOrIssue", intEL, strES, sql

End Sub

Private Sub btnCancel_Click()
    fmeProducts.Visible = False
    m_Flag = False
End Sub

Private Sub btnOK_Click()
    m_Flag = True
    m_UID = flxProducts.TextMatrix(flxProducts.row, fcsUID)
    Call cmdIssue_Click
    fmeProducts.Visible = False
End Sub

Private Sub cmdCancel_Click()

    Dim Y As Integer
    Dim sql As String
    Dim Un As String

10  On Error GoTo cmdCancel_Click_Error


20  If g.TextMatrix(1, 0) <> "" Then
25      If Not mElectronicIssue Then
30          Answer = iMsg("Print Labels?", vbQuestion + vbYesNo)
40          If TimedOut Then Unload Me: Exit Sub
50          If Answer = vbYes Then
60              For Y = 1 To g.Rows - 1
70                  If Trim$(g.TextMatrix(Y, 0)) <> "" Then
80                      Un = g.TextMatrix(Y, 0)
90                      With frmPrintBarCodeUnit
100                         .Unit = Un
110                         .Show 1
120                     End With
130                 End If
140             Next
150         End If
160         Answer = iMsg("Program Analyser?", vbQuestion + vbYesNo)
170         If TimedOut Then Unload Me: Exit Sub
180         If Answer = vbYes Then
190             For Y = 1 To g.Rows - 1
200                 If g.TextMatrix(Y, 0) <> "" Then
210                     Un = g.TextMatrix(Y, 0)
220                     sql = "IF NOT EXISTS (SELECT * FROM BBOrderComms " & _
                              "               WHERE TestRequired = 'Crossmatch' " & _
                              "               AND UnitNumber = '" & Un & "' " & _
                              "               AND SampleID = '" & frmxmatch.tLabNum & "') " & _
                              "  INSERT INTO BBOrderComms " & _
                              "  (TestRequired, UnitNumber, SampleID) VALUES " & _
                              "  ('CrossMatch', '" & Un & "', '" & frmxmatch.tLabNum & "')"
230                     CnxnBB(0).Execute sql
240                 End If
250             Next
260         End If
265     End If
270 End If

280 Unload Me

290 Exit Sub

cmdCancel_Click_Error:

    Dim strES As String
    Dim intEL As Integer

300 intEL = Erl
310 strES = Err.Description
320 LogError "frmXM", "cmdCancel_Click", intEL, strES, sql

End Sub

Private Sub cmdIssue_Click()
10      On Error GoTo cmdIssue_Click_Error
'    If m_Flag = False Then
'        If frmxmatch.tLabNum <> "" Then
''            Call GetProducts(Trim(frmxmatch.tLabNum))
''            fmeProducts.Visible = True
'            Exit Sub
'        End If
'    End If
    'MsgBox (GetElectronicIssue)
11  If Not GetElectronicIssue Then    'cmdIssue.Caption = "&Issue" Then
20      'MsgBox ("In Condition")
        XmOrIssue "I"
30  Else
        'MsgBox (" Condition")
40      XmOrIssue "V"
50  End If

cmdIssue_Click_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
    MsgBox (strES & " " & intEL)
150 LogError "frmXM", "cmdRemove_Click", intEL, strES, ""

    

End Sub

Private Sub cmdRemove_Click()

    Dim tl As Recordset
    Dim sql As String
    Dim pc As String

10  On Error GoTo cmdRemove_Click_Error

20  cmdRemove.Enabled = False

30  pc = ProductBarCodeFor(lblProduct)

40  sql = "Select * from Product where " & _
          "BarCode = '" & pc & "' " & _
          "and ISBT128 = '" & txtUnitNumber & "' "
50  Set tl = New Recordset
60  RecOpenServerBB 0, tl, sql

    'Add new record tp PRODUCT table
70  sql = "Insert into Product " & _
          "( Number, [Event], Operator, [DateTime], " & _
          "  GroupRh, BarCode, Supplier, DateExpiry, Screen, crt, cco, cen, crtr, ccor, cenr, Checked, ISBT128 ) VALUES " & _
          "( '" & UCase$(Replace(tl!Number & "", "+", "X")) & "', " & _
          "  '" & "R" & "', " & _
          "  '" & UserCode & "', " & _
          "  '" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
          "  '" & tl!GroupRh & "', " & _
          "  '" & tl!BarCode & "', " & _
          "  '" & tl!Supplier & "', " & _
          "  '" & Format(tl!DateExpiry, "dd/mmm/yyyy HH:mm") & "', " & _
          "  '" & tl!Screen & "', " & _
          "  '" & 0 & "', " & _
          "  '" & 0 & "', " & _
          "  '" & 0 & "', " & _
          "  '" & 0 & "', " & _
          "  '" & 0 & "', " & _
          "  '" & 0 & "', " & _
          "  '" & 0 & "', " & _
          "  '" & Trim$(tl!ISBT128 & "") & " ' )"
80  CnxnBB(0).Execute sql

    'Update LATEST record

90  sql = "Update Latest " & _
          "Set Event = 'R', Operator = '" & UserCode & "', Datetime =  '" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "' , GroupRh = '" & tl!GroupRh & "', " & _
          "BarCode = '" & tl!BarCode & "', Supplier = '" & tl!Supplier & "', DateExpiry = '" & Format(tl!DateExpiry, "dd/mmm/yyyy HH:mm") & "', " & _
          "Screen = '" & tl!Screen & "', crt = 0, cco = 0 , cen = 0, crtr = 0, ccor = 0, cenr = 0 " & _
          "where " & _
          "ISBT128 = '" & txtUnitNumber & "' and " & _
          "BarCode = '" & ProductBarCodeFor(lblProduct) & "' "
100 CnxnBB(0).Execute sql

110 FillG

120 Exit Sub

cmdRemove_Click_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
    MsgBox (strES & " " & intEL)
150 LogError "frmXM", "cmdRemove_Click", intEL, strES, sql

End Sub

Private Sub cmdXM_Click()

10  XmOrIssue "K"

End Sub

Private Sub cmdAddPending_Click()

    Dim sql As String
    Dim ta As Recordset
    Dim tl As Recordset
    Dim pc As String

10  On Error GoTo cmdAddPending_Click_Error

20  pc = ProductBarCodeFor(lblProduct)

30  cmdAddPending.Enabled = False

40  sql = "Select * from Latest where " & _
          "BarCode = '" & pc & "' " & _
          "and ISBT128 = '" & txtUnitNumber & "' "
50  Set ta = New Recordset
60  RecOpenServerBB 0, ta, sql
70  If ta.EOF Then
80      iMsg lblProduct & "/" & txtUnitNumber & vbCrLf & "Unit not found", vbExclamation
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110 End If

120 sql = "Select * from Latest where " & _
          "BarCode = '" & pc & "' " & _
          "and ISBT128 = '" & txtUnitNumber & "' "
130 Set tl = New Recordset
140 RecOpenServerBB 0, tl, sql
150 If tl.EOF Then
160     iMsg lblProduct & "/" & txtUnitNumber & vbCrLf & "Unit not found", vbExclamation
170     If TimedOut Then Unload Me: Exit Sub
180     Exit Sub
190 End If

200 sql = "Insert into Product " & _
          "( Number, [Event], Patid, PatName, Operator, [DateTime], " & _
          "  GroupRh, BarCode, Supplier, DateExpiry, Screen, LabNumber, crt, cco, cen, crtr, ccor, cenr, Checked, ISBT128 ) VALUES " & _
          "( '" & UCase$(Replace(ta!Number, "+", "X")) & "', " & _
          "  '" & "P" & "', " & _
          "  '" & frmxmatch.txtChart & "', " & _
          "  '" & AddTicks(frmxmatch.txtName) & "', " & _
          "  '" & UserCode & "', " & _
          "  '" & Format(Now, "dd/mmm/yyyy HH:nn:ss") & "', " & _
          "  '" & ta!GroupRh & "', " & _
          "  '" & ta!BarCode & "', " & _
          "  '" & ta!Supplier & "', " & _
          "  '" & Format(ta!DateExpiry, "dd/mmm/yyyy HH:mm") & "', " & _
          "  '" & ta!Screen & "', " & _
          "  '" & frmxmatch.tLabNum & "', " & _
          "  '" & cXM(0) & "', " & _
          "  '" & cXM(1) & "', " & _
          "  '" & cXM(2) & "', " & _
          "  '" & IIf(xReaction(1), 1, 0) & "', " & _
          "  '" & IIf(xReaction(3), 1, 0) & "', " & _
          "  '" & IIf(xReaction(5), 1, 0) & "', " & _
          "  '" & IIf(ta!Checked, 1, 0) & "', " & _
          "  '" & ta!ISBT128 & "" & "' )"

210 CnxnBB(0).Execute sql


220 tl!Number = UCase$(Replace(ta!Number, "+", "X"))
230 tl!Event = "P"
240 tl!Patid = frmxmatch.txtChart
250 tl!PatName = frmxmatch.txtName
260 tl!Operator = UserCode
270 tl!DateTime = Now
280 tl!GroupRh = ta!GroupRh
290 tl!BarCode = ta!BarCode
300 tl!Supplier = ta!Supplier
310 tl!DateExpiry = ta!DateExpiry
320 tl!Screen = ta!Screen
330 tl!LabNumber = frmxmatch.tLabNum
340 tl.Update

350 FillG

360 lblProduct = ""
370 txtUnitNumber = ""
380 lblGroupRh = ""
390 lblExpiry = ""
400 lblStatus = ""
410 lblChecked = ""
420 cmdLogAsChecked.Visible = False

430 txtUnitNumber.SetFocus

440 Exit Sub

cmdAddPending_Click_Error:

    Dim strES As String
    Dim intEL As Integer

450 intEL = Erl
460 strES = Err.Description
470 LogError "frmXM", "cmdAddPending_Click", intEL, strES, sql

End Sub

Private Sub cmdLogAsChecked_Click()

' Dim tb As Recordset
    Dim sql As String

10  On Error GoTo cmdLogAsChecked_Click_Error

    '20    Sql = "Select * from Latest where " & _
     '          "BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
     '          "and ISBT128 = '" & txtUnitNumber & "' "
    '30    Set tb = New Recordset
    '40    RecOpenServerBB 0, tb, Sql
    '50    If Not tb.EOF Then
    '60      tb!Checked = True
    '70      tb!Operator = UserCode
    '80      tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
    '90      tb.Update
    '
    '100     Sql = "Insert into Product " & _
     '            "  Select * from Latest where " & _
     '            "  BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
     '            "  and ISBT128 = '" & txtUnitNumber & "' "
    '110     CnxnBB(0).Execute Sql
    '120   End If

20  sql = "UPDATE Product SET Checked = '1'  WHERE " & _
          "ISBT128 = '" & txtUnitNumber & "' AND BarCode = '" & ProductBarCodeFor(lblProduct) & "' "
30  CnxnBB(0).Execute sql
40  sql = "UPDATE Latest SET Checked = '1'  WHERE " & _
          "ISBT128 = '" & txtUnitNumber & "' AND BarCode = '" & ProductBarCodeFor(lblProduct) & "' "
50  CnxnBB(0).Execute sql


60  cmdLogAsChecked.Visible = False
70  lblChecked = "EI suitable"    '"Group Checked"

80  Exit Sub

cmdLogAsChecked_Click_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmXM", "cmdLogAsChecked_Click", intEL, strES, sql

End Sub

Private Sub cXM_Click(Index As Integer)

    Dim n As Integer
    Dim done As Integer

10  If cXM(Index) = 1 Then
20      xReaction(Index * 2).Enabled = True
30      xReaction(Index * 2 + 1).Enabled = True
40  Else
50      xReaction(Index * 2).Enabled = False
60      xReaction(Index * 2 + 1).Enabled = False
70  End If

80  done = False

90  For n = 0 To 5
100     If xReaction(n) Then done = True
110 Next

120 If done Then
130     cmdXM.Enabled = True
140 Else
150     cmdXM.Enabled = False
160 End If

End Sub

Private Sub FillG()

    Dim tb As Recordset
    Dim sql As String
    Dim s As String

10  On Error GoTo FillG_Error

20  g.Rows = 2
30  g.AddItem ""
40  g.RemoveItem 1

50  sql = "SELECT DISTINCT DateExpiry, BarCode, ISBT128, GroupRh, Screen " & _
          "FROM Latest " & _
          "WHERE LabNumber = '" & frmxmatch.tLabNum & "' " & _
          "AND (Event = 'P')"
60  Set tb = New Recordset
70  RecOpenServerBB 0, tb, sql

80  Do While Not tb.EOF
90      s = tb!ISBT128 & vbTab & _
            Format$(tb!DateExpiry, "dd/MM/yyyy HH:mm") & vbTab & _
            Bar2Group(tb!GroupRh) & vbTab & _
            ProductWordingFor(tb!BarCode) & vbTab & _
            tb!Screen & ""
100     g.AddItem s
110     tb.MoveNext
120 Loop

130 If g.Rows > 2 Then
140     g.RemoveItem 1
150 End If

160 Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmXM", "FillG", intEL, strES, sql

End Sub


Private Sub Form_Load()

    Dim n As Integer

10  FillG

20  For n = 0 To 5
30      xReaction(n).Value = False
40  Next
50  cmdXM.Enabled = False

60  lblProduct = ""

70  cmdRemove.Enabled = False
80  cmdAddPending.Enabled = False
    
    Call FormatGrid
    fmeProducts.Visible = False
    m_Flag = False

End Sub

Private Sub g_Click()

    Dim X As Integer
    Dim Y As Integer
    Dim NewColour As Long
    Dim Counter As Integer

10  If g.MouseRow = 0 Then Exit Sub
20  If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then Exit Sub

30  g.col = 0
40  If g.CellBackColor = vbYellow Then
50      NewColour = &H80000018    'tooltip
60  Else
70      NewColour = vbYellow
80  End If

90  For X = 0 To g.Cols - 1
100     g.col = X
110     g.CellBackColor = NewColour
120 Next

130 RemoveDetails

140 Counter = 0
150 For Y = 1 To g.Rows - 1
160     g.col = 0
170     g.row = Y
180     If g.CellBackColor = vbYellow Then
190         Counter = Counter + 1
200         txtUnitNumber = g.TextMatrix(Y, 0)
210         lblExpiry = g.TextMatrix(Y, 1)
220         lblGroupRh = g.TextMatrix(Y, 2)
230         lblProduct = g.TextMatrix(Y, 3)
    '    txtUnitNumber = g.TextMatrix(Y, 4)
240         lblScreen = g.TextMatrix(Y, 4)
250     End If
260 Next

270 If Counter <> 1 Then
280     cmdRemove.Enabled = False
290     RemoveDetails
300     cmdXM.Enabled = True
310     txtUnitNumber = ""
320     lblProduct = ""
330     Exit Sub
340 End If

350 cmdRemove.Enabled = True
360 lblStatus = ""

End Sub

Private Sub lblTag_Click()
10  frmUnitNotes.txtUnitNumber = txtUnitNumber
20  frmUnitNotes.txtExpiry = lblExpiry
30  frmUnitNotes.Show 1
    'LoadDetails

End Sub

Private Sub RemoveDetails()

    Dim n As Integer

10  For n = 0 To 5
20      xReaction(n) = False
30  Next

40  For n = 0 To 2
50      cXM(n) = False
60  Next

70  lblExpiry = ""
80  lblGroupRh = ""
90  lblScreen = ""
100 lblChecked = ""
110 lblProduct = ""
120 cmdLogAsChecked.Visible = False
130 lblKell = ""

End Sub


Private Sub txtUnitNumber_Change()

10  If Trim$(txtUnitNumber) <> "" Then
20      cmdAddPending.Enabled = True
30      cmdRemove.Enabled = True
40  Else
50      cmdAddPending.Enabled = False
60      cmdRemove.Enabled = False
70  End If

80  lblExpiry = ""
90  lblGroupRh = ""
100 lblScreen = ""
110 lblKell = ""
120 lblProduct = ""
130 lblStatus = ""

End Sub

Private Sub txtUnitNumber_LostFocus()

    Dim Ps As New Products
    Dim f As Form
    Dim p As Product
    Dim s As String

10  If Len(Trim$(txtUnitNumber)) > 0 Then

20      If Left$(txtUnitNumber, 1) = "=" Then    'Barcode scanning entry
30          s = ISOmod37_2(Mid$(txtUnitNumber, 2, 13))
40          txtUnitNumber = Mid$(txtUnitNumber, 2, 13) & " " & s
50      End If

60      Ps.LoadLatestByUnitNumberISBT128 (txtUnitNumber)

70      If Ps.Count = 0 Then
80          iMsg "Unit Number not found."
90          If TimedOut Then Unload Me: Exit Sub
100         txtUnitNumber = ""
110         Exit Sub
120     ElseIf Ps.Count > 1 Then    'multiple products found
130         Set f = New frmSelectFromMultiple
140         f.ProductList = Ps
150         f.Show 1
160         Set p = f.SelectedProduct
170         Unload f
180         Set f = Nothing
190     Else
200         Set p = Ps.Item(1)
210     End If

211     If p.PackEvent = "W" Then
212         iMsg "This unit cannot be crossmatched/issued." & vbTab & vbTab & vbTab & "Unit Group Check failed."
213         If TimedOut Then Unload Me: Exit Sub
214         txtUnitNumber = ""
215         Exit Sub
216     End If

220     If p.PackEvent = "Z" Then
230         iMsg "This unit cannot be crossmatched/issued. It is already transfused as Emergency ONeg."
240         If TimedOut Then Unload Me: Exit Sub
250         txtUnitNumber = ""
260         Exit Sub
270     End If

280     FillDetails p

290     If InStr("XP", p.PackEvent) > 0 Then
300         iMsg "This unit is currently crossmatched " & _
                 "for a different patient. " & vbCrLf & _
                 "Please restock the unit before crossmatching again.", vbCritical
310         If TimedOut Then Unload Me: Exit Sub
320         RemoveDetails
330         lblStatus.Caption = ""
340         txtUnitNumber = ""
350     End If
360 End If

End Sub



Private Sub xReaction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim done As Integer

10  done = True
    'If trim$(tproduct) = "" Then done = False
    'If trim$(lproduct) = "" Then done = False
    'If trim$(tunitnum) = "" Then done = False

20  If done Then
30      cmdXM.Enabled = True
40  Else
50      cmdXM.Enabled = False
60  End If

End Sub
'
'Private Sub GetProducts(p_SampleID As String)
'    On Error GoTo ERROR_Handler
'
'    Dim sql As String
'    Dim tb As ADODB.Recordset
'    Dim Str As String
'
'    sql = "Select * from ocmRequestDetails Where SampleID = '" & p_SampleID & "' And transA = '1' And Status = 'Process'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    If Not tb Is Nothing Then
'        If Not tb.EOF Then
'            While Not tb.EOF
'                Str = "" & vbTab & tb!RequestID & vbTab & tb!SampleID & vbTab & tb!UID & vbTab & tb!TestCode
'                flxProducts.AddItem (Str)
'                tb.MoveNext
'            Wend
'        End If
'    End If
'
'    Exit Sub
'ERROR_Handler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "frmBatchProductIssue", "GetProducts", intEL, strES
'End Sub
'
'Private Sub UpdateIdentifier(Optional p_UID As String, Optional p_indentifier As String)
'    On Error GoTo ERROR_Handler
'
'    Dim sql As String
'    Dim tb As ADODB.Recordset
'    Dim l_ID As Integer
'
'    sql = "Update ocmRequestDetails Set indentifier = '" & p_indentifier & "' Where UID = '" & p_UID & "'"
'    Cnxn(0).Execute sql
'    DoEvents
'    DoEvents
'
'    sql = "Select IsNULL(Max(id),0) ID from ocmbtproductsissued"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    If Not tb Is Nothing Then
'        If Not tb.EOF Then
'            l_ID = tb!ID
'        End If
'    End If
'
'    l_ID = l_ID + 1
'
'    sql = "Insert Into ocmbtproductsissued(id,uid,identifier,units) "
'    sql = sql & "Values(" & l_ID & ",'" & p_UID & "','" & p_indentifier & "','1')"
'    Cnxn(0).Execute sql
'    DoEvents
'    DoEvents
'
'    Exit Sub
'ERROR_Handler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    MsgBox Err.Description & " " & "UpdateIdentifier"
'    LogError "frmBatchProductIssue", "UpdateIdentifier", intEL, strES
'End Sub
'
'Private Sub CountAndUpdate(p_SID As String)
'    On Error GoTo ERROR_Handler
'
'    Dim sql As String
'    Dim tb As ADODB.Recordset
'    Dim tb2 As ADODB.Recordset
'
'    sql = "Select Units, UID from ocmRequestDetails Where SampleID = '" & p_SID & "' And Status In ('Pending','Process')"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    If Not tb Is Nothing Then
'        If Not tb.EOF Then
'            While Not tb.EOF
'                    sql = "Select Count(UID) CUID from ocmbtproductsissued Where UID = '" & tb!UID & "'"
'                    Set tb2 = New Recordset
'                    RecOpenServer 0, tb2, sql
'                    If Not tb2 Is Nothing Then
'                        If Not tb2.EOF Then
'                            If tb2!CUID = tb!Units Then
'                                sql = "Update ocmRequestDetails Set Status = 'Issued' Where UID = '" & tb!UID & "'"
'                                Cnxn(0).Execute sql
'                            End If
'                        End If
'                    End If
'                tb.MoveNext
'            Wend
'        End If
'    End If
'
'    Exit Sub
'ERROR_Handler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    MsgBox Err.Description & " " & "CountAndUpdate"
'    LogError "frmBatchProductIssue", "CountAndUpdate", intEL, strES
'End Sub
