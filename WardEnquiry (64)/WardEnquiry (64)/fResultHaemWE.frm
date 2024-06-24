VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form fResultHaemWE 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Haematology"
   ClientHeight    =   7425
   ClientLeft      =   1170
   ClientTop       =   360
   ClientWidth     =   6210
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
   HelpContextID   =   10025
   Icon            =   "fResultHaemWE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7425
   ScaleWidth      =   6210
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4020
      Top             =   6960
   End
   Begin VB.TextBox tBasA 
      Height          =   285
      Left            =   1080
      TabIndex        =   46
      Top             =   4920
      Width           =   825
   End
   Begin VB.TextBox tEosP 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   45
      Top             =   4635
      Width           =   825
   End
   Begin VB.TextBox tNeutA 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   44
      Top             =   4350
      Width           =   825
   End
   Begin VB.TextBox tWBC 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   43
      Top             =   3450
      Width           =   825
   End
   Begin VB.TextBox tEosA 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   42
      Top             =   4620
      Width           =   825
   End
   Begin VB.TextBox tNeutP 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   41
      Top             =   4350
      Width           =   825
   End
   Begin VB.TextBox tMonoP 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   40
      Top             =   4065
      Width           =   825
   End
   Begin VB.TextBox tBasP 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   39
      Top             =   4920
      Width           =   825
   End
   Begin VB.TextBox tLymA 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   38
      Top             =   3780
      Width           =   825
   End
   Begin VB.TextBox tMonoA 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   37
      Top             =   4050
      Width           =   825
   End
   Begin VB.TextBox tLymP 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   36
      Top             =   3780
      Width           =   825
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
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
      Height          =   795
      Left            =   4500
      Picture         =   "fResultHaemWE.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
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
      Left            =   4530
      Picture         =   "fResultHaemWE.frx":1534
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2070
      Width           =   1245
   End
   Begin VB.TextBox tMCHC 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   19
      Top             =   2790
      Width           =   825
   End
   Begin VB.TextBox tRBC 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   18
      Top             =   1890
      Width           =   825
   End
   Begin VB.TextBox tHgb 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   17
      Top             =   2190
      Width           =   825
   End
   Begin VB.TextBox tHct 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   16
      Top             =   2790
      Width           =   825
   End
   Begin VB.TextBox tMCH 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   15
      Top             =   2490
      Width           =   825
   End
   Begin VB.TextBox tRDWCV 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   14
      Top             =   1860
      Width           =   825
   End
   Begin VB.TextBox tRDWSD 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   13
      Top             =   2190
      Width           =   825
   End
   Begin VB.TextBox tMCV 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2490
      Width           =   825
   End
   Begin VB.TextBox tPLCR 
      Height          =   285
      Left            =   2760
      TabIndex        =   31
      Top             =   5460
      Width           =   825
   End
   Begin VB.TextBox tMPV 
      Height          =   285
      Left            =   1080
      TabIndex        =   30
      Top             =   5760
      Width           =   825
   End
   Begin VB.TextBox tPdw 
      Height          =   285
      Left            =   2760
      TabIndex        =   29
      Top             =   5760
      Width           =   825
   End
   Begin VB.TextBox tPlt 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   28
      Top             =   5460
      Width           =   825
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   240
      TabIndex        =   54
      Top             =   7110
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblCnxn 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCnxn"
      Height          =   255
      Left            =   5310
      TabIndex        =   62
      Top             =   180
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   6030
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Run Date"
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
      TabIndex        =   61
      Top             =   210
      Width           =   690
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   540
      TabIndex        =   60
      Top             =   1110
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
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
      Left            =   2700
      TabIndex        =   59
      Top             =   780
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1020
      TabIndex        =   58
      Top             =   1080
      Width           =   3225
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3090
      TabIndex        =   57
      Top             =   750
      Width           =   1155
   End
   Begin VB.Label lblChartTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "NOPAS"
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
      TabIndex        =   56
      Top             =   780
      Width           =   555
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1020
      TabIndex        =   55
      Top             =   750
      Width           =   1185
   End
   Begin VB.Label lblNotValid 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Results not yet available."
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   1320
      TabIndex        =   53
      Top             =   1500
      Width           =   2265
   End
   Begin VB.Label lretics 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5100
      TabIndex        =   7
      Top             =   4140
      Width           =   720
   End
   Begin VB.Label lmonospot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5100
      TabIndex        =   6
      Top             =   4545
      Width           =   720
   End
   Begin VB.Label lesr 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5100
      TabIndex        =   5
      Top             =   3750
      Width           =   720
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Infectious Mono Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4005
      TabIndex        =   4
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Retics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4530
      TabIndex        =   3
      Top             =   4170
      Width           =   450
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "ESR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4650
      TabIndex        =   2
      Top             =   3810
      Width           =   330
   End
   Begin VB.Label Label34 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Bas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2190
      TabIndex        =   52
      Top             =   4920
      Width           =   270
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Eos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2220
      TabIndex        =   51
      Top             =   4650
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "WBC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   630
      TabIndex        =   50
      Top             =   3510
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Lymph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2070
      TabIndex        =   49
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Mono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2130
      TabIndex        =   48
      Top             =   4110
      Width           =   420
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Neut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2160
      TabIndex        =   47
      Top             =   4380
      Width           =   360
   End
   Begin VB.Label lblRunDate 
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
      Height          =   255
      Left            =   3090
      TabIndex        =   8
      Top             =   180
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
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
      Left            =   150
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblSampleID 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1020
      TabIndex        =   9
      Top             =   180
      Width           =   1200
   End
   Begin VB.Label Label10 
      Caption         =   "Hgb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   660
      TabIndex        =   27
      Top             =   2220
      Width           =   345
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "RDW SD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   2025
      TabIndex        =   26
      Top             =   2220
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "RBC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   675
      TabIndex        =   25
      Top             =   1920
      Width           =   330
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "MCV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   24
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Hct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   23
      Top             =   2820
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "MCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   2340
      TabIndex        =   22
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "MCHC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   2235
      TabIndex        =   21
      Top             =   2820
      Width           =   465
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "RDW CV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2040
      TabIndex        =   20
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "PLCR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2250
      TabIndex        =   35
      Top             =   5520
      Width           =   420
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Plt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   795
      TabIndex        =   34
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "MPV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   690
      TabIndex        =   33
      Top             =   5820
      Width           =   345
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Pdw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2340
      TabIndex        =   32
      Top             =   5820
      Width           =   315
   End
   Begin VB.Label lblComment 
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
      Height          =   945
      Left            =   240
      TabIndex        =   11
      Top             =   6150
      Width           =   5565
   End
End
Attribute VB_Name = "fResultHaemWE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnActivated As Boolean

Private Sub bcancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo cmdPrint_Click_Error

20    If iMsg("Report will be printed on" & vbCrLf & _
        WardEnqForcedPrinter & "." & vbCrLf & _
        "OK?", vbQuestion + vbYesNo) = vbYes Then
    
30      sql = "Select * from PrintPending where " & _
              "Department = 'H' " & _
              "and SampleID = '" & lblSampleID & "'"
40      Set tb = New Recordset
50      RecOpenClient 0, tb, sql
60      If tb.EOF Then tb.AddNew
70      tb!SampleID = lblSampleID
80      tb!Department = "H"
90      tb!Initiator = "Ward"
100     tb!UsePrinter = WardEnqForcedPrinter
110     tb!ThisIsCopy = 1
120     tb.Update
    
130     LogAsViewed "J", lblSampleID, lblChart
140   End If

150   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "fResultHaemWE", "cmdPrint_Click", intEL, strES, sql
  
End Sub

Private Sub ClearHaem()

10    tWBC = ""
20    tWBC.BackColor = &HFFFFFF
30    tWBC.ForeColor = &H0&

40    tRBC = ""
50    tRBC.BackColor = &HFFFFFF
60    tRBC.ForeColor = &H0&

70    tMCV = ""
80    tMCV.BackColor = &HFFFFFF
90    tMCV.ForeColor = &H0&

100   tHct = ""
110   tHct.BackColor = &HFFFFFF
120   tHct.ForeColor = &H0&

130   tRDWCV = ""
140   tRDWCV.BackColor = &HFFFFFF
150   tRDWCV.ForeColor = &H0&

160   tRDWSD = ""
170   tRDWSD.BackColor = &HFFFFFF
180   tRDWSD.ForeColor = &H0&

190   tMCH = ""
200   tMCH.BackColor = &HFFFFFF
210   tMCH.ForeColor = &H0&

220   tMCHC = ""
230   tMCHC.BackColor = &HFFFFFF
240   tMCHC.ForeColor = &H0&

250   tPlt = ""
260   tPlt.BackColor = &HFFFFFF
270   tPlt.ForeColor = &H0&

280   tMPV = ""
290   tMPV.BackColor = &HFFFFFF
300   tMPV.ForeColor = &H0&

310   tPLCR = ""
320   tPLCR.BackColor = &HFFFFFF
330   tPLCR.ForeColor = &H0&

340   tPdw = ""
350   tPdw.BackColor = &HFFFFFF
360   tPdw.ForeColor = &H0&

370   tLymA = ""
380   tLymA.BackColor = &HFFFFFF
390   tLymA.ForeColor = &H0&

400   tLymP = ""
410   tLymP.BackColor = &HFFFFFF
420   tLymP.ForeColor = &H0&

430   tMonoA = ""
440   tMonoA.BackColor = &HFFFFFF
450   tMonoA.ForeColor = &H0&

460   tMonoP = ""
470   tMonoP.BackColor = &HFFFFFF
480   tMonoP.ForeColor = &H0&

490   tNeutA = ""
500   tNeutA.BackColor = &HFFFFFF
510   tNeutA.ForeColor = &H0&

520   tNeutP = ""
530   tNeutP.BackColor = &HFFFFFF
540   tNeutP.ForeColor = &H0&

550   tEosA = ""
560   tEosA.BackColor = &HFFFFFF
570   tEosA.ForeColor = &H0&

580   tEosP = ""
590   tEosP.BackColor = &HFFFFFF
600   tEosP.ForeColor = &H0&

610   tBasA = ""
620   tBasA.BackColor = &HFFFFFF
630   tBasA.ForeColor = &H0&

640   tBasP = ""
650   tBasP.BackColor = &HFFFFFF
660   tBasP.ForeColor = &H0&

670   lesr = ""
680   lretics = ""
690   lmonospot = ""

End Sub

Private Sub Form_Activate()

10    If LogOffNow Then
20      Unload Me
30    End If

40    If blnActivated Then Exit Sub
50    blnActivated = True

60    lblChartTitle = "Chart"
  
70    LoadHaem

80    LogAsViewed "R", lblSampleID, frmMain.txtChart

90    PBar.Max = LogOffDelaySecs
100   PBar = 0
110   SingleUserUpdateLoggedOn UserName

120   Timer1.Enabled = True

End Sub

Private Sub LoadHaem()

      Dim tb As Recordset
      Dim sql As String
      Dim Sex As String
      Dim OBS As Observations
      Dim sampleDate As String

10    On Error GoTo LoadHaem_Error

20    If Trim$(lblSampleID) = "" Then Exit Sub

30    sql = "Select Sex, sampleDate from Demographics where " & _
            "SampleID = '" & lblSampleID & "'"
40    Set tb = New Recordset
50    RecOpenServer Val(lblCnxn), tb, sql
60    If Not tb.EOF Then
70      Sex = tb!Sex & ""
        sampleDate = tb!sampleDate
80    Else
90      Sex = ""
100   End If

110   sql = "Select * from HaemResults where " & _
            "SampleID = '" & lblSampleID & "'"
120   Set tb = New Recordset
130   RecOpenServer Val(lblCnxn), tb, sql

140   If tb.EOF Then
150     ClearHaem
160   ElseIf tb!Valid = 0 Or IsNull(tb!Valid) Then
170     lblNotValid.Visible = True
180     lblRunDate = tb!Rundate
190   Else
200     lblRunDate = tb!Rundate
210     lblNotValid.Visible = False
220     If Not IsNull(tb!rbc) Then
230       Colourise "RBC", tRBC, tb!rbc, Sex, lblDoB, sampleDate
240     End If
  
250     If Not IsNull(tb!Hgb) Then
260       Colourise "Hgb", tHgb, tb!Hgb, Sex, lblDoB, sampleDate
270     End If
  
280     If Not IsNull(tb!MCV) Then
290       Colourise "MCV", tMCV, tb!MCV, Sex, lblDoB, sampleDate
300     End If
  
310     If Not IsNull(tb!Hct) Then
320       Colourise "Hct", tHct, tb!Hct, Sex, lblDoB, sampleDate
330     End If
  
340     If Not IsNull(tb!RDWCV) Then
350       Colourise "RDWCV", tRDWCV, tb!RDWCV, Sex, lblDoB, sampleDate
360     End If
  
370     If Not IsNull(tb!rdwsd) Then
380       Colourise "RDWSD", tRDWSD, tb!rdwsd, Sex, lblDoB, sampleDate
390     End If
  
400     If Not IsNull(tb!mch) Then
410       Colourise "MCH", tMCH, tb!mch, Sex, lblDoB, sampleDate
420     End If
  
430     If Not IsNull(tb!mchc) Then
440       Colourise "MCHC", tMCHC, tb!mchc, Sex, lblDoB, sampleDate
450     End If
  
460     If Not IsNull(tb!plt) Then
470       Colourise "plt", tPlt, tb!plt, Sex, lblDoB, sampleDate
480     End If
  
490     If Not IsNull(tb!mpv) Then
500       Colourise "MPV", tMPV, tb!mpv, Sex, lblDoB, sampleDate
510     End If
  
520     If Not IsNull(tb!plcr) Then
530       Colourise "PLCR", tPLCR, tb!plcr, Sex, lblDoB, sampleDate
540     End If
  
550     If Not IsNull(tb!pdw) Then
560       Colourise "Pdw", tPdw, tb!pdw, Sex, lblDoB, sampleDate
570     End If
  
580     If Not IsNull(tb!WBC) Then
590       Colourise "WBC", tWBC, tb!WBC, Sex, lblDoB, sampleDate
600     End If
  
610     If Not IsNull(tb!LymA) Then
620       Colourise "LymA", tLymA, tb!LymA, Sex, lblDoB, sampleDate
630     End If
  
640     If Not IsNull(tb!LymP) Then
650       Colourise "LymP", tLymP, tb!LymP, Sex, lblDoB, sampleDate
660     End If
  
670     If Not IsNull(tb!MonoA) Then
680       Colourise "MonoA", tMonoA, tb!MonoA, Sex, lblDoB, sampleDate
690     End If
  
700     If Not IsNull(tb!MonoP) Then
710       Colourise "MonoP", tMonoP, tb!MonoP, Sex, lblDoB, sampleDate
720     End If
  
730     If Not IsNull(tb!NeutA) Then
740       Colourise "NeutA", tNeutA, tb!NeutA, Sex, lblDoB, sampleDate
750     End If
  
760     If Not IsNull(tb!NeutP) Then
770       Colourise "NeutP", tNeutP, tb!NeutP, Sex, lblDoB, sampleDate
780     End If
  
790     If Not IsNull(tb!EosA) Then
800       Colourise "EosA", tEosA, tb!EosA, Sex, lblDoB, sampleDate
810     End If
  
820     If Not IsNull(tb!EosP) Then
830       Colourise "EosP", tEosP, tb!EosP, Sex, lblDoB, sampleDate
840     End If
  
850     If Not IsNull(tb!BasA) Then
860       Colourise "BasA", tBasA, tb!BasA, Sex, lblDoB, sampleDate
870     End If
  
880     If Not IsNull(tb!BasP) Then
890       Colourise "BasP", tBasP, tb!BasP, Sex, lblDoB, sampleDate
900     End If
910     lesr = tb!ESR & ""
920     lretics = tb!RetP & ""  'tb!retics & ""
930     lmonospot = tb!monospot & ""
  
940     Set OBS = New Observations
950     Set OBS = OBS.Load(lblSampleID, "Haematology")
960     If Not OBS Is Nothing Then
970       lblComment = "Specimen Comment: " & OBS.Item(1).Comment & vbCrLf
980     End If

990     Set OBS = New Observations
1000    Set OBS = OBS.Load(lblSampleID, "Film")
1010    If Not OBS Is Nothing Then
1020      lblComment = lblComment & "Film Comment: " & OBS.Item(1).Comment
1030    End If

1040    cmdPrint.Enabled = False
1050    If UserCanPrint Then
1060      cmdPrint.Enabled = Not lblNotValid.Visible
1070    Else
1080      cmdPrint.Visible = False
1090    End If

1100  End If

1110  Screen.MousePointer = 0

1120  Exit Sub

LoadHaem_Error:

      Dim strES As String
      Dim intEL As Integer

1130  intEL = Erl
1140  strES = Err.Description
1150  LogError "fResultHaemWE", "LoadHaem", intEL, strES, sql

End Sub
Private Sub Colourise(ByVal Analyte As String, _
                      ByVal Destination As TextBox, _
                      ByVal strValue As String, _
                      ByVal Sex As String, _
                      ByVal DoB As String, ByVal sampleDate As String)
    
      Dim Value As Single

10    Value = Val(strValue)

20    Destination.Text = strValue
30    If Trim$(strValue) = "" Then
40      Destination.BackColor = &HFFFFFF
50      Destination.ForeColor = &H0&
60      Exit Sub
70    End If

80    Select Case InterpH(Value, Analyte, Sex, DoB, sampleDate)
        Case "X":
90        Destination.BackColor = vbBlack
100       Destination.ForeColor = vbWhite
110     Case "H":
120       Destination.BackColor = &HFFFF&
130       Destination.ForeColor = &HFF&
140     Case "L"
150       Destination.BackColor = &HFFFF00
160       Destination.ForeColor = &HC00000
170     Case Else
180       Destination.BackColor = &HFFFFFF
190       Destination.ForeColor = &H0&
200   End Select
  
End Sub

Private Sub Form_Deactivate()

10    Timer1.Enabled = False

End Sub

Private Sub Form_Load()

10    blnActivated = False

20    PBar.Max = LogOffDelaySecs
30    PBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    blnActivated = False

End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10    PBar = PBar + 1
  
20    If PBar = PBar.Max Then
30      LogOffNow = True
40      Unload Me
50    End If

End Sub


