VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmTransfusion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Details"
   ClientHeight    =   7515
   ClientLeft      =   450
   ClientTop       =   435
   ClientWidth     =   10110
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
   Icon            =   "frmTransfusion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7515
   ScaleWidth      =   10110
   Begin VB.CommandButton bPrintCord 
      Caption         =   "Print Co&rd Report"
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
      Left            =   8250
      Picture         =   "frmTransfusion.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   4860
      Width           =   1425
   End
   Begin VB.CommandButton cmdPrintAN 
      Caption         =   "Print &A/N Report"
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
      Left            =   8250
      Picture         =   "frmTransfusion.frx":1534
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4110
      Width           =   1425
   End
   Begin VB.CommandButton bNext 
      Caption         =   ">Next"
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
      Height          =   735
      Left            =   6480
      Picture         =   "frmTransfusion.frx":1B9E
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1140
      Width           =   885
   End
   Begin VB.CommandButton bPrevious 
      Caption         =   "<Previous"
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
      Height          =   735
      Left            =   4380
      Picture         =   "frmTransfusion.frx":1FE0
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1140
      Width           =   855
   End
   Begin VB.CommandButton btncancel 
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
      Height          =   705
      Left            =   8250
      Picture         =   "frmTransfusion.frx":2422
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6210
      Width           =   1425
   End
   Begin VB.CommandButton btninitiate 
      Appearance      =   0  'Flat
      Caption         =   "Initiate &Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5250
      Picture         =   "frmTransfusion.frx":2A8C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1140
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   1230
      TabIndex        =   50
      Top             =   7290
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Antibody screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   285
      TabIndex        =   57
      Top             =   3555
      Width           =   765
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSearch 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4380
      TabIndex        =   54
      Top             =   750
      Width           =   2985
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search Criteria (Name, Chart or DoB)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4380
      TabIndex        =   53
      Top             =   450
      Width           =   2955
   End
   Begin VB.Label lblKell 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   52
      Top             =   3990
      Width           =   1215
   End
   Begin VB.Label Label22 
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
      Left            =   2430
      TabIndex        =   51
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Sample Date"
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
      Left            =   4410
      TabIndex        =   49
      Top             =   2700
      Width           =   915
   End
   Begin VB.Label lSampleDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88/88/8888 88:88 "
      Height          =   285
      Left            =   5400
      TabIndex        =   48
      Top             =   2670
      Width           =   2190
   End
   Begin VB.Label lblIdent 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1110
      TabIndex        =   47
      Top             =   3600
      Visible         =   0   'False
      Width           =   6465
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Maiden"
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
      TabIndex        =   44
      Top             =   960
      Width           =   525
   End
   Begin VB.Label lMaiden 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1110
      TabIndex        =   43
      Top             =   960
      Width           =   2835
   End
   Begin VB.Label lAddr 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "addr3"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   2
      Left            =   1110
      TabIndex        =   42
      Top             =   1755
      Width           =   2835
   End
   Begin VB.Label lAddr 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "addr4"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   3
      Left            =   1110
      TabIndex        =   41
      Top             =   2010
      Width           =   2835
   End
   Begin VB.Label Label19 
      Caption         =   "Lab #"
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
      Left            =   4650
      TabIndex        =   40
      Top             =   2070
      Width           =   510
   End
   Begin VB.Label lLabNo 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "labnumber"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5190
      TabIndex        =   39
      Top             =   2010
      Width           =   2175
   End
   Begin VB.Label lEDD 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "edd"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   2790
      TabIndex        =   38
      Top             =   2970
      Width           =   1155
   End
   Begin VB.Label lReaction 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "prevreact"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   2790
      TabIndex        =   37
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label lPPreg 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "prevpreg"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   36
      Top             =   2910
      Width           =   735
   End
   Begin VB.Label lPTrans 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "prevtrans"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   35
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Prev.Preg"
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
      TabIndex        =   34
      Top             =   2910
      Width           =   705
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "EDD"
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
      Left            =   2370
      TabIndex        =   33
      Top             =   2970
      Width           =   345
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Reaction"
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
      Left            =   2070
      TabIndex        =   32
      Top             =   2730
      Width           =   645
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Prev.Trans"
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
      Left            =   285
      TabIndex        =   31
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
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
      Left            =   405
      TabIndex        =   27
      Top             =   4410
      Width           =   660
   End
   Begin VB.Label lTime 
      DataField       =   "DateTime"
      DataSource      =   "Data1"
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
      Left            =   4830
      TabIndex        =   30
      Top             =   6900
      Width           =   1695
   End
   Begin VB.Label lDoB 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "DoB"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   2790
      TabIndex        =   29
      Top             =   2430
      Width           =   1155
   End
   Begin VB.Label lComment 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "comment"
      DataSource      =   "Data1"
      Height          =   1305
      Left            =   1110
      TabIndex        =   28
      Top             =   4410
      Width           =   6465
      WordWrap        =   -1  'True
   End
   Begin VB.Label lauto 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "For Autologous Transfusion"
      Height          =   705
      Left            =   3690
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Group"
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
      Left            =   630
      TabIndex        =   25
      Top             =   4080
      Width           =   435
   End
   Begin VB.Label lGroupRh 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1110
      TabIndex        =   24
      Top             =   3990
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
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
      Left            =   795
      TabIndex        =   23
      Top             =   2400
      Width           =   270
   End
   Begin VB.Label lSex 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "sex"
      DataSource      =   "Data1"
      Height          =   225
      Left            =   1110
      TabIndex        =   22
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label10 
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
      Left            =   2400
      TabIndex        =   21
      Top             =   2490
      Width           =   315
   End
   Begin VB.Label lOperator 
      DataField       =   "xop"
      DataSource      =   "Data1"
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
      Left            =   6570
      TabIndex        =   20
      Top             =   6900
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Entered "
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
      Left            =   4200
      TabIndex        =   19
      Top             =   6900
      Width           =   600
   End
   Begin VB.Label lNumber 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "patnum"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   17
      Top             =   450
      Width           =   2835
   End
   Begin VB.Label lName 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "name"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   16
      Top             =   705
      Width           =   2835
   End
   Begin VB.Label lSpecial 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "specialprod"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1950
      TabIndex        =   15
      Top             =   6870
      Width           =   1695
   End
   Begin VB.Label lProcedure 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "procedure"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1950
      TabIndex        =   14
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lCondx 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "conditions"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1950
      TabIndex        =   13
      Top             =   6330
      Width           =   1695
   End
   Begin VB.Label lClinician 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "clinician"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   12
      Top             =   6030
      Width           =   2535
   End
   Begin VB.Label lWard 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ward"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   11
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label lAddr 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "addr2"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   1
      Left            =   1110
      TabIndex        =   10
      Top             =   1500
      Width           =   2835
   End
   Begin VB.Label lAddr 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "addr1"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   0
      Left            =   1110
      TabIndex        =   9
      Top             =   1245
      Width           =   2835
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Special Products"
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
      Left            =   675
      TabIndex        =   8
      Top             =   6900
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Surgical Procedure"
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
      Left            =   525
      TabIndex        =   7
      Top             =   6630
      Width           =   1350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Clinical Conditions"
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
      TabIndex        =   6
      Top             =   6360
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Clinician"
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
      Left            =   480
      TabIndex        =   5
      Top             =   6060
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Left            =   675
      TabIndex        =   4
      Top             =   5820
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   495
      TabIndex        =   3
      Top             =   1290
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   630
      TabIndex        =   2
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Number"
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
      Left            =   510
      TabIndex        =   1
      Top             =   480
      Width           =   555
   End
End
Attribute VB_Name = "frmTransfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Current As Integer
Dim Max As Integer

Dim sn As Recordset

Private pExternalCriteria As String


Private Sub FillDetails(ByVal sn As Recordset)
  
10    On Error GoTo FillDetails_Error

20    lblKell = sn!Kell & ""
30    lSampleDate = sn!SampleDate & ""
40    lNumber = sn!Patnum & ""
50    lName = sn!Name & ""
60    lMaiden = sn!maiden & ""
70    lAddr(0) = sn!Addr1 & ""
80    lAddr(1) = sn!Addr2 & ""
90    lAddr(2) = sn!Addr3 & ""
100   lAddr(3) = sn!addr4 & ""
110   lSex = sn!Sex & ""
120   If Not IsNull(sn!DoB) Then
130     lDoB = Format(sn!DoB, "dd/mm/yyyy")
140   Else
150     lDoB = ""
160   End If
170   lPTrans = sn!prevtrans & ""
180   lReaction = sn!prevreact & ""
190   lPPreg = sn!prevpreg & ""
200   If Not IsNull(sn!edd) Then
210     lEDD = Format(sn!edd, "dd/mm/yyyy")
220   Else
230     lEDD = ""
240   End If
250   lComment = sn!Comment & ""
260   lWard = sn!Ward & ""
270   lClinician = sn!Clinician & ""
280   lCondx = sn!Conditions & ""
290   lProcedure = sn!Procedure & ""
300   lSpecial = sn!specialprod & ""
310   lGroupRh = sn!fGroup & ""
320   lauto.Visible = False
330   If Not IsNull(sn!Autolog) Then
340     lauto.Visible = IIf(sn!Autolog, True, False)
350   End If

360   lblIdent.Visible = False
370   If Trim$(sn!AIDR & "") <> "" Then
380     lblIdent = sn!AIDR
390     If InStr(UCase$(sn!AIDR), "POS") <> 0 Then
400       lblIdent.BackColor = vbRed
410     Else
420       lblIdent.BackColor = vbYellow
430     End If
440     lblIdent.Visible = True
450   End If

460   lTime = Format(sn!DateTime, "dd/MM/yyyy HH:mm")
470   lOperator = sn!Operator & ""
480   lLabNo = sn!LabNumber & ""

490   Exit Sub

FillDetails_Error:

      Dim strES As String
      Dim intEL As Integer

500   intEL = Erl
510   strES = Err.Description
520   LogError "frmTransfusion", "FillDetails", intEL, strES

End Sub



Private Sub bNext_Click()

10    On Error GoTo bNext_Click_Error

20    If Current = Max Then Exit Sub

30    Current = Current + 1
  
40    lblstatus = "Record " & Format(Current) & " of " & Format(Max)

50    sn.MoveNext
60    FillDetails sn

70    bPrevious.Enabled = True
80    bNext.Enabled = Current <> Max

90    Exit Sub

bNext_Click_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmTransfusion", "bNext_Click", intEL, strES


End Sub

Private Sub bPrevious_Click()

10    On Error GoTo bPrevious_Click_Error

20    If Max = 0 Then Exit Sub
30    If Current = 1 Then Exit Sub

40    Current = Current - 1
  
50    lblstatus = "Record " & Format(Current) & " of " & Format(Max)

60    sn.MovePrevious
70    FillDetails sn

80    bPrevious.Enabled = Current <> 1
90    bNext.Enabled = True

100   Exit Sub

bPrevious_Click_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmTransfusion", "bPrevious_Click", intEL, strES


End Sub

Private Sub bPrintCord_Click()

10    PrintCordForm lLabNo, lSampleDate

End Sub

Public Sub PrintCordForm(ByVal SampleID As String, _
                         ByVal SampleDate As String)

10    PrintCordFormCavan SampleID, SampleDate

End Sub





Public Sub PrintCordFormCavan(ByVal SampleID As String, _
                              ByVal SampleDate As String)

      Dim DAT As String
      Dim sql As String

10    On Error GoTo PrintCordFormCavan_Error

20    PrintHeadingCavan SampleID

30    Printer.Font.Name = "Courier New"
40    Printer.Font.Size = 12

50    Printer.Print
60    Printer.Print
70    Printer.Print
80    Printer.Print
90    Printer.Print
100   Printer.Print
110   Printer.Print "  Spec Grp: ";
120   Printer.Font.Bold = True
130   Printer.Font.Size = 20
140   Printer.Print Trim$(Left$(lGroupRh, 2)); " Rh";
150   Printer.Font.Size = 10
160   Printer.CurrentY = Printer.CurrentY + 150
170   Printer.Print "(D) ";
180   Printer.Font.Size = 20
190   Printer.CurrentY = Printer.CurrentY - 150
200   If InStr(lGroupRh, "P") Then
210     Printer.Print "Positive  ";
220   ElseIf InStr(lGroupRh, "N") Then
230     Printer.Print "Negative  ";
240   End If
250   Printer.Font.Size = 14
260   Printer.CurrentY = Printer.CurrentY + 100
270   Printer.Print "Direct Coombs ";
280   DAT = "Not Tested"
290   sql = "select DAT0, DAT1 from patientdetails where " & _
            "labnumber = '" & lLabNo & "'"
300   Set sn = New Recordset
310   RecOpenServerBB 0, sn, sql
320   If Not sn.EOF Then
330     If Not IsNull(sn!DAT0) Then
340       If sn!DAT0 Then
350         DAT = "Positive"
360       End If
370     End If
380     If Not IsNull(sn!DAT1) Then
390       If sn!DAT1 Then
400         DAT = "Negative"
410       End If
420     End If
430   End If
440   Printer.Print DAT

450   Printer.Print
460   Printer.Print
470   If Trim$(lComment) <> "" Then
480     Printer.Print "Comment: "; lComment
490   End If

500   PrintFooterGHCavan SampleDate, SampleID

510   Printer.EndDoc

520   Exit Sub

PrintCordFormCavan_Error:

      Dim strES As String
      Dim intEL As Integer

530   intEL = Erl
540   strES = Err.Description
550   LogError "frmTransfusion", "PrintCordFormCavan", intEL, strES, sql

End Sub


Private Sub btnCancel_Click()

10    Unload Me

End Sub

Private Sub btninitiate_Click()

      Dim sql As String

10    On Error GoTo btninitiate_Click_Error

20    Screen.MousePointer = 11

30    Max = 0

40    sql = "SELECT * FROM PatientDetails WHERE " & _
            "PatNum = '" & AddTicks(lblSearch) & "' " & _
            "ORDER BY SampleDate ASC"
50    Set sn = New Recordset
60    RecOpenClientBB 0, sn, sql
70    If Not sn.EOF Then
80      sn.MoveLast
90      Max = sn.RecordCount
100   End If

110   If Max = 0 Then
120     lblstatus = "Not found."
130   Else
140     lblstatus = "Record " & Format(Max) & " of " & Format(Max)
150     FillDetails sn
160   End If

170   Current = Max

180   bPrevious.Enabled = Max > 1
190   bNext.Enabled = False

200   Screen.MousePointer = 0

210   Exit Sub

btninitiate_Click_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmTransfusion", "btninitiate_Click", intEL, strES, sql

End Sub


Private Sub cmdPrintAN_Click()

10    PrintANForm lLabNo, lSampleDate

End Sub

Public Sub PrintANForm(ByVal SampleID As String, _
                       ByVal SampleDate As String)

10    If HospName(0) = "Cavan" Then
20      PrintANFormCavan SampleID, SampleDate
30    End If

End Sub

Public Sub PrintANFormCavan(ByVal SampleID As String, _
                            ByVal SampleDate As String)


10    On Error GoTo PrintANFormCavan_Error

20    PrintHeadingCavan SampleID

30    Printer.Font.Name = "Courier New"
40    Printer.Font.Size = 12

50    Printer.Print "  Spec Grp: ";
60    Printer.Font.Bold = True
70    Printer.Font.Size = 20
80    If InStr(lGroupRh, "P") Or InStr(lGroupRh, "N") Then
90      Printer.Print Trim$(Left$(lGroupRh, 2)); " Rh";
100     Printer.Font.Size = 10
110     Printer.CurrentY = Printer.CurrentY + 150
120     Printer.Print "(D) ";
130     Printer.Font.Size = 20
140     Printer.CurrentY = Printer.CurrentY - 150
150     Printer.Print IIf(InStr(lGroupRh, "P"), "Positive", "Negative"); "  ";
160   Else
170     Printer.Print
180   End If
190   Printer.Font.Size = 14
200   Printer.CurrentY = Printer.CurrentY + 100
210   Printer.Print "Antibody Screen: "; lblIdent
220   If lComment <> "" Then
230     Printer.Print "Comment:"; lComment
240   End If

250   Printer.Print
260   Printer.Print Tab(25); "Ante Natal Report"
270   Printer.Print
280   Printer.Print
290   Printer.Print
300   Printer.Print
310   Printer.Print
320   Printer.Print
330   Printer.Print
340   Printer.Print

350   PrintFooterGHCavan SampleDate, SampleID

360   Printer.EndDoc

370   Exit Sub

PrintANFormCavan_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "frmTransfusion", "PrintANFormCavan", intEL, strES

End Sub

Public Sub PrintFooterGHCavan(ByVal SampleDate As String, ByVal SampleID As String)

      Dim sn As Recordset
      Dim sql As String

10    On Error GoTo PrintFooterGHCavan_Error

20    Printer.Font.Name = "Courier New"

30    Do While Printer.CurrentY < 7400 '7000
40      Printer.Print
50    Loop

60    Printer.ForeColor = vbRed
70    Printer.Font.Size = 4
80    Printer.Print String$(250, "-")

90    Printer.Font.Size = 10
100   Printer.Font.Bold = False

110   Printer.Print "Sample Date:"; Format(SampleDate, "dd/mm/yyyy");
  
120   Printer.Print Tab(38); "Report Date:"; Format(Now, "dd/mm/yyyy");

130   sql = "select * from patientdetails where " & _
            "labnumber = '" & SampleID & "'"
140   Set sn = New Recordset
150   RecOpenServerBB 0, sn, sql
  
160   Printer.Print "    Issued By "; TechnicianNameForCode(sn!Operator & "")

170   Exit Sub

PrintFooterGHCavan_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmTransfusion", "PrintFooterGHCavan", intEL, strES, sql

End Sub

Public Function TechnicianNameForCode(ByVal Code As String) As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo TechnicianNameForCode_Error

20    sql = "Select * from Users where " & _
            "Code = '" & AddTicks(Code) & "'"
30    RecOpenServer 0, tb, sql
40    If Not tb.EOF Then
50      TechnicianNameForCode = tb!Name & ""
60    Else
70      TechnicianNameForCode = "???"
80    End If

90    Exit Function

TechnicianNameForCode_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmTransfusion", "TechnicianNameForCode", intEL, strES, sql

End Function

Public Sub PrintHeadingCavan(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintHeadingCavan_Error

20    sql = "Select * from PatientDetails where " & _
            "LabNumber = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then Exit Sub

60    Printer.Font.Name = "Courier New"
70    Printer.Font.Size = 14
80    Printer.Font.Bold = True

90    Printer.ForeColor = vbRed
100   Printer.Print "CAVAN GENERAL HOSPITAL : Blood Transfusion Laboratory";
110   Printer.Font.Size = 10
120   Printer.CurrentY = 100
130   Printer.Print '" Phone 38833"

140   Printer.CurrentY = 320

150   Printer.Font.Size = 4
160   Printer.Print String$(250, "-")

170   Printer.ForeColor = vbBlack

180   Printer.Font.Name = "Courier New"
190   Printer.Font.Size = 12
200   Printer.Font.Bold = False

210   Printer.Print " Sample ID:";
220   Printer.Print SampleID;
  
230   Printer.Print Tab(35); "Name:";
240   Printer.Font.Bold = True
250   Printer.Font.Size = 14
260   Printer.Print tb!Name
270   Printer.Font.Size = 12
280   Printer.Font.Bold = False
  
290   Printer.Print "      Ward:";
300   Printer.Print tb!Ward & "";
  
310   Printer.Print Tab(35); " DOB:";
320   Printer.Print Format(tb!DoB, "dd/mm/yyyy");
330   Printer.Print Tab(60); "Chart #:";
340   Printer.Print tb!Patnum
 
350   If Trim$(tb!Clinician & "") <> "" Then
360     Printer.Print "Consultant:";
370     Printer.Print tb!Clinician & "";
380   Else
390     Printer.Print "        GP:";
400     Printer.Print tb!GP & "";
410   End If

420   Printer.Print Tab(35); "Addr:";
430   Printer.Print tb!Addr1 & "";
440   Printer.Print Tab(60); "    Sex:";
450   Select Case Left$(UCase$(Trim$(tb!Sex & "")), 1)
        Case "M": Printer.Print "Male"
460     Case "F": Printer.Print "Female"
470     Case Else: Printer.Print
480   End Select
  
490   Printer.Font.Bold = False
500   Printer.Print Tab(35); "     ";
510   Printer.Print tb!Addr2 & ""

520   Printer.Font.Size = 4
530   Printer.Print String$(250, "-")

540   Exit Sub

PrintHeadingCavan_Error:

      Dim strES As String
      Dim intEL As Integer

550   intEL = Erl
560   strES = Err.Description
570   LogError "frmTransfusion", "PrintHeadingCavan", intEL, strES, sql

End Sub


Private Sub Form_Activate()

10    If pExternalCriteria <> "" Then
20      lblSearch = pExternalCriteria
30      btninitiate_Click
40    End If
50    SingleUserUpdateLoggedOn UserName

End Sub



Private Sub Form_Unload(Cancel As Integer)

10    pExternalCriteria = ""

End Sub

Public Property Let ExternalCriteria(ByVal strNewValue As String)

10    pExternalCriteria = strNewValue

End Property

