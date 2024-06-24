VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmPatSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Details"
   ClientHeight    =   7605
   ClientLeft      =   450
   ClientTop       =   435
   ClientWidth     =   8550
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
   Icon            =   "7frmPatSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7605
   ScaleWidth      =   8550
   Begin VB.CommandButton bXmatch 
      Enabled         =   0   'False
      Height          =   795
      Left            =   5160
      Picture         =   "7frmPatSearch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2430
      Width           =   945
   End
   Begin VB.CommandButton bprint 
      Caption         =   "Print"
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
      Left            =   6330
      Picture         =   "7frmPatSearch.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5760
      Width           =   1245
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
      Picture         =   "7frmPatSearch.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   50
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
      Picture         =   "7frmPatSearch.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   1140
      Width           =   855
   End
   Begin VB.CommandButton bhistory 
      Appearance      =   0  'Flat
      Caption         =   "&History"
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
      Left            =   4095
      Picture         =   "7frmPatSearch.frx":1AC2
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2430
      Width           =   945
   End
   Begin VB.CommandButton btncopy 
      Caption         =   "Co&py"
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
      Left            =   6240
      Picture         =   "7frmPatSearch.frx":1F04
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2430
      Width           =   945
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
      Height          =   795
      Left            =   7320
      Picture         =   "7frmPatSearch.frx":256E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2430
      Width           =   945
   End
   Begin VB.TextBox txtsearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4380
      MaxLength       =   20
      TabIndex        =   0
      Top             =   720
      Width           =   2955
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
      Picture         =   "7frmPatSearch.frx":2BD8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   1230
      TabIndex        =   55
      Top             =   7290
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
      TabIndex        =   58
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
      TabIndex        =   57
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Dob as DD/MM/YY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   7380
      TabIndex        =   54
      Top             =   450
      Width           =   1035
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
      Left            =   3870
      TabIndex        =   53
      Top             =   6570
      Width           =   915
   End
   Begin VB.Label lSampleDate 
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
      TabIndex        =   52
      Top             =   6570
      Width           =   2460
   End
   Begin VB.Label lblIdent 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1110
      TabIndex        =   51
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
      TabIndex        =   47
      Top             =   960
      Width           =   525
   End
   Begin VB.Label lMaiden 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1110
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   43
      Top             =   2070
      Width           =   510
   End
   Begin VB.Label lLabNo 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "labnumber"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5190
      TabIndex        =   42
      Top             =   2010
      Width           =   2175
   End
   Begin VB.Label lEDD 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "edd"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   2790
      TabIndex        =   41
      Top             =   2970
      Width           =   1155
   End
   Begin VB.Label lReaction 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "prevreact"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   2790
      TabIndex        =   40
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label lPPreg 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "prevpreg"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   39
      Top             =   2910
      Width           =   735
   End
   Begin VB.Label lPTrans 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "prevtrans"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      Left            =   405
      TabIndex        =   30
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
      TabIndex        =   33
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
      TabIndex        =   32
      Top             =   2430
      Width           =   1155
   End
   Begin VB.Label lComment 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "comment"
      DataSource      =   "Data1"
      Height          =   1305
      Left            =   1110
      TabIndex        =   31
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   6900
      Width           =   600
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
      TabIndex        =   19
      Top             =   450
      Width           =   2955
   End
   Begin VB.Label lNumber 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "patnum"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   18
      Top             =   450
      Width           =   2835
   End
   Begin VB.Label lName 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "name"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   17
      Top             =   705
      Width           =   2835
   End
   Begin VB.Label lSpecial 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "specialprod"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1950
      TabIndex        =   16
      Top             =   6870
      Width           =   1695
   End
   Begin VB.Label lProcedure 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "procedure"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1950
      TabIndex        =   15
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lCondx 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "conditions"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1950
      TabIndex        =   14
      Top             =   6330
      Width           =   1695
   End
   Begin VB.Label lClinician 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "clinician"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   13
      Top             =   6030
      Width           =   2535
   End
   Begin VB.Label lWard 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ward"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1110
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   480
      Width           =   555
   End
End
Attribute VB_Name = "frmPatSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Current As Integer
Dim max As Integer

Dim sn As Recordset

Private pExternalCriteria As String

Private mFrom As Form

Private Sub FillDetails(ByVal sn As Recordset)
  
10    lblKell = sn!Kell & ""
20    lSampleDate = sn!SampleDate & ""
30    lNumber = sn!Patnum & ""
40    lName = sn!Name & ""
50    lMaiden = sn!maiden & ""
60    lAddr(0) = sn!Addr1 & ""
70    lAddr(1) = sn!Addr2 & ""
80    lAddr(2) = sn!Addr3 & ""
90    lAddr(3) = sn!addr4 & ""
100   lSex = sn!Sex & ""
110   If Not IsNull(sn!DoB) Then
120     lDoB = Format(sn!DoB, "dd/mm/yyyy")
130   Else
140     lDoB = ""
150   End If
160   lPTrans = sn!prevtrans & ""
170   lReaction = sn!prevreact & ""
180   lPPreg = sn!prevpreg & ""
190   If Not IsNull(sn!edd) Then
200     lEDD = Format(sn!edd, "dd/mm/yyyy")
210   Else
220     lEDD = ""
230   End If
240   lComment = sn!Comment & ""
250   lWard = sn!Ward & ""
260   lClinician = sn!Clinician & ""
270   lCondx = sn!Conditions & ""
280   lProcedure = sn!Procedure & ""
290   lSpecial = sn!specialprod & ""
300   lGroupRh = sn!fGroup & ""
310   lAuto.Visible = False
320   If Not IsNull(sn!Autolog) Then
330     lAuto.Visible = IIf(sn!Autolog, True, False)
340   End If

350   lblIdent.Visible = False
360   If Trim$(sn!AIDR & "") <> "" Then
370     lblIdent = sn!AIDR
380     If InStr(UCase$(sn!AIDR), "POS") <> 0 Then
390       lblIdent.BackColor = vbRed
400     Else
410       lblIdent.BackColor = vbYellow
420     End If
430     lblIdent.Visible = True
440   End If

450   lTime = Format(sn!DateTime, "dd/mm/yyyy hh:mm:ss")
460   lOperator = sn!Operator & ""
470   lLabNo = sn!LabNumber & ""

End Sub

Private Sub ClearDetails()
  
10    lSampleDate = ""
20    lNumber = ""
30    lName = ""
40    lMaiden = ""
50    lAddr(0) = ""
60    lAddr(1) = ""
70    lAddr(2) = ""
80    lAddr(3) = ""
90    lSex = ""
100   lDoB = ""
110   lPTrans = ""
120   lReaction = ""
130   lPPreg = ""
140   lEDD = ""
150   lComment = ""
160   lWard = ""
170   lClinician = ""
180   lCondx = ""
190   lProcedure = ""
200   lSpecial = ""
210   lGroupRh = ""
220   lAuto.Visible = False
230   lblIdent = ""
240   lblIdent.Visible = False
250   lTime = ""
260   lOperator = ""
270   lLabNo = ""
280   lblStatus = ""

End Sub


Private Sub bHistory_Click()

10    If lNumber = "" Then
20        If iMsg("MRN not found. Do you wish to perform a search with patient name?", vbQuestion + vbYesNo) = vbYes Then
30            fpathistory.optName = True
40            fpathistory.txtName = lName
50            fpathistory.cmdSearch = True
60            fpathistory.Show 1
70        End If
80    Else
90        fpathistory.optChart = True
100       fpathistory.txtName = lNumber
110       fpathistory.cmdSearch = True
120       fpathistory.Show 1
130   End If
End Sub

Private Sub bNext_Click()

10    If Current = max Then Exit Sub

20    Current = Current + 1
  
30    lblStatus = "Record " & Format(Current) & " of " & Format(max)

40    sn.MoveNext
50    FillDetails sn

60    bPrevious.Enabled = True
70    bNext.Enabled = Current <> max

End Sub

Private Sub bPrevious_Click()

10    If max = 0 Then Exit Sub
20    If Current = 1 Then Exit Sub

30    Current = Current - 1
  
40    lblStatus = "Record " & Format(Current) & " of " & Format(max)

50    sn.MovePrevious
60    FillDetails sn

70    bPrevious.Enabled = Current <> 1
80    bNext.Enabled = True

End Sub

Private Sub bprint_Click()

10    fPrintForm.SampleID = lLabNo
20    fPrintForm.Show 1

End Sub

Private Sub btnCancel_Click()

10    Unload Me

End Sub

Private Sub btncopy_Click()

      Dim f As Form
      Dim n As Integer

10    On Error GoTo btncopy_Click_Error

20    If mFrom Is frmxmatch Then
30      Set f = mFrom
40    Else
50      Exit Sub
60    End If

70    f.txtChart = lNumber
80    f.txtName = lName
90    f.tMaiden = lMaiden
100   For n = 0 To 3: f.tAddr(n) = lAddr(n): Next
110   If Dept <> ANTENATAL Then f.lSex = lSex
120   f.tDoB = lDoB
130   f.cWard = lWard
140   f.cClinician = lClinician
150   f.cConditions = lCondx
160   f.cProcedure = lProcedure
170   f.cSpecial = lSpecial
180   f.tComment = StripComment(lComment)
190   f.tptrans = lPTrans
200   f.tpreaction = lReaction
210   f.tedd = lEDD
220   Unload Me

230   Exit Sub

btncopy_Click_Error:

 Dim strES As String
 Dim intEL As Integer

240    intEL = Erl
250    strES = Err.Description
260    LogError "frmPatSearch", "btncopy_Click", intEL, strES

End Sub

Private Sub btninitiate_Click()

      Dim sql As String
      Dim sqlBase As String
      Dim strSearch As String

10    On Error GoTo btninitiate_Click_Error

20    If Trim$(txtsearch) = "" Then Exit Sub

30    sqlBase = "SELECT * FROM PatientDetails WHERE "

40    max = 0
'40    strSearch = Convert62Date(txtsearch, DONTCARE)
50    strSearch = txtsearch
60    If IsDate(strSearch) Then
70      txtsearch = strSearch
80      sql = sqlBase & "DoB = '" & Format(strSearch, "dd/mmm/yyyy") & "' ORDER BY SampleDate"
90      Set sn = New Recordset
100     RecOpenClientBB 0, sn, sql
110     If Not sn.EOF Then
120       sn.MoveLast
130       max = sn.RecordCount
140     End If
150   End If

160   If max = 0 Then
170     sql = sqlBase & "PatNum = '" & AddTicks(txtsearch) & "' ORDER BY SampleDate"
180     Set sn = New Recordset
190     RecOpenClientBB 0, sn, sql
200     If Not sn.EOF Then
210       sn.MoveLast
220       max = sn.RecordCount
230     End If
240   End If

250   If max = 0 Then
260     sql = sqlBase & "Name LIKE '" & AddTicks(txtsearch) & "%' ORDER BY SampleDate"
270     Set sn = New Recordset
280     RecOpenClientBB 0, sn, sql
290     If Not sn.EOF Then
300       sn.MoveLast
310       max = sn.RecordCount
320     End If
330   End If

340   If max = 0 Then
350     lblStatus = "Not found."
360   Else
370     lblStatus = "Record " & Format(max) & " of " & Format(max)
380     FillDetails sn
390     bXmatch.Enabled = True
400   End If

410   Current = max

420   bPrevious.Enabled = max > 1
430   bNext.Enabled = False

440   Exit Sub

btninitiate_Click_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "frmPatSearch", "btninitiate_Click", intEL, strES, sql

End Sub


Private Sub bxmatch_Click()

SaveSetting "NetAcquire", "Transfusion6", "Lastused", lLabNo

10    With frmxmatch
15        .tLabNum = lLabNo
20        .cmdSave.Enabled = False
30        .bHold.Enabled = False
40        .cmdOrderAutovue.Enabled = False
50        .bIssueBatch.Enabled = False
60        .bPrepare(0).Enabled = False
70        .bPrepare(1).Enabled = False
80        .cmdIssueToUnknown.Enabled = False
90        .cmdSearch.Enabled = False
100       .btnprint.Enabled = False
110       .bPrintForm.Enabled = False
120       .bPrintDAT.Enabled = False
130       .lblOrderDAT.Enabled = False
140       .cmbKell.Enabled = False
150       .cClinician.Enabled = False
160       .cConditions.Enabled = False
170       .cGP.Enabled = False
180       .dtDateRxd.Enabled = False
190       .dtTimeRxd.Enabled = False
200       .gDAT.Enabled = False
210       .cProcedure.Enabled = False
220       .cSpecial.Enabled = False
230       .cWard.Enabled = False
240       .iprevious.Enabled = False
250       .lblgrpchecker.Enabled = False
260       .lSex.Enabled = False
270       .lstfg.Enabled = False
280       .lstRG.Enabled = False
290       .FramePP.Enabled = False
300       .tAandE.Enabled = False
310       .tAddr(0).Enabled = False
320       .tAddr(1).Enabled = False
330       .tAddr(2).Enabled = False
340       .tAddr(3).Enabled = False
350       .tAge.Enabled = False
360       .tComment.Enabled = False
370       .tDoB.Enabled = False
380       .tedd.Enabled = False
390       .tident.Enabled = False
400       .tMaiden.Enabled = False
410       .txtName.Enabled = False
420       .tTypenex.Enabled = False
430       .txtChart.Enabled = False
440       .txtNOPAS.Enabled = False
450       .txtSampleTime.Enabled = False
460       .tSampleComment.Enabled = False
470       .txtSampleDate.Enabled = False

480       .tLabNum.Locked = True
490       .udLabNum.Enabled = False
500       .imgUseTime.Enabled = False


520       .Show 1
530   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    pExternalCriteria = ""

End Sub

Private Sub txtsearch_KeyUp(KeyCode As Integer, Shift As Integer)

10    ClearDetails

End Sub



Public Property Let ExternalCriteria(ByVal strNewValue As String)

10    pExternalCriteria = strNewValue

End Property

Public Property Let From(ByVal vNewValue As Form)

10    Set mFrom = vNewValue

End Property
