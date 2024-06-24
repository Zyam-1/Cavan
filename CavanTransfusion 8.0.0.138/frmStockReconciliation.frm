VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStockReconciliation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Stock Reconciliation"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   765
      HelpContextID   =   10090
      Left            =   8670
      Picture         =   "frmStockReconciliation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   765
      HelpContextID   =   10070
      Left            =   6450
      Picture         =   "frmStockReconciliation.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   765
      Left            =   7560
      Picture         =   "frmStockReconciliation.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid gExport 
      Height          =   1815
      Left            =   2580
      TabIndex        =   21
      Top             =   8340
      Visible         =   0   'False
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Height          =   8745
      Left            =   120
      TabIndex        =   8
      Top             =   1110
      Width           =   13125
      Begin MSFlexGridLib.MSFlexGrid gInStock 
         Height          =   1815
         Left            =   60
         TabIndex        =   9
         Top             =   420
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483634
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid gReturned 
         Height          =   1815
         Left            =   6600
         TabIndex        =   10
         Top             =   420
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483634
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid gDestroyed 
         Height          =   1815
         Left            =   60
         TabIndex        =   11
         Top             =   2580
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483634
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid gDispatch 
         Height          =   1815
         Left            =   6600
         TabIndex        =   12
         Top             =   2580
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483634
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid gCrossMatched 
         Height          =   1815
         Left            =   60
         TabIndex        =   13
         Top             =   4740
         Width           =   13005
         _ExtentX        =   22939
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483634
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid gTransfused 
         Height          =   1815
         Left            =   60
         TabIndex        =   14
         Top             =   6900
         Width           =   13005
         _ExtentX        =   22939
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483634
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Issued"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   2100
         TabIndex        =   26
         Top             =   4500
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Units in Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   20
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Units Returned to Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6600
         TabIndex        =   19
         Top             =   180
         Width           =   2475
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Units Destroyed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   2340
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Units Dispatched"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6600
         TabIndex        =   17
         Top             =   2340
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Units Crossmatched / "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   16
         Top             =   4500
         Width           =   2040
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Units Transfused"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   15
         Top             =   6660
         Width           =   1530
      End
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   360
      Left            =   1455
      TabIndex        =   1
      Top             =   495
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   157417473
      CurrentDate     =   39730
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   9900
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   360
      Left            =   4170
      TabIndex        =   5
      Top             =   495
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   157417473
      CurrentDate     =   39730
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   315
      Left            =   9825
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblExcelInfo 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9825
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stock Reconciliation Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   9825
      TabIndex        =   6
      Top             =   480
      Width           =   3420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   150
      X2              =   13200
      Y1              =   1050
      Y2              =   1065
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   4
      Top             =   570
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3780
      TabIndex        =   3
      Top             =   570
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Expiry Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "frmStockReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()

      Dim RecordsFound As Boolean

10    On Error GoTo cmdSearch_Click_Error

20    If DateDiff("d", dtFrom.Value, dtTo.Value) < 0 Then
30        iMsg "'To date' must be greater than 'From date'"
40        If TimedOut Then Unload Me: Exit Sub
50        Exit Sub
60    End If

70    RecordsFound = FillGrids1()
80    RecordsFound = RecordsFound Or FillGridsCDFT()
  
90    If Not RecordsFound Then
100     iMsg "Sorry, No records found" & vbCrLf & "Please change expiry date range and try again"
110     If TimedOut Then Unload Me: Exit Sub
120   End If

130   Exit Sub

cmdSearch_Click_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmStockReconciliation", "cmdSearch_Click", intEL, strES

End Sub

Private Sub cmdExit_Click()

10    Unload Me

End Sub

Private Sub cmdXL_Click()

      Dim strHeading As String
      Dim StartRow As Integer

10    On Error GoTo cmdXL_Click_Error

20    If gCrossMatched.Rows = 1 And gDestroyed.Rows = 1 And gDispatch.Rows = 1 And _
          gTransfused.Rows = 1 And gCrossMatched.Rows = 1 And gInStock.Rows = 1 Then
30        iMsg "Nothing to export"
40        If TimedOut Then Unload Me: Exit Sub
50        Exit Sub
60    End If

70    strHeading = "Stock Reconciliation Report" & vbCr & _
              "From " & dtFrom.Value & " To " & dtTo.Value & vbCr & vbCr
    
80    With gExport
90        .Clear
100       .Rows = gCrossMatched.Rows + gDestroyed.Rows + gDispatch.Rows + _
                  gInStock.Rows + gReturned.Rows + gTransfused.Rows + 10
110       .Cols = 11
120       .FixedRows = 0: .FixedCols = 0
130       .ColWidth(0) = 0
140   End With
150   StartRow = 0

160   If gInStock.Rows > 1 Then
170       gExport.AddItem "In Stock Units", StartRow: StartRow = StartRow + 1
180       CopyGrid gInStock, StartRow
190       StartRow = StartRow + gInStock.Rows + 1
200   End If

210   If gReturned.Rows > 1 Then
220       gExport.AddItem "Returned Units", StartRow: StartRow = StartRow + 1
230       CopyGrid gReturned, StartRow
240       StartRow = StartRow + gReturned.Rows + 1
250   End If
    
260   If gDestroyed.Rows > 1 Then
270       gExport.AddItem "Destroyed Units", StartRow: StartRow = StartRow + 1
280       CopyGrid gDestroyed, StartRow
290       StartRow = StartRow + gDestroyed.Rows + 1
300   End If

310   If gDispatch.Rows > 1 Then
320       gExport.AddItem "Dispatched Units", StartRow: StartRow = StartRow + 1
330       CopyGrid gDispatch, StartRow
340       StartRow = StartRow + gDispatch.Rows + 1
350   End If

360   If gCrossMatched.Rows > 1 Then
370       gExport.AddItem "Crossmatched Units", StartRow: StartRow = StartRow + 1
380       CopyGrid gCrossMatched, StartRow
390       StartRow = StartRow + gCrossMatched.Rows + 1
400   End If

410   If gTransfused.Rows > 1 Then
420       gExport.AddItem "Transfused Units", StartRow: StartRow = StartRow + 1
430       CopyGrid gTransfused, StartRow
440   End If
450   ExportFlexGrid gExport, Me, strHeading

460   Exit Sub

cmdXL_Click_Error:

      Dim strES As String
      Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "frmStockReconciliation", "cmdXL_Click", intEL, strES

End Sub

Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    InitGridgCrossMatched
30    InitGridgDestroyed
40    InitGridgDispatch
50    InitGridgInStock
60    InitGridgReturned
70    InitGridgTransfused

80    dtFrom.Value = Format(Date - 30, "dd/MM/yyyy")
90    dtTo.Value = Format(Date, "dd/MM/yyyy")

100   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmStockReconciliation", "Form_Load", intEL, strES

End Sub


Private Sub InitGridgInStock()

10    On Error GoTo InitGridgInStock_Error

20    With gInStock
30        .Rows = 2: .Cols = 7
40        .FixedRows = 1: .FixedCols = 0
50        .Rows = 1
60        .ColWidth(0) = 0
    
70        .TextMatrix(0, 1) = "Unit No.": .ColWidth(1) = 1500
80        .TextMatrix(0, 2) = "Group": .ColWidth(2) = 600
90        .TextMatrix(0, 3) = "Product": .ColWidth(3) = 2000
100       .TextMatrix(0, 4) = "Expiry": .ColWidth(4) = 1550
110       .TextMatrix(0, 5) = "Date Rec'd": .ColWidth(5) = 1650
120       .TextMatrix(0, 6) = "CHECKED BY": .ColWidth(6) = 1200
    
130   End With

140   Exit Sub

InitGridgInStock_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmStockReconciliation", "InitGridgInStock", intEL, strES

End Sub

Private Sub InitGridgReturned()

10    On Error GoTo InitGridgReturned_Error

20    With gReturned
30        .Rows = 2: .Cols = 7
40        .FixedRows = 1: .FixedCols = 0
50        .Rows = 1
60        .ColWidth(0) = 0
    
70        .TextMatrix(0, 1) = "Unit No.": .ColWidth(1) = 1500
80        .TextMatrix(0, 2) = "Group": .ColWidth(2) = 600
90        .TextMatrix(0, 3) = "Product": .ColWidth(3) = 2000
100       .TextMatrix(0, 4) = "Expiry": .ColWidth(4) = 1550
110       .TextMatrix(0, 5) = "Date Rec'd": .ColWidth(5) = 1650
120       .TextMatrix(0, 6) = "CHECKED BY": .ColWidth(6) = 1200
    
130   End With

140   Exit Sub

InitGridgReturned_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmStockReconciliation", "InitGridgReturned", intEL, strES

End Sub
Private Sub InitGridgDestroyed()

10    On Error GoTo InitGridgDestroyed_Error

20    With gDestroyed
  
30      .Rows = 2: .Cols = 7
40      .FixedRows = 1: .FixedCols = 0
50      .Rows = 1
60      .ColWidth(0) = 0
  
70      .TextMatrix(0, 1) = "Unit No.": .ColWidth(1) = 1500
80      .TextMatrix(0, 2) = "Group": .ColWidth(2) = 600
90      .TextMatrix(0, 3) = "Product": .ColWidth(3) = 2000
100     .TextMatrix(0, 4) = "Expiry": .ColWidth(4) = 1550
110     .TextMatrix(0, 5) = "Date Rec'd": .ColWidth(5) = 1650
120     .TextMatrix(0, 6) = "CHECKED BY": .ColWidth(6) = 1200
    
130   End With

140   Exit Sub

InitGridgDestroyed_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmStockReconciliation", "InitGridgDestroyed", intEL, strES

End Sub
Private Sub InitGridgDispatch()

10    On Error GoTo InitGridgDispatch_Error

20    With gDispatch
30        .Rows = 2: .Cols = 7
40        .FixedRows = 1: .FixedCols = 0
50        .Rows = 1
60        .ColWidth(0) = 0
    
70        .TextMatrix(0, 1) = "Unit No.": .ColWidth(1) = 1500
80        .TextMatrix(0, 2) = "Group": .ColWidth(2) = 600
90        .TextMatrix(0, 3) = "Product": .ColWidth(3) = 2000
100       .TextMatrix(0, 4) = "Expiry": .ColWidth(4) = 1550
110       .TextMatrix(0, 5) = "Date Rec'd": .ColWidth(5) = 1650
120       .TextMatrix(0, 6) = "CHECKED BY": .ColWidth(6) = 1200
    
130   End With

140   Exit Sub

InitGridgDispatch_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmStockReconciliation", "InitGridgDispatch", intEL, strES

End Sub
Private Sub InitGridgCrossMatched()

10    On Error GoTo InitGridgCrossMatched_Error

20    With gCrossMatched
30        .Rows = 2: .Cols = 11
40        .FixedRows = 1: .FixedCols = 0
50        .Rows = 1
60        .ColWidth(0) = 0
    
70        .TextMatrix(0, 1) = "Unit No.": .ColWidth(1) = 1500
80        .TextMatrix(0, 2) = "Group": .ColWidth(2) = 600
90        .TextMatrix(0, 3) = "Product": .ColWidth(3) = 2000
100       .TextMatrix(0, 4) = "Expiry": .ColWidth(4) = 1550
110       .TextMatrix(0, 5) = "Date X-M": .ColWidth(5) = 1650
120       .TextMatrix(0, 6) = "Patient": .ColWidth(6) = 2000
130       .TextMatrix(0, 7) = "MRN": .ColWidth(7) = 900
140       .TextMatrix(0, 8) = "DoB": .ColWidth(8) = 1000
150       .TextMatrix(0, 9) = "Ward": .ColWidth(9) = 1500
160       .TextMatrix(0, 10) = "CHECKED BY": .ColWidth(10) = 1200
170   End With

180   Exit Sub

InitGridgCrossMatched_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmStockReconciliation", "InitGridgCrossMatched", intEL, strES

End Sub
Private Sub InitGridgTransfused()

10    On Error GoTo InitGridgTransfused_Error

20    With gTransfused
30        .Rows = 2: .Cols = 11
40        .FixedRows = 1: .FixedCols = 0
50        .Rows = 1
60        .ColWidth(0) = 0
    
70        .TextMatrix(0, 1) = "Unit No.": .ColWidth(1) = 1500
80        .TextMatrix(0, 2) = "Group": .ColWidth(2) = 600
90        .TextMatrix(0, 3) = "Product": .ColWidth(3) = 2000
100       .TextMatrix(0, 4) = "Expiry": .ColWidth(4) = 1550
110       .TextMatrix(0, 5) = "Date X-M": .ColWidth(5) = 1650
120       .TextMatrix(0, 6) = "Patient": .ColWidth(6) = 2000
130       .TextMatrix(0, 7) = "MRN": .ColWidth(7) = 900
140       .TextMatrix(0, 8) = "DoB": .ColWidth(8) = 1000
150       .TextMatrix(0, 9) = "Ward": .ColWidth(9) = 1500
160       .TextMatrix(0, 10) = "CHECKED BY": .ColWidth(10) = 1200
    
170   End With

180   Exit Sub

InitGridgTransfused_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmStockReconciliation", "InitGridgTransfused", intEL, strES

End Sub



Private Function FillGrids1() As Boolean
      'Returns True if any records found

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillGrids1_Error
    
20    InitGridgCrossMatched
30    InitGridgTransfused

40    sql = "SELECT L.*,P.ward,P.DoB " & _
            "FROM Latest L " & _
            "LEFT JOIN PatientDetails P " & _
            "ON L.LabNumber = P.LabNumber " & _
            "WHERE (DateExpiry BETWEEN CONVERT(DATETIME, 'xFromDatex', 103) AND CONVERT(DATETIME, 'xToDatex', 103)) " & _
            "AND (Event = 'S' OR Event = 'X' OR Event = 'I') " & _
            "ORDER BY L.[DateTime]"
  
50    sql = Replace(sql, "xFromDatex", dtFrom.Value)
60    sql = Replace(sql, "xToDatex", dtTo.Value)

70    Set tb = New Recordset

80    RecOpenClientBB 0, tb, sql

90    With tb
100     If Not .EOF Then

110       FillGrids1 = True
120       PB.Visible = True
130       PB.Min = 0
140       PB.max = .RecordCount + 1
150       PB.Value = 0
160       tb.MoveFirst
170       Do While Not .EOF
180         s = "" & vbTab & !ISBT128 & "" & vbTab & _
                Bar2Group(!GroupRh) & "" & vbTab & _
                ProductWordingFor(!BarCode) & "" & vbTab & _
                Format(!DateExpiry & "", "dd/mmm/yyyy HH:mm") & vbTab & _
                Format(!DateTime, "dd/MM/yyyy hh:mm:ss") & ""
190         If UCase$(!Event) = "S" Or UCase$(!Event) = "X" Or UCase$(!Event) = "I" Or UCase$(!Event) = "V" Then
200           s = s & vbTab & _
                 !PatName & "" & vbTab & _
                 !Patid & "" & vbTab & _
                 !DoB & "" & vbTab & _
                 !Ward & ""
210         End If
220         Select Case UCase$(!Event)
              Case "S":
                  'Transfused Units
230               gTransfused.AddItem s, gTransfused.Rows
240     Case "X", "I", "V":
                  'Crossmatched units
250               gCrossMatched.AddItem s, gCrossMatched.Rows
260               If UCase$(!Event) = "I" Or UCase$(!Event) = "V" Then
270                 gCrossMatched.row = gCrossMatched.Rows - 1
280                 gCrossMatched.col = 1
290                 gCrossMatched.CellBackColor = vbYellow
300               End If
310         End Select
320         PB.Value = PB.Value + 1
330         .MoveNext
340       Loop
350       PB.Visible = False
360     Else
370       FillGrids1 = False
380     End If

390   End With

400   Exit Function

FillGrids1_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "frmStockReconciliation", "FillGrids1", intEL, strES, sql
440   PB.Visible = False

End Function

Private Function FillGridsCDFT() As Boolean
      'Returns True if any records found

      Dim tb As Recordset
      Dim tbD As Recordset
      Dim sql As String
      Dim s As String
      Dim RxDate As String

10    On Error GoTo FillGridsCDFT_Error
    
20    InitGridgDestroyed
30    InitGridgDispatch
40    InitGridgInStock
50    InitGridgReturned

60    sql = "SELECT L.ISBT128, L.Event, L.BarCode, L.DateExpiry, L.GroupRh " & _
            "FROM Latest L " & _
            "WHERE (L.DateExpiry BETWEEN CONVERT(DATETIME, 'xFromDatex', 103) " & _
            "       AND CONVERT(DATETIME, 'xToDatex', 103)) " & _
            "AND L.Event IN ('C', 'D', 'F', 'T', 'N') " & _
            "ORDER BY L.[DateTime]"
  
70    sql = Replace(sql, "xFromDatex", dtFrom.Value)
80    sql = Replace(sql, "xToDatex", dtTo.Value)

90    Set tb = New Recordset

100   RecOpenClientBB 0, tb, sql

110   With tb
120     If Not .EOF Then
130       FillGridsCDFT = True
    
140       PB.Visible = True
150       PB.Min = 0
160       PB.max = .RecordCount + 1
170       PB.Value = 0
180       tb.MoveFirst
190       Do While Not .EOF
200         sql = "SELECT DateTime FROM Product WHERE " & _
                  "Event = 'C' " & _
                  "AND ISBT128 = '" & !ISBT128 & "' AND DateExpiry = '" & Format(!DateExpiry, "dd/MMM/yyyy HH:mm") & "'"
210         Set tbD = New Recordset
220         RecOpenServerBB 0, tbD, sql
230         If Not tbD.EOF Then
240           RxDate = Format(tbD!DateTime, "dd/MM/yyyy HH:nn:ss")
250         Else
260           RxDate = ""
270         End If
280         s = "" & vbTab & !ISBT128 & "" & vbTab & _
                Bar2Group(!GroupRh) & "" & vbTab & _
                ProductWordingFor(!BarCode) & "" & vbTab & _
                Format(!DateExpiry & "", "dd/mmm/yyyy HH:mm") & vbTab & _
                RxDate
290         Select Case UCase$(!Event)
              Case "D":
                  'Destroyed Units
300               gDestroyed.AddItem s, gDestroyed.Rows
310     Case "F", "N":
                  'Displtched Units
320               gDispatch.AddItem s, gDispatch.Rows
330     Case "T":
                  'Returned to supplier units
340               gReturned.AddItem s, gReturned.Rows
350     Case "C":
                  'Received but not used (In stock units)
360               gInStock.AddItem s, gInStock.Rows
370         End Select
380         PB.Value = PB.Value + 1
390         .MoveNext
400       Loop
410       PB.Visible = False
420     Else
430       FillGridsCDFT = False
440     End If

450   End With

460   Exit Function

FillGridsCDFT_Error:

      Dim strES As String
      Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "frmStockReconciliation", "FillGridsCDFT", intEL, strES, sql
500   FillGridsCDFT = False

End Function

Private Sub CopyGrid(gSource As MSFlexGrid, StartRow As Integer)

10    On Error GoTo CopyGrid_Error

20    With gSource
30        .row = 0
40        .col = 0
50        .RowSel = .Rows - 1
60        .ColSel = .Cols - 1
70    End With

80    With gExport
90        .row = StartRow
100       .col = 0
110       .RowSel = StartRow + gSource.Rows - 1
120       .ColSel = gSource.Cols - 1
130       .Clip = gSource.Clip
140   End With

150   Exit Sub

CopyGrid_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmStockReconciliation", "CopyGrid", intEL, strES

End Sub

