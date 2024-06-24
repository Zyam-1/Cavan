VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTotals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Totals"
   ClientHeight    =   7800
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   9495
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
   Icon            =   "7frmTotals.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7800
   ScaleWidth      =   9495
   Begin VB.CommandButton cmdPrint 
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
      Height          =   765
      Left            =   6030
      Picture         =   "7frmTotals.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6405
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1020
      Width           =   9195
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   765
      Left            =   8070
      Picture         =   "7frmTotals.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4050
      Picture         =   "7frmTotals.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   2010
      TabIndex        =   5
      Top             =   480
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   7
      Top             =   7560
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Between Dates"
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
      Left            =   480
      TabIndex        =   6
      Top             =   210
      Width           =   2925
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

10    If Not SetFormPrinter() Then Exit Sub

20    Printer.Print
30    Printer.Print t
40    Printer.EndDoc

End Sub

Private Sub cmdStart_Click()

      Dim tb As Recordset
      Dim sqlP As String
      Dim sn As Recordset
      Dim sql As String
      Dim strstrFromTime As String
      Dim strToTime As String
      Dim s As String
      Dim strGr As String
      Dim intN As Integer

10    On Error GoTo cmdstart_Click_Error

20    t = ""

30    strstrFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
40    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

50    sql = "select count(*) as tot from product where " & _
            "datetime between '" & strstrFromTime & "' " & _
            "and '" & strToTime & "' " & _
            "and event = '"
60    Set sn = New Recordset
70    RecOpenClientBB 0, sn, sql & "C'"
80    t = t & "Total products received: " & Format(sn!Tot) & vbCrLf & vbCrLf

90    sqlP = "Select * from ProductList"
100   Set tb = New Recordset
110   RecOpenServerBB 0, tb, sqlP
120   Do While Not tb.EOF
130     Set sn = New Recordset
140     RecOpenClientBB 0, sn, sql & "C' and barcode = '" & tb!BarCode & "'"
150     If sn!Tot <> 0 Then
160       s = tb!Wording
170       t = t & vbCrLf & s & ": " & Format(sn!Tot)
180       t = t & vbCrLf
190       For intN = 1 To 8
200         strGr = Choose(intN, "51", "62", "73", "84", "95", "06", "17", "28")
210         Set sn = New Recordset
220         RecOpenClientBB 0, sn, sql & "C' and barcode = '" & tb!BarCode & "' and grouprh = '" & strGr & "'"
230         t = t & Bar2Group(strGr) & " " & _
                    Format(sn!Tot) & "  "
240       Next
250       t = t & vbCrLf & vbCrLf
260     End If
270     tb.MoveNext
280   Loop

290   t = t & vbCrLf & vbCrLf & "---Totals---" & vbCrLf

300   sql = "select count(*) as Tot from product where " & _
            "(datetime between '" & strstrFromTime & "' " & _
            "and '" & _
            strToTime & "') and event = '"
310   Set sn = New Recordset
320   RecOpenClientBB 0, sn, sql & "T' "
330   t = t & "Returned: " & Format(sn!Tot) & vbCrLf

340   Set sn = New Recordset
350   RecOpenClientBB 0, sn, sql & "D' "
360   t = t & "Destroyed: " & Format(sn!Tot) & vbCrLf

370   Set sn = New Recordset
380   RecOpenClientBB 0, sn, sql & "S' "
390   t = t & "Transfused: " & Format(sn!Tot) & vbCrLf & vbCrLf

400   Set sn = New Recordset
410   RecOpenClientBB 0, sn, sql & "X' and labnumber<>''"
420   t = t & "Cross Matches: " & Format(sn!Tot) & vbCrLf

430   sql = "select count(*) as tot from patientdetails where " & _
            "datetime between '" & strstrFromTime & "' " & _
            "and '" & _
            strToTime & "' and requestfrom = '"
440   Set sn = New Recordset
450   RecOpenClientBB 0, sn, sql & "G' and labnumber<>''"
460   t = t & "Group and Hold: " & Format(sn!Tot) & vbCrLf

470   Set sn = New Recordset
480   RecOpenClientBB 0, sn, sql & "A'"
490   t = t & "Ante-Natal: " & Format(sn!Tot) & vbCrLf
  
500   sql = "select count(distinct labnumber) as tot from patientdetails where " & _
            "(dat0 = 1 or dat1 = 1 or dat2 = 1 or dat3 = 1 " & _
            "or dat4 = 1 or dat5 = 1 or dat6 = 1 or dat7 = 1 " & _
            "or dat8 = 1 or dat9 = 1 or requestfrom = 'D') and " & _
            "(datetime between '" & strstrFromTime & _
            "' and '" & strToTime & "')"
510   Set sn = New Recordset
520   RecOpenClientBB 0, sn, sql
530   t = t & "D.A.T.: " & Format(sn!Tot) & vbCrLf

540   Exit Sub

cmdstart_Click_Error:

      Dim strES As String
      Dim intEL As Integer

550   intEL = Erl
560   strES = Err.Description
570   LogError "frmTotals", "cmdstart_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10    dtTo = Format(Now, "dd/mmm/yyyy")
20    dtFrom = Format(Now - 7, "dd/mmm/yyyy")

End Sub

