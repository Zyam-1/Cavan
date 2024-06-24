VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmViewCourier 
   Caption         =   "NetAcquire - View Courier Log"
   ClientHeight    =   5445
   ClientLeft      =   285
   ClientTop       =   1620
   ClientWidth     =   10830
   ControlBox      =   0   'False
   Icon            =   "frmViewCourier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10830
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   9690
      Picture         =   "frmViewCourier.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   345
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message to View"
      Height          =   1035
      Left            =   1680
      TabIndex        =   4
      Top             =   90
      Width           =   7785
      Begin VB.OptionButton optMessage 
         Alignment       =   1  'Right Justify
         Caption         =   "Reserve Stock (RS3 )"
         Height          =   195
         Index           =   10
         Left            =   3630
         TabIndex        =   17
         Tag             =   "RS3"
         Top             =   690
         Width           =   1905
      End
      Begin VB.OptionButton optMessage 
         Caption         =   "Sample Valid Response"
         Height          =   195
         Index           =   8
         Left            =   5610
         TabIndex        =   14
         Tag             =   "SVR"
         Top             =   450
         Width           =   1995
      End
      Begin VB.OptionButton optMessage 
         Caption         =   "Sample Valid Query"
         Height          =   195
         Index           =   9
         Left            =   5610
         TabIndex        =   13
         Tag             =   "SVQ"
         Top             =   210
         Width           =   1725
      End
      Begin VB.OptionButton optMessage 
         Caption         =   "Stock Transfer"
         Height          =   195
         Index           =   7
         Left            =   2025
         TabIndex        =   12
         Tag             =   "ST"
         Top             =   690
         Width           =   1395
      End
      Begin VB.OptionButton optMessage 
         Alignment       =   1  'Right Justify
         Caption         =   "RER Response"
         Height          =   195
         Index           =   6
         Left            =   4110
         TabIndex        =   11
         Tag             =   "EIR"
         Top             =   450
         Width           =   1425
      End
      Begin VB.OptionButton optMessage 
         Alignment       =   1  'Right Justify
         Caption         =   "RER Query"
         Height          =   195
         Index           =   5
         Left            =   4410
         TabIndex        =   10
         Tag             =   "EIQ"
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton optMessage 
         Caption         =   "Stock Update (SU3)"
         Height          =   195
         Index           =   4
         Left            =   2025
         TabIndex        =   9
         Tag             =   "SU3"
         Top             =   210
         Width           =   1800
      End
      Begin VB.OptionButton optMessage 
         Caption         =   "Fate of Unit"
         Height          =   195
         Index           =   3
         Left            =   2025
         TabIndex        =   8
         Tag             =   "FT"
         Top             =   450
         Width           =   1155
      End
      Begin VB.OptionButton optMessage 
         Alignment       =   1  'Right Justify
         Caption         =   "Return to Stock"
         Height          =   210
         Index           =   2
         Left            =   60
         TabIndex        =   7
         Tag             =   "RTS"
         Top             =   690
         Width           =   1890
      End
      Begin VB.OptionButton optMessage 
         Alignment       =   1  'Right Justify
         Caption         =   "Stock Movement"
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Tag             =   "SM"
         Top             =   450
         Width           =   1890
      End
      Begin VB.OptionButton optMessage 
         Alignment       =   1  'Right Justify
         Caption         =   "Reserve Stock (RS )"
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Tag             =   "RS3"
         Top             =   195
         Width           =   1890
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3705
      Left            =   60
      TabIndex        =   3
      Top             =   1260
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   6535
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1035
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1605
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95617025
         CurrentDate     =   38674
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95617025
         CurrentDate     =   38674
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   60
      TabIndex        =   16
      Top             =   5070
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmViewCourier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub SetGridFormat(ByVal Message As String)

10    Select Case Message
        Case "RS"
20        grd.Cols = 13
30        grd.FormatString = "<Date/Time              |<Unit No.                |<Product      |<Location  |<Expiry Date  " & _
                             "|<Unit Group  |<Stock Comment |<Patient Chart " & _
                             "|<Patient Name |<Date of Birth|<Sex      |<Patient Group" & _
                             "|<Dereservation Date/Time"
40      Case "RS3"
50        grd.Cols = 13
60        grd.FormatString = "<Date/Time              |<Unit No.                |<Product      |<Location  |<Expiry Date  " & _
                             "|<Unit Group  |<Stock Comment |<Patient Chart " & _
                             "|<Patient Name |<Date of Birth|<Sex      |<Patient Group" & _
                             "|<Dereservation Date/Time"
70      Case "SM", "RTS"
80        grd.Cols = 6
90        grd.FormatString = "<Date/Time              |<Unit No.                  " & _
                             "|<Product                                                          " & _
                             "|<Location  |<Action |<User Name "
        
100     Case "FT"
110       grd.Cols = 11
120       grd.FormatString = "<Date/Time             |<Unit No.                |<Product      |<Chart  " & _
                             "|<Patient Name  |<Date of Birth|<Sex      |<Location  |<Status         " & _
                             "|<User Name |<Date/Time "
        
        
130     Case "SU3"
140       grd.Cols = 7
150       grd.FormatString = "<Date/Time              |<Unit No.                " & _
                             "|<Product                                      " & _
                             "|<Location  |<Expiry Date  " & _
                             "|<Unit Group  |<Stock Comment                    "

160     Case "EIQ", "SVQ"
170       grd.Cols = 5
180       grd.FormatString = "<Date/Time              |<Patient Chart " & _
                             "|<Patient Name                                     |<Date of Birth|<Sex      "
        
        
190     Case "EIR"
200       grd.Cols = 8
210       grd.FormatString = "<Date/Time              |<Patient Chart |<Patient Name |<Date of Birth|<Sex      " & _
                             "|<Patient Group |<RER Status         |<RER Expiry Date   "

220     Case "ST"
230       grd.Cols = 9
240       grd.FormatString = "<Date/Time              |<Unit No.                |<Product      |<Location  |<User " & _
                             "|<Patient Chart |<Patient Name |<Date of Birth|<Sex      "

250     Case "SVR"
260       grd.Cols = 10
270       grd.FormatString = "<Date/Time              |<Patient Chart |<Patient Name |<Date of Birth|<Sex      " & _
                             "|<Patient Group |<RER Status         |<RER Expiry Date   " & _
                             "|<Sample Status |<Sample Valid Date "

280   End Select

End Sub

Private Sub FillGrid(ByVal Message As String)

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo FillGrid_Error

20    grd.Rows = 2
30    grd.AddItem ""
40    grd.RemoveItem 1

50    sql = "Select * from Courier where " & _
            "Identifier = '" & Message & "' " & _
            "and MessageTime between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
            "Order by MessageTime desc"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF

90      Select Case Message
          Case "RS", "RS3"
100         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!UnitNumber & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                tb!UnitExpiry & vbTab & _
                ReadableGroup(tb!UnitGroup & "") & vbTab & _
                tb!StockComment & vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "") & vbTab & _
                ReadableGroup(tb!PatientGroup & "") & vbTab & _
                tb!DeReservationDateTime
        
110       Case "SM", "RTS"
120         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!UnitNumber & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                tb!ActionText & vbTab & _
                tb!UserName & ""
          
130       Case "FT"
140         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!UnitNumber & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "") & vbTab & _
                tb!Location & vbTab
150         Select Case tb!SampleStatus & ""
              Case "T": s = s & "Transfused"
160           Case "A": s = s & "Aborted"
170           Case "S": s = s & "Spiked"
180           Case "U": s = s & "Unknown"
190           Case "D": s = s & "Destroyed"
200           Case Else: s = s & "???"
210         End Select
220         s = s & vbTab & tb!UserName & ""
          
          
230       Case "SU3"
240         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!UnitNumber & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                tb!UnitExpiry & vbTab & _
                ReadableGroup(tb!UnitGroup & "") & vbTab & _
                tb!StockComment & ""
        
        
250       Case "EIQ", "SVQ"
260         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "")
          
          
270       Case "EIR"
280         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "") & vbTab & _
                ReadableGroup(tb!PatientGroup & "") & vbTab & _
                ReadableStatus(tb!RERStatus & "") & vbTab & _
                tb!RERExpiry & ""
        
290       Case "ST"
300         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!UnitNumber & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                tb!UserName & vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "")
        
310       Case "SVR"
320         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "") & vbTab & _
                ReadableGroup(tb!PatientGroup & "") & vbTab & _
                ReadableStatus(tb!RERStatus & "") & vbTab & _
                tb!RERExpiry & vbTab & _
                ReadableValidity(tb!SampleStatus & "") & vbTab & _
                tb!SampleExpiry & ""
        
330     End Select
        
340     grd.AddItem s
350     tb.MoveNext

360   Loop

370   If grd.Rows > 2 Then
380     grd.RemoveItem 1
390   End If

400   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "frmViewCourier", "FillGrid", intEL, strES, sql


End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub dtFrom_CloseUp()

      Dim n As Integer

10    For n = 0 To 9
20      If optMessage(n).Value = True Then
30        FillGrid optMessage(n).Tag
40      End If
50    Next

End Sub


Private Sub dtTo_CloseUp()

      Dim n As Integer

10    For n = 0 To 9
20      If optMessage(n).Value = True Then
30        FillGrid optMessage(n).Tag
40      End If
50    Next

End Sub


Private Sub Form_Load()

10    dtFrom = Format$(Now, "dd/mm/yyyy")
20    dtTo = dtFrom

End Sub

Private Sub Form_Resize()

10    If Me.Width < 10170 Then
20      Me.Width = 10170
30    End If
40    If Me.Height < 2325 Then
50      Me.Height = 2325
60    End If

70    grd.Width = Me.Width - 370
80    grd.Height = Me.Height - 1890

End Sub


Private Sub grd_Click()

10    If grd.MouseRow = 0 Then
20      If InStr(grd.TextMatrix(0, grd.Col), "Date") <> 0 Then
30        grd.Sort = 9
40      Else
50        If SortOrder Then
60          grd.Sort = flexSortGenericAscending
70        Else
80          grd.Sort = flexSortGenericDescending
90        End If
100     End If
110     SortOrder = Not SortOrder
120   End If

End Sub

Private Sub grd_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(grd.TextMatrix(Row1, grd.Col)) Then
20      Cmp = 0
30      Exit Sub
40    End If

50    If Not IsDate(grd.TextMatrix(Row2, grd.Col)) Then
60      Cmp = 0
70      Exit Sub
80    End If

90    d1 = Format(grd.TextMatrix(Row1, grd.Col), "dd/mmm/yyyy hh:mm:ss")
100   d2 = Format(grd.TextMatrix(Row2, grd.Col), "dd/mmm/yyyy hh:mm:ss")

110   If SortOrder Then
120     Cmp = Sgn(DateDiff("s", d1, d2))
130   Else
140     Cmp = Sgn(DateDiff("s", d2, d1))
150   End If

End Sub


Private Sub optMessage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10    SetGridFormat optMessage(Index).Tag
20    FillGrid optMessage(Index).Tag

End Sub

