VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAges 
   Caption         =   "NetAcquire - Ages"
   ClientHeight    =   5820
   ClientLeft      =   2895
   ClientTop       =   750
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   4815
   Begin VB.ComboBox cmbInt 
      Height          =   315
      Left            =   5460
      TabIndex        =   12
      Text            =   "cmbInt"
      Top             =   1350
      Width           =   915
   End
   Begin VB.ComboBox cmbDMY 
      Height          =   315
      ItemData        =   "frmAges.frx":0000
      Left            =   6360
      List            =   "frmAges.frx":0002
      TabIndex        =   11
      Text            =   "cmbDMY"
      Top             =   1350
      Width           =   1035
   End
   Begin VB.ComboBox cM 
      Height          =   315
      Left            =   3060
      TabIndex        =   10
      Text            =   "cM"
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.ComboBox cY 
      Height          =   315
      Left            =   2310
      TabIndex        =   9
      Text            =   "cY"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cD 
      Height          =   315
      Left            =   3660
      TabIndex        =   8
      Text            =   "cD"
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton bRemove 
      Caption         =   "&Remove Age Range"
      Height          =   735
      Left            =   3000
      Picture         =   "frmAges.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1605
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   825
      Left            =   3000
      Picture         =   "frmAges.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4710
      Width           =   1245
   End
   Begin VB.CommandButton bAdd 
      Caption         =   "&Add Age Range"
      Height          =   735
      Left            =   330
      Picture         =   "frmAges.frx":0AB0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1605
   End
   Begin VB.CommandButton bSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   735
      Left            =   2040
      Picture         =   "frmAges.frx":0EF2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   645
      Picture         =   "frmAges.frx":155C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4710
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2775
      Left            =   330
      TabIndex        =   0
      Top             =   570
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   315
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Age From (YMD)       |<Age To (YMD)           "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lSampleType 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serum"
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
      Left            =   330
      TabIndex        =   7
      Top             =   120
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      Height          =   75
      Left            =   270
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label lParameter 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creatinine"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "frmAges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAnalyte As String
Private mSampleType As String
Private mDiscipline As String

Private Activated As Boolean

Private FromDays() As Long
Private ToDays() As Long
Private Sub AdjustG()

          Dim Y As Integer

46930     For Y = 0 To UBound(FromDays)
46940         g.TextMatrix(Y + 1, 0) = dmyFromCount(FromDays(Y))
46950         g.TextMatrix(Y + 1, 1) = dmyFromCount(ToDays(Y))
46960     Next

End Sub

Private Sub FillCoagAges()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim s As String

46970     On Error GoTo FillCoagAges_Error

46980     g.Rows = 2
46990     g.AddItem ""
47000     g.RemoveItem 1

47010     sql = "Select AgeFromDays, AgeToDays from CoagTestDefinitions where " & _
              "TestName = '" & mAnalyte & "' " & _
              "Order by AgeFromDays"
47020     Set tb = New Recordset
47030     RecOpenClient 0, tb, sql

47040     ReDim FromDays(0 To tb.RecordCount - 1)
47050     ReDim ToDays(0 To tb.RecordCount - 1)
47060     n = 0
47070     Do While Not tb.EOF
47080         FromDays(n) = tb!AgeFromDays
47090         ToDays(n) = tb!AgeToDays
47100         s = dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
47110         g.AddItem s
47120         n = n + 1
47130         tb.MoveNext
47140     Loop

47150     If g.Rows > 2 Then
47160         g.RemoveItem 1
47170     End If

47180     bRemove.Enabled = g.Rows > 2

47190     Exit Sub

FillCoagAges_Error:

          Dim strES As String
          Dim intEL As Integer

47200     intEL = Erl
47210     strES = Err.Description
47220     LogError "frmAges", "FillCoagAges", intEL, strES, sql


End Sub


Private Sub FillDMY(ByVal Days As Long, _
          ByRef Y As Long, _
          ByRef m As Long, _
          ByRef D As Long)

47230     Y = Days \ 365

47240     Days = Days - (Y * 365)

47250     m = Days \ 30.42

47260     D = Days - (m * 30.42)

End Sub

Private Sub bAdd_Click()

47270     Select Case mDiscipline
              Case "Haematology": AddHaem
47280         Case "Coagulation": AddCoag
47290     End Select

End Sub

Private Sub AddCoag()

          Dim tb As Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field

47300     On Error GoTo AddCoag_Error

47310     sql = "Select top 1 * from CoagTestDefinitions where " & _
              "TestName = '" & mAnalyte & "' " & _
              "order by AgeToDays desc"
47320     Set tb = New Recordset
47330     RecOpenClient 0, tb, sql

47340     Set tbNew = New Recordset
47350     RecOpenClient 0, tbNew, sql

47360     tbNew.AddNew
47370     For Each fld In tb.Fields
47380         If fld.Name <> "pkid" Then
47390             If fld.Name = "AgeToDays" Then
47400                 tbNew!AgeToDays = tb!AgeToDays + 1
47410             ElseIf fld.Name = "AgeFromDays" Then
47420                 tbNew!AgeFromDays = tb!AgeToDays + 1
47430             Else
47440                 tbNew(fld.Name) = tb(fld.Name)
47450             End If
47460         End If
47470     Next
47480     tbNew.Update

47490     FillCoagAges

47500     Exit Sub

AddCoag_Error:

          Dim strES As String
          Dim intEL As Integer

47510     intEL = Erl
47520     strES = Err.Description
47530     LogError "fAges", "AddCoag", intEL, strES, sql


End Sub


Private Sub AddHaem()

          Dim tb As Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field
          Dim tb2 As Recordset

47540     On Error GoTo AddHaem_Error

47550     sql = "Select top 1 * from HaemTestDefinitions where " & _
              "AnalyteName = '" & mAnalyte & "' " & _
              "order by AgeToDays desc"
            
47560     MsgBox (sql)
47570     Set tb = New Recordset
47580     RecOpenClient 0, tb, sql

47590     Set tbNew = New Recordset
47600     RecOpenClient 0, tbNew, sql

47610     tbNew.AddNew
47620     For Each fld In tb.Fields
              '        MsgBox ("In Loop")
47630         If UCase$(fld.Name) <> "COUNTER" Then
                  '             MsgBox UCase$(fld.Name) & "__First"
47640             If UCase$(fld.Name) = "AGETODAYS" Then
                      '            MsgBox UCase$(fld.Name) & "___2"
47650                 tbNew!AgeToDays = tb!AgeToDays + 1
47660             ElseIf UCase$(fld.Name) = "AGEFROMDAYS" Then
                      '            MsgBox UCase$(fld.Name) & "___elseif"
47670                 tbNew!AgeFromDays = tb!AgeToDays + 1
47680             Else
                      '            MsgBox UCase$(fld.Name) & "___else"
47690                 tbNew(fld.Name) = tb(fld.Name)
          
47700             End If
47710         End If
47720     Next
47730     sql = "SELECT Top 1 Counter FROM HaemTestDefinitions WHERE AnalyteName = '" & mAnalyte & "' ORDER BY AgeToDays Desc"
47740     Set tb2 = New Recordset
47750     RecOpenClient 0, tb2, sql
          '      MsgBox ConvertNull(tb2!Counter, 0)
47760     tbNew!Counter = ConvertNull(tb2!Counter, 0)
47770     tbNew.Update
          'MsgBox "Update"

47780     FillHaemAges

47790     Exit Sub

AddHaem_Error:

          Dim strES As String
          Dim intEL As Integer

47800     intEL = Erl
47810     strES = Err.Description
47820     MsgBox (strES)
47830     LogError "frmAges", "AddHaem", intEL, strES, sql


End Sub

Private Sub FillHaemAges()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim s As String

47840     On Error GoTo FillHaemAges_Error

47850     g.Rows = 2
47860     g.AddItem ""
47870     g.RemoveItem 1
          'MsgBox "fillhaem"
47880     sql = "Select AgeFromDays, AgeToDays from HaemTestDefinitions where " & _
              "AnalyteName = '" & mAnalyte & "' " & _
              "Order by AgeFromDays"
47890     Set tb = New Recordset
47900     RecOpenClient 0, tb, sql

47910     ReDim FromDays(0 To tb.RecordCount - 1)
47920     ReDim ToDays(0 To tb.RecordCount - 1)
47930     n = 0
47940     Do While Not tb.EOF
              'MsgBox "loop while"
47950         FromDays(n) = tb!AgeFromDays
47960         ToDays(n) = tb!AgeToDays
47970         s = dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
              '            MsgBox s
47980         g.AddItem s
47990         n = n + 1
48000         tb.MoveNext
48010     Loop

48020     If g.Rows > 2 Then
48030         g.RemoveItem 1
48040     End If

48050     bRemove.Enabled = g.Rows > 2

48060     Exit Sub

FillHaemAges_Error:

          Dim strES As String
          Dim intEL As Integer

48070     intEL = Erl
48080     strES = Err.Description
48090     LogError "frmAges", "FillHaemAges", intEL, strES, sql


End Sub


Private Sub bcancel_Click()

48100     Unload Me

End Sub

Private Sub bRemove_Click()

48110     Select Case mDiscipline
              Case "Haematology": RemoveHaem
48120         Case "Coagulation": RemoveCoag
48130     End Select

End Sub


Private Sub RemoveCoag()

          Dim Y As Integer
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String

48140     On Error GoTo RemoveCoag_Error

48150     g.Col = 0
48160     For Y = 1 To g.Rows - 2
48170         g.row = Y
48180         If g.CellBackColor = vbYellow Then
48190             Exit For
48200         End If
48210     Next
48220     Y = Y - 1

48230     If Y = 0 Then Exit Sub

48240     rFrom = FromDays(Y)
48250     rTo = ToDays(Y)

48260     sql = "Delete from CoagTestDefinitions where " & _
              "AgeFromDays = '" & rFrom & "' " & _
              "and AgeToDays = '" & rTo & "'"
48270     Cnxn(0).Execute sql

48280     sql = "Update CoagTestDefinitions " & _
              "Set AgeToDays = '" & rTo + 1 & "' " & _
              "where AgeToDays = '" & rFrom - 1 & "'"
48290     Cnxn(0).Execute sql

48300     FillCoagAges

48310     Exit Sub

RemoveCoag_Error:

          Dim strES As String
          Dim intEL As Integer

48320     intEL = Erl
48330     strES = Err.Description
48340     LogError "frmAges", "RemoveCoag", intEL, strES, sql


End Sub


Private Sub RemoveHaem()

          Dim Y As Integer
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String

48350     On Error GoTo RemoveHaem_Error

48360     g.Col = 0
48370     For Y = 1 To g.Rows - 2
48380         g.row = Y
48390         If g.CellBackColor = vbYellow Then
48400             Exit For
48410         End If
48420     Next
48430     Y = Y - 1

48440     If Y = 0 Then Exit Sub

48450     rFrom = FromDays(Y)
48460     rTo = ToDays(Y)

48470     sql = "Delete from HaemTestDefinitions where " & _
              "AgeFromDays = '" & rFrom & "' " & _
              "and AgeToDays = '" & rTo & "'"
48480     Cnxn(0).Execute sql

48490     sql = "Update HaemTestDefinitions " & _
              "Set AgeToDays = '" & rTo + 1 & "' " & _
              "where AgeToDays = '" & rFrom - 1 & "'"
48500     Cnxn(0).Execute sql

48510     FillHaemAges

48520     Exit Sub

RemoveHaem_Error:

          Dim strES As String
          Dim intEL As Integer

48530     intEL = Erl
48540     strES = Err.Description
48550     LogError "frmAges", "RemoveHaem", intEL, strES, sql


End Sub


Private Sub bSave_Click()

48560     Select Case mDiscipline
              Case "Haematology": SaveHaem
48570         Case "Coagulation": SaveCoag
48580     End Select

End Sub

Private Sub SaveHaem()

          Dim sql As String
          Dim Days As Long
          Dim n As Integer
48590     On Error GoTo SaveHaem_Error

48600     ReDim WasFrom(0 To UBound(FromDays)) As Long
48610     ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Integer

48620     g.Col = 0
48630     For n = 1 To g.Rows - 2
48640         g.row = n
48650         If g.CellBackColor = vbYellow Then
48660             Active = n - 1
48670             Exit For
48680         End If
48690     Next

48700     Days = (Val(cY) * 365.25) + (Val(cM) * 30.42) + Val(cD)
48710     If Days = 0 Then Exit Sub

48720     For n = 0 To UBound(FromDays)
48730         WasFrom(n) = FromDays(n)
48740         WasTo(n) = ToDays(n)
48750     Next

48760     ToDays(Active) = Days

48770     For n = 0 To UBound(FromDays) - 1
48780         FromDays(n + 1) = ToDays(n) + 1
48790         If ToDays(n + 1) < FromDays(n + 1) Then
48800             ToDays(n + 1) = FromDays(n + 1)
48810         End If
48820     Next

48830     For n = 0 To UBound(WasFrom)
48840         sql = "Update HaemTestDefinitions " & _
                  "Set AgeFromDays = '" & FromDays(n) & "', " & _
                  "AgeToDays = '" & ToDays(n) & "' where " & _
                  "AgeFromDays = '" & WasFrom(n) & "' " & _
                  "and AgeToDays = '" & WasTo(n) & "' " & _
                  "and AnalyteName = '" & lParameter & "'"
48850         Cnxn(0).Execute sql
48860     Next

48870     AdjustG

48880     cY.Visible = False
48890     cM.Visible = False
48900     cD.Visible = False
48910     bsave.Visible = False

48920     Exit Sub

SaveHaem_Error:

          Dim strES As String
          Dim intEL As Integer

48930     intEL = Erl
48940     strES = Err.Description
48950     LogError "frmAges", "SaveHaem", intEL, strES, sql


End Sub


Private Sub SaveCoag()

          Dim sql As String
          Dim Days As Long
          Dim n As Integer
48960     On Error GoTo SaveCoag_Error

48970     ReDim WasFrom(0 To UBound(FromDays)) As Long
48980     ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Integer

48990     g.Col = 0
49000     For n = 1 To g.Rows - 2
49010         g.row = n
49020         If g.CellBackColor = vbYellow Then
49030             Active = n - 1
49040             Exit For
49050         End If
49060     Next

49070     Days = (Val(cY) * 365.25) + (Val(cM) * 30.42) + Val(cD)
49080     If Days = 0 Then Exit Sub

49090     For n = 0 To UBound(FromDays)
49100         WasFrom(n) = FromDays(n)
49110         WasTo(n) = ToDays(n)
49120     Next

49130     ToDays(Active) = Days

49140     For n = 0 To UBound(FromDays) - 1
49150         FromDays(n + 1) = ToDays(n) + 1
49160         If ToDays(n + 1) < FromDays(n + 1) Then
49170             ToDays(n + 1) = FromDays(n + 1)
49180         End If
49190     Next

49200     For n = 0 To UBound(WasFrom)
49210         sql = "Update CoagTestDefinitions " & _
                  "Set AgeFromDays = '" & FromDays(n) & "', " & _
                  "AgeToDays = '" & ToDays(n) & "' where " & _
                  "AgeFromDays = '" & WasFrom(n) & "' " & _
                  "and AgeToDays = '" & WasTo(n) & "' " & _
                  "and TestName = '" & lParameter & "'"
49220         Cnxn(0).Execute sql
49230     Next

49240     AdjustG

49250     cY.Visible = False
49260     cM.Visible = False
49270     cD.Visible = False
49280     bsave.Visible = False

49290     Exit Sub

SaveCoag_Error:

          Dim strES As String
          Dim intEL As Integer

49300     intEL = Erl
49310     strES = Err.Description
49320     LogError "frmAges", "SaveCoag", intEL, strES, sql


End Sub

Private Sub cD_Click()

49330     bsave.Visible = True

End Sub

Private Sub cD_KeyPress(KeyAscii As Integer)

49340     KeyAscii = 0

End Sub


Private Sub cM_Click()

49350     bsave.Visible = True

End Sub

Private Sub cM_KeyPress(KeyAscii As Integer)

49360     KeyAscii = 0

End Sub


Private Sub cmbDMY_Click()

          Dim n As Integer
          Dim intTop As Integer

49370     Select Case cmbDMY
              Case "Days": intTop = 30
49380         Case "Months": intTop = 11
49390         Case "Years": intTop = 120
49400     End Select

49410     cmbInt.Clear

49420     For n = 1 To intTop
49430         cmbInt.AddItem Format$(n)
49440     Next
49450     cmbInt = "1"

End Sub


Private Sub cmbDMY_KeyPress(KeyAscii As Integer)

49460     KeyAscii = 0

End Sub

Private Sub cY_Click()

49470     If Val(cY) > 10 Then
49480         cM = "0"
49490         cD = "0"
49500     End If
        
49510     bsave.Visible = True

End Sub

Private Sub cY_KeyPress(KeyAscii As Integer)

49520     KeyAscii = 0

End Sub


Private Sub Form_Activate()

49530     If Activated Then Exit Sub

49540     Select Case mDiscipline
              Case "Haematology": FillHaemAges
49550         Case "Coagulation": FillCoagAges
49560     End Select

49570     AdjustG

49580     g.Col = 0
49590     g.row = 1
49600     g.CellBackColor = vbYellow

49610     Activated = True

End Sub

Private Sub Form_Load()

          Dim n As Integer

49620     lParameter = mAnalyte

49630     Select Case mSampleType
              Case "S": lSampleType = "Serum"
49640         Case "U": lSampleType = "Urine"
49650         Case Else: lSampleType = mSampleType
49660     End Select

49670     Activated = False

49680     cY.Clear
49690     cM.Clear
49700     cD.Clear

49710     For n = 0 To 120
49720         cY.AddItem Format$(n)
49730     Next
49740     For n = 0 To 11
49750         cM.AddItem Format$(n)
49760     Next
49770     For n = 0 To 30
49780         cD.AddItem Format$(n)
49790     Next

49800     cmbDMY.Clear
49810     cmbDMY.AddItem "Days"
49820     cmbDMY.AddItem "Months"
49830     cmbDMY.AddItem "Years"
49840     cmbDMY = "Days"

49850     cmbInt.Clear
49860     For n = 1 To 30
49870         cmbInt.AddItem Format$(n)
49880     Next

End Sub


Private Sub Form_Unload(Cancel As Integer)

49890     Activated = False

End Sub


Private Sub g_Click()

          Dim Y As Integer
          Dim OrigY As Integer

          Dim Days As Long
          Dim Months As Long
          Dim Years As Long

49900     If g.MouseRow = 0 Then Exit Sub
49910     If g.MouseRow = g.Rows - 1 Then Exit Sub

49920     OrigY = g.row

49930     g.Col = 0
49940     For Y = 1 To g.Rows - 1
49950         g.row = Y
49960         g.CellBackColor = 0
49970     Next

49980     g.row = OrigY
49990     g.CellBackColor = vbYellow

50000     FillDMY ToDays(g.row - 1), Years, Months, Days
50010     cY = Format$(Years)
50020     cM = Format$(Months)
50030     cD = Format$(Days)

50040     cY.Top = g.Top + 50 + (g.row * 315)
50050     cM.Top = g.Top + 50 + (g.row * 315)
50060     cD.Top = g.Top + 50 + (g.row * 315)

50070     cY.Visible = True
50080     cM.Visible = True
50090     cD.Visible = True
50100     bsave.Visible = True

End Sub


Public Property Let Analyte(ByVal Analyte As String)

50110     mAnalyte = Analyte

End Property
Public Property Let Discipline(ByVal Discipline As String)

50120     mDiscipline = Discipline

End Property

Public Property Let SampleType(ByVal SampleType As String)

50130     mSampleType = SampleType

End Property

