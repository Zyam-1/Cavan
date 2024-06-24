VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExternalBatch 
   Caption         =   "NetAcquire - External Batches"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtEarliest 
      Height          =   285
      Left            =   12150
      TabIndex        =   11
      Top             =   4620
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   218824705
      CurrentDate     =   39192
   End
   Begin VB.TextBox txtSorting 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5100
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmExternalBatch.frx":0000
      Top             =   810
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   795
      Left            =   12180
      Picture         =   "frmExternalBatch.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   795
      Left            =   12180
      Picture         =   "frmExternalBatch.frx":068A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7410
      Width           =   1485
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert Result"
      Height          =   795
      Left            =   12180
      Picture         =   "frmExternalBatch.frx":0ACC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox txtResult 
      Height          =   285
      Left            =   12150
      TabIndex        =   4
      Text            =   "Received"
      Top             =   1980
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   12180
      Picture         =   "frmExternalBatch.frx":0F0E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8370
      Width           =   1485
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   795
      Left            =   12180
      Picture         =   "frmExternalBatch.frx":1578
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   9045
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15954
      _Version        =   393216
      Cols            =   11
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"frmExternalBatch.frx":1882
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Earliest Sample Date"
      Height          =   375
      Left            =   12300
      TabIndex        =   10
      Top             =   4230
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Result"
      Height          =   195
      Left            =   12180
      TabIndex        =   6
      Top             =   1770
      Width           =   450
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   12180
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frmExternalBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SortOrder As Boolean

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

32510     On Error GoTo FillG_Error

32520     Screen.MousePointer = vbHourglass
32530     g.Visible = False

32540     g.Rows = 2
32550     g.AddItem ""
32560     g.RemoveItem 1

32570     sql = "SELECT D.SampleID, D.SampleDate, D.PatName, D.Chart, D.Ward, D.Clinician, D.DoB, D.GP, " & _
              "E.Analyte, E.Result, E.SendTo FROM " & _
              "Demographics AS D, ExtResults AS E WHERE " & _
              "(E.Result IS NULL OR E.Result like '') " & _
              "AND D.SampleID = E.SampleID " & _
              "AND D.SampleDate > '" & Format$(dtEarliest, "Long Date") & "' " & _
              "AND Analyte <> 'TFT' " & _
              "AND Analyte <> 'PSA' " & _
              "AND Analyte <> 'B12 + Folate' " & _
              "ORDER BY D.SampleID"
32580     Set tb = New Recordset
32590     RecOpenClient 0, tb, sql
32600     Do While Not tb.EOF
32610         s = CStr(tb!SampleID) & vbTab & _
                  Format$(tb!SampleDate, "dd/MM/yyyy") & vbTab & _
                  tb!PatName & vbTab
32620         If IsDate(tb!DoB & "") Then
32630             s = s & Format$(tb!DoB, "dd/MM/yyyy")
32640         End If
32650         s = s & vbTab & tb!Chart & vbTab & _
                  tb!Analyte & vbTab & _
                  vbTab & _
                  tb!SendTo & vbTab & _
                  tb!Ward & vbTab & _
                  tb!Clinician & vbTab & tb!GP & ""
32660         g.AddItem s
32670         tb.MoveNext
32680     Loop

32690     If g.Rows > 2 Then
32700         g.RemoveItem 1
32710     End If
32720     g.row = 0
32730     g.RowSel = 0
32740     g.Visible = True
32750     Screen.MousePointer = vbNormal

32760     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

32770     intEL = Erl
32780     strES = Err.Description
32790     LogError "frmExternalBatch", "FillG", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

32800     If cmdSave.Visible Then
32810         If iMsg("Cancel without Saving?", vbYesNo + vbQuestion) = vbNo Then
32820             Exit Sub
32830         End If
32840     End If

32850     Unload Me

End Sub

Private Sub cmdInsert_Click()

          Dim intStart As Integer
          Dim intEnd As Integer
          Dim n As Integer

32860     intStart = g.row
32870     intEnd = g.RowSel

32880     If intStart = 0 Or intEnd = 0 Then
32890         Exit Sub
32900     End If

32910     If intStart > intEnd Then
32920         n = intStart
32930         intStart = intEnd
32940         intEnd = n
32950     End If

32960     For n = intStart To intEnd
32970         g.TextMatrix(n, 6) = txtResult
32980     Next

32990     cmdSave.Visible = True

End Sub

Private Sub cmdRefresh_Click()

33000     FillG

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim n As Integer
          Dim sql As String

33010     On Error GoTo cmdSave_Click_Error

33020     For n = 1 To g.Rows - 1
33030         If Trim(g.TextMatrix(n, 6)) <> "" Then 'result
33040             sql = "Select * from ExtResults where " & _
                      "sampleid = '" & g.TextMatrix(n, 0) & "' " & _
                      "and Analyte = '" & g.TextMatrix(n, 5) & "'"
33050             Set tb = New Recordset
33060             RecOpenServer 0, tb, sql
33070             If tb.EOF Then
33080                 tb.AddNew
33090             End If
33100             tb!SampleID = g.TextMatrix(n, 0)
33110             tb!Analyte = g.TextMatrix(n, 5)
33120             tb!Result = g.TextMatrix(n, 6)
33130             tb!Date = Format(Now, "dd/mmm/yyyy")
33140             tb.Update
33150         End If
33160     Next

33170     cmdSave.Visible = False

33180     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

33190     intEL = Erl
33200     strES = Err.Description
33210     LogError "frmExternalBatch", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub cmdXL_Click()

33220     ExportFlexGrid g, Me

End Sub

Private Sub dtEarliest_CloseUp()

33230     FillG

End Sub


Private Sub Form_Activate()

33240     FillG

End Sub

Private Sub Form_Load()

          'FillG

33250     dtEarliest = Format$(Now - 42, "Short Date")

End Sub


Private Sub g_Click()

33260     g.Col = g.MouseCol

33270     If g.MouseRow = 0 Then
        
33280         txtSorting.Visible = True
33290         txtSorting.Refresh
        
33300         If g.Col = 1 Or g.Col = 3 Then
33310             g.Sort = 9
33320         Else
33330             If SortOrder Then
33340                 g.Sort = flexSortGenericAscending
33350             Else
33360                 g.Sort = flexSortGenericDescending
33370             End If
33380         End If
33390         SortOrder = Not SortOrder
        
33400     End If
        
33410     txtSorting.Visible = False

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

33420     If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
33430         Cmp = 0
33440         Exit Sub
33450     End If

33460     If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
33470         Cmp = 0
33480         Exit Sub
33490     End If

33500     d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
33510     d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

33520     If SortOrder Then
33530         Cmp = Sgn(DateDiff("s", d1, d2))
33540     Else
33550         Cmp = Sgn(DateDiff("s", d2, d1))
33560     End If


End Sub


