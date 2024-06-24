VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fLists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Listings"
   ClientHeight    =   8505
   ClientLeft      =   195
   ClientTop       =   375
   ClientWidth     =   8520
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8505
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   7200
      Picture         =   "flists.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8100
      Top             =   5400
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8100
      Top             =   6180
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7290
      Picture         =   "flists.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3870
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   1275
      Left            =   4710
      TabIndex        =   12
      Top             =   150
      Width           =   2205
      Begin VB.OptionButton o 
         Caption         =   "Specimen Sources"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   19
         Top             =   990
         Width           =   1725
      End
      Begin VB.OptionButton o 
         Caption         =   "Errors"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   540
         Width           =   735
      End
      Begin VB.OptionButton o 
         Caption         =   "Units"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   765
      End
      Begin VB.OptionButton o 
         Caption         =   "Sample Types"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   750
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   7290
      Picture         =   "flists.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7500
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7290
      Picture         =   "flists.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6030
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7290
      Picture         =   "flists.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Clinician"
      Height          =   1275
      Left            =   180
      TabIndex        =   6
      Top             =   150
      Width           =   4365
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   330
         Width           =   645
      End
      Begin VB.TextBox tText 
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
         Left            =   660
         MaxLength       =   50
         TabIndex        =   1
         Top             =   660
         Width           =   3495
      End
      Begin VB.TextBox tCode 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   0
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   300
         TabIndex        =   8
         Top             =   690
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   390
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6765
      Left            =   180
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11933
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Code   |Text                                                                                      "
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
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   7290
      Picture         =   "flists.frx":163A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2430
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   7290
      Picture         =   "flists.frx":1CA4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7020
      TabIndex        =   18
      Top             =   990
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "fLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer

Private Sub FireUp()

Dim n As Integer
Dim s As String
Dim X As Integer

If g.Row = 1 Then Exit Sub

FireCounter = FireCounter + 1
If FireCounter > 5 Then
  tmrUp.Interval = 100
End If

n = g.Row

g.Visible = False

s = ""
For X = 0 To g.Cols - 1
  s = s & g.TextMatrix(n, X) & vbTab
Next
s = Left$(s, Len(s) - 1)

g.RemoveItem n
g.AddItem s, n - 1

g.Row = n - 1
For X = 0 To g.Cols - 1
  g.Col = X
  g.CellBackColor = vbYellow
Next

If Not g.RowIsVisible(g.Row) Then
  g.TopRow = g.Row
End If

g.Visible = True

cmdSave.Visible = True

End Sub

Private Sub FireDown()

Dim n As Integer
Dim s As String
Dim X As Integer
Dim VisibleRows As Integer

If g.Row = g.Rows - 1 Then Exit Sub
n = g.Row

FireCounter = FireCounter + 1
If FireCounter > 5 Then
  tmrDown.Interval = 100
End If

VisibleRows = g.Height \ g.RowHeight(1) - 1

g.Visible = False

s = ""
For X = 0 To g.Cols - 1
  s = s & g.TextMatrix(n, X) & vbTab
Next
s = Left$(s, Len(s) - 1)

g.RemoveItem n
If n < g.Rows Then
  g.AddItem s, n + 1
  g.Row = n + 1
Else
  g.AddItem s
  g.Row = g.Rows - 1
End If

For X = 0 To g.Cols - 1
  g.Col = X
  g.CellBackColor = vbYellow
Next

If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
  If g.Row - VisibleRows + 1 > 0 Then
    g.TopRow = g.Row - VisibleRows + 1
  End If
End If

g.Visible = True

cmdSave.Visible = True

End Sub


Private Sub bAdd_Click()

tCode = Trim$(UCase$(tCode))
tText = Trim$(tText)

If tCode = "" Then
  Exit Sub
End If

If tText = "" Then Exit Sub

g.AddItem tCode & vbTab & tText

tCode = ""
tText = ""

 cmdSave.Visible = True

End Sub

Private Sub cmdcancel_Click()

Unload Me

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FireDown

tmrDown.Interval = 250
FireCounter = 0

tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FireUp

tmrUp.Interval = 250
FireCounter = 0

tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

Dim LT As String

LT = Switch(o(0), "Units", _
            o(1), "Errors", _
            o(2), "Sample Types.", _
            o(3), "Specimen Sources")
            
Printer.Print

Printer.Print "List of "; LT

g.Col = 0
g.Row = 1
g.ColSel = g.Cols - 1
g.RowSel = g.Rows - 1

Printer.Print g.Clip

Printer.EndDoc
Screen.MousePointer = 0

End Sub

Private Sub cmdsave_Click()

Dim LT As String
Dim Y As Integer
Dim tb As Recordset
Dim sql As String

LT = Switch(o(0), "UN", _
            o(1), "ER", _
            o(2), "ST", _
            o(3), "MB")
            
For Y = 1 To g.Rows - 1
  If g.TextMatrix(Y, 0) <> "" Then
    sql = "Select * from Lists where " & _
          "ListType = '" & LT & "' " & _
          "and Code = '" & g.TextMatrix(Y, 0) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If tb.EOF Then
      tb.AddNew
    End If
    tb!Code = g.TextMatrix(Y, 0)
    tb!ListType = LT
    tb!Text = g.TextMatrix(Y, 1)
    tb!ListOrder = Y
    tb!InUse = 1
    tb.Update
  End If
Next

FillG

tCode = ""
tText = ""
tCode.SetFocus
cmdMoveUp.Enabled = False
cmdMoveDown.Enabled = False
cmdSave.Visible = False
cmdDelete.Enabled = False

End Sub

Private Sub FillG()

Dim LT As String
Dim s As String
Dim tb As Recordset
Dim sql As String

LT = Switch(o(0), "UN", _
            o(1), "ER", _
            o(2), "ST", _
            o(3), "MB")
            
g.Rows = 2
g.AddItem ""
g.RemoveItem 1

sql = "Select * from Lists where " & _
      "ListType = '" & LT & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  s = tb!Code & vbTab & tb!Text & ""
  g.AddItem s
  tb.MoveNext
Loop

If g.Rows > 2 Then
  g.RemoveItem 1
End If

End Sub




Private Sub cmdDelete_Click()

Dim strLTCode As String
Dim strLTText As String
Dim Y As Integer
Dim sql As String
Dim s As String

strLTCode = Switch(o(0), "UN", _
                   o(1), "ER", _
                   o(2), "ST", _
                   o(3), "MB")
strLTText = Switch(o(0), "Units", _
                   o(1), "Errors", _
                   o(2), "Sample Types", _
                   o(3), "Specimen Sources")
            
g.Col = 0
For Y = 1 To g.Rows - 1
  g.Row = Y
  If g.CellBackColor = vbYellow Then
    s = "Delete " & g.TextMatrix(Y, 1) & vbCrLf & _
        "From " & strLTText & " ?"
    If iMsg(s, vbQuestion + vbYesNo) = vbYes Then
      sql = "Delete from Lists where " & _
            "ListType = '" & strLTCode & "' " & _
            "and Code = '" & g.TextMatrix(Y, 0) & "'"
      Cnxn(0).Execute sql
    End If
    Exit For
  End If
Next

cmdDelete.Enabled = False
FillG

End Sub

Private Sub cmdXL_Click()

ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()

If Activated Then Exit Sub

Activated = True

FillG

End Sub

Private Sub Form_Load()

g.Font.Bold = True

Activated = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdSave.Visible Then
  If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
    Cancel = True
    Exit Sub
  End If
End If

End Sub


Private Sub g_Click()

Static SortOrder As Boolean
Dim X As Integer
Dim Y As Integer
Dim ySave As Integer

ySave = g.Row

g.Visible = False
g.Col = 0
For Y = 1 To g.Rows - 1
  g.Row = Y
  If g.CellBackColor = vbYellow Then
    For X = 0 To g.Cols - 1
      g.Col = X
      g.CellBackColor = 0
    Next
    Exit For
  End If
Next
g.Row = ySave
g.Visible = True

If g.MouseRow = 0 Then
  If SortOrder Then
    g.Sort = flexSortGenericAscending
  Else
    g.Sort = flexSortGenericDescending
  End If
  SortOrder = Not SortOrder
  Exit Sub
End If

For X = 0 To g.Cols - 1
  g.Col = X
  g.CellBackColor = vbYellow
Next

cmdMoveUp.Enabled = True
cmdMoveDown.Enabled = True
cmdDelete.Enabled = True

End Sub



Private Sub o_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

FillG

FrameAdd.Caption = "Add New " & Left$(o(Index).Caption, Len(o(Index).Caption) - 1)

tCode = ""
tText = ""
If tCode.Visible Then
  tCode.SetFocus
End If

End Sub

Private Sub tmrDown_Timer()

FireDown

End Sub


Private Sub tmrUp_Timer()

FireUp

End Sub


