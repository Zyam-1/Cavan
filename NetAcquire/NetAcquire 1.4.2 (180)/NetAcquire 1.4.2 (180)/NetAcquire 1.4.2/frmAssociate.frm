VERSION 5.00
Begin VB.Form frmAssociate 
   Caption         =   "NetAcquire - Associated Sample ID's"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&save"
      Height          =   525
      Left            =   2310
      Picture         =   "frmAssociate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4230
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   240
      TabIndex        =   4
      Top             =   1530
      Width           =   3285
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   525
         Left            =   2070
         Picture         =   "frmAssociate.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtAdd 
         Height          =   285
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   525
      Left            =   2310
      Picture         =   "frmAssociate.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox lstAss 
      Height          =   2205
      IntegralHeight  =   0   'False
      Left            =   450
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000018&
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
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   210
      TabIndex        =   9
      Top             =   330
      Width           =   3315
   End
   Begin VB.Label lblChart 
      BackColor       =   &H80000018&
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
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   210
      TabIndex        =   8
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label lblSID 
      BackColor       =   &H80000018&
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
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmAssociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetAss()

          Dim sql As String
          Dim tb As Recordset

56970     On Error GoTo GetAss_Error

56980     lstAss.Clear

56990     sql = "SELECT AssID FROM AssociatedIDs WHERE " & _
              "SampleID = '" & Val(lblSID) & "' " & _
              "OR AssID = '" & Val(lblSID) & "'"
57000     Set tb = New Recordset
57010     RecOpenServer 0, tb, sql
57020     Do While Not tb.EOF
57030         lstAss.AddItem tb!AssID ' - sysOptMicroOffset(0)
57040         tb.MoveNext
57050     Loop
        
57060     If lstAss.ListCount = 0 Then 'AssID unknown - look in previous for possibility
57070         sql = "SELECT SampleID FROM Demographics WHERE " & _
                  "SampleID = '" & Val(lblSID) - 1 & "' " & _
                  "AND Chart = '" & lblChart & "' " & _
                  "AND PatName = '" & AddTicks(lblName) & "'"
57080         Set tb = New Recordset
57090         RecOpenServer 0, tb, sql
57100         If tb.EOF Then 'not in previous
57110             txtAdd = ""
57120         Else
57130             txtAdd = CStr(Val(tb!SampleID))
57140         End If
57150     Else 'AssID already known
57160         GetAssRecurse
57170     End If

57180     Exit Sub

GetAss_Error:

          Dim strES As String
          Dim intEL As Integer

57190     intEL = Erl
57200     strES = Err.Description
57210     LogError "frmAssociate", "GetAss", intEL, strES, sql


End Sub

Private Sub GetAssRecurse()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim Y As Integer
          Dim FoundAss As Boolean
          Dim FoundSID As Boolean

57220     On Error GoTo GetAssRecurse_Error

57230     For n = lstAss.ListCount - 1 To 0 Step -1

57240         sql = "SELECT SampleID, AssID FROM AssociatedIDs WHERE " & _
                  "SampleID = '" & Val(lstAss.List(n)) & "' OR AssID  = '" & Val(lstAss.List(n)) & "'"
57250         Set tb = New Recordset
57260         RecOpenServer 0, tb, sql
57270         Do While Not tb.EOF
57280             FoundAss = False
57290             FoundSID = False
57300             For Y = lstAss.ListCount - 1 To 0 Step -1
57310                 If lstAss.List(Y) = tb!SampleID Then FoundSID = True
57320                 If lstAss.List(Y) = tb!AssID Then FoundAss = True
57330                 If FoundAss And FoundSID Then Exit For
57340             Next
57350             If Not FoundAss Then lstAss.AddItem tb!AssID ' - sysOptMicroOffset(0)
57360             If Not FoundSID Then lstAss.AddItem tb!SampleID ' - sysOptMicroOffset(0)
57370             If Not FoundAss Or Not FoundSID Then GetAssRecurse
57380             tb.MoveNext
57390         Loop

57400     Next

57410     Exit Sub

GetAssRecurse_Error:

          Dim strES As String
          Dim intEL As Integer

57420     intEL = Erl
57430     strES = Err.Description
57440     LogError "frmAssociate", "GetAssRecurse", intEL, strES, sql


End Sub

Private Sub cmdAdd_Click()

          Dim tb As Recordset
          Dim sql As String

57450     On Error GoTo cmdAdd_Click_Error

57460     If Trim$(txtAdd) = "" Then Exit Sub

57470     sql = "SELECT SampleID, AssID FROM AssociatedIDs WHERE " & _
              "( SampleID = '" & Val(lblSID) & "' " & _
              "  AND AssID = '" & Val(txtAdd) & "' ) " & _
              "OR " & _
              "( SampleID = '" & Val(txtAdd) & "' " & _
              "  AND AssID = '" & Val(lblSID) & "' ) "
57480     Set tb = New Recordset
57490     RecOpenServer 0, tb, sql
57500     If tb.EOF Then
57510         tb.AddNew
57520         tb!SampleID = Val(lblSID) ' + sysOptMicroOffset(0)
57530         tb!AssID = Val(txtAdd) ' + sysOptMicroOffset(0)
57540         tb.Update
57550     End If

57560     lstAss.AddItem txtAdd
57570     txtAdd = ""
57580     GetAssRecurse

57590     Exit Sub

cmdAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

57600     intEL = Erl
57610     strES = Err.Description
57620     LogError "frmAssociate", "cmdadd_Click", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

57630     Unload Me

End Sub

Private Sub cmdRemove_Click()

          Dim n As Integer
          Dim sql As String

57640     On Error GoTo cmdRemove_Click_Error

57650     If lstAss.SelCount = 0 Then Exit Sub

57660     For n = lstAss.ListCount - 1 To 0 Step -1
57670         If lstAss.Selected(n) Then
          
57680             sql = "DELETE FROM AssociatedIDs WHERE " & _
                      "( SampleID = '" & Val(lstAss.List(n)) & "' " & _
                      "  AND AssID = '" & Val(lblSID) & "' ) " & _
                      "OR " & _
                      "( AssID = '" & Val(lstAss.List(n)) & "' " & _
                      "  AND SampleID = '" & Val(lblSID) & "' ) "
57690             Cnxn(0).Execute sql
           
57700             lstAss.RemoveItem n
        
57710         End If
57720     Next

57730     Exit Sub

cmdRemove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

57740     intEL = Erl
57750     strES = Err.Description
57760     LogError "frmAssociate", "cmdRemove_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

57770     GetAss

End Sub

