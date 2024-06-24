VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioControlDefinitions 
   Caption         =   "NetAcquire - Biochemistry QC"
   ClientHeight    =   4635
   ClientLeft      =   765
   ClientTop       =   405
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4980
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   3420
      Picture         =   "frmBioControlDefinitions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3330
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3420
      Picture         =   "frmBioControlDefinitions.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2010
      Width           =   1245
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   915
      Left            =   3900
      Picture         =   "frmBioControlDefinitions.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   300
      Width           =   765
   End
   Begin VB.TextBox txtAliasName 
      Height          =   285
      Left            =   1620
      TabIndex        =   4
      Top             =   900
      Width           =   2235
   End
   Begin VB.TextBox txtControlName 
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   300
      Width           =   2235
   End
   Begin MSFlexGridLib.MSFlexGrid grdControl 
      Height          =   2775
      Left            =   270
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Control Name    |<Alias Name       "
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Alias Name (From Analyser)"
      Height          =   465
      Left            =   450
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Control Name (Friendly Name)"
      Height          =   465
      Left            =   270
      TabIndex        =   1
      Top             =   210
      Width           =   1290
   End
End
Attribute VB_Name = "frmBioControlDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String

5710      On Error GoTo FillG_Error

5720      grdControl.Rows = 2
5730      grdControl.AddItem ""
5740      grdControl.RemoveItem 1

5750      sql = "Select distinct ControlName, AliasName " & _
              "from BioQCDefs"
5760      Set tb = New Recordset
5770      RecOpenServer 0, tb, sql
5780      Do While Not tb.EOF
5790          grdControl.AddItem tb!ControlName & vbTab & tb!AliasName & ""
5800          tb.MoveNext
5810      Loop

5820      If grdControl.Rows > 2 Then
5830          grdControl.RemoveItem 1
5840      End If

5850      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

5860      intEL = Erl
5870      strES = Err.Description
5880      LogError "frmBioControlDefinitions", "FillG", intEL, strES, sql


End Sub

Private Sub cmdAdd_Click()

          Dim n As Integer
          Dim Found As Boolean

5890      If txtControlName = "" Then
5900          iMsg "Enter Control Name"
5910          Exit Sub
5920      End If
5930      If txtAliasName = "" Then
5940          iMsg "Enter Alias Name"
5950          Exit Sub
5960      End If

5970      Found = False
5980      For n = 1 To grdControl.Rows - 1
5990          If txtControlName = grdControl.TextMatrix(n, 0) Or _
                  txtAliasName = grdControl.TextMatrix(n, 1) Then
           
6000              grdControl.TextMatrix(n, 0) = txtControlName
6010              grdControl.TextMatrix(n, 1) = txtAliasName
          
6020              Found = True
6030              Exit For
6040          End If
6050      Next
6060      If Not Found Then
6070          grdControl.AddItem txtControlName & vbTab & txtAliasName
6080      End If
6090      If grdControl.Rows = 3 And grdControl.TextMatrix(1, 0) = "" Then
6100          grdControl.RemoveItem 1
6110      End If

6120      txtControlName = ""
6130      txtAliasName = ""

6140      cmdSave.Enabled = True

End Sub

Private Sub cmdCancel_Click()

6150      Unload Me

End Sub


Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Y As Integer

6160      On Error GoTo cmdSave_Click_Error

6170      For Y = 1 To grdControl.Rows - 1
6180          If grdControl.TextMatrix(Y, 1) <> "" Then
6190              sql = "Select * from BioQCDefs where " & _
                      "AliasName = '" & grdControl.TextMatrix(Y, 1) & "'"
6200              Set tb = New Recordset
6210              RecOpenServer 0, tb, sql
6220              If tb.EOF Then tb.AddNew
6230              tb!ControlName = grdControl.TextMatrix(Y, 0)
6240              tb!AliasName = grdControl.TextMatrix(Y, 1)
6250              tb.Update
        
6260              sql = "Update BiochemistryQC " & _
                      "set ControlName = '" & grdControl.TextMatrix(Y, 0) & "' " & _
                      "where AliasName ='" & grdControl.TextMatrix(Y, 1) & "'"
6270          End If
6280      Next
6290      cmdSave.Enabled = False

6300      Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

6310      intEL = Erl
6320      strES = Err.Description
6330      LogError "frmBioControlDefinitions", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

6340      FillG

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

6350      If cmdSave.Enabled Then
6360          If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
6370              Cancel = True
6380          End If
6390      End If

End Sub


Private Sub grdControl_Click()

          Static SortOrder As Boolean

6400      If grdControl.MouseRow = 0 Then
6410          If SortOrder Then
6420              grdControl.Sort = flexSortGenericAscending
6430          Else
6440              grdControl.Sort = flexSortGenericDescending
6450          End If
6460          SortOrder = Not SortOrder
6470          Exit Sub
6480      End If

6490      txtControlName = grdControl.TextMatrix(grdControl.row, 0)
6500      txtAliasName = grdControl.TextMatrix(grdControl.row, 1)

End Sub


