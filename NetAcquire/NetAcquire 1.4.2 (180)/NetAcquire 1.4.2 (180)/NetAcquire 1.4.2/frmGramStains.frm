VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGramStains 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Gram Stains / Wet Prep"
   ClientHeight    =   5340
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1100
      Left            =   8940
      Picture         =   "frmGramStains.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   1100
      Left            =   8940
      Picture         =   "frmGramStains.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   1100
      Left            =   8940
      Picture         =   "frmGramStains.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1200
   End
   Begin VB.ComboBox cmbQuantity 
      Height          =   315
      Left            =   8760
      TabIndex        =   3
      Text            =   "cmbQuantity"
      Top             =   240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox cmbWetPrep 
      Height          =   315
      Left            =   8790
      TabIndex        =   2
      Text            =   "cmbWetPrep"
      Top             =   660
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox cmbGramStains 
      Height          =   315
      Left            =   7320
      TabIndex        =   1
      Text            =   "cmbGramStains"
      Top             =   240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   180
      TabIndex        =   0
      Top             =   5055
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3675
      Left            =   180
      TabIndex        =   4
      Top             =   1260
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   6482
      _Version        =   393216
      Rows            =   13
      Cols            =   4
      RowHeightMin    =   275
   End
   Begin VB.Label lblSampleID 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1260
      TabIndex        =   10
      Top             =   870
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   900
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gram Stains / Wet Prep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2433
      TabIndex        =   5
      Top             =   180
      Width           =   4095
   End
   Begin VB.Menu mnuLists 
      Caption         =   "&Lists"
      Begin VB.Menu mnuGramStains 
         Caption         =   "&Gram Stains"
      End
      Begin VB.Menu mnuWetPrep 
         Caption         =   "&Wet Prep"
      End
      Begin VB.Menu mnuQuantity 
         Caption         =   "&Quantity"
      End
   End
End
Attribute VB_Name = "frmGramStains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ChangesMade As Boolean
Public SampleID As String



Private Sub InitGrid()

61510     On Error GoTo InitGrid_Error

61520     With Grid
61530         .Clear
61540         .Rows = 13
61550         .Cols = 4
61560         .FixedRows = 1
61570         .FixedCols = 0
          
61580         .TextMatrix(0, 0) = "Gram Stains": .ColWidth(0) = "3000": .ColAlignment(0) = flexAlignLeftCenter
61590         .TextMatrix(0, 1) = "Quantity": .ColWidth(1) = "1200": .ColAlignment(1) = flexAlignCenterCenter
61600         .TextMatrix(0, 2) = "Wet Prep": .ColWidth(2) = "3000": .ColAlignment(2) = flexAlignLeftCenter
61610         .TextMatrix(0, 3) = "Quantity": .ColWidth(3) = "1200": .ColAlignment(3) = flexAlignCenterCenter
          
61620         .row = 0: .Col = 0: .CellAlignment = flexAlignCenterCenter
61630         .row = 0: .Col = 1: .CellAlignment = flexAlignCenterCenter
61640         .row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
61650         .row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
61660     End With

61670     Exit Sub

InitGrid_Error:

          Dim strES As String
          Dim intEL As Integer

61680     intEL = Erl
61690     strES = Err.Description
61700     LogError "frmGramStains", "InitGrid", intEL, strES

End Sub


Private Sub FillGrid()

          Dim IDs As New IdentResults
          Dim i As Integer

61710     On Error GoTo FillGrid_Error

61720     IDs.Load SampleID
61730     If Not IDs Is Nothing Then
61740         For i = 1 To IDs.Count
61750             If UCase$(IDs(i).TestType) = "GRAMSTAIN" Then
61760                 Grid.TextMatrix(i, 0) = IDs(i).TestName
61770                 Grid.TextMatrix(i, 1) = IDs(i).Result
61780             End If
61790         Next

61800         For i = 1 To IDs.Count
61810             If UCase$(IDs(i).TestType) = "WETPREP" Then
61820                 Grid.TextMatrix(i, 2) = IDs(i).TestName
61830                 Grid.TextMatrix(i, 3) = IDs(i).Result
61840             End If
61850         Next
61860     End If

61870     Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

61880     intEL = Erl
61890     strES = Err.Description
61900     LogError "frmGramStains", "FillGrid", intEL, strES

End Sub

Private Function DataIsValid() As Boolean
          Dim i As Integer

61910     DataIsValid = True
61920     With Grid
61930         For i = 1 To .Rows - 1
61940             If (.TextMatrix(i, 0) <> "" And .TextMatrix(i, 1) = "") Or (.TextMatrix(i, 0) = "" And .TextMatrix(i, 1) <> "") Then
61950                 DataIsValid = False
61960                 If .TextMatrix(i, 0) = "" Then
61970                     .row = i: .Col = 0
61980                     .CellBackColor = vbRed
61990                 ElseIf .TextMatrix(i, 1) = "" Then
62000                     .row = i: .Col = 1
62010                     .CellBackColor = vbRed
62020                 End If
62030             End If
62040             If (.TextMatrix(i, 2) <> "" And .TextMatrix(i, 3) = "") Or (.TextMatrix(i, 2) = "" And .TextMatrix(i, 3) <> "") Then
62050                 DataIsValid = False
62060                 If .TextMatrix(i, 2) = "" Then
62070                     .row = i: .Col = 2
62080                     .CellBackColor = vbRed
62090                 ElseIf .TextMatrix(i, 3) = "" Then
62100                     .row = i: .Col = 3
62110                     .CellBackColor = vbRed
62120                 End If
62130             End If
62140         Next i
62150     End With

End Function


Private Function ItemExists(Item As String) As Boolean
          Dim i As Integer
62160     On Error GoTo ItemExists_Error

62170     ItemExists = False
62180     For i = 1 To Grid.Rows - 1
62190         If Grid.TextMatrix(i, Grid.Col) = Item And Grid.TextMatrix(i, Grid.Col) <> "" And Grid.row <> i Then
62200             ItemExists = True
62210             Exit For
62220         End If
62230     Next i

62240     Exit Function

ItemExists_Error:

          Dim strES As String
          Dim intEL As Integer

62250     intEL = Erl
62260     strES = Err.Description
62270     LogError "frmGramStains", "ItemExists", intEL, strES

End Function


Private Sub cmbGramStains_Click()
62280     If ItemExists(cmbGramStains.Text) Then
62290         iMsg "Test already exists in the list", vbInformation
62300         Exit Sub
62310     End If
62320     Grid.TextMatrix(Grid.row, Grid.Col) = cmbGramStains.Text
62330     cmbGramStains.Visible = False
62340     cmdSave.Enabled = True

End Sub

Private Sub cmbGramStains_KeyPress(KeyAscii As Integer)

62350     KeyAscii = 0

End Sub


Private Sub cmbGramStains_LostFocus()

62360     cmbGramStains.Visible = False
End Sub

Private Sub cmbQuantity_Click()
62370     Grid.TextMatrix(Grid.row, Grid.Col) = cmbQuantity.Text
62380     cmbQuantity.Visible = False
62390     cmdSave.Enabled = True
End Sub

Private Sub cmbQuantity_KeyPress(KeyAscii As Integer)

62400     KeyAscii = 0

End Sub


Private Sub cmbQuantity_LostFocus()
62410     cmbQuantity.Visible = False
End Sub

Private Sub cmbWetPrep_Click()
62420     Grid.TextMatrix(Grid.row, Grid.Col) = cmbWetPrep.Text
62430     cmbWetPrep.Visible = False
62440     cmdSave.Enabled = True
End Sub

Private Sub cmbWetPrep_KeyPress(KeyAscii As Integer)

62450     KeyAscii = 0

End Sub


Private Sub cmbWetPrep_LostFocus()
62460     cmbWetPrep.Visible = False
End Sub

Private Sub cmdExit_Click()
62470     Unload Me
End Sub

Private Sub cmdReset_Click()
62480     If ChangesMade Then
62490         If iMsg("Changes will not be saved. Continue?", vbQuestion + vbYesNo) = vbNo Then
62500             Exit Sub
62510         End If
62520     End If
62530     InitGrid
62540     ChangesMade = False
End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim i As Integer

62550     On Error GoTo cmdSave_Click_Error

62560     If Not DataIsValid Then
62570         iMsg "Input values are not correct", vbInformation
62580         cmdSave.Enabled = False
62590         Exit Sub
62600     End If
        
62610     sql = "DELETE FROM Identification " & _
              "WHERE SampleID = '" & lblSampleID & "' "
62620     Cnxn(0).Execute sql

          Dim ID As New IdentResult
62630     With ID
62640         For i = 1 To Grid.Rows - 1
                  '        If Grid.TextMatrix(i, 0) <> "" Then
                  'SAVE GRAM STAINS
62650             .SampleID = lblSampleID
62660             .TestType = "GramStain"
62670             .TestName = Grid.TextMatrix(i, 0)
62680             .Result = Grid.TextMatrix(i, 1)
62690             .Valid = 1
62700             .Printed = 0
62710             .UserName = UserName
62720             .TestDateTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
62730             .Save
                  '       End If
                  'SAVE WET PREP
62740             .SampleID = lblSampleID
62750             .TestType = "WetPrep"
62760             .TestName = Grid.TextMatrix(i, 2)
62770             .Result = Grid.TextMatrix(i, 3)
62780             .Valid = 1
62790             .Printed = 0
62800             .UserName = UserName
62810             .TestDateTime = Format(Now, "dd/MMM/yyyy hh:m:ss")
62820             .Save
62830         Next i
62840     End With
62850     ChangesMade = False
62860     cmdSave.Enabled = False

62870     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

62880     intEL = Erl
62890     strES = Err.Description
62900     LogError "frmGramStains", "cmdSave_Click", intEL, strES, sql

End Sub




Private Sub Form_Load()

62910     lblSampleID = SampleID
62920     InitGrid
62930     FillGrid
62940     FillGenericList cmbGramStains, "FG", True
62950     FillGenericList cmbWetPrep, "WP", True
62960     FillGenericList cmbQuantity, "GQ", True
62970     cmdSave.Enabled = False
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

62980     If cmdSave.Enabled Then
62990         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
63000             Cancel = True
63010         End If
63020     End If

End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
63030     cmbGramStains.Visible = False
63040     cmbWetPrep.Visible = False
63050     cmbQuantity.Visible = False
63060     Grid.row = Grid.row
63070     Grid.Col = Grid.Col
63080     Grid.CellBackColor = vbWhite
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim Combo As ComboBox
          Dim ComboLeft As Integer
          Dim ComboTop As Integer
          Dim i As Integer

63090     On Error GoTo Grid_MouseUp_Error

63100     Select Case Grid.Col
              Case 0: Set Combo = cmbGramStains
63110         Case 1: Set Combo = cmbQuantity
63120         Case 2: Set Combo = cmbWetPrep
63130         Case 3: Set Combo = cmbQuantity
63140     End Select

63150     ComboTop = Grid.Top + Grid.row * 275
63160     ComboLeft = Grid.Left + 50

63170     For i = 0 To Grid.Col - 1
63180         ComboLeft = ComboLeft + Grid.ColWidth(i)
63190     Next i

63200     If Not Combo Is Nothing Then
63210         Combo.Move ComboLeft, ComboTop, Grid.ColWidth(Grid.Col)
63220         If Grid.TextMatrix(Grid.row, Grid.Col) <> "" Then
63230             Combo.Text = Grid.TextMatrix(Grid.row, Grid.Col)
63240         End If
63250         Combo.Visible = True
63260     End If

63270     ChangesMade = True

63280     Exit Sub

Grid_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

63290     intEL = Erl
63300     strES = Err.Description
63310     LogError "frmGramStains", "Grid_MouseUp", intEL, strES


End Sub

Private Sub mnuGramStains_Click()

63320     With frmListsGeneric
63330         .ListType = "FG"
63340         .ListTypeName = "Gram Stain"
63350         .ListTypeNames = "Gram Stains"
63360         .Show 1
63370     End With

63380     FillGenericList cmbGramStains, "FG", True

End Sub


Private Sub mnuQuantity_Click()

63390     With frmListsGeneric
63400         .ListType = "GQ"
63410         .ListTypeName = "Quantity"
63420         .ListTypeNames = "Quantities"
63430         .Show 1
63440     End With

63450     FillGenericList cmbQuantity, "GQ", True

End Sub


Private Sub mnuWetPrep_Click()

63460     With frmListsGeneric
63470         .ListType = "WP"
63480         .ListTypeName = "Wet Prep"
63490         .ListTypeNames = "Wet Preps"
63500         .Show 1
63510     End With

63520     FillGenericList cmbWetPrep, "WP", True

End Sub


