VERSION 5.00
Begin VB.Form frmAntibioticLists 
   Caption         =   "NetAcquire"
   ClientHeight    =   7755
   ClientLeft      =   345
   ClientTop       =   570
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   7740
      Picture         =   "frmAntibioticLists.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6720
      Width           =   765
   End
   Begin VB.CommandButton cmdOrganisms 
      Caption         =   "Organisms"
      Height          =   465
      Left            =   5400
      TabIndex        =   16
      Top             =   180
      Width           =   1725
   End
   Begin VB.CommandButton cmdNewSite 
      Caption         =   "New Site"
      Height          =   465
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "New Site"
      Top             =   690
      Width           =   1725
   End
   Begin VB.CommandButton cmdNewOrganismGroup 
      Caption         =   "New Organism Group"
      Height          =   465
      Left            =   3600
      TabIndex        =   14
      ToolTipText     =   "New Organism"
      Top             =   180
      Width           =   1725
   End
   Begin VB.CommandButton cmdNewAntibiotic 
      Caption         =   "New Antibiotic"
      Height          =   465
      Left            =   5400
      TabIndex        =   13
      Top             =   690
      Width           =   1725
   End
   Begin VB.CommandButton cmdRemoveFromSecondary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Width           =   765
   End
   Begin VB.CommandButton cmdTransferToSecondary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   765
   End
   Begin VB.CommandButton cmdRemoveFromPrimary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2850
      Width           =   765
   End
   Begin VB.Frame Frame4 
      Caption         =   "Secondary List"
      Height          =   2655
      Left            =   3990
      TabIndex        =   8
      Top             =   4860
      Width           =   3525
      Begin VB.CommandButton cmdMoveUpSec 
         Caption         =   "Move Up"
         Height          =   825
         Left            =   2550
         Picture         =   "frmAntibioticLists.frx":1330
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   660
         Width           =   795
      End
      Begin VB.CommandButton cmdMoveDownSec 
         Caption         =   "Move Down"
         Height          =   825
         Left            =   2550
         Picture         =   "frmAntibioticLists.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1500
         Width           =   795
      End
      Begin VB.ListBox lstSecondary 
         Height          =   2205
         Left            =   240
         TabIndex        =   9
         Top             =   270
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdTransferToPrimary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":1BB4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2250
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Caption         =   "Available Antibiotics"
      Height          =   5715
      Left            =   210
      TabIndex        =   5
      Top             =   1770
      Width           =   2715
      Begin VB.ListBox lstAvailable 
         Height          =   5130
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   2385
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Primary List"
      Height          =   2595
      Left            =   3990
      TabIndex        =   3
      Top             =   1770
      Width           =   3525
      Begin VB.CommandButton cmdMoveUpPri 
         Caption         =   "Move Up"
         Height          =   825
         Left            =   2520
         Picture         =   "frmAntibioticLists.frx":1FF6
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   795
      End
      Begin VB.CommandButton cmdMoveDownPri 
         Caption         =   "Move Down"
         Height          =   825
         Left            =   2520
         Picture         =   "frmAntibioticLists.frx":2438
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1440
         Width           =   795
      End
      Begin VB.ListBox lstPrimary 
         Height          =   2205
         Left            =   210
         TabIndex        =   4
         Top             =   270
         Width           =   2295
      End
   End
   Begin VB.Frame fraOrg 
      Height          =   1425
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   2715
      Begin VB.ComboBox cmbSite 
         Height          =   315
         Left            =   90
         TabIndex        =   18
         Text            =   "cmbSite"
         Top             =   990
         Width           =   2415
      End
      Begin VB.ComboBox cmbOrganismGroup 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Site"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   810
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Organism Group"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmAntibioticLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearLists()

50140     lstPrimary.Clear
50150     lstSecondary.Clear

End Sub

Private Sub FillAvailable()

          Dim sql As String
          Dim tb As Recordset

50160     On Error GoTo FillAvailable_Error

50170     lstAvailable.Clear

50180     sql = "Select * from Antibiotics " & _
              "order by ListOrder asc"
50190     Set tb = New Recordset
50200     RecOpenServer 0, tb, sql

50210     Do While Not tb.EOF
50220         lstAvailable.AddItem Trim$(tb!AntibioticName & "")
50230         tb.MoveNext
50240     Loop

50250     Exit Sub

FillAvailable_Error:

          Dim strES As String
          Dim intEL As Integer

50260     intEL = Erl
50270     strES = Err.Description
50280     LogError "frmAntibioticLists", "FillAvailable", intEL, strES, sql


End Sub

Private Sub FillLists()

          Dim n As Integer
          Dim strAB As String
          Dim Ax As ABDefinition
          Dim Axs As New ABDefinitions

50290     On Error GoTo FillLists_Error

50300     lstPrimary.Clear
50310     lstSecondary.Clear

50320     Axs.Load cmbSite, cmbOrganismGroup
50330     For Each Ax In Axs
        
50340         strAB = Ax.AntibioticName
50350         Select Case Ax.PriSec
                  Case "P": lstPrimary.AddItem strAB
50360             Case "S": lstSecondary.AddItem strAB
50370         End Select
          
50380         For n = 0 To lstAvailable.ListCount - 1
50390             If lstAvailable.List(n) = strAB Then
50400                 lstAvailable.RemoveItem n
50410                 Exit For
50420             End If
50430         Next
        
50440     Next

50450     Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

50460     intEL = Erl
50470     strES = Err.Description
50480     LogError "frmAntibioticLists", "FillLists", intEL, strES

End Sub

Private Sub FillOrganismGroups()

          Dim tb As Recordset
          Dim sql As String

50490     On Error GoTo FillOrganismGroups_Error

50500     cmbOrganismGroup.Clear
50510     sql = "Select * from Lists where " & _
              "ListType = 'OR' and InUse = 1 " & _
              "order by ListOrder"
50520     Set tb = New Recordset
50530     RecOpenServer 0, tb, sql
50540     Do While Not tb.EOF
50550         cmbOrganismGroup.AddItem tb!Text & ""
50560         tb.MoveNext
50570     Loop

50580     Exit Sub

FillOrganismGroups_Error:

          Dim strES As String
          Dim intEL As Integer

50590     intEL = Erl
50600     strES = Err.Description
50610     LogError "frmAntibioticLists", "FillOrganismGroups", intEL, strES, sql


End Sub

Private Sub FillSites()

          Dim tb As Recordset
          Dim sql As String

50620     On Error GoTo FillSites_Error

50630     sql = "Select * from Lists where " & _
              "ListType = 'SI' and InUse = 1 " & _
              "order by ListOrder"
50640     Set tb = New Recordset
50650     RecOpenServer 0, tb, sql
50660     Do While Not tb.EOF
50670         cmbSite.AddItem tb!Text & ""
50680         tb.MoveNext
50690     Loop
50700     cmbSite.AddItem "Generic"

50710     Exit Sub

FillSites_Error:

          Dim strES As String
          Dim intEL As Integer

50720     intEL = Erl
50730     strES = Err.Description
50740     LogError "frmAntibioticLists", "FillSites", intEL, strES, sql


End Sub

Private Sub cmbOrganismGroup_Click()

50750     ClearLists

50760     FillAvailable

50770     cmbSite = "Generic"

50780     FillLists

End Sub


Private Sub cmbSite_Click()

50790     ClearLists

50800     If cmbOrganismGroup = "" Then Exit Sub

50810     FillLists

End Sub


Private Sub cmdCancel_Click()

50820     Unload Me

End Sub

Private Sub cmdMoveDownPri_Click()

          Dim n As Integer
          Dim s As String

50830     If lstPrimary.ListIndex = -1 Then Exit Sub
50840     If lstPrimary.ListIndex = lstPrimary.ListCount - 1 Then Exit Sub

50850     n = lstPrimary.ListIndex

50860     s = lstPrimary

50870     lstPrimary.RemoveItem n
50880     If n < lstPrimary.ListCount Then
50890         lstPrimary.AddItem s, n + 1
50900         lstPrimary.Selected(n + 1) = True
50910     Else
50920         lstPrimary.AddItem s
50930         lstPrimary.Selected(lstPrimary.ListCount - 1) = True
50940     End If

50950     SaveListOrder

End Sub

Private Sub cmdMoveDownSec_Click()

          Dim n As Integer
          Dim s As String

50960     If lstSecondary.ListIndex = -1 Then Exit Sub
50970     If lstSecondary.ListIndex = lstSecondary.ListCount - 1 Then Exit Sub

50980     n = lstSecondary.ListIndex

50990     s = lstSecondary

51000     lstSecondary.RemoveItem n
51010     If n < lstSecondary.ListCount Then
51020         lstSecondary.AddItem s, n + 1
51030         lstSecondary.Selected(n + 1) = True
51040     Else
51050         lstSecondary.AddItem s
51060         lstSecondary.Selected(lstSecondary.ListCount - 1) = True
51070     End If

51080     SaveListOrder

End Sub


Private Sub cmdMoveUpPri_Click()

          Dim n As Integer
          Dim s As String

51090     If lstPrimary.ListIndex < 1 Then Exit Sub

51100     n = lstPrimary.ListIndex

51110     s = lstPrimary

51120     lstPrimary.RemoveItem n
51130     lstPrimary.AddItem s, n - 1

51140     lstPrimary.Selected(n - 1) = True

51150     SaveListOrder


End Sub

Private Sub cmdMoveUpSec_Click()

          Dim n As Integer
          Dim s As String

51160     If lstSecondary.ListIndex < 1 Then Exit Sub

51170     n = lstSecondary.ListIndex

51180     s = lstSecondary

51190     lstSecondary.RemoveItem n
51200     lstSecondary.AddItem s, n - 1

51210     lstSecondary.Selected(n - 1) = True

51220     SaveListOrder

End Sub


Private Sub cmdNewAntibiotic_Click()

51230     frmNewAntibiotics.Show 1
51240     FillAvailable

End Sub

Private Sub cmdNewOrganismGroup_Click()

51250     frmMicroLists.o(1) = True
51260     frmMicroLists.Show 1
51270     FillOrganismGroups

End Sub

Private Sub cmdNewSite_Click()

51280     frmMicroSites.Show 1
51290     FillOrganismGroups

End Sub


Private Sub cmdOrganisms_Click()

51300     frmOrganisms.Show 1

End Sub

Private Sub cmdRemoveFromPrimary_Click()

          Dim n As Integer
          Dim Ax As ABDefinition
          Dim Axs As New ABDefinitions

51310     For n = 0 To lstPrimary.ListCount - 1
51320         If lstPrimary.Selected(n) Then
51330             Set Ax = New ABDefinition
51340             Ax.AntibioticName = lstPrimary.List(n)
51350             Ax.ListOrder = 999
51360             Ax.OrganismGroup = cmbOrganismGroup
51370             Ax.PriSec = "P"
51380             Ax.Site = cmbSite
51390             Axs.Delete Ax
51400         End If
51410     Next

51420     FillAvailable
51430     FillLists

End Sub

Private Sub cmdRemoveFromSecondary_Click()

          Dim n As Integer
          Dim Ax As ABDefinition
          Dim Axs As New ABDefinitions

51440     For n = 0 To lstSecondary.ListCount - 1
51450         If lstSecondary.Selected(n) Then
51460             Set Ax = New ABDefinition
51470             Ax.AntibioticName = lstSecondary.List(n)
51480             Ax.ListOrder = 999
51490             Ax.OrganismGroup = cmbOrganismGroup
51500             Ax.PriSec = "P"
51510             Ax.Site = cmbSite
51520             Axs.Delete Ax
51530         End If
51540     Next

51550     FillAvailable
51560     FillLists

End Sub


Private Sub SaveListOrder()

          Dim n As Integer
          Dim Ax As ABDefinition
          Dim Axs As New ABDefinitions

51570     On Error GoTo SaveListOrder_Error

51580     For n = 0 To lstPrimary.ListCount - 1
51590         Set Ax = New ABDefinition
51600         Ax.AntibioticName = lstPrimary.List(n)
51610         Ax.OrganismGroup = cmbOrganismGroup
51620         Ax.Site = cmbSite
51630         Ax.ListOrder = n
51640         Ax.PriSec = "P"
51650         Axs.Save Ax
51660     Next

51670     For n = 0 To lstSecondary.ListCount - 1
51680         Set Ax = New ABDefinition
51690         Ax.AntibioticName = lstSecondary.List(n)
51700         Ax.OrganismGroup = cmbOrganismGroup
51710         Ax.Site = cmbSite
51720         Ax.ListOrder = n
51730         Ax.PriSec = "S"
51740         Axs.Save Ax
51750     Next

51760     Exit Sub

SaveListOrder_Error:

          Dim strES As String
          Dim intEL As Integer

51770     intEL = Erl
51780     strES = Err.Description
51790     LogError "frmAntibioticLists", "SaveListOrder", intEL, strES

End Sub

Private Sub cmdTransferToPrimary_Click()

          Dim n As Integer
          Dim Ax As ABDefinition
          Dim Axs As New ABDefinitions

51800     For n = 0 To lstAvailable.ListCount - 1
51810         If lstAvailable.Selected(n) Then
51820             Set Ax = New ABDefinition
51830             Ax.AntibioticName = lstAvailable.List(n)
51840             Ax.ListOrder = 999
51850             Ax.OrganismGroup = cmbOrganismGroup
51860             Ax.PriSec = "P"
51870             Ax.Site = cmbSite
51880             Axs.Save Ax
51890         End If
51900     Next

51910     FillAvailable
51920     FillLists

End Sub


Private Sub cmdTransferToSecondary_Click()

          Dim n As Integer
          Dim Ax As ABDefinition
          Dim Axs As New ABDefinitions

51930     For n = 0 To lstAvailable.ListCount - 1
51940         If lstAvailable.Selected(n) Then
51950             Set Ax = New ABDefinition
51960             Ax.AntibioticName = lstAvailable.List(n)
51970             Ax.ListOrder = 999
51980             Ax.OrganismGroup = cmbOrganismGroup
51990             Ax.PriSec = "S"
52000             Ax.Site = cmbSite
52010             Axs.Save Ax
52020         End If
52030     Next

52040     FillAvailable
52050     FillLists

End Sub


Private Sub Form_Load()

52060     FillAvailable

52070     FillOrganismGroups

52080     FillSites
52090     cmbSite = "Generic"

End Sub


