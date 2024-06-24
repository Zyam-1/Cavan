VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ABDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AntibioticName As String
Private m_ListOrder As Integer
Private m_AllowIfPregnant As Boolean
Private m_AllowIfOutPatient As Boolean
Private m_AllowIfChild As Boolean
Private m_Code As String
Private m_AllowIfPenicillinAllergy As Boolean
Private m_OrganismGroup As String
Private m_Site As String
Private m_PriSec As String
Private m_ViewInGrid As Boolean
Public Property Get AntibioticName() As String

10        AntibioticName = m_AntibioticName

End Property

Public Property Get Code() As String

20        Code = m_Code

End Property


Public Function Load(ByVal ABName As String) As ABDefinition

          Dim tb As Recordset
          Dim sql As String
          Dim AB As New ABDefinition

30        On Error GoTo Load_Error

40        sql = "SELECT Code, AntibioticName, ListOrder, " & _
              "COALESCE(AllowIfPregnant, 0) AllowIfPregnant, " & _
              "COALESCE(AllowIfOutPatient, 0) AllowIfOutPatient, " & _
              "COALESCE(AllowIfChild, 0) AllowIfChild, " & _
              "COALESCE(AllowIfPenAll, 0) AllowIfPenAll, " & _
              "COALESCE(ViewInGrid, 0) ViewInGrid " & _
              "FROM Antibiotics " & _
              "WHERE AntibioticName = '" & ABName & "'"
50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If Not tb.EOF Then
80            With AB
90                .Code = tb!Code & ""
100               .AntibioticName = tb!AntibioticName & ""
110               .OrganismGroup = ""
120               .Site = ""
130               .ListOrder = 0
140               .PriSec = ""
150               .AllowIfPregnant = tb!AllowIfPregnant
160               .AllowIfOutPatient = tb!AllowIfOutPatient
170               .AllowIfChild = tb!AllowIfChild
180               .AllowIfPenicillinAllergy = tb!AllowIfPenAll
190               .ViewInGrid = tb!ViewInGrid
200           End With
210           Set Load = AB
220       Else
230           Set Load = Nothing
240       End If

250       Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "ABDefinition", "Load", intEL, strES, sql

End Function

Public Property Get Site() As String

290       Site = m_Site

End Property


Public Property Get ListOrder() As Integer

300       ListOrder = m_ListOrder

End Property


Public Property Get AllowIfPregnant() As Boolean

310       AllowIfPregnant = m_AllowIfPregnant

End Property

Public Property Get AllowIfOutPatient() As Boolean

320       AllowIfOutPatient = m_AllowIfOutPatient

End Property


Public Property Get AllowIfChild() As Boolean

330       AllowIfChild = m_AllowIfChild

End Property



Public Property Get ViewInGrid() As Boolean

340       ViewInGrid = m_ViewInGrid

End Property


Public Property Get AllowIfPenicillinAllergy() As Boolean

350       AllowIfPenicillinAllergy = m_AllowIfPenicillinAllergy

End Property

Public Property Get PriSec() As String

360       PriSec = m_PriSec

End Property

Public Property Get OrganismGroup() As String

370       OrganismGroup = m_OrganismGroup

End Property


Public Property Let Site(ByVal sNewValue As String)

380       m_Site = Trim$(sNewValue)

End Property
Public Property Let PriSec(ByVal sNewValue As String)

390       m_PriSec = sNewValue

End Property

Public Property Let Code(ByVal sNewValue As String)

400       m_Code = sNewValue

End Property


Public Property Let OrganismGroup(ByVal sNewValue As String)

410       m_OrganismGroup = Trim$(sNewValue)

End Property

Public Property Let AntibioticName(ByVal sNewValue As String)

420       m_AntibioticName = Trim$(sNewValue)

End Property


Public Property Let ListOrder(ByVal iNewValue As Integer)

430       m_ListOrder = iNewValue

End Property


Public Property Let AllowIfPregnant(ByVal iNewValue As Boolean)

440       m_AllowIfPregnant = iNewValue

End Property

Public Property Let AllowIfPenicillinAllergy(ByVal iNewValue As Boolean)

450       m_AllowIfPenicillinAllergy = iNewValue

End Property


Public Property Let AllowIfOutPatient(ByVal iNewValue As Boolean)

460       m_AllowIfOutPatient = iNewValue

End Property


Public Property Let AllowIfChild(ByVal iNewValue As Boolean)

470       m_AllowIfChild = iNewValue

End Property



Public Property Let ViewInGrid(ByVal iNewValue As Boolean)

480       m_ViewInGrid = iNewValue

End Property




