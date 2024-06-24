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

10    AntibioticName = m_AntibioticName

End Property

Public Property Get Code() As String

10    Code = m_Code

End Property


Public Function Load(ByVal ABName As String) As ABDefinition

      Dim tb As Recordset
      Dim sql As String
      Dim AB As New ABDefinition

10    On Error GoTo Load_Error

20    sql = "SELECT Code, AntibioticName, ListOrder, " & _
            "COALESCE(AllowIfPregnant, 0) AllowIfPregnant, " & _
            "COALESCE(AllowIfOutPatient, 0) AllowIfOutPatient, " & _
            "COALESCE(AllowIfChild, 0) AllowIfChild, " & _
            "COALESCE(AllowIfPenAll, 0) AllowIfPenAll, " & _
            "COALESCE(ViewInGrid, 0) ViewInGrid " & _
            "FROM Antibiotics " & _
            "WHERE AntibioticName = '" & ABName & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60      With AB
70        .Code = tb!Code & ""
80        .AntibioticName = tb!AntibioticName & ""
90        .OrganismGroup = ""
100       .Site = ""
110       .ListOrder = 0
120       .PriSec = ""
130       .AllowIfPregnant = tb!AllowIfPregnant
140       .AllowIfOutPatient = tb!AllowIfOutPatient
150       .AllowIfChild = tb!AllowIfChild
160       .AllowIfPenicillinAllergy = tb!AllowIfPenAll
170       .ViewInGrid = tb!ViewInGrid
180     End With
190     Set Load = AB
200   Else
210     Set Load = Nothing
220   End If

230   Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "ABDefinition", "Load", intEL, strES, sql

End Function

Public Property Get Site() As String

10    Site = m_Site

End Property


Public Property Get ListOrder() As Integer

10    ListOrder = m_ListOrder

End Property


Public Property Get AllowIfPregnant() As Boolean

10    AllowIfPregnant = m_AllowIfPregnant

End Property

Public Property Get AllowIfOutPatient() As Boolean

10    AllowIfOutPatient = m_AllowIfOutPatient

End Property


Public Property Get AllowIfChild() As Boolean

10    AllowIfChild = m_AllowIfChild

End Property



Public Property Get ViewInGrid() As Boolean

10    ViewInGrid = m_ViewInGrid

End Property


Public Property Get AllowIfPenicillinAllergy() As Boolean

10    AllowIfPenicillinAllergy = m_AllowIfPenicillinAllergy

End Property

Public Property Get PriSec() As String

10    PriSec = m_PriSec

End Property

Public Property Get OrganismGroup() As String

10    OrganismGroup = m_OrganismGroup

End Property


Public Property Let Site(ByVal sNewValue As String)

10    m_Site = Trim$(sNewValue)

End Property
Public Property Let PriSec(ByVal sNewValue As String)

10    m_PriSec = sNewValue

End Property

Public Property Let Code(ByVal sNewValue As String)

10    m_Code = sNewValue

End Property


Public Property Let OrganismGroup(ByVal sNewValue As String)

10    m_OrganismGroup = Trim$(sNewValue)

End Property

Public Property Let AntibioticName(ByVal sNewValue As String)

10    m_AntibioticName = Trim$(sNewValue)

End Property


Public Property Let ListOrder(ByVal iNewValue As Integer)

10    m_ListOrder = iNewValue

End Property


Public Property Let AllowIfPregnant(ByVal iNewValue As Boolean)

10    m_AllowIfPregnant = iNewValue

End Property

Public Property Let AllowIfPenicillinAllergy(ByVal iNewValue As Boolean)

10    m_AllowIfPenicillinAllergy = iNewValue

End Property


Public Property Let AllowIfOutPatient(ByVal iNewValue As Boolean)

10    m_AllowIfOutPatient = iNewValue

End Property


Public Property Let AllowIfChild(ByVal iNewValue As Boolean)

10    m_AllowIfChild = iNewValue

End Property



Public Property Let ViewInGrid(ByVal iNewValue As Boolean)

10    m_ViewInGrid = iNewValue

End Property




