VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SampleID As String
Private m_Code As String
Private m_Result As String
Private m_Operator As String
Private m_Units As String
Private m_RunDate As String
Private m_RunTime As String
Private m_Flags As String
Private m_SampleType As String
Private m_Analyser As String
Private m_Low As Single
Private m_FlagLow As Single
Private m_PlausibleLow As Single
Private m_High As Single
Private m_FlagHigh As Single
Private m_PlausibleHigh As Single
Private m_Printformat As Integer
Private m_LongName As String
Private m_ShortName As String
Private m_DoDelta As Boolean
Private m_DeltaLimit As Single
Private m_DeltaDaysBackLimit As Integer
Private m_Category As String

Private m_ResultGroup As String

Private m_Valid As Boolean
Private m_Printed As Boolean

Private m_h As Boolean
Private m_s As Boolean
Private m_o As Boolean
Private m_l As Boolean
Private m_g As Boolean
Private m_J As Boolean

Private m_Authorised As Boolean

Private m_Comment As String

Private m_DefIndex As Long

Private m_PrintRefRange As Boolean

Public Sub Update(ByVal Discipline As String)
          'Discipline is either "Bio", "Imm" or "End"

          Dim sql As String
          Dim lngSampleID As Long

2810      lngSampleID = Format$(Val(m_SampleID))
2820      sql = "IF EXISTS (SELECT * FROM " & Discipline & "Results WHERE " & _
              "           SampleID = '" & lngSampleID & "' " & _
              "           AND Code = '" & m_Code & "') " & _
              "  UPDATE " & Discipline & "Results " & _
              "  SET Result = '" & m_Result & "' " & _
              "  WHERE SampleID = '" & lngSampleID & "' " & _
              "  AND Code = '" & m_Code & "' " & _
              "ELSE " & _
              "  INSERT INTO " & Discipline & "Results " & _
              "  ([sampleid], [Code], [Result], [Valid], [Printed], [RunTime], [RunDate], " & _
              "  [Operator], [Flags], [Units], [SampleType], [Analyser], [Authorised], " & _
              "  [Comment], [DefIndex]) "
2830      sql = sql & "  VALUES (" & _
              "  '" & m_SampleID & "', " & _
              "  '" & m_Code & "', " & _
              "  '" & m_Result & "', " & _
              "  '" & m_Valid & "', " & _
              "  '" & m_Printed & "', " & _
              "  '" & m_RunTime & "', " & _
              "  '" & m_RunDate & "', " & _
              "  '" & m_Operator & "', " & _
              "  '" & m_Flags & "', " & _
              "  '" & m_Units & "', " & _
              "  '" & m_SampleType & "', " & _
              "  '" & m_Analyser & "', " & _
              "  '" & m_Authorised & "', " & _
              "  '" & m_Comment & "', " & _
              "  '" & m_DefIndex & "') "
2840      Cnxn(0).Execute sql

2850      sql = "Delete from " & Discipline & "Requests where " & _
              "SampleID = '" & lngSampleID & "' " & _
              "and Code = '" & m_Code & "'"
2860      Cnxn(0).Execute sql
End Sub


Public Property Get SampleID() As String

2870      SampleID = m_SampleID

End Property

Public Property Let SampleID(ByVal SampleID As String)

2880      m_SampleID = SampleID

End Property

Public Property Get Code() As String

2890      Code = m_Code

End Property

Public Property Let Code(ByVal Code As String)

2900      m_Code = Code

End Property

Public Property Get Result() As String

2910      Result = m_Result

End Property

Public Property Let Result(ByVal Result As String)

2920      m_Result = Result

End Property

Public Property Get Operator() As String

2930      Operator = m_Operator

End Property

Public Property Let Operator(ByVal Operator As String)

2940      m_Operator = Operator

End Property

Public Property Get Units() As String

2950      Units = m_Units

End Property

Public Property Let Units(ByVal Units As String)

2960      m_Units = Units

End Property

Public Property Get Rundate() As String

2970      Rundate = m_RunDate

End Property

Public Property Let Rundate(ByVal Rundate As String)

2980      m_RunDate = Rundate

End Property

Public Property Get RunTime() As String

2990      RunTime = m_RunTime

End Property

Public Property Let RunTime(ByVal RunTime As String)

3000      m_RunTime = RunTime

End Property

Public Property Get Flags() As String

3010      Flags = m_Flags

End Property

Public Property Let Flags(ByVal Flags As String)

3020      m_Flags = Flags

End Property

Public Property Get SampleType() As String

3030      SampleType = m_SampleType

End Property

Public Property Let SampleType(ByVal SampleType As String)

3040      m_SampleType = SampleType

End Property

Public Property Get Analyser() As String

3050      Analyser = m_Analyser

End Property

Public Property Let Analyser(ByVal Analyser As String)

3060      m_Analyser = Analyser

End Property

Public Property Get Low() As Single

3070      Low = m_Low

End Property

Public Property Let Low(ByVal Low As Single)

3080      m_Low = Low

End Property

Public Property Get FlagLow() As Single

3090      FlagLow = m_FlagLow

End Property

Public Property Let FlagLow(ByVal FlagLow As Single)

3100      m_FlagLow = FlagLow

End Property

Public Property Get PlausibleLow() As Single

3110      PlausibleLow = m_PlausibleLow

End Property

Public Property Let PlausibleLow(ByVal PlausibleLow As Single)

3120      m_PlausibleLow = PlausibleLow

End Property

Public Property Get High() As Single

3130      High = m_High

End Property

Public Property Let High(ByVal High As Single)

3140      m_High = High

End Property

Public Property Get FlagHigh() As Single

3150      FlagHigh = m_FlagHigh

End Property

Public Property Let FlagHigh(ByVal FlagHigh As Single)

3160      m_FlagHigh = FlagHigh

End Property

Public Property Get PlausibleHigh() As Single

3170      PlausibleHigh = m_PlausibleHigh

End Property

Public Property Let PlausibleHigh(ByVal PlausibleHigh As Single)

3180      m_PlausibleHigh = PlausibleHigh

End Property

Public Property Get Printformat() As Integer

3190      Printformat = m_Printformat

End Property

Public Property Let Printformat(ByVal Printformat As Integer)

3200      m_Printformat = Printformat

End Property

Public Property Get LongName() As String

3210      LongName = m_LongName

End Property

Public Property Let LongName(ByVal LongName As String)

3220      m_LongName = LongName

End Property

Public Property Get ShortName() As String

3230      ShortName = m_ShortName

End Property

Public Property Let ShortName(ByVal ShortName As String)

3240      m_ShortName = ShortName

End Property

Public Property Get DoDelta() As Boolean

3250      DoDelta = m_DoDelta

End Property

Public Property Let DoDelta(ByVal DoDelta As Boolean)

3260      m_DoDelta = DoDelta

End Property

Public Property Get DeltaLimit() As Single

3270      DeltaLimit = m_DeltaLimit

End Property

Public Property Let DeltaLimit(ByVal DeltaLimit As Single)

3280      m_DeltaLimit = DeltaLimit

End Property



Public Property Get Category() As String

3290      Category = m_Category

End Property

Public Property Let Category(ByVal Category As String)

3300      m_Category = Category

End Property

Public Property Get ResultGroup() As String

3310      ResultGroup = m_ResultGroup

End Property

Public Property Let ResultGroup(ByVal ResultGroup As String)

3320      m_ResultGroup = ResultGroup

End Property

Public Property Get Valid() As Boolean

3330      Valid = m_Valid

End Property

Public Property Let Valid(ByVal Valid As Boolean)

3340      m_Valid = Valid

End Property

Public Property Get Printed() As Boolean

3350      Printed = m_Printed

End Property

Public Property Let Printed(ByVal Printed As Boolean)

3360      m_Printed = Printed

End Property

Public Property Get H() As Boolean

3370      H = m_h

End Property

Public Property Let H(ByVal H As Boolean)

3380      m_h = H

End Property

Public Property Get s() As Boolean

3390      s = m_s

End Property

Public Property Let s(ByVal s As Boolean)

3400      m_s = s

End Property

Public Property Get o() As Boolean

3410      o = m_o

End Property

Public Property Let o(ByVal o As Boolean)

3420      m_o = o

End Property

Public Property Get l() As Boolean

3430      l = m_l

End Property

Public Property Let l(ByVal l As Boolean)

3440      m_l = l

End Property

Public Property Get g() As Boolean

3450      g = m_g

End Property

Public Property Let g(ByVal g As Boolean)

3460      m_g = g

End Property

Public Property Get J() As Boolean

3470      J = m_J

End Property

Public Property Let J(ByVal J As Boolean)

3480      m_J = J

End Property

Public Property Get Authorised() As Boolean

3490      Authorised = m_Authorised

End Property

Public Property Let Authorised(ByVal Authorised As Boolean)

3500      m_Authorised = Authorised

End Property

Public Property Get Comment() As String

3510      Comment = m_Comment

End Property

Public Property Let Comment(ByVal Comment As String)

3520      m_Comment = Comment

End Property

Public Property Get DefIndex() As Long

3530      DefIndex = m_DefIndex

End Property

Public Property Let DefIndex(ByVal DefIndex As Long)

3540      m_DefIndex = DefIndex

End Property

Public Property Get PrintRefRange() As Boolean

3550      PrintRefRange = m_PrintRefRange

End Property

Public Property Let PrintRefRange(ByVal PrintRefRange As Boolean)

3560      m_PrintRefRange = PrintRefRange

End Property

Public Property Get DeltaDaysBackLimit() As Integer

3570      DeltaDaysBackLimit = m_DeltaDaysBackLimit

End Property

Public Property Let DeltaDaysBackLimit(ByVal DeltaDaysBackLimit As Integer)

3580      m_DeltaDaysBackLimit = DeltaDaysBackLimit

End Property

