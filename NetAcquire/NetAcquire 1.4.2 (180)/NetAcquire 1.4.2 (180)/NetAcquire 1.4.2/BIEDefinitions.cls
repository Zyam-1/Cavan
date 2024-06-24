VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolBIEDefinitions As New Collection

Private mBIEDefinitions As New BIEDefinitions
Public Function ItemByCode(ByVal Code As String) As BIEDefinition

          Dim b As BIEDefinition
          Dim RetVal As BIEDefinition

1910      Set RetVal = Nothing
1920      For Each b In mcolBIEDefinitions
1930          If b.Code = Code Then
1940              Set RetVal = b
1950              Exit For
1960          End If
1970      Next
1980      Set ItemByCode = RetVal

End Function

Public Function ItemByShortName(ByVal ShortName As String) As BIEDefinition

          Dim b As BIEDefinition
          Dim RetVal As BIEDefinition

1990      Set RetVal = Nothing
2000      For Each b In mcolBIEDefinitions
2010          If b.ShortName = ShortName Then
2020              Set RetVal = b
2030              Exit For
2040          End If
2050      Next
2060      Set ItemByShortName = RetVal

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
2070      Set NewEnum = mcolBIEDefinitions.[_NewEnum]

End Function

Public Sub Load()

          Dim BD As BIEDefinition
          Dim tb As Recordset
          Dim sql As String

2080      On Error GoTo Load_Error

2090      sql = "SELECT LongName, ShortName, " & _
              "COALESCE(DoDelta, 0) DoDelta, " & _
              "COALESCE(DeltaLimit, 0)DeltaLimit, " & _
              "COALESCE(PrintPriority, 999) PrintPriority, " & _
              "COALESCE(DP, 0) DP, " & _
              "BarCode, Units, " & _
              "COALESCE(H, 0) H, " & _
              "COALESCE(S, 0) S, " & _
              "COALESCE(L, 0) L, " & _
              "COALESCE(O, 0) O, " & _
              "COALESCE(G, 0) G, " & _
              "COALESCE(J, 0) J, " & _
              "Category, Code, " & _
              "COALESCE(Printable, 1) Printable, "
2100      sql = sql & _
              "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
              "COALESCE(PlausibleHigh, 999) PlausibleHigh, " & _
              "COALESCE(KnownToAnalyser, 1) KnownToAnalyser, " & _
              "COALESCE(SampleType, 'S') SampleType, " & _
              "COALESCE(InUse, 1) InUse, "
2110      sql = sql & _
              "COALESCE(MaleLow, 0) MaleLow, " & _
              "COALESCE(MaleHigh, 999) MaleHigh, " & _
              "COALESCE(FemaleLow, 0) FemaleLow, " & _
              "COALESCE(FemaleHigh, 999) FemaleHigh, " & _
              "COALESCE(FlagMaleLow, 0) FlagMaleLow, " & _
              "COALESCE(FlagMaleHigh, 999) FlagMaleHigh, " & _
              "COALESCE(FlagFemaleLow, 0) FlagFemaleLow, " & _
              "COALESCE(FlagFemaleHigh, 999) FlagFemaleHigh, " & _
              "COALESCE(LControlLow, 0) LControlLow, " & _
              "COALESCE(LControlHigh, 0) LControlHigh, " & _
              "COALESCE(NControlLow, 0) NControlLow, " & _
              "COALESCE(NControlHigh, 0) NControlHigh, " & _
              "COALESCE(HControlLow, 0) HControlLow, " & _
              "COALESCE(HControlHigh, 0) HControlHigh, "
2120      sql = sql & _
              "COALESCE(AgeFromDays, 0) AgeFromDays, " & _
              "COALESCE(AgeToDays, 43830) AgeToDays, " & _
              "COALESCE(AutoValLow, 0) AutoValLow, " & _
              "COALESCE(AutoValHigh, 0) AutoValHigh, " & _
              "Hospital, Analyser, ImmunoCode, " & _
              "COALESCE(SplitList, 1) SplitList, " & _
              "COALESCE(EOD, 1) EOD, " & _
              "COALESCE(LIH, 0) LIH, "
          '"ActiveFromDate, ActiveToDate,
2130      sql = sql & "ArchitectCode, COALESCE(BIE, 'B') BIE, " & _
              "COALESCE(ReRunDays, 0) ReRunDays, " & _
              "AgeFromText, AgeToText, " & _
              "COALESCE(PrintRefRange, 1) PrintRefRange, " & _
              "COALESCE(PrintSplit, 1) PrintSplit " & _
              "FROM BioTestDefinitions"
2140      Set tb = New Recordset
2150      RecOpenServer 0, tb, sql
2160      Do While Not tb.EOF
2170          Set BD = New BIEDefinition
2180          With BD
2190              .LongName = tb!LongName & ""
2200              .ShortName = tb!ShortName & ""
2210              .DoDelta = tb!DoDelta
2220              .DeltaLimit = tb!DeltaLimit
2230              .PrintPriority = tb!PrintPriority
2240              .DP = tb!DP
2250              .BarCode = tb!BarCode & ""
2260              .Units = tb!Units & ""
2270              .H = tb!H
2280              .s = tb!s
2290              .l = tb!l
2300              .o = tb!o
2310              .g = tb!g
2320              .J = tb!J
2330              .Category = tb!Category & ""
2340              .Code = tb!Code & ""
2350              .Printable = tb!Printable
2360              .PlausibleLow = tb!PlausibleLow
2370              .PlausibleHigh = tb!PlausibleHigh
2380              .KnownToAnalyser = tb!KnownToAnalyser
2390              .SampleType = tb!SampleType & ""
2400              .InUse = tb!InUse
2410              .MaleLow = tb!MaleLow
2420              .MaleHigh = tb!MaleHigh
2430              .FemaleLow = tb!FemaleLow
2440              .FemaleHigh = tb!FemaleHigh
2450              .FlagMaleLow = tb!FlagMaleLow
2460              .FlagMaleHigh = tb!FlagMaleHigh
2470              .FlagFemaleLow = tb!FlagFemaleLow
2480              .FlagFemaleHigh = tb!FlagFemaleHigh
2490              .LControlLow = tb!LControlLow
2500              .LControlHigh = tb!LControlHigh
2510              .NControlLow = tb!NControlLow
2520              .NControlHigh = tb!NControlHigh
2530              .HControlLow = tb!HControlLow
2540              .HControlHigh = tb!HControlHigh
2550              .AgeFromDays = tb!AgeFromDays
2560              .AgeToDays = tb!AgeToDays
2570              .AutoValLow = tb!AutoValLow
2580              .AutoValHigh = tb!AutoValHigh
2590              .Hospital = tb!Hospital & ""
2600              .Analyser = tb!Analyser & ""
2610              .ImmunoCode = tb!ImmunoCode & ""
2620              .SplitList = tb!SplitList
2630              .EOD = tb!EOD
2640              .LIH = tb!LIH
                  '570       .ActiveFromDate = tb!ActiveFromDate
                  '580       .ActiveToDate = tb!ActiveToDate
2650              .ArchitectCode = tb!ArchitectCode & ""
2660              .BIE = tb!BIE
2670              .ReRunDays = tb!ReRunDays
2680              .AgeFromText = tb!AgeFromText & ""
2690              .AgeToText = tb!AgeToText & ""
2700              .PrintRefRange = tb!PrintRefRange
2710              .PrintSplit = tb!PrintSplit
2720              mcolBIEDefinitions.Add BD
2730          End With
2740          tb.MoveNext
2750      Loop
2760      Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

2770      intEL = Erl
2780      strES = Err.Description
2790      LogError "BIEDefinitions", "Load", intEL, strES, sql

End Sub


Private Sub Class_Initialize()

2800      Load

End Sub


