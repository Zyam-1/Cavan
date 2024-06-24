Attribute VB_Name = "modWardPrint"
Option Explicit

Public Type PrintLine
  Analyte As String * 19
  Result As String * 6
  Flag As String * 3
  Units As String * 7
  NormalRange As String * 11
  Fasting As String * 9
End Type

Public Function CheckGentTobra(ByVal SampleID As String, _
                               ByVal PatName As String, _
                               ByVal LongName As String, _
                               ByVal Code As String, _
                               ByVal Result As String) _
                               As String

      Dim tb As Recordset
      Dim tbR As Recordset
      Dim sql As String
      Dim AssValue As String
      Dim S As String

      'Gentamicin and Tobramicin

10    On Error GoTo CheckGentTobra_Error

20    S = ""

30    sql = "SELECT DISTINCT D.SampleID " & _
            "FROM Demographics D WHERE " & _
            "D.SampleID IN " & _
            "  (  SELECT SampleID FROM BioResults WHERE " & _
            "     (SampleID = '" & Val(SampleID) - 1 & "' " & _
            "      OR SampleID = '" & Val(SampleID) + 1 & "') " & _
            "      AND Code = '" & Code & "'  ) " & _
            "AND D.PatName = '" & AddTicks(PatName) & "' " & _
            "AND (D.SampleID = '" & Val(SampleID) - 1 & "' " & _
            "     OR SampleID = '" & Val(SampleID) + 1 & "')"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      sql = "Select Result from BioResults where " & _
              "SampleID = '" & tb!SampleID & "' " & _
              "and Code = '" & Code & "'"
80      Set tbR = New Recordset
90      RecOpenServer 0, tbR, sql
100     If Not tbR.EOF Then
110       AssValue = tbR!Result & ""
120       If Val(AssValue) < Val(Result) Or InStr(AssValue, "<") <> 0 Then
130         S = LongName & " Trough" & vbTab & Format$(AssValue, "0.0")
140       Else
150         S = LongName & " Peak" & vbTab & Format$(AssValue, "0.0")
160       End If
170     End If
180   End If
190   CheckGentTobra = S

200   Exit Function

CheckGentTobra_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "modWardPrint", "CheckGentTobra", intEL, strES, sql

End Function

Public Sub PrintHeadingCavan(ByVal SampleID As String, _
                             ByVal Cn As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim clinicianfound As Integer
      Dim GPText As String
      Dim OBS As Observations

10    On Error GoTo PrintHeadingCavan_Error

20    sql = "SELECT Demographics.*, GPs.Addr0 AS GPAddr0, GPs.Addr1 AS GPAddr1 " & _
            "FROM Demographics LEFT OUTER JOIN GPS " & _
            "ON Demographics.GP = GPs.Text " & _
            "WHERE Demographics.SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient Cn, tb, sql

50    If tb.EOF Then Exit Sub

60    Printer.Font.Name = "Courier New"
70    Printer.Font.Size = 12
80    Printer.Print

90    Printer.Print "Lab. No. : "; tb!SampleID;
100   Printer.Print Tab(35); "   Hosp. No: ";
110   Printer.Print Left$(HospName(Cn), 1) & " ";
120   Printer.Print tb!Chart & ""

130   Printer.Print "Clinician: "; 'cons/ward or gp/addr
140   If Trim$(tb!Clinician & "") <> "" Then
150     Printer.Print tb!Clinician;
160     clinicianfound = True
170   Else
180     GPText = tb!GP & ""
190     If Trim$(GPText) <> "" Then
200       Printer.Print GPText;
210     End If
220     clinicianfound = False
230   End If

240   Printer.Print Tab(35); "       Name: ";
250   Printer.Font.Bold = True
260   Printer.Print tb!PatName & ""
270   Printer.Font.Bold = False

280   Printer.Print "Report to: ";
290   If clinicianfound Then
300     Printer.Print tb!Ward & "";
310   Else
320     Printer.Print tb!GPAddr0 & "";
330   End If
340   Printer.Print Tab(42); "Addr: "; tb!Addr0 & ""

350   Printer.Print Tab(12);
360   If Not clinicianfound Then
370     Printer.Print tb!GPAddr1 & "";
380   End If
390   Printer.Print Tab(35); "       Addr: "; tb!Addr1 & ""
400   Printer.Print Tab(35); "   DoB(Age): ";
410   If Not IsNull(tb!DoB) Then
420     Printer.Print tb!DoB;
430   End If
440   Printer.Print " ("; tb!Age & ""; ") Sex:"; tb!Sex & ""

450   Printer.Print "Comment  : ";

460   Set OBS = New Observations
470   Set OBS = OBS.Load(SampleID, "Demographic")
480   If Not OBS Is Nothing Then
490     Printer.Print Left$(OBS.Item(1).Comment, 60)
500   Else
510     Printer.Print
520   End If

530   Printer.Print String$(75, "-")

540   Exit Sub

PrintHeadingCavan_Error:

      Dim strES As String
      Dim intEL As Integer

550   intEL = Erl
560   strES = Err.Description
570   LogError "modWardPrint", "PrintHeadingCavan", intEL, strES, sql

End Sub



