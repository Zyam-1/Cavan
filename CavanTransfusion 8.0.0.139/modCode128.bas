Attribute VB_Name = "modCode128"
Option Explicit

Public Function Code128(ByVal Chaine As String) As String
  
        'Parameters : a string
        'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
        '         * an empty string if the supplied parameter is no good
  
        Dim i As Integer
        Dim CheckSum As Long
        Dim Mini As Integer
        Dim Dummy As Integer
        Dim TableB As Boolean
  
10      Code128 = ""
20      If Len(Chaine) > 0 Then
        'Check for valid characters
30        For i = 1 To Len(Chaine)
40          Select Case Asc(Mid$(Chaine, i, 1))
            Case 32 To 126, 203
50          Case Else
60            i = 0
70            Exit For
80          End Select
90        Next
          'Calculation of the code string with optimized use of tables B and C
100       Code128 = ""
110       TableB = True
120       If i > 0 Then
130         i = 1 ' i% become the string index
140         Do While i <= Len(Chaine)
150           If TableB Then
                'See if interesting to switch to table C
                'yes for 4 digits at start or end, else if 6 digits
160             Mini = IIf(i = 1 Or i + 3 = Len(Chaine), 4, 6)
170             GoSub testnum
180             If Mini < 0 Then ' Choice of table C
190               If i = 1 Then 'Starting with table C
200                 Code128 = Chr$(210)
210               Else ' Switch to table C
220                 Code128 = Code128 & Chr$(204)
230               End If
240               TableB = False
250             Else
260               If i = 1 Then Code128 = Chr$(209) 'Starting with table B
270             End If
280           End If
290           If Not TableB Then
                'We are on table C, try to process 2 digits
300             Mini = 2
310             GoSub testnum
320             If Mini < 0 Then 'OK for 2 digits, process it
330               Dummy = Val(Mid$(Chaine, i, 2))
340               Dummy = IIf(Dummy < 95, Dummy + 32, Dummy + 105)
350               Code128 = Code128 & Chr$(Dummy)
360               i = i + 2
370             Else ' We haven't 2 digits, switch to table B
380               Code128 = Code128 & Chr$(205)
390               TableB = True
400             End If
410           End If
420           If TableB Then
                ' Process 1 digit with table B
430             Code128 = Code128 & Mid$(Chaine, i, 1)
440             i = i + 1
450           End If
460         Loop
            ' Calculation of the checksum
470         For i = 1 To Len(Code128)
480           Dummy = Asc(Mid$(Code128, i, 1))
490           Dummy = IIf(Dummy < 127, Dummy - 32, Dummy - 105)
500           If i = 1 Then CheckSum = Dummy
510           CheckSum = (CheckSum + (i - 1) * Dummy) Mod 103
520         Next
            ' Calculation of the checksum ASCII code
530         CheckSum = IIf(CheckSum < 95, CheckSum + 32, CheckSum + 105)
            ' Add the checksum and the STOP
540         Code128 = Code128 & Chr$(CheckSum) & Chr$(211)
550       End If
560     End If
570     Exit Function
testnum:
        'if the mini characters from i are numeric, then mini=0
580     Mini = Mini - 1
590     If i + Mini <= Len(Chaine) Then
600       Do While Mini >= 0
610         If Asc(Mid$(Chaine, i + Mini, 1)) < 48 Or Asc(Mid$(Chaine, i + Mini, 1)) > 57 Then Exit Do
620         Mini = Mini - 1
630       Loop
640     End If
650   Return

End Function
Public Function ISOmod37_2(ByVal DIN As String) As String

      Dim J As Integer
      Dim Sum As Long
      Dim CharValue As Long
      Dim CheckValue As Long
      Dim ISOCharTable As String
      Dim ThisChar As String

10    ISOCharTable = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ*"
20    Sum = 0
30    For J = 1 To Len(DIN)
40      ThisChar = Mid(DIN, J, 1)
50      Select Case ThisChar
          Case "0" To "9": CharValue = Asc(ThisChar) - 48
60        Case "A" To "Z": CharValue = Asc(ThisChar) - 55
70        Case "*": CharValue = 36
80      End Select
90      Sum = ((Sum + CharValue) * 2) Mod 37
100   Next
110   CheckValue = (38 - Sum) Mod 37
120   ISOmod37_2 = Mid$(ISOCharTable, CheckValue + 1, 1)

End Function
