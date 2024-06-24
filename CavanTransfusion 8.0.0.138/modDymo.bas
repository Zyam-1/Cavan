Attribute VB_Name = "modDymo"

Option Explicit
Public Sub PrintDymo(ByVal PatName As String, _
                     ByVal Chart As String, _
                     ByVal DoB As String, _
                     ByVal LabNumber As String)

          Dim Px As Printer
          Dim PxFound As Boolean
          Dim strOriginalPrinter As String
          Dim n As Integer

10    On Error Resume Next

20    strOriginalPrinter = Printer.DeviceName

30    PxFound = False
40    For Each Px In Printers
50      If Px.DeviceName = "DYMO LabelWriter 330 Turbo" Then
60          Set Printer = Px
70          PxFound = True
80          Exit For
90      End If
100   Next
110   If Not PxFound Then Exit Sub

120   For n = 1 To 2

130     Printer.Font.Name = "Courier New"
140     Printer.Font.Size = 8
150     Printer.Font.Bold = True
160     Printer.Print " "; Left$(PatName & Space$(23), 23); " ";
170     Printer.Font.Bold = False
180     Printer.Print Chart

190     Printer.Font.Bold = True
200     Printer.Print " "; LabNumber;
210     Printer.Font.Bold = False
220     If IsDate(DoB) Then
230         Printer.Print Tab(17); "DoB:"; Format(DoB, "dd/mm/yyyy")
240     End If

250     Printer.Print
260     Printer.Print
270     Printer.Font.Size = 16
280     Printer.Font.Bold = True
290     Printer.Print " "; LabNumber; Tab(10); LabNumber

300     Printer.EndDoc

310   Next


320   Printer.Print " "; LabNumber; Tab(10); LabNumber
330   Printer.Print
340   Printer.Print " "; LabNumber; Tab(10); LabNumber
350   Printer.EndDoc

360   For Each Px In Printers
370     If Px.DeviceName = strOriginalPrinter Then
380         Set Printer = Px
390         Exit For
400     End If
410   Next

End Sub
Public Sub PrintPDFCavan(ByVal UnitNumber As String, _
                         ByVal PatName As String, _
                         ByVal Chart As String, _
                         ByVal DoB As String, _
                         ByVal Ward As String, _
                         ByVal PatGroup As String, _
                         ByVal Product As String, _
                         ByVal UnitGroup As String, _
                         ByVal UnitExpiry As String, _
                         ByVal IsCompatWith As String, _
                         ByVal SampleDate As String, _
                         ByVal strSex As String, _
                         ByVal strPatientGroup As String, _
                         ByVal PatSurName As String, _
                         ByVal PatForeName As String)

      Dim Px As Printer
      Dim PxFound As Boolean
      Dim strOriginalPrinter As String
      Dim s As String
      Dim Y As Integer
      Dim SurName As String
      Dim ForeName As String
      Dim Prod() As String
      Dim UseBy As String
      Dim Generic As String
      Dim vbQRObj As vbQRCode
      Dim QRSize As Integer
      Dim intScale As Integer
      Dim off As Single
      Dim Matrix() As Integer
      Dim Y1, X1 As Long
      Dim lngColor As Long

      'Print3BarCodes LabNumber

10    On Error GoTo PrintPDFCavan_Error

20    strOriginalPrinter = Printer.DeviceName

30    PxFound = False
40    For Each Px In Printers
          '50      If InStr(UCase$(Px.DeviceName), "ZEBRA 105SL (300 DPI)") Then
50        If UCase$(Px.DeviceName) = UCase(TransfusionPDF) Then
60            Set Printer = Px
70            PxFound = True
80            Exit For
90        End If
100   Next
110   If Not PxFound Then Exit Sub

120   Printer.ScaleMode = vbMillimeters
130   frmXMLabel.Picture1.Picture = frmXMLabel.Picture1.Image

140   Printer.Font.Name = "Courier New"
150   Printer.Font.Size = 8

160   If PatSurName = "" And PatForeName = "" Then
170       Y = InStr(PatName, " ")
180       If Y <> 0 Then
190           SurName = Left$(PatName, Y - 1)
200           ForeName = Mid$(PatName, Y + 1)
210       Else
220           SurName = PatName
230           ForeName = ""
240       End If
250       SurName = UCase$(SurName)
260   Else
270       SurName = PatSurName
280       ForeName = PatForeName
290   End If


      'Minimum of (SampleDate+72Hours) and UnitExpiry
300   If DateDiff("n", UnitExpiry, DateAdd("h", 72, SampleDate)) > 0 Then
310       UseBy = Format$(UnitExpiry, "dd/MM/yyyy hh:nn")
320   Else
330       UseBy = Format$(DateAdd("h", 72, SampleDate), "dd/MM/yyyy at HH:nn")
340   End If
      'Zyam added a blank line for Printing label 26-01-24
      Printer.Print ""
      'Zyam
350   Printer.Print "Unit Number ";
360   Printer.Font.Bold = True
370   Printer.Print UnitNumber
380   Printer.Font.Bold = False

390   Printer.Font.Size = 4
400   Printer.Print String$(60, "-")

410   Printer.Font.Size = 6
420   Printer.Print IsCompatWith

430   Printer.Font.Size = 8
440   Printer.Print " Chart No:"; Chart
450   Printer.Print "  Surname:"; SurName
460   Printer.Print " Forename:"; ForeName
470   Printer.Print "      DoB:"; DoB
480   Printer.Print "     Ward:"; Ward
490   Printer.Print "Blood Grp:"; PatGroup

500   Printer.Font.Size = 6
510   s = "C2|"
520   s = s & Chart & "|"
530   s = s & SurName & "|"
540   s = s & ForeName & "|"
550   s = s & DoB & "|"
560   s = s & strSex & "|"
570   s = s & strPatientGroup & "|"
580   s = s & UnitNumber
      'g.row = selectedrows(d)
      '.g.col = 6
      'If Len(Trim$(g)) = 0 Then    'Not scanned-in unit.
      '    .g.col = 0
      '    strBTCunit = "=" & Left(g, 13) & Asc(Right(g, 1)) + getCheckDigitFactor(Right(g, 1))
      '    s = s & strBTCunit
      'Else
      '    s = s & g
      'End If

590   frmXMLabel.Picture1.Cls  'Clear Picture Box
600   Set vbQRObj = New vbQRCode

610   vbQRObj.FindBestMask = 1
620   vbQRObj.ShowMarkers = 0
630   vbQRObj.QuietZone = 0

640   If (vbQRObj.Encode(s)) Then
650       QRSize = vbQRObj.Size
660       intScale = Int(frmXMLabel.Picture1.ScaleWidth / QRSize)
670       off = (frmXMLabel.Picture1.ScaleWidth - intScale * QRSize) / 2
680       Matrix() = vbQRObj.Matrix()
690       For Y1 = 0 To QRSize - 1
700           For X1 = 0 To QRSize - 1
710               lngColor = vbWhite
720               If Matrix(Y1, X1) = 1 Then lngColor = vbBlack
730               frmXMLabel.Picture1.Line (off + X1 * intScale, off + Y1 * intScale)-Step(intScale, intScale), lngColor, BF
740           Next
750       Next
760   End If

770   If frmXMLabel.Picture1.Image <> 0 Then    'check if there is a picture to print
780       Printer.PaintPicture frmXMLabel.Picture1.Image, _
                               10, 25, 15, 15
790   End If
800   Set vbQRObj = Nothing

810   Printer.Font.Name = "Courier New"

820   Printer.Print
830   Printer.Print
840   Printer.Print
850   Printer.Print
860   Printer.Print
870   Printer.Print
880   Printer.Print


890   Generic = ProductGenericForWording(Product)
900   Printer.Font.Size = 6
910   If Generic = "Red Cells" Then
920       Printer.Print "Do not commence Transfusion after"
930   Else
940       Printer.Print
950   End If

960   Printer.Font.Size = 8

970   If Generic = "Red Cells" Then
980       Printer.Print UseBy
990   Else
1000      Printer.Print
1010  End If

1020  Printer.Print " Removed from Blood Bank"
1030  Printer.Print "   By:"
1040  Printer.Print " Time:"
1050  Printer.Print " Date:"
      '1010  Printer.Print
1060  Printer.Print "   ";

      'Use first two words of product
1070  Prod = Split(Product, " ")
1080  If UBound(Prod) > 0 Then
1090      Product = Prod(0) & " " & Prod(1)
1100  End If

      '870   Printer.Font.Name = "Code 128"
      '880   Printer.Font.Size = 12
      '890   Printer.Print Code128("=" & UCase(UnitNumber))
      '900   Printer.Font.Name = "Courier New"

1110  Printer.Font.Size = 6
1120  Printer.Print ; Product; " ";
1130  Printer.Font.Bold = True
1140  Printer.Print UnitNumber; " "
1150  Printer.Print "   "; UnitGroup; "   ";
1160  Printer.Print UnitExpiry
1170  Printer.Font.Bold = False
1180  Printer.Print "   "; IsCompatWith
1190  Printer.Print "   "; PatName; " ("; Chart; ")"
1200  Printer.Font.Size = 8

1210  Printer.EndDoc


1220  Printer.Font.Name = "Courier New"
1230  Printer.Font.Size = 8

1240  If PatSurName = "" And PatForeName = "" Then
1250      Y = InStr(PatName, " ")
1260      If Y <> 0 Then
1270          SurName = Left$(PatName, Y - 1)
1280          ForeName = Mid$(PatName, Y + 1)
1290      Else
1300          SurName = PatName
1310          ForeName = ""
1320      End If
1330      SurName = UCase$(SurName)
1340  Else
1350      SurName = PatSurName
1360      ForeName = PatForeName
1370  End If


1380  Printer.Print "Unit Number ";
1390  Printer.Font.Bold = True
1400  Printer.Print UnitNumber
1410  Printer.Font.Bold = False

1420  Printer.Font.Size = 4
1430  Printer.Print String$(60, "-")

1440  Printer.Font.Size = 6
1450  Printer.Print IsCompatWith

1460  Printer.Font.Size = 8
1470  Printer.Print " Chart No:"; Chart
1480  Printer.Print "  Surname:"; SurName
1490  Printer.Print " Forename:"; ForeName
1500  Printer.Print "      DoB:"; DoB
1510  Printer.Print "     Ward:"; Ward
1520  Printer.Print "Blood Grp:"; PatGroup



      '1440  Printer.Font.Name = "Code PDF417"
      '1450  Printer.Font.Size = 8
      '1460  s = "C|"
      '1470  s = s & Chart & "|"
      '1480  s = s & SurName & "|"
      '1490  s = s & ForeName & "|"
      '1500  s = s & UnitNumber & "|"
      '1510  s = s & DoB & " "
      '
      '1520  CodeBarre = PDF417(s, -1, 4, CodeErr)
      '
      '1530  If CodeErr = 0 Then
      '
      '1540    Printer.Print
      '1550    Printer.Print
      '1560    Printer.Print
      '1570    Printer.Print
      '
      '1580    CB = Split(CodeBarre, vbCr)
      '1590    For Y = 0 To UBound(CB)
      '1600        CB(Y) = Replace(CB(Y), vbLf, "")
      '1610        Printer.Print "          "; CB(Y)
      '1620    Next
      '
      '1630  End If


1530  Printer.Font.Size = 6
1540  s = "C2|"
1550  s = s & Chart & "|"
1560  s = s & SurName & "|"
1570  s = s & ForeName & "|"
1580  s = s & DoB & "|"
1590  s = s & strSex & "|"
1600  s = s & strPatientGroup & "|"
1610  s = s & UnitNumber
      'g.row = selectedrows(d)
      '.g.col = 6
      'If Len(Trim$(g)) = 0 Then    'Not scanned-in unit.
      '    .g.col = 0
      '    strBTCunit = "=" & Left(g, 13) & Asc(Right(g, 1)) + getCheckDigitFactor(Right(g, 1))
      '    s = s & strBTCunit
      'Else
      '    s = s & g
      'End If

1620  frmXMLabel.Picture1.Cls  'Clear Picture Box
1630  Set vbQRObj = New vbQRCode

1640  vbQRObj.FindBestMask = 1
1650  vbQRObj.ShowMarkers = 0
1660  vbQRObj.QuietZone = 0

1670  If (vbQRObj.Encode(s)) Then
1680      QRSize = vbQRObj.Size
1690      intScale = Int(frmXMLabel.Picture1.ScaleWidth / QRSize)
1700      off = (frmXMLabel.Picture1.ScaleWidth - intScale * QRSize) / 2
1710      Matrix() = vbQRObj.Matrix()
1720      For Y1 = 0 To QRSize - 1
1730          For X1 = 0 To QRSize - 1
1740              lngColor = vbWhite
1750              If Matrix(Y1, X1) = 1 Then lngColor = vbBlack
1760              frmXMLabel.Picture1.Line (off + X1 * intScale, off + Y1 * intScale)-Step(intScale, intScale), lngColor, BF
1770          Next
1780      Next
1790  End If

1800  If frmXMLabel.Picture1.Image <> 0 Then    'check if there is a picture to print
1810      Printer.PaintPicture frmXMLabel.Picture1.Image, _
                               10, 25, 10, 10
1820  End If
1830  Set vbQRObj = Nothing

      'Printer.Font.Name = "Courier New"

      'Printer.Print
1840  Printer.Print
1850  Printer.Print
1860  Printer.Print
1870  Printer.Print
1880  Printer.Print
1890  Printer.Print


1900  Printer.Font.Name = "Courier New"

1910  Printer.Font.Size = 6
1920  If Generic = "Red Cells" Then
1930      Printer.Print "Do not commence Transfusion after"
1940  Else
1950      Printer.Print
1960  End If

1970  Printer.Font.Size = 8

1980  If Generic = "Red Cells" Then
1990      Printer.Print UseBy
2000  Else
2010      Printer.Print
2020  End If

      '1570  Printer.Print
      '1580  Printer.Print


      'Use first two words of product
2030  Prod = Split(Product, " ")
2040  If UBound(Prod) > 0 Then
2050      Product = Prod(0) & " " & Prod(1)
2060  End If

2070  Printer.Print "  ";
2080  Printer.Font.Name = "Code 128"
2090  Printer.Font.Size = 20    '12/24

2100  Printer.Print Code128(UCase(UnitNumber))
2110  Printer.Font.Name = "Courier New"
2120  Printer.Font.Size = 8    '12
      '1680  Printer.Print '
2130  Printer.Print    '
2140  Printer.Print

2150  Printer.Font.Size = 6
2160  Printer.Print "   ";
2170  Printer.Print ; Product; " ";
2180  Printer.Font.Bold = True
2190  Printer.Print UnitNumber; " "
2200  Printer.Print "   "; UnitGroup; "   ";
2210  Printer.Print UnitExpiry
2220  Printer.Font.Bold = False
2230  Printer.Print "   "; IsCompatWith
2240  Printer.Print "   "; PatName; " ("; Chart; ")"
2250  Printer.Font.Size = 8

2260  Printer.ScaleMode = vbTwips

2270  Printer.EndDoc


2280  For Each Px In Printers
2290      If Px.DeviceName = strOriginalPrinter Then
2300          Set Printer = Px
2310          Exit For
2320      End If
2330  Next

2340  Exit Sub

PrintPDFCavan_Error:

      Dim strES As String
      Dim intEL As Integer

2350  intEL = Erl
2360  strES = Err.Description
2370  LogError "modDymo", "PrintPDFCavan", intEL, strES

End Sub

Private Function PDF417(ByVal Chaine As String, _
                        Optional ByRef sécu As Integer, _
                        Optional ByRef nbcol As Integer, _
                        Optional ByRef CodeErr As Integer) _
                        As String
      'V 1.0.1

      'Parameters : The string to encode.
      '             The hoped sécurity level, -1 = automatic.
      '             The hoped number of data MC columns, -1 = automatic.
      '             A variable which will can retrieve an error number.
      'Return : * a string which, printed with the PDF417.TTF font, gives the bar code.
      '         * an empty string if the given parameters aren't good.
      '         * sécu% contain le really used sécurity level.
      '         * NbCol% contain the really used number of data CW columns.
      '         * Codeerr% is 0 if no error occured, else :
      '           0  : No error
      '           1  : Chaine$ is empty
      '           2  : Chaine$ contain too many datas, we go beyong the 928 CWs.
      '           3  : Number of CWs per row too small, we go beyong 90 rows.
      '           10 : The sécurity level has being lowers not to exceed the 928 CWs. (It's not an error, only a warning.)

      'Global variables
          Dim i As Integer
          Dim j As Integer
          Dim K As Integer
          Dim IndexChaine As Integer
          Dim Dummy As String
          Dim Flag As Boolean
          'Splitting into blocks
          Dim Liste%(), IndexListe%
          'Data compaction
          Dim Longueur%, ChaineMC$, total
          '"text" mode processing
          Dim ListeT%(), IndexListeT%, CurTable%, ChaineT$, NewTable%
          'Reed Solomon codes
          Dim MCcorrection%()
          'MC de cotés gauche et droit / Left and right side CWs
          Dim C1%, C2%, C3%
          'Sous programme QuelMode / Sub routine QuelMode
          Dim mode%, CodeASCII%
          'Sous programme QuelleTable / Sub routine QuelleTable
          Dim Table%
          'Sous programme Modulo / Sub routine Modulo
          Dim ChaineMod$, Diviseur&, ChaineMult$, Nombre&
          'Tables
          Dim ASCII$
          'This string describe the ASCII code for the "text" mode.
          'ASCII$ contain 95 fields of 4 digits which correspond to char. ASCII values 32 to 126. These fields are :
          '  2 digits indicating the table(s) (1 or several) where this char. is located. (Table numbers : 1, 2, 4 and 8)
          '  2 digits indicating the char. number in the table
          '  Sample : 0726 at the beginning of the string :
          '  The Char. having code 32 is in the tables 1, 2 and 4 at row 26
10    ASCII$ = "07260810082008151218042104100828082308241222042012131216121712190400040104020403040404050406040704080409121408000801042308020825080301000101010201030104010501060107010801090110011101120113011401150116011701180119012001210122012301240125080408050806042408070808020002010202020302040205020602070208020902100211021202130214021502160217021802190220022102220223022402250826082108270809"
          Dim CoefRS$(8)
          'CoefRS$ contient 8 chaines représentant les coefficients sur 3 chiffres des polynomes de calcul des codes de reed Solomon
          'CoefRS$ contain 8 strings describing the factors of the polynomial equations for the reed Solomon codes.
20    CoefRS$(0) = "027917"
30    CoefRS$(1) = "522568723809"
40    CoefRS$(2) = "237308436284646653428379"
50    CoefRS$(3) = "274562232755599524801132295116442428295042176065"
60    CoefRS$(4) = "361575922525176586640321536742677742687284193517273494263147593800571320803133231390685330063410"
70    CoefRS$(5) = "539422006093862771453106610287107505733877381612723476462172430609858822543376511400672762283184440035519031460594225535517352605158651201488502648733717083404097280771840629004381843623264543"
80    CoefRS$(6) = "521310864547858580296379053779897444400925749415822093217208928244583620246148447631292908490704516258457907594723674292272096684432686606860569193219129186236287192775278173040379712463646776171491297763156732095270447090507048228821808898784663627378382262380602754336089614087432670616157374242726600269375898845454354130814587804034211330539297827865037517834315550086801004108539"
90    CoefRS$(7) = "524894075766882857074204082586708250905786138720858194311913275190375850438733194280201280828757710814919089068569011204796605540913801700799137439418592668353859370694325240216257284549209884315070329793490274877162749812684461334376849521307291803712019358399908103511051008517225289470637731066255917269463830730433848585136538906090002290743199655903329049802580355588188462010134628320479130739071263318374601192605142673687234722384177752607640455193689707805641048060732621895544261852655309697755756060231773434421726528503118049795032144500238836394280566319009647550073914342126032681331792620060609441180791893754605383228749760213054297134054834299922191910532609829189020167029872449083402041656505579481173404251688095497555642543307159924558648055497010"
100   CoefRS$(8) = "352077373504035599428207409574118498285380350492197265920155914299229643294871306088087193352781846075327520435543203666249346781621640268794534539781408390644102476499290632545037858916552041542289122272383800485098752472761107784860658741290204681407855085099062482180020297451593913142808684287536561076653899729567744390513192516258240518794395768848051610384168190826328596786303570381415641156237151429531207676710089168304402040708575162864229065861841512164477221092358785288357850836827736707094008494114521002499851543152729771095248361578323856797289051684466533820669045902452167342244173035463651051699591452578037124298332552043427119662777475850764364578911283711472420245288594394511327589777699688043408842383721521560644714559062145873663713159672729"
110   CoefRS$(8) = CoefRS$(8) & "624059193417158209563564343693109608563365181772677310248353708410579870617841632860289536035777618586424833077597346269757632695751331247184045787680018066407369054492228613830922437519644905789420305441207300892827141537381662513056252341242797838837720224307631061087560310756665397808851309473795378031647915459806590731425216548249321881699535673782210815905303843922281073469791660162498308155422907817187062016425535336286437375273610296183923116667751353062366691379687842037357720742330005039923311424242749321054669316342299534105667488640672576540316486721610046656447171616464190531297321762752533175134014381433717045111020596284736138646411877669141919045780407164332899165726600325498655357752768223849647063310863251366304282738675410389244031121303263"
          Dim CodageMC$(2)
          'CodageMC$ contain the 3 sets of the 929 MCs. Each MC est described in the PDF417.TTF font by 3 char. composing 3 time 5 bits. The first bit which is always 1
          ' and the last one which is always 0 are into the separator character.
120   CodageMC$(0) = "urAxfsypyunkxdwyozpDAulspBkeBApAseAkprAuvsxhypnkutwxgzfDAplsfBkfrApvsuxyfnkptwuwzflspsyfvspxyftwpwzfxyyrxufkxFwymzonAudsxEyolkucwdBAoksucidAkokgdAcovkuhwxazdnAotsugydlkoswugjdksosidvkoxwuizdtsowydswowjdxwoyzdwydwjofAuFsxCyodkuEwxCjclAocsuEickkocgckcckEcvAohsuayctkogwuajcssogicsgcsacxsoiycwwoijcwicyyoFkuCwxBjcdAoEsuCicckoEguCbcccoEaccEoEDchkoawuDjcgsoaicggoabcgacgDobjcibcFAoCsuBicEkoCguBbcEcoCacEEoCDcECcascagcaacCkuAroBaoBDcCBtfkwpwyezmnAtdswoymlktcwwojFBAmksFAkmvkthwwqzFnAmtstgyFlkmswFksFkgFvkmxwtizFtsmwyFswFsiFxwmyzFwyFyzvfAxpsyuyvdkxowyujqlAvcsxoiqkkvcgxobqkcvcamfAtFswmyqvAmdktEwwmjqtkvgwxqjhlAEkkmcgtEbhkkqsghkcEvAmhstayhvAEtkmgwtajhtkqwwvijhssEsghsgExsmiyhxsEwwmijhwwqyjhwiEyyhyyEyjhyjvFkxmwytjqdAvEsxmiqckvEgxmbqccvEaqcEqcCmFktCwwljqhkmEstCigtAEckvaitCbgskEccmEagscqgamEDEcCEhkmawtDjgxkEgsmaigwsqiimabgwgEgaEgDEiwmbjgywEiigyiEibgybgzjqFAvCsxliqEkvCgxlbqEcvCaqEEvCDqECqEBEFAmCstBighAEEkmCgtBbggkqagvDbggcEEEmCDggEqaDgg"
130   CodageMC$(0) = CodageMC$(0) & "CEasmDigisEagmDbgigqbbgiaEaDgiDgjigjbqCkvBgxkrqCcvBaqCEvBDqCCqCBECkmBgtArgakECcmBagacqDamBDgaEECCgaCECBEDggbggbagbDvAqvAnqBBmAqEBEgDEgDCgDBlfAspsweyldksowClAlcssoiCkklcgCkcCkECvAlhssqyCtklgwsqjCsslgiCsgCsaCxsliyCwwlijCwiCyyCyjtpkwuwyhjndAtoswuincktogwubncctoancEtoDlFksmwwdjnhklEssmiatACcktqismbaskngglEaascCcEasEChklawsnjaxkCgstrjawsniilabawgCgaawaCiwlbjaywCiiayiCibCjjazjvpAxusyxivokxugyxbvocxuavoExuDvoCnFAtmswtirhAnEkxviwtbrgkvqgxvbrgcnEEtmDrgEvqDnEBCFAlCssliahACEklCgslbixAagknagtnbiwkrigvrblCDiwcagEnaDiwECEBCaslDiaisCaglDbiysaignbbiygrjbCaDaiDCbiajiCbbiziajbvmkxtgywrvmcxtavmExtDvmCvmBnCktlgwsrraknCcxtrracvnatlDraEnCCraCnCBraBCCklBgskraakCCclBaiikaacnDalBDiicrbaCCCiiEaaCCCBaaBCDglBrabgCDaijgabaCDDijaabDCDrijrvlcxsqvlExsnvlCvlBnBctkqrDcnBEtknrDEvlnrDCnBBrDBCBclAqaDcCBElAnibcaDEnBnibErDnCBBibCaDBibBaDqibqibnxsfvkltkfnAmnAlCAoaBoiDoCAlaBlkpkBdAkosBckkogsebBcckoaBcEkoDBhkkqwsfjBgskqiBggkqbBgaBgDBiwkrjBiiBibBjjlpAsuswhil"
140   CodageMC$(0) = CodageMC$(0) & "oksuglocsualoEsuDloCBFAkmssdiDhABEksvisdbDgklqgsvbDgcBEEkmDDgElqDBEBBaskniDisBagknbDiglrbDiaBaDBbiDjiBbbDjbtukwxgyirtucwxatuEwxDtuCtuBlmkstgnqklmcstanqctvastDnqElmCnqClmBnqBBCkklgDakBCcstrbikDaclnaklDbicnraBCCbiEDaCBCBDaBBDgklrDbgBDabjgDbaBDDbjaDbDBDrDbrbjrxxcyyqxxEyynxxCxxBttcwwqvvcxxqwwnvvExxnvvCttBvvBllcssqnncllEssnrrcnnEttnrrEvvnllBrrCnnBrrBBBckkqDDcBBEkknbbcDDEllnjjcbbEnnnBBBjjErrnDDBjjCBBqDDqBBnbbqDDnjjqbbnjjnxwoyyfxwmxwltsowwfvtoxwvvtmtslvtllkossfnlolkmrnonlmlklrnmnllrnlBAokkfDBolkvbDoDBmBAljbobDmDBljbmbDljblDBvjbvxwdvsuvstnkurlurltDAubBujDujDtApAAokkegAocAoEAoCAqsAqgAqaAqDAriArbkukkucshakuEshDkuCkuBAmkkdgBqkkvgkdaBqckvaBqEkvDBqCAmBBqBAngkdrBrgkvrBraAnDBrDAnrBrrsxcsxEsxCsxBktclvcsxqsgnlvEsxnlvCktBlvBAlcBncAlEkcnDrcBnEAlCDrEBnCAlBDrCBnBAlqBnqAlnDrqBnnDrnwyowymwylswotxowyvtxmswltxlksosgfltoswvnvoltmkslnvmltlnvlAkokcfBloksvDnoBlmAklbroDnmBllbrmDnlAkvBlvDnvbrvyzeyzdwyexyuwydxytswetwuswdvxutwtvxtkselsuksdntulstrvu"
150   CodageMC$(1) = "ypkzewxdAyoszeixckyogzebxccyoaxcEyoDxcCxhkyqwzfjutAxgsyqiuskxggyqbuscxgausExgDusCuxkxiwyrjptAuwsxiipskuwgxibpscuwapsEuwDpsCpxkuywxjjftApwsuyifskpwguybfscpwafsEpwDfxkpywuzjfwspyifwgpybfwafywpzjfyifybxFAymszdixEkymgzdbxEcymaxEEymDxECxEBuhAxasyniugkxagynbugcxaaugExaDugCugBoxAuisxbiowkuigxbbowcuiaowEuiDowCowBdxAoysujidwkoygujbdwcoyadwEoyDdwCdysozidygozbdyadyDdzidzbxCkylgzcrxCcylaxCEylDxCCxCBuakxDgylruacxDauaExDDuaCuaBoikubgxDroicubaoiEubDoiCoiBcykojgubrcycojacyEojDcyCcyBczgojrczaczDczrxBcykqxBEyknxBCxBBuDcxBquDExBnuDCuDBobcuDqobEuDnobCobBcjcobqcjEobncjCcjBcjqcjnxAoykfxAmxAluBoxAvuBmuBloDouBvoDmoDlcbooDvcbmcblxAexAduAuuAtoBuoBtwpAyeszFiwokyegzFbwocyeawoEyeDwoCwoBthAwqsyfitgkwqgyfbtgcwqatgEwqDtgCtgBmxAtiswrimwktigwrbmwctiamwEtiDmwCmwBFxAmystjiFwkmygtjbFwcmyaFwEmyDFwCFysmziFygmzbFyaFyDFziFzbyukzhghjsyuczhahbwyuEzhDhDyyuCyuBwmkydgzErxqkwmczhrxqcyvaydDxqEwmCxqCwmBxqBtakwngydrviktacwnavicxrawnDviEtaCviCtaBviBmiktbgwnrqykmictb"
160   CodageMC$(1) = CodageMC$(1) & "aqycvjatbDqyEmiCqyCmiBqyBEykmjgtbrhykEycmjahycqzamjDhyEEyChyCEyBEzgmjrhzgEzahzaEzDhzDEzrytczgqgrwytEzgngnyytCglzytBwlcycqxncwlEycnxnEytnxnCwlBxnBtDcwlqvbctDEwlnvbExnnvbCtDBvbBmbctDqqjcmbEtDnqjEvbnqjCmbBqjBEjcmbqgzcEjEmbngzEqjngzCEjBgzBEjqgzqEjngznysozgfgfyysmgdzyslwkoycfxloysvxlmwklxlltBowkvvDotBmvDmtBlvDlmDotBvqbovDvqbmmDlqblEbomDvgjoEbmgjmEblgjlEbvgjvysegFzysdwkexkuwkdxkttAuvButAtvBtmBuqDumBtqDtEDugbuEDtgbtysFwkFxkhtAhvAxmAxqBxwekyFgzCrwecyFaweEyFDweCweBsqkwfgyFrsqcwfasqEwfDsqCsqBliksrgwfrlicsraliEsrDliCliBCykljgsrrCycljaCyEljDCyCCyBCzgljrCzaCzDCzryhczaqarwyhEzananyyhCalzyhBwdcyEqwvcwdEyEnwvEyhnwvCwdBwvBsncwdqtrcsnEwdntrEwvntrCsnBtrBlbcsnqnjclbEsnnnjEtrnnjClbBnjBCjclbqazcCjElbnazEnjnazCCjBazBCjqazqCjnaznzioirsrfyziminwrdzzililyikzygozafafyyxozivivyadzyxmyglitzyxlwcoyEfwtowcmxvoyxvwclxvmwtlxvlslowcvtnoslmvrotnmsllvrmtnlvrllDoslvnbolDmrjonbmlDlrjmnblrjlCbolDvajoCbmizoajmCblizmajlizlCbvajvzieifwrFzzididyiczygeaFzywuy"
170   CodageMC$(1) = CodageMC$(1) & "gdihzywtwcewsuwcdxtuwstxttskutlusktvnutltvntlBunDulBtrbunDtrbtCDuabuCDtijuabtijtziFiFyiEzygFywhwcFwshxsxskhtkxvlxlAxnBxrDxCBxaDxibxiCzwFcyCqwFEyCnwFCwFBsfcwFqsfEwFnsfCsfBkrcsfqkrEsfnkrCkrBBjckrqBjEkrnBjCBjBBjqBjnyaozDfDfyyamDdzyalwEoyCfwhowEmwhmwElwhlsdowEvsvosdmsvmsdlsvlknosdvlroknmlrmknllrlBboknvDjoBbmDjmBblDjlBbvDjvzbebfwnpzzbdbdybczyaeDFzyiuyadbhzyitwEewguwEdwxuwgtwxtscustuscttvustttvtklulnukltnrulntnrtBDuDbuBDtbjuDbtbjtjfsrpyjdwrozjcyjcjzbFbFyzjhjhybEzjgzyaFyihyyxwEFwghwwxxxxschssxttxvvxkkxllxnnxrrxBBxDDxbbxjFwrmzjEyjEjbCzjazjCyjCjjBjwCowCmwClsFowCvsFmsFlkfosFvkfmkflArokfvArmArlArvyDeBpzyDdwCewauwCdwatsEushusEtshtkdukvukdtkvtAnuBruAntBrtzDpDpyDozyDFybhwCFwahwixsEhsgxsxxkcxktxlvxAlxBnxDrxbpwnuzboybojDmzbqzjpsruyjowrujjoijobbmyjqybmjjqjjmwrtjjmijmbbljjnjjlijlbjkrsCusCtkFukFtAfuAftwDhsChsaxkExkhxAdxAvxBuzDuyDujbuwnxjbuibubDtjbvjjusrxijugrxbjuajuDbtijvibtbjvbjtgrwrjtajtDbsrjtrjsqjsnBxjDxiDxbbxgnyrbxabxDDwrbxrbwqbwn"
180   CodageMC$(2) = "pjkurwejApbsunyebkpDwulzeDspByeBwzfcfjkprwzfEfbspnyzfCfDwplzzfBfByyrczfqfrwyrEzfnfnyyrCflzyrBxjcyrqxjEyrnxjCxjBuzcxjquzExjnuzCuzBpzcuzqpzEuznpzCdjAorsufydbkonwudzdDsolydBwokzdAyzdodrsovyzdmdnwotzzdldlydkzynozdvdvyynmdtzynlxboynvxbmxblujoxbvujmujlozoujvozmozlcrkofwuFzcnsodyclwoczckyckjzcucvwohzzctctycszylucxzyltxDuxDtubuubtojuojtcfsoFycdwoEzccyccjzchchycgzykxxBxuDxcFwoCzcEycEjcazcCycCjFjAmrstfyFbkmnwtdzFDsmlyFBwmkzFAyzFoFrsmvyzFmFnwmtzzFlFlyFkzyfozFvFvyyfmFtzyflwroyfvwrmwrltjowrvtjmtjlmzotjvmzmmzlqrkvfwxpzhbAqnsvdyhDkqlwvczhBsqkyhAwqkjhAiErkmfwtFzhrkEnsmdyhnsqtymczhlwEkyhkyEkjhkjzEuEvwmhzzhuzEthvwEtyzhthtyEszhszyduExzyvuydthxzyvtwnuxruwntxrttbuvjutbtvjtmjumjtgrAqfsvFygnkqdwvEzglsqcygkwqcjgkigkbEfsmFygvsEdwmEzgtwqgzgsyEcjgsjzEhEhyzgxgxyEgzgwzycxytxwlxxnxtDxvbxmbxgfkqFwvCzgdsqEygcwqEjgcigcbEFwmCzghwEEyggyEEjggjEazgizgFsqCygEwqCjgEigEbECygayECjgajgCwqBjgCigCbEBjgDjgBigBbCrklfwspzCnsldyClwlczCkyCkjzCuCvwlhzzCtCtyCszyFuCx"
190   CodageMC$(2) = CodageMC$(2) & "zyFtwfuwftsrusrtljuljtarAnfstpyankndwtozalsncyakwncjakiakbCfslFyavsCdwlEzatwngzasyCcjasjzChChyzaxaxyCgzawzyExyhxwdxwvxsnxtrxlbxrfkvpwxuzinArdsvoyilkrcwvojiksrciikgrcbikaafknFwtmzivkadsnEyitsrgynEjiswaciisiacbisbCFwlCzahwCEyixwagyCEjiwyagjiwjCazaiziyzifArFsvmyidkrEwvmjicsrEiicgrEbicaicDaFsnCyihsaEwnCjigwrajigiaEbigbCCyaayCCjiiyaajiijiFkrCwvljiEsrCiiEgrCbiEaiEDaCwnBjiawaCiiaiaCbiabCBjaDjibjiCsrBiiCgrBbiCaiCDaBiiDiaBbiDbiBgrAriBaiBDaAriBriAqiAnBfskpyBdwkozBcyBcjBhyBgzyCxwFxsfxkrxDfklpwsuzDdsloyDcwlojDciDcbBFwkmzDhwBEyDgyBEjDgjBazDizbfAnpstuybdknowtujbcsnoibcgnobbcabcDDFslmybhsDEwlmjbgwDEibgiDEbbgbBCyDayBCjbiyDajbijrpkvuwxxjjdArosvuijckrogvubjccroajcEroDjcCbFknmwttjjhkbEsnmijgsrqinmbjggbEajgabEDjgDDCwlljbawDCijiwbaiDCbjiibabjibBBjDDjbbjjjjjFArmsvtijEkrmgvtbjEcrmajEErmDjECjEBbCsnlijasbCgnlbjagrnbjaabCDjaDDBibDiDBbjbibDbjbbjCkrlgvsrjCcrlajCErlDjCCjCBbBgnkrjDgbBajDabBDjDDDArbBrjDrjBcrkqjBErknjBCjBBbAqjBqbAnjBnjAorkfjAmjAlb"
200   CodageMC$(2) = CodageMC$(2) & "AfjAvApwkezAoyAojAqzBpskuyBowkujBoiBobAmyBqyAmjBqjDpkluwsxjDosluiDoglubDoaDoDBmwktjDqwBmiDqiBmbDqbAljBnjDrjbpAnustxiboknugtxbbocnuaboEnuDboCboBDmsltibqsDmgltbbqgnvbbqaDmDbqDBliDniBlbbriDnbbrbrukvxgxyrrucvxaruEvxDruCruBbmkntgtwrjqkbmcntajqcrvantDjqEbmCjqCbmBjqBDlglsrbngDlajrgbnaDlDjrabnDjrDBkrDlrbnrjrrrtcvwqrtEvwnrtCrtBblcnsqjncblEnsnjnErtnjnCblBjnBDkqblqDknjnqblnjnnrsovwfrsmrslbkonsfjlobkmjlmbkljllDkfbkvjlvrsersdbkejkubkdjktAeyAejAuwkhjAuiAubAdjAvjBuskxiBugkxbBuaBuDAtiBviAtbBvbDuklxgsyrDuclxaDuElxDDuCDuBBtgkwrDvglxrDvaBtDDvDAsrBtrDvrnxctyqnxEtynnxCnxBDtclwqbvcnxqlwnbvEDtCbvCDtBbvBBsqDtqBsnbvqDtnbvnvyoxzfvymvylnwotyfrxonwmrxmnwlrxlDsolwfbtoDsmjvobtmDsljvmbtljvlBsfDsvbtvjvvvyevydnwerwunwdrwtDsebsuDsdjtubstjttvyFnwFrwhDsFbshjsxAhiAhbAxgkirAxaAxDAgrAxrBxckyqBxEkynBxCBxBAwqBxqAwnBxnlyoszflymlylBwokyfDxolyvDxmBwlDxlAwfBwvDxvtzetzdlyenyulydnytBweDwuBwdbxuDwtbxttzFlyFnyhBwFDwhbwxAiqAinAyokjfAymAylAifAyvkzekzdAyeByuAydBytszp"
210   CodeErr% = 0
220   If Chaine$ = "" Then CodeErr% = 1: Exit Function
          'Split the string in character blocks of the same type : numeric , text, byte
          'The first column of the array Liste% contain the char. number, the second one contain the mode switch
230   IndexChaine% = 1
240   GoSub QuelMode
250   Do
260     ReDim Preserve Liste%(1, IndexListe%)
270     Liste%(1, IndexListe%) = mode%
280     Do While Liste%(1, IndexListe%) = mode%
290         Liste%(0, IndexListe%) = Liste%(0, IndexListe%) + 1
300         IndexChaine% = IndexChaine% + 1
310         If IndexChaine% > Len(Chaine$) Then Exit Do
320         GoSub QuelMode
330     Loop
340     IndexListe% = IndexListe% + 1
350   Loop Until IndexChaine% > Len(Chaine$)
          'We retain "numeric" mode only if it's earning, else "text" mode or even "byte" mode
          'The efficiency limits have been pre-defined according to the previous mode and/or the next mode.
360   For i% = 0 To IndexListe% - 1
370     If Liste%(1, i%) = 902 Then
380         If i% = 0 Then    'C'est le premier bloc / It's the first block
390             If IndexListe% > 1 Then    'et il y en a d'autres derrière / And there is other blocks behind
400                 If Liste%(1, i% + 1) = 900 Then
                        'Premier bloc et suivi par un bloc de type "texte" / First block and followed by a "text" type block
410                     If Liste%(0, i%) < 8 Then Liste%(1, i%) = 900
420                 ElseIf Liste%(1, i% + 1) = 901 Then
                        'Premier bloc et suivi par un bloc de type "octet" / First block and followed by a "byte" type block
430                     If Liste%(0, i%) = 1 Then Liste%(1, i%) = 901
440                 End If
450             End If
460         Else
                'It's not the first block
470             If i% = IndexListe% - 1 Then
                    'It's the last one
480                 If Liste%(1, i% - 1) = 900 Then
                        'It's  preceded by a "text" type block
490                     If Liste%(0, i%) < 7 Then Liste%(1, i%) = 900
500                 ElseIf Liste%(1, i% - 1) = 901 Then
                        'It's  preceded by a "byte" type block
510                     If Liste%(0, i%) = 1 Then Liste%(1, i%) = 901
520                 End If
530             Else
                    'It's not the last block
540                 If Liste%(1, i% - 1) = 901 And Liste%(1, i% + 1) = 901 Then
                        'Framed by "byte" type blocks
550                     If Liste%(0, i%) < 4 Then Liste%(1, i%) = 901
560                 ElseIf Liste%(1, i% - 1) = 900 And Liste%(1, i% + 1) = 901 Then
                        'Preceded by "text" and followed by "byte" (If the reverse it's never interesting to change)
570                     If Liste%(0, i%) < 5 Then Liste%(1, i%) = 900
580                 ElseIf Liste%(1, i% - 1) = 900 And Liste%(1, i% + 1) = 900 Then
                        'Framed by "text" type blocks
590                     If Liste%(0, i%) < 8 Then Liste%(1, i%) = 900
600                 End If
610             End If
620         End If
630     End If
640   Next
650   GoSub Regroupe
          'Maintain "text" mode only if it's earning
660   For i% = 0 To IndexListe% - 1
670     If Liste%(1, i%) = 900 And i% > 0 Then
            'It's not the first (If first, never interesting to change)
680         If i% = IndexListe% - 1 Then    ' It's the last one
690             If Liste%(1, i% - 1) = 901 Then
                    'It's  preceded by a "byte" type block
700                 If Liste%(0, i%) = 1 Then Liste%(1, i%) = 901
710             End If
720         Else
                ' It's not the last one
730             If Liste%(1, i% - 1) = 901 And Liste%(1, i% + 1) = 901 Then
                    ' Framed by "byte" type blocks
740                 If Liste%(0, i%) < 5 Then Liste%(1, i%) = 901
750             ElseIf (Liste%(1, i% - 1) = 901 And Liste%(1, i% + 1) <> 901) Or (Liste%(1, i% - 1) <> 901 And Liste%(1, i% + 1) = 901) Then
                    ' A "byte" block ahead or behind
760                 If Liste%(0, i%) < 3 Then Liste%(1, i%) = 901
770             End If
780         End If
790     End If
800   Next
810   GoSub Regroupe
          'Now we compress datas into the MCs, the MCs are stored in 3 char.
          'in a large string : ChaineMC$
820   IndexChaine% = 1
830   For i% = 0 To IndexListe% - 1
        'Donc 3 modes de compactage / Thus 3 compaction modes
840     Select Case Liste%(1, i%)
        Case 900    'Texte
850         ReDim ListeT%(1, Liste%(0, i%))
            'ListeT% will contain the table number(s) (1 ou several) and the value of each char.
            'Table number encoded in the 4 less weight bits, that is in decimal 1, 2, 4, 8
860         For IndexListeT% = 0 To Liste%(0, i%) - 1
870             GoSub QuelleTable
880             ListeT%(0, IndexListeT%) = Table%
890             ListeT%(1, IndexListeT%) = CodeASCII%
900         Next
910         CurTable% = 1    'Table par défaut / Default table
920         ChaineT$ = ""
            'Les données sont stockées sur 2 car. dans la chaine TableT$ / Datas are stored in 2 char. in the string TableT$
930         For j% = 0 To Liste%(0, i%) - 1
940             If (ListeT%(0, j%) And CurTable%) > 0 Then
                    'Le car. est dans la table courante / The char. is in the current table
950                 ChaineT$ = ChaineT$ & Format(ListeT%(1, j%), "00")
960             Else
                    'Faut changer de table / Obliged to change the table
970                 Flag = False    'True si on change de table pour un seul car. / True if we change the table only for 1 char.
980                 If j% = Liste%(0, i%) - 1 Then
990                     Flag = True
1000                Else
1010                    If (ListeT%(0, j%) And ListeT%(0, j% + 1)) = 0 Then Flag = True    'Pas de table commune avec le car. suivant / No common table with the next char.
1020                End If
1030                If Flag Then
                        'On change de table pour 1 seul car., Chercher un commutateur fugitif
                        'We change only for 1 char., Look for a temporary switch
1040                    If (ListeT%(0, j%) And 1) > 0 And CurTable% = 2 Then
                            'Table 2 vers 1 pour 1 car. --> T_MAJ / Table 2 to 1 for 1 char. --> T_UPP
1050                        ChaineT$ = ChaineT$ & "27" & Format(ListeT%(1, j%), "00")
1060                    ElseIf (ListeT%(0, j%) And 8) > 0 Then
                            'Table 1 ou 2 ou 4 vers table 8 pour 1 car. --> T_PON / Table 1 or 2 or 4 to table 8 for 1 char. --> T_PUN
1070                        ChaineT$ = ChaineT$ & "29" & Format(ListeT%(1, j%), "00")
1080                    Else
                            'Pas de commutateur fugitif / No temporary switch available
1090                        Flag = False
1100                    End If
1110                End If
1120                If Not Flag Then    'On re-teste flag qui a peut-être changé ci-dessus ! donc ELSE pas possible / We test again flag which is perhaps changed ! Impossible tio use ELSE statement
                        '
                        'We must use a bi-state switch
                        'Looking for the new table to use
1130                    If j% = Liste%(0, i%) - 1 Then
1140                        NewTable% = ListeT%(0, j%)
1150                    Else
1160                        NewTable% = ListeT%(0, j%) And ListeT%(0, j%)
1170                    End If
                        'Maintain the first if several tables are possible
1180                    Select Case NewTable%
                        Case 3, 5, 7, 9, 11, 13, 15
1190                        NewTable% = 1
1200                    Case 6, 10, 14
1210                        NewTable% = 2
1220                    Case 12
1230                        NewTable% = 4
1240                    End Select
                        'Select the switch, on occasion we must use 2 switchs consecutively
1250                    Select Case CurTable%
                        Case 1
1260                        Select Case NewTable%
                            Case 2
1270                            ChaineT$ = ChaineT$ & "27"
1280                        Case 4
1290                            ChaineT$ = ChaineT$ & "28"
1300                        Case 8
1310                            ChaineT$ = ChaineT$ & "2825"
1320                        End Select
1330                    Case 2
1340                        Select Case NewTable%
                            Case 1
1350                            ChaineT$ = ChaineT$ & "2828"
1360                        Case 4
1370                            ChaineT$ = ChaineT$ & "28"
1380                        Case 8
1390                            ChaineT$ = ChaineT$ & "2825"
1400                        End Select
1410                    Case 4
1420                        Select Case NewTable%
                            Case 1
1430                            ChaineT$ = ChaineT$ & "28"
1440                        Case 2
1450                            ChaineT$ = ChaineT$ & "27"
1460                        Case 8
1470                            ChaineT$ = ChaineT$ & "25"
1480                        End Select
1490                    Case 8
1500                        Select Case NewTable%
                            Case 1
1510                            ChaineT$ = ChaineT$ & "29"
1520                        Case 2
1530                            ChaineT$ = ChaineT$ & "2927"
1540                        Case 4
1550                            ChaineT$ = ChaineT$ & "2928"
1560                        End Select
1570                    End Select
1580                    CurTable% = NewTable%
1590                    ChaineT$ = ChaineT$ & Format(ListeT%(1, j%), "00")    'On ajoute enfin le car. / At last we add the char.
1600                End If
1610            End If
1620        Next
1630        If Len(ChaineT$) Mod 4 > 0 Then ChaineT$ = ChaineT$ & "29"    'Bourrage si nb de car. impair / Padding if number of char. is odd
            'Now translate the string ChaineT$ into CWs
1640        If i% > 0 Then ChaineMC$ = ChaineMC$ & "900"    'Set up the switch exept for the first block because "text" is the default
1650        For j% = 1 To Len(ChaineT$) Step 4
1660            ChaineMC$ = ChaineMC$ & Format(Mid$(ChaineT$, j%, 2) * 30 + Mid$(ChaineT$, j% + 2, 2), "000")
1670        Next
1680    Case 901    'Octet
            ' Select the switch between the 3 possible
1690        If Liste%(0, i%) = 1 Then
                '1 seul octet, c'est immédiat
1700            ChaineMC$ = ChaineMC$ & "913" & Format(Asc(Mid$(Chaine$, IndexChaine%, 1)), "000")
1710        Else
                'Select the switch for perfect multiple of 6 bytes or no
1720            If Liste%(0, i%) Mod 6 = 0 Then
1730                ChaineMC$ = ChaineMC$ & "924"
1740            Else
1750                ChaineMC$ = ChaineMC$ & "901"
1760            End If
1770            j% = 0
1780            Do While j% < Liste%(0, i%)
1790                Longueur% = Liste%(0, i%) - j%
1800                If Longueur% >= 6 Then
                        'Take groups of 6
1810                    Longueur% = 6
1820                    total = 0
1830                    For K% = 0 To Longueur% - 1
1840                        total = total + (Asc(Mid$(Chaine$, IndexChaine% + j% + K%, 1)) * 256 ^ (Longueur% - 1 - K%))
1850                    Next
1860                    ChaineMod$ = Format(total, "general number")
1870                    Dummy$ = ""
1880                    Do
1890                        Diviseur& = 900
1900                        GoSub Modulo
1910                        Dummy$ = Format(Diviseur&, "000") & Dummy$
1920                        ChaineMod$ = ChaineMult$
1930                        If ChaineMult$ = "0" Then Exit Do
1940                    Loop
1950                    ChaineMC$ = ChaineMC$ & Dummy$
1960                Else
                        ' If it remain a group of less than 6 bytes
1970                    For K% = 0 To Longueur% - 1
1980                        ChaineMC$ = ChaineMC$ & Format(Asc(Mid$(Chaine$, IndexChaine% + j% + K%, 1)), "000")
1990                    Next
2000                End If
2010                j% = j% + Longueur%
2020            Loop
2030        End If
2040    Case 902    ' Numeric
2050        ChaineMC$ = ChaineMC$ & "902"
2060        j% = 0
2070        Do While j% < Liste%(0, i%)
2080            Longueur% = Liste%(0, i%) - j%
2090            If Longueur% > 44 Then Longueur% = 44
2100            ChaineMod$ = "1" & Mid$(Chaine$, IndexChaine% + j%, Longueur%)
2110            Dummy$ = ""
2120            Do
2130                Diviseur& = 900
2140                GoSub Modulo
2150                Dummy$ = Format(Diviseur&, "000") & Dummy$
2160                ChaineMod$ = ChaineMult$
2170                If ChaineMult$ = "0" Then Exit Do
2180            Loop
2190            ChaineMC$ = ChaineMC$ & Dummy$
2200            j% = j% + Longueur%
2210        Loop
2220    End Select
2230    IndexChaine% = IndexChaine% + Liste%(0, i%)
2240  Next
          'ChaineMC$ contain the MC list (on 3 digits) depicting the datas
          'Now we take care of the correction level
2250  Longueur% = Len(ChaineMC$) / 3
2260  If sécu% < 0 Then
        'Fixing auto. the correction level according to the standard recommendations
2270    If Longueur% < 41 Then
2280        sécu% = 2
2290    ElseIf Longueur% < 161 Then
2300        sécu% = 3
2310    ElseIf Longueur% < 321 Then
2320        sécu% = 4
2330    Else
2340        sécu% = 5
2350    End If
2360  End If
          'On s'occupe maintenant du nombre de MC par ligne / Now we take care of the number of CW per row
2370  Longueur% = Longueur% + 1 + (2 ^ (sécu% + 1))
2380  If nbcol% > 30 Then nbcol% = 30
2390  If nbcol% < 1 Then
        '
        'With a 3 modules high font, for getting a "square" bar code
        'x = nb. of col. | Width by module = 69 + 17x | Height by module = 3t / x (t is the total number of MCs)
        'Thus we have 69 + 17x = 3t/x <=> 17x²+69x-3t=0 - Discriminant is 69²-4*17*-3t = 4761+204t thus x=SQR(discr.)-69/2*17
2400    nbcol% = (Sqr(204# * Longueur% + 4761) - 69) / (34 / 1.3)   '1.3 = coeff. de pondération déterminé au pif après essais / 1.3 = balancing factor determined at a guess after tests
2410    If nbcol% = 0 Then nbcol% = 1
2420  End If
          'If we go beyong 928 CWs we try to reduce the correction level
2430  Do While sécu% > 0
        'Calculation of the total number of CW with the padding
2440    Longueur% = Len(ChaineMC$) / 3 + 1 + (2 ^ (sécu% + 1))
2450    Longueur% = (Longueur% \ nbcol% + IIf(Longueur% Mod nbcol% > 0, 1, 0)) * nbcol%
2460    If Longueur% < 929 Then Exit Do
        'We must reduce security level
2470    sécu% = sécu% - 1
2480    CodeErr% = 10
2490  Loop
2500  If Longueur% > 928 Then CodeErr% = 2: Exit Function
2510  If Longueur% / nbcol% > 90 Then CodeErr% = 3: Exit Function
          'Calcul du rembourrage / Padding calculation
2520  Longueur% = Len(ChaineMC$) / 3 + 1 + (2 ^ (sécu% + 1))
2530  i% = 0
2540  If Longueur% \ nbcol% < 3 Then
2550    i% = nbcol% * 3 - Longueur%   'Il faut au moins 3 lignes dans le code / A bar code must have at least 3 row
2560  Else
2570    If Longueur% Mod nbcol% > 0 Then i% = nbcol% - (Longueur% Mod nbcol%)
2580  End If
          'On ajoute le rembourrage / We add the padding
2590  Do While i% > 0
2600    ChaineMC$ = ChaineMC$ & "900"
2610    i% = i% - 1
2620  Loop
          'On ajoute le descripteur de longueur / We add the length descriptor
2630  ChaineMC$ = Format(Len(ChaineMC$) / 3 + 1, "000") & ChaineMC$
          'On s'occupe maintenant des codes de Reed Solomon / Now we take care of the Reed Solomon codes
2640  Longueur% = Len(ChaineMC$) / 3
2650  K% = 2 ^ (sécu% + 1)
2660  ReDim MCcorrection%(K% - 1)
2670  total = 0
2680  For i% = 0 To Longueur% - 1
2690    total = (Mid$(ChaineMC$, i% * 3 + 1, 3) + MCcorrection%(K% - 1)) Mod 929
2700    For j% = K% - 1 To 0 Step -1
2710        If j% = 0 Then
2720            MCcorrection%(j%) = (929 - (total * Mid$(CoefRS$(sécu%), j% * 3 + 1, 3)) Mod 929) Mod 929
2730        Else
2740            MCcorrection%(j%) = (MCcorrection%(j% - 1) + 929 - (total * Mid$(CoefRS$(sécu%), j% * 3 + 1, 3)) Mod 929) Mod 929
2750        End If
2760    Next
2770  Next
2780  For j% = 0 To K% - 1
2790    If MCcorrection%(j%) <> 0 Then MCcorrection%(j%) = 929 - MCcorrection%(j%)
2800  Next
          'On va ajouter les codes de correction à la chaine / We add theses codes to the string
2810  For i% = K% - 1 To 0 Step -1
2820    ChaineMC$ = ChaineMC$ & Format(MCcorrection%(i%), "000")
2830  Next
          'La chaine des MC est terminée
          'Calcul des paramètres pour les MC de cotés gauche et droit
          'The MC string is finished
          'Calculation of parameters for the left and right side MCs
2840  C1% = (Len(ChaineMC$) / 3 / nbcol% - 1) \ 3
2850  C2% = sécu% * 3 + (Len(ChaineMC$) / 3 - 1) Mod 3
2860  C3% = nbcol% - 1
          'On encode chaque ligne / We encode each row
2870  For i% = 0 To Len(ChaineMC$) / 3 / nbcol% - 1
2880    Dummy$ = Mid$(ChaineMC$, i% * nbcol% * 3 + 1, nbcol% * 3)
2890    K% = (i% \ 3) * 30
2900    Select Case i% Mod 3
        Case 0
2910        Dummy$ = Format(K% + C1%, "000") & Dummy$ & Format(K% + C3%, "000")
2920    Case 1
2930        Dummy$ = Format(K% + C2%, "000") & Dummy$ & Format(K% + C1%, "000")
2940    Case 2
2950        Dummy$ = Format(K% + C3%, "000") & Dummy$ & Format(K% + C2%, "000")
2960    End Select
2970    PDF417$ = PDF417$ & "+*"    'Commencer par car. de start et séparateur / Start with a start char. and a separator
2980    For j% = 0 To Len(Dummy$) / 3 - 1
2990        PDF417$ = PDF417$ & Mid$(CodageMC$(i% Mod 3), Mid$(Dummy$, j% * 3 + 1, 3) * 3 + 1, 3) & "*"
3000    Next
3010    PDF417$ = PDF417$ & "-" & Chr$(13) & Chr$(10)    'Ajouter car. de stop et CRLF / Add a stop char. and a CRLF
3020  Next
3030  Exit Function
Regroupe:
          'Regrouper les blocs de même type / Bring together same type blocks
3040  If IndexListe% > 1 Then
3050    i% = 1
3060    Do While i% < IndexListe%
3070        If Liste%(1, i% - 1) = Liste%(1, i%) Then
                'Regroupement / Bringing together
3080            Liste%(0, i% - 1) = Liste%(0, i% - 1) + Liste%(0, i%)
3090            j% = i% + 1
                'Réduction de la liste / Decrease the list
3100            Do While j% < IndexListe%
3110                Liste%(0, j% - 1) = Liste%(0, j%)
3120                Liste%(1, j% - 1) = Liste%(1, j%)
3130                j% = j% + 1
3140            Loop
3150            IndexListe% = IndexListe% - 1
3160            i% = i% - 1
3170        End If
3180        i% = i% + 1
3190    Loop
3200  End If
3210  Return
QuelMode:
3220  CodeASCII% = Asc(Mid$(Chaine$, IndexChaine%, 1))
3230  Select Case CodeASCII%
          Case 48 To 57
3240    mode% = 902
3250  Case 9, 10, 13, 32 To 126
3260    mode% = 900
3270  Case Else
3280    mode% = 901
3290  End Select
3300  Return
QuelleTable:
3310  CodeASCII% = Asc(Mid$(Chaine$, IndexChaine% + IndexListeT%, 1))
3320  Select Case CodeASCII%
          Case 9    'HT
3330    Table% = 12
3340    CodeASCII% = 12
3350  Case 10    'LF
3360    Table% = 12
3370    CodeASCII% = 15
3380  Case 13    'CR
3390    Table% = 12
3400    CodeASCII% = 11
3410  Case Else
3420    Table% = Mid$(ASCII$, CodeASCII% * 4 - 127, 2)
3430    CodeASCII% = Mid$(ASCII$, CodeASCII% * 4 - 125, 2)
3440  End Select
3450  Return
Modulo:
          'ChaineMod$ représente un très grand nombre sur plus de 9 chiffres
          'Diviseur& est le diviseur, contient le résultat au retour
          'ChaineMult$ contient au retour le résultat de la division entière
          '
          'ChaineMod$ depict a very large number having more than de 9 digits
          'Diviseur& is the divisor, contain the result after return
          'ChaineMult$ contain after return the result of the integer division
3460  ChaineMult$ = ""
3470  Do
3480    Nombre& = Val(Left$(ChaineMod$, 9))
3490    If Len(ChaineMod$) > 9 Then ChaineMod$ = Mid$(ChaineMod$, 10) Else ChaineMod$ = ""
3500    ChaineMod$ = Format(Nombre& Mod Diviseur&, "general number") & ChaineMod$
3510    ChaineMult$ = ChaineMult$ & Format(Nombre& \ Diviseur&, "general number")
3520    If Len(ChaineMod$) <= 9 Then
3530        If Val(ChaineMod$) < Diviseur& Then Exit Do
3540    End If
3550  Loop
3560  Diviseur& = Val(ChaineMod$)
3570  Return
End Function

Public Sub PrintBarCodesN(ByVal DataToPrint As String, _
                          ByVal HowMany As Integer, _
                          ByVal Name As String, _
                          ByVal Chart As String, _
                          ByVal DoB As String, _
                          ByVal Group As String)
          Dim Px As Printer
          Dim strOriginalPrinter As String
          Dim n As Integer

10    On Error Resume Next

20    strOriginalPrinter = Printer.DeviceName

30    If Not SetLabelPrinter() Then
40      Exit Sub
50    End If

60    For n = 1 To HowMany

70      Printer.Orientation = vbPRORPortrait
80      Printer.Font.Name = "Code 128"
90      Printer.CurrentX = 0
100     Printer.CurrentY = 100
110     Printer.CurrentX = 200
120     Printer.Font.Size = 26
130     Printer.Font.Bold = False
140     Printer.Print Code128(DataToPrint)
150     Printer.Font.Bold = True
160     Printer.Font.Size = 12
170     Printer.Font.Name = "Courier New"
180     Printer.Print "  "; DataToPrint
190     If Trim$(Name) <> "" Then
200         Printer.Font.Size = 10
210         Printer.Print Name
220         If Trim$(Chart) <> "" Then
230             Printer.Print "Ch:"; Chart;
240         End If
250         If Trim$(DoB) <> "" Then
260             Printer.Print " DoB:"; DoB;
270         End If
280     ElseIf Trim$(Group) <> "" Then
290         Printer.Font.Size = 16
300         Printer.Print "   Group:"; Group;
310     End If

320     Printer.EndDoc

330   Next

340   For Each Px In Printers
350     If Px.DeviceName = strOriginalPrinter Then
360         Set Printer = Px
370         Exit For
380     End If
390   Next

End Sub


