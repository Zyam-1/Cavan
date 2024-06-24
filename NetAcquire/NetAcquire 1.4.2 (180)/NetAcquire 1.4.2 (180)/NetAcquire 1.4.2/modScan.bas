Attribute VB_Name = "modScan"
Option Explicit

Public Sub Scan(ByVal SampleID As String, _
                ByVal PixelType As Integer, _
                ByVal Resolution As Integer)

          Dim X As Long
          Dim FilePath As String
          Dim n As Integer
          Dim ScannedName As String
          Dim tb As Recordset
          Dim sql As String
          Dim src() As Byte
          Dim a() As Byte
          Static Source As Long
          Dim FileNum As Integer

20080     On Error GoTo Scan_Error

20090     SampleID = Replace(SampleID, "/", "-")

20100     If Trim$(SampleID) = "" Then
20110         iMsg SampleID & "?", vbExclamation
20120         Exit Sub
20130     End If
20140     If Source = 0 Then
20150         X = TWAIN_SelectImageSource(frmScan.hWnd)
20160         Source = X
20170     End If

20180     X = TWAIN_OpenDefaultSource()
20190     X = TWAIN_State()
20200     If X <> 4 Then
20210         iMsg "Error", vbExclamation
20220         Exit Sub
20230     End If

20240     X = TWAIN_SetCurrentResolution(Resolution)
20250     X = TWAIN_GetCurrentResolution()
20260     If X <> Resolution Then
20270         iMsg "Errir", vbExclamation
20280     End If

20290     X = TWAIN_NegotiatePixelTypes(PixelType)

20300     FilePath = GetOptionSetting("ScanPath", "")

20310     For n = 0 To 25
20320         ScannedName = SampleID & Chr$(Asc("A") + n) & ".jpg"
20330         sql = "SELECT SampleID FROM ScannedImages WHERE " & _
                    "ScannedName = '" & ScannedName & "'"
20340         Set tb = New Recordset
20350         RecOpenClient 0, tb, sql
20360         If tb.EOF Then
20370             Exit For
20380         End If
20390     Next

20400     X = TWAIN_GetHideUI()
20410     TWAIN_SetHideUI 1
20420     X = TWAIN_GetHideUI()

20430     X = TWAIN_AcquireToFilename(frmScan.hWnd, FilePath & ScannedName & ".jpg")

          
20440     Close #1
20450     Open FilePath & ScannedName & ".jpg" For Binary As #1
20460     ReDim src(0 To LOF(1) - 1)
20470     Get #1, , src
20480     Close #1

          'a = Compress(src)

20490     sql = "SELECT * FROM ScannedImages WHERE 0 = 1"
20500     Set tb = New Recordset
20510     RecOpenClient 0, tb, sql
20520     tb.AddNew
20530     tb!ScannedImage = src
20540     tb!SampleID = SampleID
20550     tb!ScannedName = ScannedName
20560     tb.Update

20570     Kill FilePath & "*.jpg"

20580     Exit Sub

Scan_Error:

          Dim strES As String
          Dim intEL As Integer

20590     intEL = Erl
20600     strES = Err.Description
20610     LogError "modScan", "Scan", intEL, strES, sql, FilePath

End Sub




Public Sub SetViewScans(ByVal SampleID As String, _
                        ByVal cmdButton As CommandButton)

          Dim sql As String
          Dim tb As Recordset

20620     On Error GoTo SetViewScans_Error

20630     sql = "SELECT TOP 1 * FROM ScannedImages WHERE SampleID = '" & SampleID & "'"
20640     Set tb = New Recordset
20650     RecOpenClient 0, tb, sql
20660     cmdButton.Visible = Not tb.EOF

20670     Exit Sub

SetViewScans_Error:

          Dim strES As String
          Dim intEL As Integer

20680     intEL = Erl
20690     strES = Err.Description
20700     LogError "modScan", "SetViewScans", intEL, strES, sql


End Sub



