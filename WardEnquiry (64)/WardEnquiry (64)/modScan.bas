Attribute VB_Name = "modScan"
Option Explicit

'Public Sub Scan(ByVal SampleID As String, _
'                ByVal PixelType As Integer, _
'                ByVal Resolution As Integer)
'
'    Dim X As Long
'    Dim FilePath As String
'    Dim n As Integer
'    Dim ScannedName As String
'    Dim tb As Recordset
'    Dim sql As String
'    Dim src() As Byte
'    Dim a() As Byte
'    Static Source As Long
'    Dim FileNum As Integer
'
'    On Error GoTo Scan_Error
'
'    SampleID = Replace(SampleID, "/", "-")
'
'    If Trim$(SampleID) = "" Then
'        iMsg SampleID & "?", vbExclamation
'        Exit Sub
'    End If
'    If Source = 0 Then
'        X = TWAIN_SelectImageSource(frmScan.hwnd)
'        Source = X
'    End If
'
'    X = TWAIN_OpenDefaultSource()
'    X = TWAIN_State()
'    If X <> 4 Then
'        iMsg "Error", vbExclamation
'        Exit Sub
'    End If
'
'    X = TWAIN_SetCurrentResolution(Resolution)
'    X = TWAIN_GetCurrentResolution()
'    If X <> Resolution Then
'        iMsg "Errir", vbExclamation
'    End If
'
'    X = TWAIN_NegotiatePixelTypes(PixelType)
'
'    FilePath = GetOptionSetting("ScanPath", "")
'
'    For n = 0 To 25
'        ScannedName = SampleID & Chr$(Asc("A") + n) & ".jpg"
'        sql = "SELECT SampleID FROM ScannedImages WHERE " & _
'              "ScannedName = '" & ScannedName & "'"
'        Set tb = New Recordset
'        RecOpenClient 0, tb, sql
'        If tb.EOF Then
'            Exit For
'        End If
'    Next
'
'    X = TWAIN_GetHideUI()
'    TWAIN_SetHideUI 1
'    X = TWAIN_GetHideUI()
'
'    X = TWAIN_AcquireToFilename(frmScan.hwnd, FilePath & ScannedName & ".jpg")
'
'
'    Close #1
'    Open FilePath & ScannedName & ".jpg" For Binary As #1
'    ReDim src(0 To LOF(1) - 1)
'    Get #1, , src
'    Close #1
'
'    'a = Compress(src)
'
'    sql = "SELECT * FROM ScannedImages WHERE 0 = 1"
'    Set tb = New Recordset
'    RecOpenClient 0, tb, sql
'    tb.AddNew
'    tb!ScannedImage = src
'    tb!SampleID = SampleID
'    tb!ScannedName = ScannedName
'    tb.Update
'
'    Kill FilePath & "*.jpg"
'
'    Exit Sub
'
'Scan_Error:
'
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "modScan", "Scan", intEL, strES, sql, FilePath
'
'End Sub




Public Sub SetViewScans(ByVal SampleID As String, _
                        ByVal cmdButton As CommandButton)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo SetViewScans_Error

20        sql = "SELECT TOP 1 * FROM ScannedImages WHERE SampleID = '" & SampleID & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        cmdButton.Visible = Not tb.EOF

60        Exit Sub

SetViewScans_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "modScan", "SetViewScans", intEL, strES, sql


End Sub



