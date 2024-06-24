Attribute VB_Name = "modCompression"
Option Explicit

Private Declare Function qlz_compress Lib "quick32.dll" (ByRef Source As Byte, ByRef Destination As Byte, ByVal Length As Long) As Long
Private Declare Function qlz_decompress Lib "quick32.dll" (ByRef Source As Byte, ByRef Destination As Byte) As Long
Private Declare Function qlz_size_decompressed Lib "quick32.dll" (ByRef Source As Byte) As Long
Private Declare Function qlz_size_source Lib "quick32.dll" (ByRef Source As Byte) As Long

' If the Visual Basic IDE cannot find quick32.dll even though it's in the system32 directory,
' try adding a path to the quick32.dll file name in the declarations.
' This should never be neccessary though.

Function Compress(Source() As Byte) As Byte()
          Dim dst() As Byte
          Dim R As Long
60140     ReDim dst(0 To UBound(Source) * 1.2 + 36000)
60150     R = qlz_compress(Source(0), dst(0), UBound(Source) + 1)
60160     ReDim Preserve dst(0 To R - 1)
60170     Compress = dst
End Function

Public Function GetSize(Source() As Byte) As Long
60180     GetSize = qlz_size_decompressed(Source(0))
End Function

Public Function Decompress(Source() As Byte) As Byte()
          Dim dst() As Byte
          Dim R As Long
          Dim size As Long
60190     size = GetSize(Source)
60200     If size < 20 * 1000000 Then    ' Visual Basic can crash if you allocate too long strings
60210         ReDim dst(0 To size - 1)
60220         R = qlz_decompress(Source(0), dst(0))
60230         ReDim Preserve dst(0 To R - 1)
60240         Decompress = dst
60250     End If
End Function





