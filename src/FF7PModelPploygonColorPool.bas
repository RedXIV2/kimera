Attribute VB_Name = "FF7PModelPolygonColorPool"
Option Explicit
Sub ReadPColors(ByVal NFile As Integer, ByVal offset As Long, ByRef PColors() As color, ByVal NumPColors As Long)
    ReDim PColors(NumPColors - 1)
    Get NFile, offset, PColors
End Sub
Sub WritePColors(ByVal NFile As Integer, ByVal offset As Long, ByRef PColors() As color)
    Put NFile, offset, PColors
End Sub
Sub MergePColors(ByRef pc1() As color, ByRef pc2() As color)
    Dim NumPColorsPC1 As Integer
    Dim NumPColorsPC2 As Integer

    NumPColorsPC1 = UBound(pc1) + 1
    NumPColorsPC2 = UBound(pc2) + 1
    ReDim Preserve pc1(NumPColorsPC1 + NumPColorsPC2 - 1)

    CopyMemory pc1(NumPColorsPC1), pc2(0), NumPColorsPC2 * 4
End Sub
