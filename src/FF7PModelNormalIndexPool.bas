Attribute VB_Name = "FF7PModelNormalIndexPool"
Option Explicit
Sub ReadNormalIndex(ByVal NFile As Integer, ByVal offset As Long, ByRef NormalIndex() As Long, ByVal NumNormalIndex As Long)
    ReDim NormalIndex(NumNormalIndex - 1)
    Get NFile, offset, NormalIndex
End Sub
Sub WriteNormalIndex(ByVal NFile As Integer, ByVal offset As Long, ByRef NormalIndex() As Long)
    Put NFile, offset, NormalIndex
End Sub
Sub MergeNormalIndex(ByRef ni1() As Long, ByRef ni2() As Long)
    Dim NumNormalIndexNI1 As Integer
    Dim NumNormalIndexNI2 As Integer
    
    NumNormalIndexNI1 = UBound(ni1) + 1
    NumNormalIndexNI2 = UBound(ni2) + 1
    ReDim Preserve ni1(NumNormalIndexNI1 + NumNormalIndexNI2 - 1)
    
    CopyMemory ni1(NumNormalIndexNI1), ni2(0), NumNormalIndexNI2 * 12
End Sub
