Attribute VB_Name = "Utils"
Option Explicit
Type Point2D
    x As Single
    y As Single
End Type

Type Point3D
    x As Single
    y As Single
    z As Single
End Type

Type color
    B As Byte
    g As Byte
    r As Byte
    a As Byte
End Type

Type order_pair
    d As Single
End Type

Type int_vector
    length As Integer
    vector() As Integer
End Type

Type Quaternion
    x As Double
    y As Double
    z As Double
    w As Double
End Type


Public Const PI = 3.14159265358979
Public Const PIOVER180 = PI / 180#
Public Const QUAT_NORM_TOLERANCE = 0.00001
Public Const INFINITY_SINGLE = 3.4028234E+38

Public Const EulRepYes = 1
Public Const EulParOdd = 1
Public Const EulFrmR = 1
Public Const FLT_EPSILON = 1.192092896E-07
Private Const MAX_DELTA_SQUARED As Single = 0.001 * 0.001
Private OnBits(0 To 31) As Long
'---------------------------------------------------------------------------------------------------------
'--------------------------------------------ARITHMETIC/LOGIC---------------------------------------------
'---------------------------------------------------------------------------------------------------------
Function IsNan(ByVal val As Single) As Boolean
    Dim raw_val As Long

    CopyMemory raw_val, val, 4

    IsNan = ((raw_val And &H7F800000) = &H7F800000)
End Function
Function GetDegreesFromRaw(ByVal val As Long, ByVal key As Integer) As Single
    GetDegreesFromRaw = (CSng(val) / CSng(2 ^ (12 - key))) * 360#
End Function

Function GetRawFromDegrees(ByVal val As Single, ByVal key As Integer) As Long
    GetRawFromDegrees = CLng((val / 360#) * CSng(2 ^ (12 - key)))
End Function
Function GetBitInteger(ByVal val As Integer, ByVal bit_index As Integer)
    GetBitInteger = IIf((val And 2 ^ bit_index) <> 0, 1, 0)
End Function
Function SetBitInteger(ByVal val As Integer, ByVal bit_index As Integer, ByVal bit_val As Integer)
    If bit_val = 0 Then
        SetBitInteger = val And (Not (2 ^ bit_index))
    Else
        SetBitInteger = val Or (2 ^ bit_index)
    End If
End Function
Function InvertBitInteger(ByVal val As Integer, ByVal bit_index As Integer)
    If GetBitInteger(val, bit_index) = 1 Then
        InvertBitInteger = SetBitInteger(val, bit_index, 0)
    Else
        InvertBitInteger = SetBitInteger(val, bit_index, 1)
    End If
End Function
Function NormalizeAngle180(ByVal val As Single) As Single
    Dim dec As Single

    If val > 0 Then
        dec = 360#
    Else
        dec = -360#
    End If

    NormalizeAngle180 = val
    While (NormalizeAngle180 > 0# And val > 0#) Or (NormalizeAngle180 < 0# And val < 0#)
        NormalizeAngle180 = NormalizeAngle180 - dec
    Wend

    If Abs(NormalizeAngle180) > Abs(NormalizeAngle180 + dec) Then
        NormalizeAngle180 = NormalizeAngle180 + dec
    End If

    If NormalizeAngle180 >= 180# Then
        NormalizeAngle180 = NormalizeAngle180 - 360#
    End If
End Function

' arc sine
' error if value is outside the range [-1,1]
Function ASin(value As Double) As Double
    If False Then
        If Abs(value) <> 1# Then
            ASin = atan2(value, Sqr(1# - value * value))
        Else
            ASin = 1.5707963267949 * Sgn(value)
        End If
    Else
        If (Sqr(1# - value * value) <= 0.000000000001) And (Sqr(1# - value * value) >= -0.000000000001) Then
            ASin = PI / 2
        Else
            ASin = Atn(value / Sqr(-value * value + 1#))
        End If

    End If
End Function
' arc cosine
' error if NUMBER is outside the range [-1,1]
Function ACos(ByVal number As Double) As Double
    On Error Resume Next

    If number = 1 Then
        ACos = 0
        Exit Function
    End If

    ACos = Atn(-number / Sqr(-number * number + 1)) + 2 * Atn(1)

    On Error GoTo 0
End Function
' arc cotangent
' error if NUMBER is zero
Function ACot(value As Double) As Double
    ACot = Atn(1 / value)
End Function
' arc secant
' error if value is inside the range [-1,1]
Function ASec(value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    ' ASec = ACos(1 / value)
    If Abs(value) <> 1 Then
        ASec = 1.5707963267949 - Atn((1 / value) / Sqr(1 - 1 / (value * value)))
    Else
        ASec = 3.14159265358979 * Sgn(value)
    End If
End Function
' arc cosecant
' error if value is inside the range [-1,1]
Function ACsc(value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    ' ACsc = ASin(1 / value)
    If Abs(value) <> 1 Then
        ACsc = Atn((1 / value) / Sqr(1 - 1 / (value * value)))
    Else
        ACsc = 1.5707963267949 * Sgn(value)
    End If
End Function
Public Function atan2(ByVal y As Double, ByVal x As Double) As Double
    If y > 0 Then
      If x >= y Then
        atan2 = Atn(y / x)
      ElseIf x <= -y Then
        atan2 = Atn(y / x) + PI
      Else
        atan2 = PI / 2 - Atn(x / y)
      End If
    Else
      If x >= -y Then
        atan2 = Atn(y / x)
      ElseIf x <= y Then
        atan2 = Atn(y / x) - PI
      Else
        atan2 = -Atn(x / y) - PI / 2
      End If
    End If
End Function
Public Function DegToRad(ByVal x As Double) As Double
    DegToRad = x * PI / 180#
End Function
Public Function RadToDeg(ByVal x As Double) As Double
    RadToDeg = x * 180# / PI
End Function

Function Min(ByVal x As Double, ByVal y As Double) As Integer
    If x > y Then
        Min = y
    Else
        Min = x
    End If
End Function
Public Function max(ByVal x As Double, ByVal y As Double) As Single
    If x < y Then
        max = y
    Else
        max = x
    End If
End Function
Public Function Min3(ByVal a As Integer, ByVal B As Integer, ByVal C As Integer) As Integer
    If a > B Then
        If B > C Then
            Min3 = C
        Else
            Min3 = B
        End If
    Else
        If a > C Then
            Min3 = C
        Else
            Min3 = a
        End If
    End If
End Function
Public Function LShiftLong(ByVal value As Long, _
    ByVal Shift As Integer) As Long
    Dim BI As Integer

    If Shift = 0 Then
        LShiftLong = value
    Else
        MakeOnBits

        If (value And (2 ^ (31 - Shift))) Then GoTo OverFlow

        LShiftLong = ((value And OnBits(31 - Shift)) * (2 ^ Shift))
        'LShiftLong = value
        'For bi = 0 To Shift - 1
        '    LShiftLong = LShiftLong * 2
        'Next bi
    End If
    Exit Function

OverFlow:

    LShiftLong = ((value And OnBits(31 - (Shift + 1))) * _
       (2 ^ (Shift))) Or &H80000000

End Function

Public Function RShiftLong(ByVal value As Long, _
   ByVal Shift As Integer) As Long
    Dim hi As Long
    MakeOnBits
    If (value And &H80000000) Then hi = &H40000000

    RShiftLong = (value And &H7FFFFFFE) \ (2 ^ Shift)
    RShiftLong = (RShiftLong Or (hi \ (2 ^ (Shift - 1))))
    'Dim bi As Integer
    'RShiftLong = value
    'For bi = 0 To Shift - 1
    '    RShiftLong = RShiftLong \ 2
    'Next bi
End Function

Public Function ExtendSignInteger(ByVal val As Integer, ByVal length As Integer) As Integer
    Dim aux_res As Long

    If length <> 12 Then
        aux_res = aux_res
    End If

    If (val And 2 ^ (length - 1)) <> 0 Then
        aux_res = 2 ^ 16 - 1
        aux_res = aux_res Xor ((2 ^ length) - 1)
        aux_res = aux_res Or val

        CopyMemory ExtendSignInteger, aux_res, 2
    Else
        ExtendSignInteger = val
    End If
End Function
Private Sub MakeOnBits()
    Dim j As Integer, _
        v As Long

    For j = 0 To 30

        v = v + (2 ^ j)
        OnBits(j) = v

    Next j

    OnBits(j) = v + &H80000000

End Sub
Public Function GetBitBlockV(ByRef vect() As Byte, ByVal nBits As Integer, ByRef FBit As Long) As Integer
    Dim temp_val As Integer
    temp_val = GetBitBlockVUnsigned(vect, nBits, FBit)
    GetBitBlockV = GetSignExtendedShort(temp_val, nBits)
End Function
'The value is considered unsigned
Public Function GetBitBlockVUnsigned(ByRef vect() As Byte, ByVal nBits As Integer, ByRef FBit As Long) As Integer
    Dim base_byte As Long
    Dim BI As Long
    Dim res As Long
    Dim num_bytes As Long
    Dim unaligned_by_bits As Long
    Dim is_aligned As Boolean
    Dim clean_end As Boolean
    Dim first_aligned_byte As Long
    Dim last_aligned_byte As Long
    Dim end_bits As Long

    Dim aux_res As Integer


    If nBits > 0 Then
        base_byte = FBit \ 8
        unaligned_by_bits = FBit Mod 8

        If unaligned_by_bits + nBits > 8 Then
            is_aligned = (unaligned_by_bits = 0)

            end_bits = (FBit + nBits) Mod 8
            clean_end = (end_bits = 0)

            num_bytes = (nBits - IIf(is_aligned, 0, 8 - unaligned_by_bits) - IIf(clean_end, 0, end_bits)) \ 8 + IIf(is_aligned, 0, 1) + IIf(clean_end, 0, 1)
            last_aligned_byte = num_bytes - IIf(clean_end, 0, 1) - 1
            first_aligned_byte = 0

            res = 0
            'Unaligned prefix
            'Stored at the begining of the byte
            If Not is_aligned Then
                res = CLng(vect(base_byte))
                res = res And ((2 ^ (8 - unaligned_by_bits)) - 1)
                first_aligned_byte = 1
            End If

            'Aligned bytes
            For BI = first_aligned_byte To last_aligned_byte
                res = res * 256
                res = res Or CLng(vect(base_byte + BI))
            Next BI

            'Sufix
            'Stored at the end of the byte
            If Not clean_end Then
                res = res * (2 ^ end_bits)
                res = res Or ((CLng(vect(base_byte + last_aligned_byte + 1)) _
                               \ (2 ^ (8 - end_bits))) _
                               And ((2 ^ end_bits) - 1))
            End If
        Else
            res = CLng(vect(base_byte))
            res = res \ (2 ^ (8 - (unaligned_by_bits + nBits)))
            res = res And ((2 ^ nBits) - 1)
        End If

        CopyMemory GetBitBlockVUnsigned, res, 2

        FBit = FBit + nBits
    Else
        GetBitBlockVUnsigned = 0
    End If
End Function

Public Sub PutBitBlockV(ByRef vect() As Byte, ByVal nBits As Integer, ByRef FBit As Long, ByVal value As Long)
    Dim base_byte As Long
    Dim BI As Long
    Dim res As Long
    Dim num_bytes As Long
    Dim unaligned_by_bits As Long
    Dim is_aligned As Boolean
    Dim clean_end As Boolean
    Dim first_aligned_byte As Long
    Dim last_aligned_byte As Long
    Dim end_bits As Long

    Dim aux_val As Long

    'Deal with it as some raw positive value. Divisions can't be used for bit shifting negative values, since they round towards 0 instead of minus infinity
    value = value And ((2 ^ nBits) - 1)

    If nBits > 0 Then
        base_byte = FBit \ 8
        unaligned_by_bits = FBit Mod 8

        If unaligned_by_bits + nBits > 8 Then
            is_aligned = (unaligned_by_bits = 0)

            end_bits = (FBit + nBits) Mod 8
            clean_end = (end_bits = 0)

            num_bytes = (nBits - IIf(is_aligned, 0, 8 - unaligned_by_bits) - IIf(clean_end, 0, end_bits)) \ 8 + IIf(is_aligned, 0, 1) + IIf(clean_end, 0, 1)
            last_aligned_byte = num_bytes - IIf(clean_end, 0, 1) - 1
            first_aligned_byte = 0

            ReDim Preserve vect(base_byte + num_bytes - 1)

            'Unaligned prefix
            If Not is_aligned Then
                aux_val = value \ 2 ^ (nBits - (8 - unaligned_by_bits))
                aux_val = aux_val And ((2 ^ (8 - unaligned_by_bits)) - 1)
                vect(base_byte) = vect(base_byte) Or aux_val
                first_aligned_byte = 1
            End If

            'Aligned bytes
            For BI = first_aligned_byte To last_aligned_byte
                aux_val = value \ 2 ^ ((last_aligned_byte - BI) * 8 + end_bits)
                vect(base_byte + BI) = aux_val And 255
            Next BI

            'Sufix
            If Not clean_end Then
                aux_val = value And (2 ^ (end_bits) - 1)
                vect(base_byte + last_aligned_byte + 1) = aux_val * 2 ^ (8 - end_bits)
            End If
        Else
            If UBound(vect) < base_byte Then
                ReDim Preserve vect(base_byte)
                vect(base_byte) = 0
            End If
            aux_val = value And (2 ^ (nBits) - 1)
            aux_val = aux_val * 2 ^ (8 - (unaligned_by_bits + nBits))
            vect(base_byte) = vect(base_byte) Or aux_val
        End If
    End If

    FBit = FBit + nBits
End Sub

Public Function GetSignExtendedShort(ByVal src As Long, ByVal valLength As Integer) As Integer
    Dim tempV As Integer
    Dim BI As Integer
    If valLength > 0 Then
        If valLength < 16 Then
            GetSignExtendedShort = ExtendSignInteger(src, valLength)
        Else
            CopyMemory GetSignExtendedShort, src, 2
        End If
    Else
        GetSignExtendedShort = 0
    End If
End Function

Public Function GetSignExtendedLong(ByVal src As Long, ByVal valLength As Integer) As Long
    Dim tempV As Long
    If valLength > 0 Then
        If Not ((src And 2 ^ (valLength - 1)) = 0) Then
            'If (valLength < 16) Then
                tempV = Not ((2 ^ valLength) - 1)
                tempV = tempV Or src
            'Else
             '   tempV = src
            'End If

            'CopyMemory GetSignExtendedShort, tempV, 2
            GetSignExtendedLong = tempV
        Else
            'CopyMemory GetSignExtendedShort, src, 2
            GetSignExtendedLong = src
        End If
    Else
        GetSignExtendedLong = 0
    End If
End Function
Public Function SignExtendBits(ByVal in_val As Long, ByVal num_bits As Integer) As Long
    Dim sign_bits As Long
    If (in_val And 2 ^ (num_bits - 1)) = 0 Then
        SignExtendBits = in_val
    Else
        sign_bits = -1 And (Not (2 ^ num_bits - 1))
        SignExtendBits = sign_bits Or in_val
    End If
End Function

Public Sub QuickSortNumericDescending(ByRef narray() As order_pair, inLow As Long, inHi As Long)

   Dim pivot As order_pair
   Dim tmpSwap As order_pair
   Dim tmpLow As Long
   Dim tmpHi  As Long

   tmpLow = inLow
   tmpHi = inHi

   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)

      While (narray(tmpLow).d > pivot.d And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend

      While (pivot.d > narray(tmpHi).d And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If

   Wend

   If (inLow < tmpHi) Then QuickSortNumericDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortNumericDescending narray(), tmpLow, inHi

End Sub
Public Sub TransposeMatrix(ByRef mat() As Double)
    Dim i As Integer
    Dim j As Integer
    Dim temp As Double
    Dim order As Integer

    order = Sqr(UBound(mat))
    For i = 0 To order - 1
        For j = 0 To i
            temp = mat(i * order + j)
            mat(i * order + j) = mat(i + j * order)
            mat(i + j * order) = temp
        Next j
    Next i

End Sub
Public Sub GetSubMatrix(ByRef mat() As Double, ByVal i As Integer, ByVal j As Integer, ByRef mat_out() As Double)
    Dim i2 As Integer
    Dim j2 As Integer

    Dim order As Integer
    Dim pos As Integer

    order = Sqr(UBound(mat))
    Dim test As String

    ReDim mat_out((order - 1) ^ 2)

    For i2 = 0 To order - 1
        If i2 <> i Then
            For j2 = 0 To order - 1
                If j2 <> j Then
                    pos = i2 + j2 * (order - 1)
                    If i2 > i Then pos = pos - 1
                    If j2 > j Then pos = pos - order + 1
                    mat_out(pos) = mat(i2 + j2 * order)
                End If
            Next j2
        End If
    Next i2
End Sub
Public Sub GetAtachedMatrix(ByRef mat() As Double, ByRef mat_out() As Double)
    Dim i As Integer
    Dim j As Integer

    Dim order As Integer
    Dim mat_aux() As Double

    order = Sqr(UBound(mat))
    ReDim mat_out(order ^ 2)

    For i = 0 To order - 1
        For j = 0 To order - 1
            GetSubMatrix mat, i, j, mat_aux
            mat_out(i + j * order) = ((-1) ^ (i + j)) * GetMatrixDeterminant(mat_aux)
        Next j
    Next i
End Sub
Public Function GetMatrixDeterminant(ByRef mat() As Double) As Double
    Dim i As Integer
    Dim i2 As Integer
    Dim j As Integer

    Dim order As Integer
    Dim det_aux As Double

    Dim mat_aux() As Double

    order = Sqr(UBound(mat))

    If order > 2 Then
        For i = 0 To order - 1
            If mat(i) <> 0 Then
                GetSubMatrix mat, i, 0, mat_aux
                det_aux = GetMatrixDeterminant(mat_aux) * ((-1) ^ (i)) * mat(i)
                GetMatrixDeterminant = GetMatrixDeterminant + det_aux

            End If
        Next i
    Else
        GetMatrixDeterminant = mat(0) * mat(3) - mat(1) * mat(2)
    End If
End Function
Public Sub InvertMatrix(ByRef mat() As Double)
    Dim i As Integer
    Dim j As Integer

    Dim order As Integer
    Dim mat_aux() As Double

    Dim det As Double

    order = Sqr(UBound(mat))

    det = GetMatrixDeterminant(mat)

    GetAtachedMatrix mat, mat_aux

    For i = 0 To order - 1
        For j = 0 To order - 1
            mat(i + j * order) = mat_aux(i + j * order) / det
        Next j
    Next i
    TransposeMatrix mat
End Sub
Public Sub MultiplyMatrix(ByRef matA() As Double, ByRef matB() As Double, ByRef matRes() As Double)
    Dim i As Integer
    Dim j As Integer
    Dim j2 As Integer

    Dim order As Integer

    order = Sqr(UBound(matA))

    For i = 0 To order - 1
        For j = 0 To order - 1
            matRes(i + j * order) = 0
            For j2 = 0 To order - 1
                matRes(i + j * order) = matRes(i + j * order) + matA(i + j2 * order) * matB(j2 + j * order)
            Next j2
        Next j
    Next i

End Sub
Public Sub MultiplyPoint3DByOGLMatrix(ByRef matA() As Double, ByRef p_in As Point3D, ByRef p_out As Point3D)
    With p_out
        .x = p_in.x * matA(0) + p_in.y * matA(4) + p_in.z * matA(8) + matA(12)
        .y = p_in.x * matA(1) + p_in.y * matA(5) + p_in.z * matA(9) + matA(13)
        .z = p_in.x * matA(2) + p_in.y * matA(6) + p_in.z * matA(10) + matA(14)
    End With
End Sub

Public Sub BuildRotationMatrixWithQuaternions(ByVal alpha As Double, ByVal Beta As Double, ByVal Gamma As Double, ByRef mat_res() As Double)
    Dim quat_x As Quaternion
    Dim quat_y As Quaternion
    Dim quat_z As Quaternion
    Dim quat_xy As Quaternion
    Dim quat_xyz As Quaternion
    Dim px As Point3D
    Dim py As Point3D
    Dim pz As Point3D
    px.x = 1
    px.y = 0
    px.z = 0

    py.x = 0
    py.y = 1
    py.z = 0

    pz.x = 0
    pz.y = 0
    pz.z = 1

    'BuildQuaternionFromEuler -Alpha, -Beta, -Gamma, quat_xyz 'quat_x
    BuildQuaternionFromAxis px, alpha, quat_x
    BuildQuaternionFromAxis py, Beta, quat_y
    BuildQuaternionFromAxis pz, Gamma, quat_z

    MultiplyQuaternions quat_y, quat_x, quat_xy
    MultiplyQuaternions quat_xy, quat_z, quat_xyz

    'BuildQuaternionFromEuler -alpha, -beta, -gamma, quat_xyz
    BuildMatrixFromQuaternion quat_xyz, mat_res
End Sub

Public Function RotateVectorAlpha(ByVal alpha As Single, ByRef vect As Point3D) As Point3D
    Dim quat_alpha As Quaternion
    Dim res As Point3D

    BuildQuaternionFromEuler alpha, 0, 0, quat_alpha

    res = RotatePointByQuaternion(quat_alpha, vect)
    RotateVectorAlpha = res
End Function

Public Function RotateVectorBeta(ByVal Beta As Single, ByRef vect As Point3D) As Point3D
    Dim quat_beta As Quaternion
    Dim res As Point3D

    BuildQuaternionFromEuler 0, Beta, 0, quat_beta

    res = RotatePointByQuaternion(quat_beta, vect)
    RotateVectorBeta = res
End Function
Public Function RotateVectorGamma(ByVal Gamma As Single, ByRef vect As Point3D) As Point3D
    Dim quat_gamma As Quaternion
    Dim res As Point3D

    BuildQuaternionFromEuler 0, 0, Gamma, quat_gamma

    res = RotatePointByQuaternion(quat_gamma, vect)
    RotateVectorGamma = res
End Function
Public Function GetQuaternionConjugate(ByRef quat As Quaternion) As Quaternion
    With GetQuaternionConjugate
        .x = -quat.x
        .y = -quat.y
        .z = -quat.z
        .w = quat.w
    End With
End Function
Public Function RotatePointByQuaternion(ByRef quat As Quaternion, ByRef vect As Point3D) As Point3D
    Dim vect_quat As Quaternion
    Dim quat_conj As Quaternion
    Dim quat_aux As Quaternion

    With vect_quat
        .x = vect.x
        .y = vect.y
        .z = vect.z
        .w = 1
    End With

    MultiplyQuaternions quat, vect_quat, quat_aux
    quat_conj = GetQuaternionConjugate(quat)
    MultiplyQuaternions quat_aux, quat_conj, vect_quat

    With RotatePointByQuaternion
        .x = vect_quat.x
        .y = vect_quat.y
        .z = vect_quat.z
    End With
End Function
Public Sub BuildRotationMatrixWithQuaternionsXYZ(ByVal alpha As Double, ByVal Beta As Double, ByVal Gamma As Double, ByRef mat_res() As Double)
    Dim quat_x As Quaternion
    Dim quat_y As Quaternion
    Dim quat_z As Quaternion
    Dim quat_xy As Quaternion
    Dim quat_xyz As Quaternion
    Dim px As Point3D
    Dim py As Point3D
    Dim pz As Point3D
    px.x = 1
    px.y = 0
    px.z = 0

    py.x = 0
    py.y = 1
    py.z = 0

    pz.x = 0
    pz.y = 0
    pz.z = 1

    BuildQuaternionFromAxis px, alpha, quat_x
    BuildQuaternionFromAxis py, Beta, quat_y
    BuildQuaternionFromAxis pz, Gamma, quat_z

    MultiplyQuaternions quat_x, quat_y, quat_xy
    MultiplyQuaternions quat_xy, quat_z, quat_xyz

    BuildMatrixFromQuaternion quat_xyz, mat_res
End Sub

'Convert from Euler Angles
Public Sub BuildQuaternionFromEuler(ByVal alpha As Double, ByVal Beta As Double, ByVal Gamma As Double, ByRef quat_res As Quaternion)
    'Basically we create 3 Quaternions, one for pitch, one for yaw, one for roll
    'and multiply those together.

    Dim quat_x As Quaternion
    Dim quat_y As Quaternion
    Dim quat_z As Quaternion
    Dim quat_xy As Quaternion
    Dim quat_xyz As Quaternion
    Dim px As Point3D
    Dim py As Point3D
    Dim pz As Point3D
    px.x = 1
    px.y = 0
    px.z = 0

    py.x = 0
    py.y = 1
    py.z = 0

    pz.x = 0
    pz.y = 0
    pz.z = 1

    BuildQuaternionFromAxis px, alpha, quat_x
    BuildQuaternionFromAxis py, Beta, quat_y
    BuildQuaternionFromAxis pz, Gamma, quat_z

    MultiplyQuaternions quat_y, quat_x, quat_xy
    MultiplyQuaternions quat_xy, quat_z, quat_res

    NormalizeQuaternion quat_res
End Sub


Public Sub NormalizeQuaternion(ByRef quat As Quaternion)
    'Don't normalize if we don't have to
    Dim mag As Double
    Dim mag2 As Double
    Dim test As Double
    With quat
        mag2 = .w * .w + .x * .x + .y * .y + .z * .z
        If Abs(mag2 - 1#) > QUAT_NORM_TOLERANCE Then
            mag = Sqr(mag2)
            .w = .w / mag
            .x = .x / mag
            .y = .y / mag
            .z = .z / mag
        End If

        If .w > 1# Then
            .w = 1
        End If
    End With
End Sub
'Convert Quaternion to Matrix
Public Sub BuildMatrixFromQuaternion(ByRef quat As Quaternion, ByRef mat_res() As Double)
    Dim X2 As Double
    Dim Y2 As Double
    Dim z2 As Double
    Dim xy As Double
    Dim xz As Double
    Dim yz As Double
    Dim wx As Double
    Dim wy As Double
    Dim wz As Double

    With quat
        X2 = .x * .x
        Y2 = .y * .y
        z2 = .z * .z
        xy = .x * .y
        xz = .x * .z
        yz = .y * .z
        wx = .w * .x
        wy = .w * .y
        wz = .w * .z
    End With

    'This calculation would be a lot more complicated for non-unit length quaternions
    'Note: The constructor of Matrix4 expects the Matrix in column-major format like expected by
    'OpenGL
    mat_res(0) = 1# - 2# * (Y2 + z2)
    mat_res(4) = 2# * (xy - wz)
    mat_res(8) = 2# * (xz + wy)
    mat_res(12) = 0#
    mat_res(1) = 2# * (xy + wz)
    mat_res(5) = 1# - 2# * (X2 + z2)
    mat_res(9) = 2# * (yz - wx)
    mat_res(13) = 0#
    mat_res(2) = 2# * (xz - wy)
    mat_res(6) = 2# * (yz + wx)
    mat_res(10) = 1# - 2# * (X2 + Y2)
    mat_res(14) = 0#
    mat_res(3) = 0#
    mat_res(7) = 0#
    mat_res(11) = 0#
    mat_res(15) = 1#
End Sub
'Convert from Axis Angle
Public Sub BuildQuaternionFromAxis(ByRef vec As Point3D, ByVal angle As Double, ByRef res_quat As Quaternion)
    Dim sinAngle As Double
    angle = angle * PIOVER180 / 2

    sinAngle = Sin(angle)

    With res_quat
        .x = (vec.x * sinAngle)
        .y = (vec.y * sinAngle)
        .z = (vec.z * sinAngle)
        .w = Cos(angle)
    End With
End Sub
'Multiplying q1 with q2 applies the rotation q2 to q1
Public Sub MultiplyQuaternions(ByRef quat_a As Quaternion, ByRef quat_b As Quaternion, ByRef quat_res As Quaternion)
    With quat_a
        quat_res.x = .w * quat_b.x + .x * quat_b.w + .y * quat_b.z - .z * quat_b.y
        quat_res.y = .w * quat_b.y + .y * quat_b.w + .z * quat_b.x - .x * quat_b.z
        quat_res.z = .w * quat_b.z + .z * quat_b.w + .x * quat_b.y - .y * quat_b.x
        quat_res.w = .w * quat_b.w - .x * quat_b.x - .y * quat_b.y - .z * quat_b.z
    End With
End Sub
Public Function ConvertQ(ByVal heading As Double, ByVal attitude As Double, ByVal bank As Double) As Quaternion
    Dim c1, c2, c1c2, s1, s2, s1s2, c3, s3, w, h, a, B As Double
    Dim PI As Double
    PI = 4 * Atn(1)

    h = heading * (PI / 360#)
    a = attitude * (PI / 360#)
    B = bank * (PI / 360#)
    c1 = Cos(h)
    c2 = Cos(a)
    c3 = Cos(B)
    s1 = Sin(h)
    s2 = Sin(a)
    s3 = Sin(B)
    ConvertQ.w = c1 * c2 * c3 - s1 * s2 * s3
    ConvertQ.x = s1 * s2 * c3 + c1 * c2 * s3
    ConvertQ.y = s1 * c2 * c3 + c1 * s2 * s3
    ConvertQ.z = c1 * s2 * c3 - s1 * c2 * s3
End Function
Public Function ConvertQZYX(ByVal theta_x As Double, ByVal theta_y As Double, ByVal theta_z As Double) As Quaternion
    Dim cos_z_2, cos_y_2, cos_x_2 As Double
    Dim sin_z_2, sin_y_2, sin_x_2 As Double
    cos_z_2 = Cos(0.5 * theta_z)
    cos_y_2 = Cos(0.5 * theta_y)
    cos_x_2 = Cos(0.5 * theta_x)

    sin_z_2 = Sin(0.5 * theta_z)
    sin_y_2 = Sin(0.5 * theta_y)
    sin_x_2 = Sin(0.5 * theta_x)

    With ConvertQZYX
        .w = cos_z_2 * cos_y_2 * cos_x_2 + sin_z_2 * sin_y_2 * sin_x_2
        .x = cos_z_2 * cos_y_2 * sin_x_2 - sin_z_2 * sin_y_2 * cos_x_2
        .y = cos_z_2 * sin_y_2 * cos_x_2 + sin_z_2 * cos_y_2 * sin_x_2
        .z = sin_z_2 * cos_y_2 * cos_x_2 - cos_z_2 * sin_y_2 * sin_x_2
    End With
End Function
'Returns the euler angles from a rotation quaternion
Public Function GetEulerAnglesFromQuaternion(ByRef quat As Quaternion, ByVal homogenous As Boolean) As Point3D
    Dim sqw As Double
    Dim sqx As Double
    Dim sqy As Double
    Dim sqz As Double
    Dim test As Double


    With quat
        test = .x * .y + .z * .w
        If test > 0.499 Then ' singularity at north pole
            GetEulerAnglesFromQuaternion.x = 2 * atan2(.x, .w)
            GetEulerAnglesFromQuaternion.y = PI / 2
            GetEulerAnglesFromQuaternion.z = 0
        Else
            If (test < -0.499) Then ' singularity at south pole
                GetEulerAnglesFromQuaternion.x = -2# * atan2(.x, .w)
                GetEulerAnglesFromQuaternion.y = -PI / 2#
                GetEulerAnglesFromQuaternion.z = 0#
            Else
                sqx = .x * .x
                sqy = .y * .y
                sqz = .z * .z
                GetEulerAnglesFromQuaternion.x = atan2(2# * .y * .w - 2# * .x * .z, 1 - 2# * sqy - 2# * sqz)
                GetEulerAnglesFromQuaternion.y = ASin(2# * test)
                GetEulerAnglesFromQuaternion.z = atan2(2# * .x * .w - 2# * .y * .z, 1 - 2# * sqx - 2# * sqz)
            End If
        End If
    End With

    With GetEulerAnglesFromQuaternion
        Dim aux As Double
        .x = RadToDeg(.x)
        .y = RadToDeg(.y)
        .z = RadToDeg(.z)
    End With
End Function

Public Sub GetAxisAngleFromQuaternion(ByRef quat As Quaternion, ByRef x_axis As Double, ByRef y_axis As Double, ByRef z_axis As Double, ByRef angle)
    Dim s As Double
    If (quat.w > 1) Then
        NormalizeQuaternion quat
    End If
    With quat
        angle = 2 * ACos(.w)
        s = Sqr(1# - .w * .w) ' assuming quaternion normalised then w is less than 1, so term always positive.
        If (s < 0.001) Then ' test to avoid divide by zero, s is always positive due to sqrt
            ' if s close to zero then direction of axis not important
            x_axis = .x ' if it is important that axis is normalised then replace with x=1; y=z=0;
            y_axis = .y
            z_axis = .z
        Else
            x_axis = .x / s ' normalise axis
            y_axis = .y / s
            z_axis = .z / s
        End If
    End With
End Sub
Public Function GetEulerFromAxisAngle(ByVal axis_x As Double, ByVal axis_y As Double, ByVal axis_z As Double, ByVal angle As Double) As Point3D
    Dim s As Double
    Dim C As Double
    Dim t As Double
    s = Sin(angle)
    C = Cos(angle)
    t = 1# - C
    '  if axis is not already normalised then uncomment this
    ' double magnitude = Math.sqrt(x*x + y*y + z*z);
    ' if (magnitude==0) throw error;
    ' x /= magnitude;
    ' y /= magnitude;
    ' z /= magnitude;
    With GetEulerFromAxisAngle
        If ((axis_x * axis_y * t + axis_z * s) > 0.998) Then ' // north pole singularity detected
            .x = 2# * atan2(axis_x * Sin(angle / 2#), Cos(angle / 2#))
            .y = PI / 2#
            .z = 0
        ElseIf ((axis_x * axis_y * t + axis_z * s) < -0.998) Then ' // south pole singularity detected
            .x = -2 * atan2(axis_x * Sin(angle / 2#), Cos(angle / 2#))
            .y = -PI / 2
            .z = 0
        Else
            .x = atan2(axis_y * s - axis_x * axis_z * t, 1# - (axis_y * axis_y + axis_z * axis_z) * t)
            .y = ASin(axis_x * axis_y * t + axis_z * s)
            .z = atan2(axis_x * s - axis_y * axis_z * t, 1# - (axis_x * axis_x + axis_z * axis_z) * t)
        End If

        .x = RadToDeg(.x)
        .y = RadToDeg(.y)
        .z = RadToDeg(.z)
    End With
End Function
Public Function GetQuaternionFromEulerXYZr(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Quaternion
    GetQuaternionFromEulerXYZr = GetQuaternionFromEulerUniversal(DegToRad(x), DegToRad(y), DegToRad(z), 2, 1, 0, 2, 1, 0, 1)
End Function

Public Function GetQuaternionFromEulerYXZr(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Quaternion
    GetQuaternionFromEulerYXZr = GetQuaternionFromEulerUniversal(DegToRad(x), DegToRad(y), DegToRad(z), 2, 0, 1, 2, 0, 0, 1)
End Function
Public Function GetQuaternionFromEulerUniversal(ByVal y As Double, ByVal x As Double, ByVal z As Double, ByVal i As Integer, ByVal j As Integer, ByVal k As Integer, ByVal h As Integer, ByVal n As Integer, ByVal s As Integer, ByVal f As Integer) As Quaternion
    Dim a(2) As Double
    Dim ti As Double
    Dim tj As Double
    Dim th As Double
    Dim ci As Double
    Dim cj As Double
    Dim ch As Double
    Dim si As Double
    Dim sj As Double
    Dim sh As Double
    Dim cc As Double
    Dim cs As Double
    Dim sc As Double
    Dim ss As Double

    Dim t As Double

    If f = EulFrmR Then
        t = x
        x = z
        z = t
    End If
    If n = EulParOdd Then
        y = -y
    End If
    ti = x * 0.5
    tj = y * 0.5
    th = z * 0.5
    ci = Cos(ti)
    cj = Cos(tj)
    ch = Cos(th)
    si = Sin(ti)
    sj = Sin(tj)
    sh = Sin(th)
    cc = ci * ch
    cs = ci * sh
    sc = si * ch
    ss = si * sh
    With GetQuaternionFromEulerUniversal
        If s = EulRepYes Then
            a(i) = cj * (cs + sc)  'Could speed up with
            a(j) = sj * (cc + ss)  'trig identities.
            a(k) = sj * (cs - sc)
            .w = cj * (cc - ss)
        Else
            a(i) = cj * sc - sj * cs
            a(j) = cj * ss + sj * cc
            a(k) = cj * cs - sj * sc
            .w = cj * cc + sj * ss
        End If
        If n = EulParOdd Then
            a(j) = -a(j)
        End If
        .x = a(0)
        .y = a(1)
        .z = a(2)
    End With
End Function
Public Function GetEulerXYZrFromMatrix(ByRef mat() As Double) As Point3D
    GetEulerXYZrFromMatrix = GetEulerFormMatrixUniversal(mat, 2, 1, 0, 2, 1, 0, 1)
End Function

Public Function GetEulerYXZrFromMatrix(ByRef mat() As Double) As Point3D
    GetEulerYXZrFromMatrix = GetEulerFormMatrixUniversal(mat, 2, 0, 1, 2, 0, 0, 1)
End Function
'Thanks to Ken Shoemaker and his article at Graphics Gems Collection IV (http://tog.acm.org/resources/GraphicsGems/)
Public Function GetEulerFormMatrixUniversal(ByRef mat() As Double, ByVal i As Integer, ByVal j As Integer, ByVal k As Integer, ByVal h As Integer, ByVal n As Integer, ByVal s As Integer, ByVal f As Integer) As Point3D
    Dim sy As Double
    Dim cy As Double

    Dim t As Double

    With GetEulerFormMatrixUniversal
        If s = EulRepYes Then
            sy = Sqr(mat(i + 4 * j) * mat(i + 4 * j) + mat(i + 4 * k) * mat(i + 4 * k))
            If sy > 16# * FLT_EPSILON Then
                .x = atan2(mat(i + 4 * j), mat(i + 4 * k))
                .y = atan2(sy, mat(i + 4 * i))
                .z = atan2(mat(j + 4 * i), -mat(k + 4 * i))
            Else
                .x = atan2(-mat(j + 4 * k), mat(j + 4 * j))
                .y = atan2(sy, mat(i + 4 * i))
                .z = 0
            End If
        Else
            cy = Sqr(mat(i + 4 * i) * mat(i + 4 * i) + mat(j + 4 * i) * mat(j + 4 * i))
            If cy > 16# * FLT_EPSILON Then
                .x = atan2(mat(k + 4 * j), mat(k + 4 * k))
                .y = atan2(-mat(k + 4 * i), cy)
                .z = atan2(mat(j + 4 * i), mat(i + 4 * i))
            Else
                .x = atan2(-mat(j + 4 * k), mat(j + 4 * j))
                .y = atan2(-mat(k + 4 * i), cy)
                .z = 0
            End If
        End If
        If n = EulParOdd Then
            .x = -.x
            .y = -.y
            .z = -.z
        End If
        If f = EulFrmR Then
            t = .x
            .x = .z
            .z = t
        End If
        'ea.w = order

        .x = RadToDeg(.x)
        .y = RadToDeg(.y)
        .z = RadToDeg(.z)
    End With
End Function
Public Function QuaternionsSlerp(ByRef qa As Quaternion, ByRef qb As Quaternion, ByVal alpha As Double) As Quaternion
    Dim cosHalfTheta, halfTheta, sinHalfTheta As Double
    Dim ratioA As Double
    Dim ratioB As Double
    Dim qb2 As Quaternion

    With qa
        cosHalfTheta = .w * qb.w + .x * qb.x + .y * qb.y + .z * qb.z
        If (cosHalfTheta < 0) Then
            qb2.w = -qb.w
            qb2.x = -qb.x
            qb2.y = -qb.y
            qb2.z = qb.z
            cosHalfTheta = -cosHalfTheta
        Else
            qb2.w = qb.w
            qb2.x = qb.x
            qb2.y = qb.y
            qb2.z = qb.z
        End If
    End With

    With QuaternionsSlerp
        If (Abs(cosHalfTheta) >= 1) Then
            .w = qa.w
            .x = qa.x
            .y = qa.y
            .z = qa.z
        Else
            halfTheta = ACos(cosHalfTheta)
            sinHalfTheta = Sqr(1# - cosHalfTheta * cosHalfTheta)

            If (Abs(sinHalfTheta) < 0.001) Then
                .w = (qa.w * 0.5 + qb2.w * 0.5)
                .x = (qa.x * 0.5 + qb2.x * 0.5)
                .y = (qa.y * 0.5 + qb2.y * 0.5)
                .z = (qa.z * 0.5 + qb2.z * 0.5)
            End If
            ratioA = Sin((1# - alpha) * halfTheta) / sinHalfTheta
            ratioB = Sin(alpha * halfTheta) / sinHalfTheta
            'calculate Quaternion.
            .w = (qa.w * ratioA + qb2.w * ratioB)
            .x = (qa.x * ratioA + qb2.x * ratioB)
            .y = (qa.y * ratioA + qb2.y * ratioB)
            .z = (qa.z * ratioA + qb2.z * ratioB)
        End If
    End With

    NormalizeQuaternion QuaternionsSlerp
End Function
Public Function QuaternionsDot(ByRef q1 As Quaternion, ByRef q2 As Quaternion) As Double
    QuaternionsDot = q1.x * q2.x + q1.y * q2.y + q1.z * q2.z + q1.w * q2.w
End Function

Public Function QuaternionLerp(ByRef q1 As Quaternion, ByRef q2 As Quaternion, ByVal t As Double) As Quaternion
    Dim one_minus_t As Double
    With QuaternionLerp
        one_minus_t = 1# - t
        .x = q1.x * one_minus_t + q2.x * t
        .y = q1.y * one_minus_t + q2.y * t
        .z = q1.z * one_minus_t + q2.z * t
        .w = q1.w * one_minus_t + q2.w * t
    End With
    NormalizeQuaternion QuaternionLerp
End Function
'http://willperone.net/Code/quaternion.php
Public Function QuaternionSlerp2(ByRef q1 As Quaternion, ByRef q2 As Quaternion, ByVal t As Double) As Quaternion
    Dim q3 As Quaternion
    Dim dot As Double
    Dim angle As Double
    Dim one_minus_t As Double
    Dim sin_angle As Double
    Dim sin_angle_by_t As Double
    Dim sin_angle_by_one_t As Double


    dot = QuaternionsDot(q1, q2)

    '  dot = cos(theta)
    '    if (dot < 0), q1 and q2 are more than 90 degrees apart,
    '    so we can invert one to reduce spinning
    If dot < 0 Then
        dot = -dot
        With q3
            .x = -q2.x
            .y = -q2.y
            .z = -q2.z
            .w = -q2.w
        End With
    Else
        With q3
            .x = q2.x
            .y = q2.y
            .z = q2.z
            .w = q2.w
        End With
    End If

    If dot < 0.95 Then
        angle = ACos(dot)
        one_minus_t = 1# - t
        sin_angle = Sin(angle)
        sin_angle_by_t = Sin(angle * t)
        sin_angle_by_one_t = Sin(angle * one_minus_t)
        With QuaternionSlerp2
            .x = ((q1.x * sin_angle_by_one_t) + q3.x * sin_angle_by_t) / sin_angle
            .y = ((q1.y * sin_angle_by_one_t) + q3.y * sin_angle_by_t) / sin_angle
            .z = ((q1.z * sin_angle_by_one_t) + q3.z * sin_angle_by_t) / sin_angle
            .w = ((q1.w * sin_angle_by_one_t) + q3.w * sin_angle_by_t) / sin_angle
        End With
    Else ' if the angle is small, use linear interpolation
        QuaternionSlerp2 = QuaternionLerp(q1, q3, t)
    End If
End Function
'Normalizes angles to the 0 - 360 range.
Public Sub NormalizeEulerAngles(ByRef angles As Point3D)
    With angles
        If .x > 360# Then
            While .x > 360#
                .x = .x - 360#
            Wend
        ElseIf .x < 0# Then
            While .x < 0#
                .x = .x + 360#
            Wend
        End If
        If .y > 360# Then
            While .y > 360#
                .y = .y - 360#
            Wend
        ElseIf .y < 0# Then
            While .y < 0#
                .y = .y + 360#
            Wend
        End If
        If .z > 360# Then
            While .z > 360#
                .z = .z - 360#
            Wend
        ElseIf .z < 0# Then
            While .z < 0#
                .z = .z + 360#
            Wend
        End If
    End With
End Sub
'Return true if val1 > val2
Public Function CompareLongs(ByVal val1 As Long, ByVal val2 As Long) As Boolean
    CompareLongs = IIf((val1 Xor val2) < 0, _
                        val1 < 0, _
                        val1 > val2)
End Function
'---------------------------------------------------------------------------------------------------------
'--------------------------------------FILE SYSTEM--------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Public Function FileExist(asPath As String) As Boolean
    If UCase(Dir(asPath)) = UCase(TrimPath(asPath)) Then
        FileExist = True
    Else
        FileExist = False
    End If
End Function
Public Function TrimPath(ByVal asPath As String) As String
    If Len(asPath) = 0 Then Exit Function
    Dim x As Integer

    Do
        x = InStr(asPath, "\")
        If x = 0 Then Exit Do
        asPath = Right(asPath, Len(asPath) - x)
    Loop
    TrimPath = asPath
End Function

Public Function GetCommLine() As String
    Dim RetStr As Long, SLen As Long
    Dim buffer As String
    'Get a pointer to a string, which contains the command line
    RetStr = GetCommandLine
    'Get the length of that string
    SLen = lstrlen(RetStr)

    If SLen > 0 Then
        'Create a buffer
        GetCommLine = Space$(SLen)
        'Copy to the buffer
        CopyMemory ByVal GetCommLine, ByVal RetStr, SLen

        GetCommLine = Right$(GetCommLine, SLen - 1)

        While Left$(GetCommLine, 1) <> "" + Chr$(34)
            GetCommLine = Right$(GetCommLine, Len(GetCommLine) - 1)
        Wend

        If Len(GetCommLine) > 3 Then
            GetCommLine = Right$(GetCommLine, Len(GetCommLine) - 3)
            GetCommLine = Left$(GetCommLine, Len(GetCommLine) - 1)
        Else
            GetCommLine = ""
        End If
    End If
End Function
Public Function GetPathFromString(ByVal fileName As String) As String
    Dim ci As Integer

    For ci = 1 To Len(fileName)
        If Left$(Right$(fileName, 1), Len(fileName)) = "\" Then
            GetPathFromString = Left$(fileName, Len(fileName))
            Exit Function
        End If
        fileName = Left$(fileName, Len(fileName) - 1)
    Next ci
End Function
'---------------------------------------------------------------------------------------------------------
'--------------------------------------------GRAPHICS-----------------------------------------------------
'---------------------------------------------------------------------------------------------------------
'Function creaDC(ByVal x As Long, ByVal y As Long) As Long
'    Dim hBITMAP As Long, hdc As Long, tam As tagSIZE, tipo As Integer, error As String
'    error = "Error de gr�ficos!"
'    tipo = 0 + 0 + 16
'    While hdc < 1
'        hdc = CreateCompatibleDC(GetDC(0))
'    Wend
'    hBIT MAP = CreateCompatibleBitmap(GetDC(0), x, y)
'    If hBITMAP = 0 Then
'        MsgBox "No se puede crear el bitmap", tipo, error
'        DeleteDC hdc
'        Exit Function
'    End If
'    SelectObject hdc, hBITMAP

'    DeleteObject hBITMAP
'    BitBlt hdc, 0, 0, x, y, creaDC.Handle, 0, 0, Blackness
'End Function
Sub Draw_Buffer(ByVal hdc As Long, ByRef buffer() As Long, ByVal width As Integer, ByVal height As Integer)
    Dim i As Integer
    Dim j As Integer

    ''Debug.Print "Dibujando buffer de " + Str$(width) + " X" + Str$(height) + " pixels en " + Str$(hdc)
    For i = 0 To width
        For j = 0 To height
            SetPixel hdc, i, j, buffer(i, j)
            'DoEvents
        Next j
    Next i
    ''Debug.Print "Draw_buffer finalizado"
End Sub
Sub Interpolate_Buffer(ByRef dest() As Long, ByRef src() As Long, ByVal width As Integer, ByVal height As Integer)
    Dim x As Integer
    Dim y As Integer
    Dim r_temp As Byte
    Dim g_temp As Byte
    Dim b_temp As Byte
    Dim c1 As Long
    Dim c2 As Long

    ''Debug.Print "Interpolando buffer de " + Str$(width) + " X" + Str$(height) + " pixels"

    For x = 0 To width
        For y = 0 To height
                dest(x * 2, y * 2) = src(x, y)
        Next y
    Next x
    DoEvents
    For x = 1 To width * 2 - 1 Step 2
        For y = 0 To height * 2 Step 2

            c1 = dest(x - 1, y)

            c2 = dest(x + 1, y)

            r_temp = ((c1 And &HFF) + (c2 And &HFF)) / 2
            g_temp = ((c1 And 65280) + (c2 And 65280)) / 2 ^ 9
            b_temp = ((c1 And 16711680) + (c2 And 16711680)) / 2 ^ 17
            dest(x, y) = RGB(r_temp, g_temp, b_temp)
            'DoEvents
        Next y
    Next x
    DoEvents
    For x = 0 To width * 2
        For y = 1 To height * 2 - 1 Step 2

            c1 = dest(x, y - 1)

            c2 = dest(x, y + 1)

            r_temp = ((c1 And &HFF) + (c2 And &HFF)) / 2
            g_temp = ((c1 And 65280) + (c2 And 65280)) / 2 ^ 9
            b_temp = ((c1 And 16711680) + (c2 And 16711680)) / 2 ^ 17
            dest(x, y) = RGB(r_temp, g_temp, b_temp)
            'DoEvents
        Next y
    Next x
    DoEvents

    ''Debug.Print "Interpolate_buffer finalizado"
End Sub
Function CombineColor(ByRef a As color, ByRef B As color) As color
    With CombineColor
        .a = (a.a * 1# + B.a) / 2
        .r = (a.r * 1# + B.r) / 2
        .g = (a.g * 1# + B.g) / 2
        .B = (a.B * 1# + B.B) / 2
    End With
End Function
Function ConvertTwipsToPixels(lngTwips As Long, _
   lngDirection As Long) As Long

   'Handle to device
   Dim lngDC As Long
   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   'lngDC = GetDC(0)

   'If (lngDirection = 0) Then       'Horizontal
      'lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   'Else                            'Vertical
      'lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   'End If
   'lngDC = ReleaseDC(0, lngDC)
   ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch

End Function
Public Function getBrightness(ByVal r As Byte, ByVal g As Byte, ByVal B As Byte) As Long
    Dim r_t, g_t, b_t As Integer
    r_t = r
    g_t = g
    b_t = B
    getBrightness = (r_t + g_t + b_t) / 3
End Function
Public Function getRed(ByVal col As Long)
    getRed = (col And &HFF)
End Function
Public Function getGreen(ByVal col As Long)
    getGreen = RShiftLong(col, 8) And &HFF
End Function
Public Function getBlue(ByVal col As Long)
    getBlue = RShiftLong(col, 16) And &HFF
End Function
Public Function ConvertRGB555ToRGB888(ByVal src As Integer) As Long
    Dim r As Byte, B As Byte, g As Byte

    r = LShiftLong(src And 31!, 3)
    If r > 0 Then r = r + 7
    g = LShiftLong(RShiftLong(src And 992!, 5), 3)
    If g > 0 Then g = g + 7
    B = LShiftLong(RShiftLong(src And 31744!, 10), 3)
    If B > 0 Then B = B + 7

    ConvertRGB555ToRGB888 = RGB(r, g, B)
End Function
Public Function GetLongFromRGB(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte) As Long
    GetLongFromRGB = red * 2 ^ 16 Or green * 2 ^ 8 Or blue
End Function
Public Sub DrawBox(ByVal max_x As Double, ByVal max_y As Double, ByVal max_z As Double, ByVal min_x As Double, ByVal min_y As Double, ByVal min_z As Double, ByVal red As Single, ByVal green As Single, ByVal blue As Single)
    glColor3f red, green, blue

    glBegin GL_LINES
        glVertex3f max_x, max_y, max_z
        glVertex3f max_x, max_y, min_z
        glVertex3f max_x, max_y, max_z
        glVertex3f max_x, min_y, max_z
        glVertex3f max_x, max_y, max_z
        glVertex3f min_x, max_y, max_z

        glVertex3f min_x, min_y, min_z
        glVertex3f min_x, min_y, max_z
        glVertex3f min_x, min_y, min_z
        glVertex3f min_x, max_y, min_z
        glVertex3f min_x, min_y, min_z
        glVertex3f max_x, min_y, min_z

        glVertex3f max_x, min_y, min_z
        glVertex3f max_x, max_y, min_z
        glVertex3f max_x, min_y, min_z
        glVertex3f max_x, min_y, max_z

        glVertex3f min_x, max_y, min_z
        glVertex3f min_x, max_y, max_z
        glVertex3f min_x, max_y, min_z
        glVertex3f max_x, max_y, min_z

        glVertex3f min_x, min_y, max_z
        glVertex3f min_x, max_y, max_z
        glVertex3f min_x, min_y, max_z
        glVertex3f max_x, min_y, max_z
    glEnd
End Sub
Public Sub ComputeTransformedBoxBoundingBox(ByRef MV_matrix() As Double, _
            ByRef p_min As Point3D, ByRef p_max As Point3D, _
            ByRef p_min_trans As Point3D, ByRef p_max_trans As Point3D)
    Dim box_pointsV(7) As Point3D
    Dim p_aux_trans As Point3D
    Dim PI As Integer

    p_max_trans.x = -INFINITY_SINGLE
    p_max_trans.y = -INFINITY_SINGLE
    p_max_trans.z = -INFINITY_SINGLE

    p_min_trans.x = INFINITY_SINGLE
    p_min_trans.y = INFINITY_SINGLE
    p_min_trans.z = INFINITY_SINGLE

    box_pointsV(0) = p_min
    With box_pointsV(1)
        .x = p_min.x
        .y = p_min.y
        .z = p_max.z
    End With
    With box_pointsV(2)
        .x = p_min.x
        .y = p_max.y
        .z = p_min.z
    End With
    With box_pointsV(3)
        .x = p_min.x
        .y = p_max.y
        .z = p_max.z
    End With
    box_pointsV(4) = p_max
    With box_pointsV(5)
        .x = p_max.x
        .y = p_max.y
        .z = p_min.z
    End With
    With box_pointsV(6)
        .x = p_max.x
        .y = p_min.y
        .z = p_max.z
    End With
    With box_pointsV(7)
        .x = p_max.x
        .y = p_min.y
        .z = p_min.z
    End With

    For PI = 0 To 7
        MultiplyPoint3DByOGLMatrix MV_matrix, box_pointsV(PI), p_aux_trans
        With p_aux_trans
            If p_max_trans.x < .x Then p_max_trans.x = .x
            If p_max_trans.y < .y Then p_max_trans.y = .y
            If p_max_trans.z < .z Then p_max_trans.z = .z

            If p_min_trans.x > .x Then p_min_trans.x = .x
            If p_min_trans.y > .y Then p_min_trans.y = .y
            If p_min_trans.z > .z Then p_min_trans.z = .z
        End With
    Next PI
End Sub
Sub GetViewportWorldBox(ByRef p_min As Point3D, ByRef p_max As Point3D)
    Dim MV_matrix(16) As Double
    Dim P_matrix(16) As Double
    Dim PMV_matrix(16) As Double

    Dim p_min_aux As Point3D
    Dim p_max_aux As Point3D

    With p_min_aux
        .x = 0
        .y = 0
        .z = 0
    End With

    With p_max_aux
        .x = 1
        .y = 1
        .z = 0
    End With

    glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)
    glGetDoublev GL_PROJECTION_MATRIX, P_matrix(0)

    MultiplyMatrix P_matrix, MV_matrix, PMV_matrix
    InvertMatrix PMV_matrix

    MultiplyPoint3DByOGLMatrix PMV_matrix, p_min_aux, p_min
    MultiplyPoint3DByOGLMatrix PMV_matrix, p_max_aux, p_max
End Sub
Function ComputeSceneRadius(ByRef p_min As Point3D, ByRef p_max As Point3D) As Double
    Dim center_model As Point3D
    Dim model_radius As Single
    Dim distance_origin As Single
    Dim origin As Point3D
    Dim distance_radius As Single

    center_model.x = (p_min.x + p_max.x) / 2
    center_model.y = (p_min.y + p_max.y) / 2
    center_model.z = (p_min.z + p_max.z) / 2
    origin.x = 0
    origin.y = 0
    origin.z = 0
    model_radius = CalculateDistance(p_min, p_max) / 2
    distance_origin = CalculateDistance(center_model, origin)
    ComputeSceneRadius = model_radius + distance_origin
End Function
Sub SetCameraModelView(ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByVal alpha As Single, ByVal Beta As Single, ByVal Gamma As Single, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW
    glLoadIdentity

    glTranslatef cx, cy, CZ

    BuildRotationMatrixWithQuaternionsXYZ alpha, Beta, _
        Gamma, rot_mat
    glMultMatrixd rot_mat(0)

    glScalef redX, redY, redZ
End Sub
Sub SetCameraModelViewQuat(ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByRef quat As Quaternion, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW
    glLoadIdentity

    glTranslatef cx, cy, CZ

    BuildMatrixFromQuaternion quat, rot_mat
    glMultMatrixd rot_mat(0)

    glScalef redX, redY, redZ
End Sub
Sub ConcatenateCameraModelView(ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByVal alpha As Single, ByVal Beta As Single, ByVal Gamma As Single, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW

    glTranslatef cx, cy, CZ

    BuildRotationMatrixWithQuaternionsXYZ alpha, Beta, _
        Gamma, rot_mat
    glMultMatrixd rot_mat(0)

    glScalef redX, redY, redZ
End Sub
Sub ConcatenateCameraModelViewQuat(ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByRef quat As Quaternion, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW

    glTranslatef cx, cy, CZ

    BuildMatrixFromQuaternion quat, rot_mat
    glMultMatrixd rot_mat(0)

    glScalef redX, redY, redZ
End Sub
Sub SetCameraAroundModel(ByRef p_min As Point3D, ByRef p_max As Point3D, ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByVal alpha As Single, ByVal Beta As Single, ByVal Gamma As Single, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim width As Integer
    Dim height As Integer
    Dim scene_radius As Single
    Dim vp(4) As Long

    glGetIntegerv GL_VIEWPORT, vp(0)
    width = vp(2)
    height = vp(3)

    glMatrixMode GL_PROJECTION
    glLoadIdentity

    scene_radius = ComputeSceneRadius(p_min, p_max)
    gluPerspective 60, width / height, max(0.1, -CZ - scene_radius), max(0.1, -CZ + scene_radius)

    SetCameraModelView cx, cy, CZ, alpha, Beta, Gamma, redX, redY, redZ
End Sub
Sub SetCameraAroundModelQuat(ByRef p_min As Point3D, ByRef p_max As Point3D, ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByRef quat As Quaternion, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim width As Integer
    Dim height As Integer
    Dim scene_radius As Single
    Dim vp(4) As Long

    glGetIntegerv GL_VIEWPORT, vp(0)
    width = vp(2)
    height = vp(3)

    glMatrixMode GL_PROJECTION
    glLoadIdentity

    scene_radius = ComputeSceneRadius(p_min, p_max)
    gluPerspective 60, width / height, max(0.1, -CZ - scene_radius), max(0.1, -CZ + scene_radius)

    SetCameraModelViewQuat cx, cy, CZ, quat, redX, redY, redZ
End Sub
Sub SetCameraInfinite(ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByVal alpha As Single, ByVal Beta As Single, ByVal Gamma As Single, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim vp(4) As Long
    Dim width As Long
    Dim height As Long
    Dim rot_mat(16) As Double

    glGetIntegerv GL_VIEWPORT, vp(0)
    width = vp(2)
    height = vp(3)

    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluPerspective 60, width / height, 0.1, 1000000

    SetCameraModelView cx, cy, CZ, alpha, Beta, Gamma, redX, redY, redZ
End Sub
'---------------------------------------------------------------------------------------------------------
'--------------------------------------------GEOMETRIC----------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Public Function Normalize(ByRef v As Point3D) As Point3D
    Dim l As Single

    l = Sqr(v.x ^ 2 + v.y ^ 2 + v.z ^ 2)
    If l > 0 Then
        l = 1 / l
        With Normalize
            .x = v.x * l
            .y = v.y * l
            .z = v.z * l
        End With
    Else
        With Normalize
            .x = 0
            .y = 0
            .z = 0
        End With
    End If
End Function
Public Function VectorProduct(ByRef vect1 As Point3D, ByRef vect2 As Point3D) As Point3D
    With vect1
        VectorProduct.x = .y * vect2.z - .z * vect2.y
        VectorProduct.y = .z * vect2.x - .x * vect2.z
        VectorProduct.z = .x * vect2.y - .y * vect2.x
    End With
End Function
Public Function CalculateNormal(ByRef p1 As Point3D, ByRef p2 As Point3D, ByRef p3 As Point3D) As Point3D
    Dim Qx, Qy, Qz, px, py, pz As Single

    px = p2.x - p1.x
    py = p2.y - p1.y
    pz = p2.z - p1.z
    Qx = p3.x - p1.x
    Qy = p3.y - p1.y
    Qz = p3.z - p1.z
    CalculateNormal.x = py * Qz - pz * Qy
    CalculateNormal.y = pz * Qx - px * Qz
    CalculateNormal.z = px * Qy - py * Qx
End Function
Public Function CalculatePoint2LineProjectionPosition(ByRef q As Point3D, ByRef p1 As Point3D, ByRef p2 As Point3D) As Single
    Dim alpha As Single
    Dim VD As Point3D

    VD.x = p2.x - p1.x
    VD.y = p2.y - p1.y
    VD.z = p2.z - p1.z

    alpha = (VD.x * (q.x - p1.x) + VD.y * (q.y - p1.y) + VD.z * (q.z - p1.z)) / _
            (VD.x ^ 2 + VD.y ^ 2 + VD.z ^ 2)

    If alpha > 1 Then alpha = 1
    If alpha < -1 Then alpha = -1

    CalculatePoint2LineProjectionPosition = alpha
End Function
Public Function CalculateLinePoint(ByVal alpha As Single, ByRef p1 As Point3D, ByRef p2 As Point3D) As Point3D
    With CalculateLinePoint
        .x = p1.x + (p2.x - p1.x) * alpha
        .y = p1.y + (p2.y - p1.y) * alpha
        .z = p1.z + (p2.z - p1.z) * alpha
    End With
End Function
Public Function CalculatePoint2LineProjection(ByRef q As Point3D, ByRef p1 As Point3D, ByRef p2 As Point3D) As Point3D
    Dim alpha As Single

    alpha = CalculatePoint2LineProjectionPosition(q, p1, p2)

    CalculatePoint2LineProjection = CalculateLinePoint(alpha, p1, p2)
End Function

Public Function CalculateDistance(ByRef p1 As Point3D, ByRef p2 As Point3D) As Single
    CalculateDistance = Sqr((p2.x - p1.x) ^ 2 + (p2.y - p1.y) ^ 2 + (p2.z - p1.z) ^ 2)
End Function
Public Function InterpolateColor(ByRef c1 As color, ByRef c2 As color, ByVal alpha As Single) As color
    With InterpolateColor
        .r = c2.r * alpha + c1.r * (1# - alpha)
        .g = c2.g * alpha + c1.g * (1# - alpha)
        .B = c2.B * alpha + c1.B * (1# - alpha)
        .a = c2.a * alpha + c1.a * (1# - alpha)
    End With
End Function
Public Function InterpolatePoint2D(ByRef p1 As Point2D, ByRef p2 As Point2D, ByVal alpha As Single) As Point2D
    With InterpolatePoint2D
        .x = p2.x * alpha + p1.x * (1# - alpha)
        .y = p2.y * alpha + p1.y * (1# - alpha)
    End With
End Function
Public Function CompareSimilarPoints3D(ByRef a As Point3D, ByRef B As Point3D) As Boolean
    CompareSimilarPoints3D = ComparePoints3D(a, B)
    If Not CompareSimilarPoints3D Then
        Dim dx As Single
        Dim dy As Single
        Dim dz As Single

        Dim dist_square As Single

        With a
            dx = .x - B.x
            dy = .y - B.y
            dz = .z - B.z
        End With
        dist_square = dx * dx + dy * dy + dz * dz

        CompareSimilarPoints3D = dist_square <= MAX_DELTA_SQUARED
    End If
End Function
Public Function ComparePoints3D(ByRef a As Point3D, ByRef B As Point3D) As Boolean
    ComparePoints3D = (a.x = B.x) And (a.y = B.y) And (a.z = B.z)
End Function
Public Function ComparePoints2D(ByRef a As Point2D, ByRef B As Point2D) As Boolean
    ComparePoints2D = (a.x = B.x) And (a.y = B.y)
End Function
Public Function CompareColors(ByRef a As color, ByRef B As color) As Boolean
    CompareColors = (a.r = B.r) And (a.r = B.r) And (a.g = B.g) And (a.a = B.a)
End Function

Public Function IsLexicographicallyGreater(ByVal str1 As String, ByVal str2 As String) As Boolean
    Dim len1 As Integer
    Dim len2 As Integer
    Dim ci As Integer
    Dim min_len As Integer
    Dim c1 As Integer
    Dim c2 As Integer

    len1 = Len(str1)
    len2 = Len(str2)
    min_len = IIf(len1 > len2, len2, len1)
    IsLexicographicallyGreater = False
    For ci = 1 To min_len
        c1 = Asc(Mid$(str1, ci, 1))
        c2 = Asc(Mid$(str2, ci, 1))
        If (c1 > c2) Then
            IsLexicographicallyGreater = True
            Exit For
        ElseIf (c1 < c2) Then
            Exit For
        End If
    Next ci
End Function

Public Function IsPoint3DUnderPlane(ByRef point As Point3D, ByVal a As Single, ByVal B As Single, _
                                    ByVal C As Single, ByVal d As Single) As Boolean

    Dim orthogonal_projection As Point3D
    Dim vect As Point3D
    Dim vect_norm As Point3D

    orthogonal_projection = GetPoint3DOrthogonalProjection(point, a, B, C, d)

    With vect
        .x = orthogonal_projection.x - point.x
        .y = orthogonal_projection.y - point.y
        .z = orthogonal_projection.z - point.z
    End With

    vect_norm = Normalize(vect)

    With vect_norm
        IsPoint3DUnderPlane = Not (Abs(a - .x) < 0.0001 And _
                                    Abs(B - .y) < 0.0001 And _
                                    Abs(C - .z) < 0.0001) And _
                                    Not CalculateDistance(point, orthogonal_projection) < 0.0001
    End With
End Function
Public Function IsPoint3DAbovePlane(ByRef point As Point3D, ByVal a As Single, ByVal B As Single, _
                                    ByVal C As Single, ByVal d As Single) As Boolean

    Dim orthogonal_projection As Point3D
    Dim vect As Point3D
    Dim vect_norm As Point3D

    orthogonal_projection = GetPoint3DOrthogonalProjection(point, a, B, C, d)

    With vect
        .x = orthogonal_projection.x - point.x
        .y = orthogonal_projection.y - point.y
        .z = orthogonal_projection.z - point.z
    End With

    vect_norm = Normalize(vect)

    With vect_norm
        IsPoint3DAbovePlane = (Abs(a - .x) < 0.0001 And _
                                    Abs(B - .y) < 0.0001 And _
                                    Abs(C - .z) < 0.0001) And _
                                    Not CalculateDistance(point, orthogonal_projection) < 0.0001
    End With
End Function
Public Function GetPoint3DOrthogonalProjection(ByRef point As Point3D, ByVal a As Single, _
                                    ByVal B As Single, ByVal C As Single, ByVal d As Single) As Point3D

    Dim alpha As Single

    With point
        alpha = (-a * .x - B * .y - C * .z - d) / (a * a + B * B + C * C)
    End With

    With GetPoint3DOrthogonalProjection
        .x = point.x + alpha * a
        .y = point.y + alpha * B
        .z = point.z + alpha * C
    End With
End Function

Public Function GetPointAlphaInLine(ByRef point As Point3D, ByRef p1 As Point3D, ByRef p2 As Point3D) As Single
    If p2.x - p1.x <> 0 Then
        GetPointAlphaInLine = (point.x - p1.x) / (p2.x - p1.x)
    ElseIf p2.y - p1.y <> 0 Then
        GetPointAlphaInLine = (point.y - p1.y) / (p2.y - p1.y)
    Else
        GetPointAlphaInLine = (point.z - p1.z) / (p2.z - p1.z)
    End If
End Function

Public Function EqualSignSingle(ByVal num1 As Single, ByVal num2 As Single)
    EqualSignSingle = (num1 <= 0 And num2 <= 0) Or (num1 >= 0 And num2 >= 0)
End Function

Public Function GetPointInLine(ByRef p1 As Point3D, ByRef p2 As Point3D, ByVal alpha As Single) As Point3D
    With GetPointInLine
        .x = p1.x + (p2.x - p1.x) * alpha
        .y = p1.y + (p2.y - p1.y) * alpha
        .z = p1.z + (p2.z - p1.z) * alpha
    End With
End Function
Public Function GetPointInLine2D(ByRef p1 As Point2D, ByRef p2 As Point2D, ByVal alpha As Single) As Point2D
    With GetPointInLine2D
        .x = p1.x + (p2.x - p1.x) * alpha
        .y = p1.y + (p2.y - p1.y) * alpha
    End With
End Function

Public Function GetPointMirroredRelativeToPlane(ByRef point As Point3D, ByVal a As Single, _
                                                ByVal B As Single, ByVal C As Single, _
                                                ByVal d As Single) As Point3D
    Dim alpha As Single

    With point
        alpha = (-a * .x - B * .y - C * .z - d) / (a * a + B * B + C * C)
    End With

    With GetPointMirroredRelativeToPlane
        .x = point.x + 2 * alpha * a
        .y = point.y + 2 * alpha * B
        .z = point.z + 2 * alpha * C
    End With
End Function

Public Sub ComputePlaneABCD(ByRef PlaneVect1 As Point3D, ByRef PlaneVect2 As Point3D, _
                            ByRef PlanePoint As Point3D, ByRef a As Single, ByRef B As Single, _
                            ByRef C As Single, ByRef d As Single)
    Dim normal_plane As Point3D

    normal_plane = VectorProduct(PlaneVect1, PlaneVect2)
    normal_plane = Normalize(normal_plane)

    With normal_plane
        a = .x
        B = .y
        C = .z
    End With

    d = ComputePlaneD(normal_plane, PlanePoint)
End Sub
Public Function ComputePlaneD(ByRef normal As Point3D, ByRef PlanePoint As Point3D) As Single
    With PlanePoint
        ComputePlaneD = -normal.x * .x - normal.y * .y - normal.z * .z
    End With
End Function

Public Sub ComputeTransformedPlaneVectors(ByRef vect1_in As Point3D, ByRef vect2_in As Point3D, _
                                    ByRef rot_quat As Quaternion, _
                                    ByRef vect1_out As Point3D, ByRef vect2_out As Point3D)
    vect1_out = RotatePointByQuaternion(rot_quat, vect1_in)
    vect2_out = RotatePointByQuaternion(rot_quat, vect2_in)
End Sub

Public Function ComputeVectorsAngleCos(ByRef vec_1 As Point3D, ByRef vec_2 As Point3D) As Double
    Dim x1, Y1, z1 As Double
    Dim X2, Y2, z2 As Double

    With vec_1
        x1 = CDbl(vec_1.x)
        Y1 = CDbl(vec_1.y)
        z1 = CDbl(vec_1.z)
    End With

    With vec_2
        X2 = CDbl(vec_2.x)
        Y2 = CDbl(vec_2.y)
        z2 = CDbl(vec_2.z)
    End With

    ComputeVectorsAngleCos = (x1 * X2 + Y1 * Y2 + z1 * z2) / _
                            (Sqr(x1 * x1 + Y1 * Y1 + z1 * z1) * Sqr(X2 * X2 + Y2 * Y2 + z2 * z2))

    'Debug.Print "v1 = ("; Str$(x1); ","; Str$(Y1); ","; Str$(z1); "), v2 = ("; Str$(X2); ","; Str$(Y2); ","; Str$(z2); "), res = "; Str$(ComputeVectorsAngleCos)
End Function
Public Function AreVectorsParalel(ByRef vec1 As Point3D, ByRef vec2 As Point3D) As Boolean
    Dim norm_vec1 As Point3D
    Dim norm_vec2 As Point3D

    norm_vec1 = Normalize(vec1)
    norm_vec2 = Normalize(vec2)
    AreVectorsParalel = CompareSimilarPoints3D(norm_vec1, norm_vec2)
    If Not AreVectorsParalel Then
        With norm_vec2
            .x = -.x
            .y = -.y
            .z = -.z
        End With
        AreVectorsParalel = CompareSimilarPoints3D(norm_vec1, norm_vec2)
    End If
End Function
Public Function GetVectorToPlaneIntersection(ByRef v1 As Point3D, ByRef v2 As Point3D, ByVal a As Double, ByVal B As Double, ByVal C As Double, ByVal d As Double, ByRef alpha_out As Double) As Boolean
    Dim triangle_normal As Point3D
    Dim plane_normal As Point3D

    Dim lambda_mult_plane As Double
    Dim k_plane As Double

    triangle_normal = VectorProduct(v1, v2)
    With plane_normal
        .x = a
        .y = B
        .z = C
    End With

    alpha_out = 0
    GetVectorToPlaneIntersection = AreVectorsParalel(triangle_normal, plane_normal)
    If Not GetVectorToPlaneIntersection Then
        'If they aren't, find the cut point.
        With v1
            lambda_mult_plane = -a * CDbl(.x) - B * CDbl(.y) - C * CDbl(.z)
            k_plane = lambda_mult_plane - d
        End With

        With v2
            lambda_mult_plane = lambda_mult_plane + a * CDbl(.x) + B * CDbl(.y) + C * CDbl(.z)
        End With
        If (Abs(lambda_mult_plane) > 0.0000001 And k_plane <> 0) Then
            alpha_out = k_plane / lambda_mult_plane
        End If
    End If
End Function

Public Function GetFirstIndexOccurrenceLong(ByRef vect() As Long, ByVal first_index As Long, ByVal last_index As Long, ByVal value As Long) As Long
    Dim elem_index As Long
    Dim num_elems As Long

    num_elems = UBound(vect) + 1

    For elem_index = first_index To last_index
        If vect(elem_index) = value Then
            Exit For
        End If
    Next elem_index

    If vect(elem_index) = value Then
        GetFirstIndexOccurrenceLong = elem_index
    Else
        GetFirstIndexOccurrenceLong = -1
    End If
End Function

Public Sub EchangeVectorElementsLong(ByRef vect() As Long, ByVal index1 As Long, ByVal index2 As Long)
    Dim aux_val As Long

    aux_val = vect(index1)
    vect(index1) = vect(index2)
    vect(index2) = aux_val
End Sub

Public Sub PrintTableLong(ByRef vect() As Long)
    Dim num_elems As Long
    Dim ei As Long

    Debug.Print "{"
    num_elems = UBound(vect) + 1
    For ei = 0 To num_elems - 1
        Debug.Print " "; Str$(vect(ei)); ","
    Next ei
    Debug.Print "}"
End Sub
