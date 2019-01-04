Attribute VB_Name = "FF7DAAnimationsPackAnimationFrameBone"
Option Explicit
Const MAX_NEGLECTABLE_ROTATION_DIFFERENCE = 0
Type DAFrameBone
    AccumAlphaS As Integer
    AccumBetaS As Integer
    AccumGammaS As Integer
    
    AccumAlphaL As Long
    AccumBetaL As Long
    AccumGammaL As Long
    
    alpha As Single
    Beta As Single
    Gamma As Single
End Type
'For raw rotations
Function ReadDAUncompressedFrameBoneRotation(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte) As Integer
    Dim value As Integer
    
    value = GetBitBlockV(AnimationStream, 12 - key, offsetBit)
    'Convert to 12-bits value
    value = value * (2 ^ key)
    ReadDAUncompressedFrameBoneRotation = value
End Function
Function ReadDAFrameBoneRotationDelta(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte) As Integer
    Dim dLength As Integer
    Dim value As Integer
    Dim sign As Long
    Dim aux_sign_val As Integer
    If GetBitBlockVUnsigned(AnimationStream, 1, offsetBit) = 1 Then

        dLength = GetBitBlockVUnsigned(AnimationStream, 3, offsetBit)

        Select Case (dLength)
            Case 0:
                'Minimum bone rotation decrement
                value = -1
            Case 7:
                'Just like the first frame
                value = GetBitBlockV(AnimationStream, 12 - key, offsetBit)
            Case Else:
                value = GetBitBlockV(AnimationStream, dLength, offsetBit)
                
                'Invert the value of the last bit
                aux_sign_val = 2 ^ (dLength - 1)
            
                If value < 0 Then
                    value = value - aux_sign_val
                Else
                    value = value + aux_sign_val
                End If
        End Select
        'Convert to 12-bits value
        value = value * (2 ^ key)
        ReadDAFrameBoneRotationDelta = value
    Else
        ReadDAFrameBoneRotationDelta = 0
    End If
End Function
'For bone rotations of the first frame
Sub ReadDAUncompressedFrameBone(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef bone As DAFrameBone)
    With bone
        .AccumAlphaS = ReadDAUncompressedFrameBoneRotation(AnimationStream, offsetBit, key)
        .AccumBetaS = ReadDAUncompressedFrameBoneRotation(AnimationStream, offsetBit, key)
        .AccumGammaS = ReadDAUncompressedFrameBoneRotation(AnimationStream, offsetBit, key)
        
        .AccumAlphaL = IIf(.AccumAlphaS < 0, .AccumAlphaS + &H1000, .AccumAlphaS)
        .AccumBetaL = IIf(.AccumBetaS < 0, .AccumBetaS + &H1000, .AccumBetaS)
        .AccumGammaL = IIf(.AccumGammaS < 0, .AccumGammaS + &H1000, .AccumGammaS)
        
        .alpha = GetDegreesFromRaw(.AccumAlphaL, 0)
        .Beta = GetDegreesFromRaw(.AccumBetaL, 0)
        .Gamma = GetDegreesFromRaw(.AccumGammaL, 0)
    End With
End Sub
'For bone rotations of all the other frames
Sub ReadDAFrameBone(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef bone As DAFrameBone, ByRef LastFrameBone As DAFrameBone)
    With bone
        .AccumAlphaS = LastFrameBone.AccumAlphaS + ReadDAFrameBoneRotationDelta(AnimationStream, offsetBit, key)
        .AccumBetaS = LastFrameBone.AccumBetaS + ReadDAFrameBoneRotationDelta(AnimationStream, offsetBit, key)
        .AccumGammaS = LastFrameBone.AccumGammaS + ReadDAFrameBoneRotationDelta(AnimationStream, offsetBit, key)
        
        .AccumAlphaL = IIf(.AccumAlphaS < 0, .AccumAlphaS + &H1000, .AccumAlphaS)
        .AccumBetaL = IIf(.AccumBetaS < 0, .AccumBetaS + &H1000, .AccumBetaS)
        .AccumGammaL = IIf(.AccumGammaS < 0, .AccumGammaS + &H1000, .AccumGammaS)
        
        .alpha = GetDegreesFromRaw(.AccumAlphaL, 0)
        .Beta = GetDegreesFromRaw(.AccumBetaL, 0)
        .Gamma = GetDegreesFromRaw(.AccumGammaL, 0)
    End With
End Sub
Sub WriteDAFrameBoneRotationDelta(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByVal value As Long)
    Dim dLength As Integer
    Dim Value_out As Long
    Dim aux_sign_val As Long
    Dim foundQ As Boolean
    Dim value_out_reduced As Integer
    Dim precision_divisor As Long
    
    'Remove sign to prevent bad rounding on shift
    Value_out = (value And (2 ^ (12 - key) - 1))
    
    If (Value_out = 0) Then
        PutBitBlockV AnimationStream, 1, offsetBit, 0
    Else
        PutBitBlockV AnimationStream, 1, offsetBit, 1
        
        If Value_out = -(2 ^ key) Then
            'Minimum subtraction given the key precision
            PutBitBlockV AnimationStream, 3, offsetBit, 0
        Else
            'Find the lowest data length that can hold the requiered precision.
            dLength = 1
            foundQ = False
            While Not foundQ And dLength < 7
                foundQ = (Value_out And (Not ((2 ^ dLength) - 1))) = 0
                dLength = dLength + 1
            Wend
            dLength = IIf(foundQ, dLength - 1, 7)
            
            'Write data length
            PutBitBlockV AnimationStream, 3, offsetBit, dLength
    
            If foundQ Then
                'Write compressed (dLength < 7)
                Value_out = InvertBitInteger(Value_out, dLength - 1)
                PutBitBlockV AnimationStream, dLength, offsetBit, Value_out
            Else
                'Write raw (dLength = 7)
                PutBitBlockV AnimationStream, 12 - key, offsetBit, Value_out
            End If
        End If
    End If
End Sub
'For raw rotations
Sub WriteDAUncompressedFrameBoneRotation(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByVal value As Long)
    Dim Value_out As Long
    
    'Reduce precision to key bits
    'Remove sign to prevent bad rounding on shift
    'Value_out = (value And (2 ^ 12 - 1)) \ (2 ^ key)
    
    PutBitBlockV AnimationStream, 12 - key, offsetBit, value
End Sub
'For bone rotations of the first frame
Sub WriteDAUncompressedFrameBone(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef bone As DAFrameBone)
    With bone
        'Debug.Print "           Angle delta (U) "; Str$(.alpha); ", "; Str$(.Beta); ", "; Str$(.Gamma)
        WriteDAUncompressedFrameBoneRotation AnimationStream, offsetBit, key, GetRawFromDegrees(.alpha, key)
        WriteDAUncompressedFrameBoneRotation AnimationStream, offsetBit, key, GetRawFromDegrees(.Beta, key)
        WriteDAUncompressedFrameBoneRotation AnimationStream, offsetBit, key, GetRawFromDegrees(.Gamma, key)
    End With
End Sub
'For bone rotations of all the other frames
Sub WriteDAFrameBone(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef bone As DAFrameBone, ByRef LastFrameBone As DAFrameBone) ', ByRef AnimationCarry As Point3D)
    'Dim ang_diff As Point3D
    Dim raw_diff_x As Integer
    Dim raw_diff_y As Integer
    Dim raw_diff_z As Integer
    'With ang_diff
        '.x = bone.alpha - LastFrameBone.alpha + AnimationCarry.x
        '.y = bone.Beta - LastFrameBone.Beta + AnimationCarry.y
        '.z = bone.Gamma - LastFrameBone.Gamma + AnimationCarry.z
        
        'raw_diff_x = GetRawFromDegrees(.x, key)
        'raw_diff_y = GetRawFromDegrees(.y, key)
        'raw_diff_z = GetRawFromDegrees(.z, key)
        raw_diff_x = GetRawFromDegrees(bone.alpha, key) - GetRawFromDegrees(LastFrameBone.alpha, key)
        raw_diff_y = GetRawFromDegrees(bone.Beta, key) - GetRawFromDegrees(LastFrameBone.Beta, key)
        raw_diff_z = GetRawFromDegrees(bone.Gamma, key) - GetRawFromDegrees(LastFrameBone.Gamma, key)
        
        WriteDAFrameBoneRotationDelta AnimationStream, offsetBit, key, raw_diff_x
        WriteDAFrameBoneRotationDelta AnimationStream, offsetBit, key, raw_diff_y
        WriteDAFrameBoneRotationDelta AnimationStream, offsetBit, key, raw_diff_z
        
        'AnimationCarry.x = .x - GetDegreesFromRaw(raw_diff_x, key)
        'AnimationCarry.y = .y - GetDegreesFromRaw(raw_diff_y, key)
        'AnimationCarry.z = .z - GetDegreesFromRaw(raw_diff_z, key)
    'End With
End Sub
Sub CopyDAFrameBone(ByRef bone_in As DAFrameBone, ByRef bone_out As DAFrameBone)
    With bone_in
        bone_out.alpha = bone_in.alpha
        bone_out.Beta = bone_in.Beta
        bone_out.Gamma = bone_in.Gamma
    End With
End Sub

Sub NormalizeDAAnimationsPackAnimationFrameBone(ByRef bone As DAFrameBone, ByRef next_frame_bone As DAFrameBone)
    With next_frame_bone
        NormalizeDAAnimationsPackAnimationFrameBoneComponent bone.alpha, .alpha
        NormalizeDAAnimationsPackAnimationFrameBoneComponent bone.Beta, .Beta
        NormalizeDAAnimationsPackAnimationFrameBoneComponent bone.Gamma, .Gamma
    End With
End Sub
Sub NormalizeDAAnimationsPackAnimationFrameBoneComponent(ByVal val As Single, ByRef next_val As Single)
    Dim delta As Single
    
    delta = next_val - val
    If Abs(delta) > 180# Or delta = 180# Then
        delta = NormalizeAngle180(delta)
        
        next_val = val + delta
        If next_val - val >= 180# Then
            Debug.Print "WTF!"
        End If
    End If
End Sub

'For bone rotations of the first frame
Sub CheckWriteDAUncompressedFrameBone(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef ref_bone As DAFrameBone)
    Dim aux As Integer
    Dim aux2 As Integer
    
    With ref_bone
        aux = GetRawFromDegrees(.alpha, key)
        aux2 = ReadDAUncompressedFrameBoneRotation(AnimationStream, offsetBit, key)
        If GetRawFromDegrees(.alpha, key) <> aux2 Then
            Debug.Print "error"
        End If
        aux = GetRawFromDegrees(.Beta, key)
        aux2 = ReadDAUncompressedFrameBoneRotation(AnimationStream, offsetBit, key)
        If GetRawFromDegrees(.Beta, key) <> aux2 Then
            Debug.Print "error"
        End If
        If GetRawFromDegrees(.Gamma, key) <> ReadDAUncompressedFrameBoneRotation(AnimationStream, offsetBit, key) Then
            Debug.Print "error"
        End If
    End With
End Sub
'For bone rotations of all the other frames
Sub CheckWriteDAFrameBone(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef ref_bone As DAFrameBone, ByRef ref_last_frame_bone As DAFrameBone)
    Dim aux As Integer
    Dim aux2 As Integer
    With ref_bone
        aux = GetRawFromDegrees(.alpha, key) - GetRawFromDegrees(ref_last_frame_bone.alpha, key)
        aux2 = ReadDAFrameBoneRotationDelta(AnimationStream, offsetBit, key)
        If aux <> aux2 Then
            Debug.Print "Rotation delta mismatch detected"
        End If
        aux = GetRawFromDegrees(.Beta, key) - GetRawFromDegrees(ref_last_frame_bone.Beta, key)
        aux2 = ReadDAFrameBoneRotationDelta(AnimationStream, offsetBit, key)
        If aux <> aux2 Then
            Debug.Print "Rotation delta mismatch detected"
        End If
        aux = GetRawFromDegrees(.Gamma, key) - GetRawFromDegrees(ref_last_frame_bone.Gamma, key)
        aux2 = ReadDAFrameBoneRotationDelta(AnimationStream, offsetBit, key)
        If aux <> aux2 Then
            Debug.Print "Rotation delta mismatch detected"
        End If
    End With
End Sub
