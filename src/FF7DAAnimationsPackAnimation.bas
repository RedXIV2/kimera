Attribute VB_Name = "FF7DAAnimationsPackAnimation"
Option Explicit
Type DAAnimation
    NumBonesModel As Long           'Number of bones for the model + 1 (root transformation). Unreliable.
    NumFrames1 As Long              'Number of frames (conservative). Usually worng (smaller than the actual number).
    BlockLength As Long

    NumFrames2 As Integer           'Number of frames. Usually wrong (higher than the actual number).
    MissingNumFrames2 As Boolean    'The animation has no secondary frames count. Only RSAA seems to use it.
    AnimationLength As Integer      'Don't use this field EVER. It would be interpreted as a signed value.
    AnimationLengthLong As Long     'This isn't part of the actual structure, it's used just to overcome the lack of unsigned shorts support
    key As Byte
    
    Frames() As DAFrame
    UnknownData() As Byte
End Type
Sub ReadDAAnimation(ByVal NFile As Integer, ByRef offset As Long, ByVal BonesVectorLength As Integer, ByRef Animation As DAAnimation)
    Dim fi As Integer
    Dim AnimationStream() As Byte
    Dim offsetBit As Long
    Dim num_bones As Integer
    Dim sanity_check As Boolean
    Dim err_msg As String
    Dim missing_frame_counter As Boolean
    Dim last_offsetBit As Long
    
    On Error GoTo ErrorHand
    
    ''Debug.Print "+Base offset at byte" + Str$(offset)
    With Animation
        Get NFile, offset, .NumBonesModel
        Get NFile, offset + 4, .NumFrames1
        Get NFile, offset + 8, .BlockLength
        ''Debug.Print .BlockLength
        If .BlockLength < 11 Then
            If .BlockLength > 0 Then
                ReDim .UnknownData(.BlockLength - 1)
                Get NFile, offset + 12, .UnknownData
            End If
            offset = offset + 12 + .BlockLength
            .NumFrames2 = 0
            'Debug.Print "Empty slot!"
            Exit Sub
        End If
        Get NFile, offset + 12, .NumFrames2
        Get NFile, offset + 14, .AnimationLength
        CopyMemory .AnimationLengthLong, .AnimationLength, 2
        Get NFile, offset + 16, .key
        
        'Hack for reading animations with missing secondary frame counter (which can't be actually used by FF7)
        If .NumFrames2 = .BlockLength - 5 Then
            Get NFile, offset + 12, .AnimationLength
            CopyMemory .AnimationLengthLong, .AnimationLength, 2
            Get NFile, offset + 14, .key
            ReDim AnimationStream(.AnimationLengthLong)
            Get NFile, offset + 15, AnimationStream
            .MissingNumFrames2 = True
        Else
            ReDim AnimationStream(.AnimationLengthLong)
            Get NFile, offset + 17, AnimationStream
            .MissingNumFrames2 = False
        End If
        
        sanity_check = True
        
        If .NumFrames1 <> .NumFrames2 Then
            Debug.Print "WARNING!!! NumFrames1 is different from NumFrames2"
        End If
        'From now on, let's ignore the frame counter and just read as many frames a possible.
        .NumFrames2 = 9999
        
        If Not (.key = 0 Or .key = 2 Or .key = 4) Then
            err_msg = "ERROR!!! Invalid key " + Str$(.key)
            sanity_check = False
        End If
        
        If (Not sanity_check) Then
            MsgBox err_msg + ". Animation skipped"
            Debug.Print err_msg
            .NumFrames2 = 0
            ReDim .UnknownData(.BlockLength - 1)
            Get NFile, offset + 12, .UnknownData
            offset = offset + .BlockLength + 12
            Exit Sub
        End If
        
        ReDim .Frames(.NumFrames2 - 1)
        
        offsetBit = 0
        'Debug.Print "   -First frame at byte" + Str$(offset + 17)
        'Debug.Print "   Frame 0"
        
        ReadDAUncompressedFrame AnimationStream, offsetBit, .key, BonesVectorLength, .Frames(0)
        For fi = 1 To .NumFrames2 - 1
            'If we ran out of data while reading the frame, it means this frame doesn't
            last_offsetBit = offsetBit
            If Not ReadDAFrame(AnimationStream, offsetBit, .key, BonesVectorLength, .Frames(fi), .Frames(fi - 1)) Then
                .NumFrames2 = fi
                ReDim Preserve .Frames(.NumFrames2 - 1)
                offsetBit = last_offsetBit
                Exit For
            End If
        Next fi
        
        If (.BlockLength - .AnimationLengthLong > 5) Then
            ReDim .UnknownData(.BlockLength - .AnimationLengthLong - 1)
            'Debug.Print ".AnimationLengthLong = "; .AnimationLengthLong; "unkonwdata = "; UBound(.UnknownData) + 1; " total = "; .AnimationLengthLong + UBound(.UnknownData) + 1
            Get NFile, offset + 5 + .AnimationLengthLong + 12, .UnknownData
        End If

        
        offset = offset + .BlockLength + 12
    End With
    Exit Sub
ErrorHand:
    'Debug.Print "Error "; Err; " reading frame "; fi
    offset = offset + Animation.BlockLength + 12
    Animation.NumFrames2 = fi
    Err.Raise (1000 + fi)
End Sub
Sub WriteDAAnimation(ByVal NFile As Integer, ByRef offset As Long, ByRef Animation As DAAnimation)
    Dim AnimationStream() As Byte
    Dim offsetBit As Long
    Dim fi As Integer
    Dim AnimationCarry() As Point3D
    Dim BI As Integer
    Dim num_bones As Integer
    Dim offset_correction As Long
    
    With Animation
        offset_correction = 0
        
        Put NFile, offset, .NumBonesModel
        Put NFile, offset + 4, .NumFrames1
        If .BlockLength < 11 Or .NumFrames2 = 0 Then
            'The animation is either an empty slot or somehow corrupted.
            Put NFile, offset + 8, .BlockLength
            Put NFile, offset + 12, .UnknownData
        Else
            If Not .MissingNumFrames2 Then
                Put NFile, offset + 12, .NumFrames2
            Else
                offset_correction = -2
            End If
            'We don't know yet the value of AnimationLength, so write it later
            
            'Find highest key without exceding the maximum animation length
            .key = 0
            Do
                ReDim AnimationStream(0)
                offsetBit = 0
                'num_bones = UBound(.Frames(0).Bones) + 1
                'ReDim AnimationCarry(num_bones - 1)
                'For BI = 0 To num_bones - 1
                '    AnimationCarry(BI).x = 0
                '    AnimationCarry(BI).y = 0
                '    AnimationCarry(BI).z = 0
                'Next BI
                WriteDAUncompressedFrame AnimationStream, offsetBit, .key, .Frames(0)
                For fi = 1 To .NumFrames2 - 1
                    'Debug.Print "   Frame "; Str$(fi)
                    WriteDAFrame AnimationStream, offsetBit, .key, .Frames(fi), .Frames(fi - 1) ', AnimationCarry
                Next fi
                
                .BlockLength = UBound(AnimationStream) + 1 + 5
                .AnimationLengthLong = UBound(AnimationStream) + 1
                .key = .key + 2
            Loop Until .AnimationLengthLong <= 65535 Or .key > 4
            .key = .key - 2
            Put NFile, offset + 16 + offset_correction, .key
            If .AnimationLengthLong > 65535 Then
                MsgBox "Can't save this animation because it's too big for FF7 battle animations format", vbOKOnly, "Error saving animation"
                'Err.Raise 1000
                Exit Sub
            End If
            If SafeArrayGetDim(.UnknownData) > 0 Then
                Put NFile, offset + 12 + 5 + .AnimationLengthLong + offset_correction, .UnknownData
                .BlockLength = .BlockLength + UBound(.UnknownData) + 1 + offset_correction
            End If
            Put NFile, offset + 8, .BlockLength
            CopyMemory .AnimationLength, .AnimationLengthLong, 2
            Put NFile, offset + 14 + offset_correction, .AnimationLength
            Put NFile, offset + 17 + offset_correction, AnimationStream
        End If
        offset = offset + 12 + .BlockLength
    End With
End Sub
Sub CreateEmptyDAAnimationsPackAnimation(ByRef Anim As DAAnimation, ByVal NumBones As Integer)
    With Anim
        .NumFrames1 = 1
        .NumFrames2 = 1
        ReDim .Frames(0)
        CreateEmptyDAAnimationsPackAnimation1stFrame .Frames(0), NumBones
    End With
End Sub
'Ensure the animation delta between two consecutive frames stays always on the (-180º, 180º) boundary (it's not possible to encode values outside)
Sub NormalizeDAAnimationsPackAnimation(ByRef Anim As DAAnimation)
    Dim fi As Integer
    
    For fi = 0 To Anim.NumFrames2 - 2
        NormalizeDAAnimationsPackAnimationFrame Anim.Frames(fi), Anim.Frames(fi + 1)
    Next fi
End Sub

Sub CheckWriteDAAnimation(ByVal NFile As Integer, ByRef offset As Long, ByRef Animation As DAAnimation)
    Dim fi As Integer
    Dim AnimationStream() As Byte
    Dim offsetBit As Long
    Dim num_bones As Integer
    Dim sanity_check As Boolean
    Dim err_msg As String
    Dim missing_frame_counter As Boolean
    Dim last_offsetBit As Long
    
    Dim NumBonesModel As Long           'Number of bones for the model + 1 (root transformation). NOT RELIABLE
    Dim NumFrames1 As Long              'Usually wrong, so refrain from using it.
    Dim BlockLength As Long

    Dim NumFrames2 As Integer           'This one is the real deal.
    Dim AnimationLength As Integer      'Don't use this field EVER. It would be interpreted as a signed value.
    Dim AnimationLengthLong As Long     'This isn't part of the actual structure, it's used just to overcome the lack of unsigned shorts support
    Dim key As Byte
    
    Dim Frames() As DAFrame
    
    With Animation
        Get NFile, offset, NumBonesModel
        If NumBonesModel <> .NumBonesModel Then
            Debug.Print "Error"
        End If
        Get NFile, offset + 4, NumFrames1
        If NumFrames1 <> .NumFrames1 Then
            Debug.Print "Error"
        End If
        Get NFile, offset + 8, BlockLength
        If BlockLength <> .BlockLength Then
            Debug.Print "Error"
        End If
        ''Debug.Print .BlockLength
        If BlockLength < 11 Then

            offset = offset + 12 + BlockLength
            NumFrames2 = 0
            If NumFrames2 <> .NumFrames2 Then
                Debug.Print "Error"
            End If
            
            Exit Sub
        End If
        Get NFile, offset + 12, NumFrames2
        If NumFrames2 <> .NumFrames2 Then
               Debug.Print "Error"
        End If
        Get NFile, offset + 14, AnimationLength
        If AnimationLength <> .AnimationLength Then
            Debug.Print "Error"
        End If
        CopyMemory AnimationLengthLong, AnimationLength, 2
        Get NFile, offset + 16, .key
        If key <> .key Then
            Debug.Print "Error"
        End If
        
        'Hack for reading animations with missing secondary frame counter (which can't be actually used by FF7)
        If NumFrames2 = BlockLength - 5 Then
            Get NFile, offset + 12, AnimationLength
            If AnimationLength <> .AnimationLength Then
               Debug.Print "Error"
            End If
            CopyMemory AnimationLengthLong, AnimationLength, 2
            Get NFile, offset + 14, .key
            If key <> .key Then
               Debug.Print "Error"
            End If
            ReDim AnimationStream(AnimationLengthLong)
            Get NFile, offset + 15, AnimationStream
            missing_frame_counter = True
        Else
            ReDim AnimationStream(AnimationLengthLong)
            Get NFile, offset + 17, AnimationStream
            missing_frame_counter = False
        End If
        
        offsetBit = 0
        
        num_bones = UBound(.Frames(0).Bones) + 1
        CheckWriteDAUncompressedFrame AnimationStream, offsetBit, key, .Frames(0)
        For fi = 1 To NumFrames2 - 1
            If fi = 3 Then
                fi = fi
            End If
            last_offsetBit = offsetBit
            CheckWriteDAFrame AnimationStream, offsetBit, key, .Frames(fi), .Frames(fi - 1)
        Next fi
        
        offset = offset + .BlockLength + 12
    End With
End Sub
