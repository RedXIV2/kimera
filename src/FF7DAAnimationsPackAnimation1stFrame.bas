Attribute VB_Name = "FF7DAAnimationsPackAnimationFrame"
Option Explicit
Type DAFrame
    X_start As Long
    Y_start As Long
    Z_start As Long
    Bones() As DAFrameBone
End Type
Sub ReadDAUncompressedFrame(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByVal BonesVectorLength As Integer, ByRef Frame As DAFrame)
    Dim BI As Integer
    Dim tempV As Long
    Dim aux As Integer
    
    aux = offsetBit
    With Frame
        '.BonesVectorLength = BonesVectorLength '+ IIf(NumBones = 1, 0, 1)
        'NumBones = IIf(NumBones = 2, 1, NumBones)  'Some single bone models have bones counter of 2 instead of 1, so adjust the value for convenience.
        
        .X_start = GetBitBlockV(AnimationStream, 16, offsetBit)
        
        .Y_start = GetBitBlockV(AnimationStream, 16, offsetBit)
        
        .Z_start = GetBitBlockV(AnimationStream, 16, offsetBit)
        
        'Debug.Print "       Position delta "; Str$(.X_start); ", "; Str$(.Y_start); ", "; Str$(.Z_start)
        
        ReDim .Bones(BonesVectorLength - 1)
        For BI = 0 To BonesVectorLength - 1
            'Debug.Print "       Bone "; Str$(bi)
            ReadDAUncompressedFrameBone AnimationStream, offsetBit, key, .Bones(BI)
        Next BI
        ''Debug.Print "diff: "; offsetBit - aux
    End With
End Sub
Function ReadDAFrame(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByVal BonesVectorLength As Integer, ByRef Frame As DAFrame, ByRef LastFrame As DAFrame) As Boolean
    Dim BI As Integer
    Dim oi As Integer
    Dim offLength As Integer
    
    On Error GoTo OutOfDataHandler

    With Frame
        For oi = 0 To 2
            Select Case (GetBitBlockV(AnimationStream, 1, offsetBit) And 1)
                Case 0:
                    offLength = 7
                Case 1:
                    offLength = 16
                Case Else:
                    'Debug.Print "What?!"
            End Select

            Select Case (oi)
                Case 0:
                    .X_start = GetBitBlockV(AnimationStream, offLength, offsetBit) + LastFrame.X_start
                Case 1:
                    .Y_start = GetBitBlockV(AnimationStream, offLength, offsetBit) + LastFrame.Y_start
                Case 2:
                    .Z_start = GetBitBlockV(AnimationStream, offLength, offsetBit) + LastFrame.Z_start
                Case Else:
                    'Debug.Print "What?!"
            End Select

        Next oi
        
        'Debug.Print "       Position delta "; Str$(.X_start); ", "; Str$(.Y_start); ", "; Str$(.Z_start)
    
        ReDim .Bones(BonesVectorLength - 1)
        
        For BI = 0 To BonesVectorLength - 1
            'Debug.Print "       Bone "; Str$(bi)
            ReadDAFrameBone AnimationStream, offsetBit, key, .Bones(BI), LastFrame.Bones(BI)
        Next BI
        ''Debug.Print "diff: "; offsetBit - aux
    End With
    ReadDAFrame = True
    Exit Function
OutOfDataHandler:
    ReadDAFrame = False
End Function
Sub WriteDAUncompressedFrame(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Integer, ByRef Frame As DAFrame)
    Dim BI As Integer
    Dim tempV As Long
    Dim num_bones As Integer
    
    With Frame
        '.NumBones = NumBones + IIf(NumBones = 1, 0, 1)
        
        'Debug.Print "       Position delta "; Str$(.X_start); ", "; Str$(.Y_start); ", "; Str$(.Z_start)
        
        PutBitBlockV AnimationStream, 16, offsetBit, .X_start
        
        PutBitBlockV AnimationStream, 16, offsetBit, .Y_start
        
        PutBitBlockV AnimationStream, 16, offsetBit, .Z_start
        
        num_bones = UBound(.Bones) + 1
        For BI = 0 To num_bones - 1
            'Debug.Print "       Bone "; Str$(bi)
            WriteDAUncompressedFrameBone AnimationStream, offsetBit, key, .Bones(BI)
        Next BI
    End With
End Sub
Sub WriteDAFrame(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef Frame As DAFrame, ByRef LastFrame As DAFrame) ', ByRef AnimationCarry() As Point3D)
    Dim BI As Integer
    Dim oi As Integer
    Dim offLength As Integer
    Dim off_delta As Long
    Dim num_bones As Integer
    
    With Frame
        'Debug.Print "       Position delta "; Str$(.X_start - LastFrame.X_start); ", "; Str$(.Y_start - LastFrame.Y_start); ", "; Str$(.Z_start - LastFrame.Z_start)
        For oi = 0 To 2
            Select Case (oi)
                Case 0:
                    off_delta = .X_start - LastFrame.X_start
                Case 1:
                    off_delta = .Y_start - LastFrame.Y_start
                Case 2:
                    off_delta = .Z_start - LastFrame.Z_start
                Case Else:
                    'Debug.Print "What?!"
            End Select
            
            If (off_delta < 2 ^ (7 - 1) And off_delta >= -2 ^ (7 - 1)) Then
                offLength = 7
                PutBitBlockV AnimationStream, 1, offsetBit, 0
            Else
                offLength = 16
                PutBitBlockV AnimationStream, 1, offsetBit, 1
            End If
            
            PutBitBlockV AnimationStream, offLength, offsetBit, off_delta
        Next oi
        
        num_bones = UBound(.Bones) + 1
        For BI = 0 To num_bones - 1
            'Debug.Print "       Bone "; Str$(bi)
            WriteDAFrameBone AnimationStream, offsetBit, key, .Bones(BI), LastFrame.Bones(BI) ', AnimationCarry(BI)
        Next BI
    End With
End Sub
Sub CreateEmptyDAAnimationsPackAnimation1stFrame(ByRef Frame As DAFrame, ByVal NumBones As Integer)
    With Frame
        ReDim .Bones(NumBones - 1)
    End With
End Sub
Sub CopyDAFrame(ByRef frame_in As DAFrame, frame_out As DAFrame)
    Dim BI As Integer
    Dim num_bones As Integer
    With frame_in
        frame_out.X_start = .X_start
        frame_out.Y_start = .Y_start
        frame_out.Z_start = .Z_start
        
        num_bones = UBound(frame_in.Bones) + 1
        ReDim frame_out.Bones(num_bones - 1)
        For BI = 0 To num_bones - 1
            CopyDAFrameBone .Bones(BI), frame_out.Bones(BI)
        Next BI
    End With
End Sub
Sub NormalizeDAAnimationsPackAnimationFrame(ByRef Frame As DAFrame, ByRef next_frame As DAFrame)
    Dim BI As Integer
    Dim num_bones As Integer
    
    num_bones = UBound(Frame.Bones) + 1
    For BI = 0 To num_bones - 1
        NormalizeDAAnimationsPackAnimationFrameBone Frame.Bones(BI), next_frame.Bones(BI)
    Next BI
End Sub

Sub CheckWriteDAUncompressedFrame(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef ref_frame As DAFrame)
    Dim BI As Integer
    Dim tempV As Long
    Dim aux As Integer
    
    Dim X_start As Long
    Dim Y_start As Long
    Dim Z_start As Long
    
    Dim num_bones As Integer
    
    aux = offsetBit
    With ref_frame
        X_start = GetBitBlockV(AnimationStream, 16, offsetBit)
        If X_start <> .X_start Then
            Debug.Print "Error"
        End If
        
        Y_start = GetBitBlockV(AnimationStream, 16, offsetBit)
        If Y_start <> .Y_start Then
            Debug.Print "Error"
        End If
        
        Z_start = GetBitBlockV(AnimationStream, 16, offsetBit)
        If Z_start <> .Z_start Then
            Debug.Print "Error"
        End If
        
        'Debug.Print "       Position delta "; Str$(.X_start); ", "; Str$(.Y_start); ", "; Str$(.Z_start)
        
        num_bones = UBound(.Bones) + 1
        For BI = 0 To num_bones - 1
            'Debug.Print "       Bone "; Str$(bi)
            CheckWriteDAUncompressedFrameBone AnimationStream, offsetBit, key, .Bones(BI)
        Next BI
        ''Debug.Print "diff: "; offsetBit - aux
    End With
End Sub
Sub CheckWriteDAFrame(ByRef AnimationStream() As Byte, ByRef offsetBit As Long, ByVal key As Byte, ByRef ref_frame As DAFrame, ByRef ref_last_frame As DAFrame)
    Dim BI As Integer
    Dim oi As Integer
    Dim offLength As Integer
    
    Dim X_start As Long
    Dim Y_start As Long
    Dim Z_start As Long
    
    Dim num_bones As Integer

    With ref_frame
        For oi = 0 To 2
            Select Case (GetBitBlockV(AnimationStream, 1, offsetBit) And 1)
                Case 0:
                    offLength = 7
                Case 1:
                    offLength = 16
                Case Else:
                    'Debug.Print "What?!"
            End Select

            Select Case (oi)
                Case 0:
                    X_start = GetBitBlockV(AnimationStream, offLength, offsetBit) + ref_last_frame.X_start
                    If X_start <> .X_start Then
                        Debug.Print "Error"
                    End If
                Case 1:
                    Y_start = GetBitBlockV(AnimationStream, offLength, offsetBit) + ref_last_frame.Y_start
                    If Y_start <> .Y_start Then
                        Debug.Print "Error"
                    End If
                Case 2:
                    Z_start = GetBitBlockV(AnimationStream, offLength, offsetBit) + ref_last_frame.Z_start
                    If Z_start <> .Z_start Then
                        Debug.Print "Error"
                    End If
                Case Else:
                    'Debug.Print "What?!"
            End Select

        Next oi
        
        'Debug.Print "       Position delta "; Str$(.X_start); ", "; Str$(.Y_start); ", "; Str$(.Z_start)
    
        num_bones = UBound(.Bones) + 1
        For BI = 0 To num_bones - 1
            'Debug.Print "       Bone "; Str$(bi)
            CheckWriteDAFrameBone AnimationStream, offsetBit, key, .Bones(BI), ref_last_frame.Bones(BI)
        Next BI
        ''Debug.Print "diff: "; offsetBit - aux
    End With
End Sub
