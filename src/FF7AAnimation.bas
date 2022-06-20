Attribute VB_Name = "FF7AAnimation"
Option Explicit
Type AAnimation
    'A Animation format by Mirex and Aali
    'http://wiki.qhimm.com/FF7/Field_Module#.22A.22_Field_Animation_Files_for_PC_by_Mirex_.28Edits_by_Aali.29
    version As Long     'Must be one for FF7 to load it
    NumFrames As Long
    NumBones As Long
    RotationOrder(2) As Byte 'Rotation order determines in which order rotations _
                                are applied, 0 means alpha rotation, 1 beta rotation and _
                                2 gamma rotation
    unused As Byte
    runtime_data(4) As Long
    Frames() As AFrame
    AFile As String
End Type
Function ReadAAnimationNBones(ByVal fileName As String) As Integer
    Dim fi As Long
    Dim fileNumber As Integer

    On Error GoTo errorH
    If fileName = "" Then
        ReadAAnimationNBones = -1
        Exit Function
    End If

    fileNumber = FreeFile
    Open fileName For Binary As fileNumber

    Get fileNumber, 9, ReadAAnimationNBones

    Close fileNumber
    Exit Function
errorH:
    MsgBox "Error opening " + fileName, vbCritical, "A Error " + Str$(Err)
End Function
Sub ReadAAnimation(ByRef obj As AAnimation, ByVal fileName As String)
    Dim fi As Long
    Dim fileNumber As Integer

    On Error GoTo errorH

    fileNumber = FreeFile
    Open fileName For Binary As fileNumber

    With obj
        .AFile = Right$(fileName, Len(fileName) - Len(GetPathFromString(fileName)))
        Get fileNumber, 1, .version
        Get fileNumber, 5, .NumFrames
        Get fileNumber, 9, .NumBones
        Get fileNumber, 13, .RotationOrder
        Get fileNumber, 16, .unused
        Get fileNumber, 17, .runtime_data

        'If .NumFrames > 1 Then
        '    .NumFrames = .NumFrames * 2
        'End If

        ReDim .Frames(.NumFrames)

        ReadAFrame fileNumber, 25 + 12 + fi * (.NumBones * 12 + 24), .NumBones, .Frames(0)

        For fi = 0 To .NumFrames - 1
            ReadAFrame fileNumber, 25 + 12 + fi * (.NumBones * 12 + 24), .NumBones, .Frames(fi)
        Next fi
    End With
    Close fileNumber
    Exit Sub
errorH:
    MsgBox "Error opening " + fileName, vbCritical, "A Error " + Str$(Err)
End Sub
Sub WriteAAnimation(ByRef obj As AAnimation, ByVal fileName As String)
    Dim fi As Integer
    Dim fileNumber As Integer

    On Error GoTo errorH

    If fileName = "" Then _
        fileName = "dummy_animation.a"
    fileNumber = FreeFile
    Open fileName For Output As fileNumber
    Close fileNumber
    Open fileName For Binary As fileNumber

    With obj
        Put fileNumber, 1, .version
        Put fileNumber, 5, .NumFrames
        Put fileNumber, 9, .NumBones
        'Is there any animation using another rotation order?
        .RotationOrder(0) = 1
        .RotationOrder(1) = 0
        .RotationOrder(2) = 2
        Put fileNumber, 13, .RotationOrder
        Put fileNumber, 16, .unused
        Put fileNumber, 17, .runtime_data

        For fi = 0 To .NumFrames - 1
            WriteAFrame fileNumber, 25 + 12 + (fi) * (.NumBones * 12 + 24), .NumBones, .Frames(fi)
        Next fi
    End With
    Close fileNumber
    Exit Sub
errorH:
    MsgBox "Error writting " + fileName, vbCritical, "A Error " + Str$(Err)
End Sub
Sub CreateCompatibleEmptyAAnimation(ByRef obj As AAnimation, ByVal NumBones As Integer)
    Dim i As Integer

    With obj
        .AFile = ""
        .NumBones = NumBones
        .NumFrames = 1
        ReDim .Frames(1)
        ReDim .Frames(0).Rotations(NumBones)
    End With
    For i = 0 To NumBones - 1
        With obj.Frames(0).Rotations(i)
            .alpha = 0
            .Beta = 0
            .Gamma = 0
        End With
    Next i
End Sub
Function RemoveFrameAAnimation(ByRef obj As AAnimation, ByVal frame_index As Integer) As Boolean
    Dim fi As Integer

    With obj
        If .NumFrames > 1 Then
            For fi = frame_index To .NumFrames - 2
                .Frames(fi) = .Frames(fi + 1)
            Next fi
            .NumFrames = .NumFrames - 1
            ReDim Preserve .Frames(.NumFrames - 1)

            RemoveFrameAAnimation = True
        Else
            RemoveFrameAAnimation = False
        End If
    End With
End Function
Function IsFrameBrokenAAnimation(ByRef obj As AAnimation, ByVal frame_index As Integer) As Boolean
    IsFrameBrokenAAnimation = IsBrokenAAFrame(obj.Frames(frame_index), obj.NumBones)
End Function
