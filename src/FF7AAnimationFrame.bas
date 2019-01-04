Attribute VB_Name = "FF7AAnimationFrame"
Option Explicit
Type AFrame
    'A Frame format by Mirex and Aali
    'http://wiki.qhimm.com/FF7/Field_Module#.22A.22_Field_Animation_Files_for_PC_by_Mirex_.28Edits_by_Aali.29
    RootRotationAlpha As Single
    RootRotationBeta As Single
    RootRotationGamma As Single
    RootTranslationX As Single
    RootTranslationY As Single
    RootTranslationZ As Single
    Rotations() As ARotation
End Type
Sub ReadAFrame(ByVal NFile As Integer, ByVal offset As Long, ByVal NumBones As Integer, ByRef obj As AFrame)
    Dim BI As Integer
    
    With obj
        ReDim obj.Rotations(NumBones)
        
        Get NFile, offset, .RootRotationAlpha
        Get NFile, offset + 4, .RootRotationBeta
        Get NFile, offset + 8, .RootRotationGamma
        Get NFile, offset + 12, .RootTranslationX
        Get NFile, offset + 16, .RootTranslationY
        Get NFile, offset + 20, .RootTranslationZ
        For BI = 0 To NumBones - 1
            ReadARotation NFile, offset + 24 + BI * 12, .Rotations(BI)
        Next BI
    End With
End Sub
Sub WriteAFrame(ByVal NFile As Integer, ByVal offset As Long, ByVal NumBones As Integer, ByRef obj As AFrame)
    Dim BI As Integer
    
    With obj
        Put NFile, offset, .RootRotationAlpha
        Put NFile, offset + 4, .RootRotationBeta
        Put NFile, offset + 8, .RootRotationGamma
        Put NFile, offset + 12, .RootTranslationX
        Put NFile, offset + 16, .RootTranslationY
        Put NFile, offset + 20, .RootTranslationZ
        For BI = 0 To NumBones - 1
            WriteARotation NFile, offset + 24 + BI * 12, .Rotations(BI)
        Next BI
    End With
End Sub
Sub CopyAFrame(ByRef frame_in As AFrame, ByRef frame_out As AFrame)
    Dim NumBones As Integer
    Dim BI As Integer
    
    With frame_in
        NumBones = UBound(.Rotations) + 1
        ReDim frame_out.Rotations(NumBones - 1)
    
        frame_out.RootRotationAlpha = .RootRotationAlpha
        frame_out.RootRotationBeta = .RootRotationBeta
        frame_out.RootRotationGamma = .RootRotationGamma
        frame_out.RootTranslationX = .RootTranslationX
        frame_out.RootTranslationY = .RootTranslationY
        frame_out.RootTranslationZ = .RootTranslationZ
        For BI = 0 To NumBones - 1
            CopyARotation .Rotations(BI), frame_out.Rotations(BI)
        Next BI
    End With
End Sub
Function IsBrokenAAFrame(ByRef Frame As AFrame, ByVal num_bones As Integer) As Boolean
    Dim BI As Integer
    
    IsBrokenAAFrame = False
    
    With Frame
        If IsNan(.RootTranslationX) Then
            IsBrokenAAFrame = True
        End If
        If IsNan(.RootTranslationY) Then
            IsBrokenAAFrame = True
        End If
        If IsNan(.RootTranslationZ) Then
            IsBrokenAAFrame = True
        End If
        
        If IsNan(.RootRotationAlpha) Then
            IsBrokenAAFrame = True
        ElseIf .RootRotationAlpha > 9999# Then
            IsBrokenAAFrame = True
        ElseIf .RootRotationAlpha < -9999# Then
            IsBrokenAAFrame = True
        End If
        If IsNan(.RootRotationBeta) Then
            IsBrokenAAFrame = True
        ElseIf .RootRotationBeta > 9999# Then
            IsBrokenAAFrame = True
        ElseIf .RootRotationBeta < -9999# Then
            IsBrokenAAFrame = True
        End If
        If IsNan(.RootRotationGamma) Then
            IsBrokenAAFrame = True
        ElseIf .RootRotationGamma > 9999# Then
            IsBrokenAAFrame = True
        ElseIf .RootRotationGamma < -9999# Then
            IsBrokenAAFrame = True
        End If
        
        If Not IsBrokenAAFrame Then
            For BI = 0 To num_bones - 1
                IsBrokenAAFrame = IsBrokenARotation(.Rotations(BI))
                If IsBrokenAAFrame Then
                    Exit For
                End If
            Next BI
        End If
    End With
End Function
