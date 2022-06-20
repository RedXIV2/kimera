Attribute VB_Name = "FF7AAnimationRotation"
Option Explicit
Type ARotation
    alpha As Single
    Beta As Single
    Gamma As Single
End Type
Sub ReadARotation(ByVal NFile As Integer, ByVal offset As Long, ByRef rot As ARotation)
    Get NFile, offset, rot
End Sub
Sub WriteARotation(ByVal NFile As Integer, ByVal offset As Long, ByRef rot As ARotation)
    Put NFile, offset, rot
End Sub
Sub CopyARotation(ByRef rot_in As ARotation, ByRef rot_out As ARotation)
    With rot_in
        rot_out.alpha = .alpha
        rot_out.Beta = .Beta
        rot_out.Gamma = .Gamma
    End With
End Sub

Function IsBrokenARotation(ByRef rot As ARotation) As Boolean
    IsBrokenARotation = False

    With rot
        If IsNan(.alpha) Then
            IsBrokenARotation = True
        ElseIf .alpha > 9999# Then
            IsBrokenARotation = True
        ElseIf .alpha < -9999# Then
            IsBrokenARotation = True
        End If
        If IsNan(.Beta) Then
            IsBrokenARotation = True
        ElseIf .Beta > 9999# Then
            IsBrokenARotation = True
        ElseIf .Beta < -9999# Then
            IsBrokenARotation = True
        End If
        If IsNan(.Gamma) Then
            IsBrokenARotation = True
        ElseIf .Gamma > 9999# Then
            IsBrokenARotation = True
        ElseIf .Gamma < -9999# Then
            IsBrokenARotation = True
        End If
    End With
End Function


