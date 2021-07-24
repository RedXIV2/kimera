Attribute VB_Name = "FF7PModelBoundingBox"
Option Explicit
Type PBoundingBox
    max_x As Single
    max_y As Single
    max_z As Single
    min_x As Single
    min_y As Single
    min_z As Single
End Type
Sub ReadBoundingBox(ByVal NFile As Integer, ByVal offset As Long, ByRef BoundingBox As PBoundingBox)
    With BoundingBox
        Get NFile, offset, .max_x
        Get NFile, offset + 4, .max_y
        Get NFile, offset + 4 * 2, .max_z
        Get NFile, offset + 4 * 3, .min_x
        Get NFile, offset + 4 * 4, .min_y
        Get NFile, offset + 4 * 5, .min_z
    End With
End Sub
Sub WriteBoundingBox(ByVal NFile As Integer, ByVal offset As Long, ByRef BoundingBox As PBoundingBox)
    With BoundingBox
        Put NFile, offset, .max_x
        Put NFile, offset + 4, .max_y
        Put NFile, offset + 4 * 2, .max_z
        Put NFile, offset + 4 * 3, .min_x
        Put NFile, offset + 4 * 4, .min_y
        Put NFile, offset + 4 * 5, .min_z
    End With
End Sub
Sub MergeBoundingBox(ByRef b1 As PBoundingBox, ByRef b2 As PBoundingBox)
    With b1
        If .max_x < b2.max_x Then .max_x = b2.max_x
        If .max_y < b2.max_y Then .max_y = b2.max_y
        If .max_z < b2.max_z Then .max_z = b2.max_z

        If .min_x > b2.min_x Then .min_x = b2.min_x
        If .min_y > b2.min_y Then .min_y = b2.min_y
        If .min_z > b2.min_z Then .min_z = b2.min_z
    End With
End Sub
Function ComputeDiameter(ByRef BoundingBox As PBoundingBox) As Single
    Dim diffx As Single
    Dim diffy As Single
    Dim diffz As Single

    With BoundingBox
        diffx = .max_x - .min_x
        diffy = .max_y - .min_y
        diffz = .max_z - .min_z
    End With

    If diffx > diffy Then
        If diffx > diffz Then
            ComputeDiameter = diffx
        Else
            ComputeDiameter = diffz
        End If
    Else
        If diffy > diffz Then
            ComputeDiameter = diffy
        Else
            ComputeDiameter = diffz
        End If
    End If

End Function
Sub DrawPBoundingBox(ByRef BoundingBox As PBoundingBox)
    With BoundingBox
        DrawBox .max_x, .max_y, .max_z, .min_x, .min_y, .min_z, 0, 1, 0
    End With
End Sub
