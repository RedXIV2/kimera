Attribute VB_Name = "Module5"
 Option Explicit
 Type color
    red As GLbyte
    green As GLbyte
    blue As GLbyte
    mask As GLbyte
End Type

Type Pallette
    nbColors As GLint
    colors() As color
End Type
Sub ConvertFF7Color2RGB(ByVal src As GLshort, ByRef dest As color)
    With dest
        .red = LShiftLong(src And 31!, 3)
        If .red > 0 Then .red = .red + 7
        .green = LShiftLong(RShiftLong(src And 992!, 5), 3)
        If .green > 0 Then .green = .green + 7
        .blue = LShiftLong(RShiftLong(src And 31744!, 10), 3)
        If .blue > 0 Then .blue = .blue + 7
        
        If src < 0 Then
            .mask = 1
        Else
            .mask = 0
        End If
        
    End With
End Sub
Function ConvertRGB2FF7Color(ByRef src As color) As GLushort
    With src
        If .mask = 1 Then
            ConvertRGB2FF7Color = &H8000
        Else
            ConvertRGB2FF7Color = 0
        End If
        
        ConvertRGB2FF7Color = ConvertRGB2FF7Color Or _
                              RShiftLong(.red, 3) And &H1F Or _
                              LShiftLong(RShiftLong(.green, 3) And &H1F, 5) Or _
                              LShiftLong(RShiftLong(.blue, 3) And &H1F, 10)
    End With
End Function
Sub ConvertWinColor2RGB(ByVal src As GLint, ByRef dest As color)
    With dest
        .red = (src And &HFF)
        .green = RShiftLong(src, 8) And &HFF
        .blue = RShiftLong(src, 16) And &HFF
        .mask = 0
    End With
End Sub
Sub Load_Pallette(ByVal n_file As Integer, ByVal offset As GLuint, ByVal nbColors As GLuint, ByRef pal As Pallette)
    Dim c_i As GLuint
    Dim c_temp As GLushort
    
    With pal
        .nbColors = nbColors
        ReDim .colors(nbColors)
    
        For c_i = 0 To nbColors
            Get n_file, offset + c_i * 2, c_temp
            ConvertFF7Color2RGB c_temp, .colors(c_i)
            If .colors(c_i).red = 0 And .colors(c_i).green = 255 And .colors(c_i).blue = 0 Then _
                Debug.Print "Paleta " + Str$(c_i \ 256) + " Color Verde" + Str$(c_i Mod 256) + " Antes " + Str$(c_temp)
            
            If .colors(c_i).red = 63 And .colors(c_i).green = 55 And .colors(c_i).blue = 39 Then _
                Debug.Print "Paleta " + Str$(c_i \ 256) + " Color Marron" + Str$(c_i Mod 256) + " Antes " + Str$(c_temp)
        Next c_i
    End With
End Sub
Sub Save_Pallette(ByVal n_file As Integer, ByVal offset As GLuint, ByRef pal As Pallette)
    Dim c_i As GLuint
    Dim c_temp As GLushort
    
    With pal
        For c_i = 0 To .nbColors - 1
            c_temp = ConvertRGB2FF7Color(.colors(c_i))
            Put n_file, offset + c_i * 2, c_temp
        Next c_i
    End With
End Sub
Function addColor(ByRef pal As Pallette, ByVal pal_i As Integer, ByRef color As color) As Byte
    'ReDim Preserve pal.colors(UBound(pal.colors()) + 1)
    Dim c_i As Long
    Dim c_if As Long
    Dim rv As Long
    Dim bv As Long
    Dim gv As Long
    Dim passed As Boolean
    Dim min_cost_c As Long
    Dim cost_c As Long
    
    rv = 1
    bv = 1
    gv = 1
    With color
        If .red > .blue Then
            rv = 256
            If .red > .green Then
                rv = 65536
                If .blue > .green Then
                    bv = 256
                Else
                    gv = 256
                End If
            Else
                gv = 65536
            End If
        Else
            bv = 256
            If .blue > .green Then
                bv = 65536
                If .red > .green Then
                    rv = 256
                Else
                    gv = 256
                End If
            Else
                gv = 65536
            End If
        End If
    End With
    
    min_cost_c = 2147483647
    passed = False
    For c_i = pal_i * 256 To pal_i * 256 + 255
        With pal.colors(c_i)
            If .red = 0 And .blue = 0 And .green = 0 And .mask = 0 Then
                If passed Then
                    .red = color.red
                    .blue = color.blue
                    .green = color.green
                    .mask = color.mask
                    c_if = c_i
                    Exit For
                Else
                    passed = True
                End If
            End If
            If .mask = color.mask Then
                cost_c = Abs(.red * 1# - color.red) * rv + Abs(.blue * 1# - color.blue) * bv + Abs(.green * 1# - color.green) * gv
                If cost_c < min_cost_c Then
                    min_cost_c = cost_c
                    c_if = c_i
                End If
            End If
        End With
    Next c_i
    
    addColor = c_if Mod 256
End Function
