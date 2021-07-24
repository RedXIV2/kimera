Attribute VB_Name = "FF7PModelHundretsPool"
'PHundret structure kindly provided by Aali.
'http://forums.qhimm.com/index.php?topic=4194.msg110965#msg110965
Option Explicit
Type PHundret
    field_0 As Long
    field_4 As Long
    field_8 As Long                     'Render state part value (if it's actually changed)
    field_C As Long                     'Change render state part?
                                        'Masks for fields 8 and C:
                                            '0x1: V_WIREFRAME
                                            '0x2: V_TEXTURE
                                            '0x4: V_LINEARFILTER
                                            '0x8: V_PERSPECTIVE
                                            '0x10: V_TMAPBLEND
                                            '0x20: V_WRAP_U
                                            '0x40: V_WRAP_V
                                            '0x80: V_UNKNOWN80
                                            '0x100: V_COLORKEY
                                            '0x200: V_DITHER
                                            '0x400: V_ALPHABLEND
                                            '0x800: V_ALPHATEST
                                            '0x1000: V_ANTIALIAS
                                            '0x2000: V_CULLFACE
                                            '0x4000: V_NOCULL
                                            '0x8000: V_DEPTHTEST
                                            '0x10000: V_DEPTHMASK
                                            '0x20000: V_SHADEMODE
                                            '0x40000: V_SPECULAR
                                            '0x80000: V_LIGHTSTATE
                                            '0x100000: V_FOG
                                            '0x200000: V_TEXADDR
    TexID As Long                       'Texture identifier for the corresponding group. For consistency sake should be the same as on the group,
                                        'but this is the one FF7 actually uses.
    texture_set_pointer As Long         'This should be filled in real time
    field_18 As Long
    field_1C As Long
    field_20 As Long
    shademode As Long
    lightstate_ambient As Long
    field_2C As Long
    lightstate_material_pointer As Long 'This should be filled in real time
    srcblend As Long
    destblend As Long
    field_3C As Long
    alpharef As Long
    blend_mode As Long  '0 - Average, source color / 2 + destination color / 2.
                        '1 - Additive, source color + destination color.
                        '2 - Subtractive, broken and unused but it should be destination color - source color.
                        '3 - Not sure, but it seems broken and is never used.
                        '4 - No blending (FF8 only)
    zsort As Long       'Filled in real time
    field_4C As Long
    field_50 As Long
    field_54 As Long
    field_58 As Long
    vertex_alpha As Long
    field_60 As Long
End Type
Sub ReadHundrets(ByVal NFile As Integer, ByVal offset As Long, ByRef hundrets() As PHundret, ByVal NumHundrets As Long)
    ReDim hundrets(NumHundrets - 1)
    Get NFile, offset, hundrets
    Dim hi As Integer
    For hi = 0 To NumHundrets - 1
        With hundrets(hi)
            'Debug.Print "field_0 = "; .field_0
            'Debug.Print "field_4 = "; .field_4
            'Debug.Print "field_8 = "; .field_8
            'Debug.Print "field_C = "; .field_C
            'Debug.Print "TexID = "; .TexID
            'Debug.Print "texture_set_pointer = "; .texture_set_pointer
            'Debug.Print "field_18 = "; .field_18
            'Debug.Print "field_1C = "; .field_1C
            'Debug.Print "field_20 = "; .field_20
            'Debug.Print "shademode = "; .shademode
            'Debug.Print "lightstate_ambient = "; .lightstate_ambient
            'Debug.Print "field_2C = "; .field_2C
            'Debug.Print "lightstate_material_pointer = "; .lightstate_material_pointer
            'Debug.Print "field_2C = "; .field_2C
            'Debug.Print "srcblend = "; .srcblend
            'Debug.Print "destblend = "; .destblend
            'Debug.Print "field_3C = "; .field_3C
            'Debug.Print "alpharef = "; .alpharef
            'Debug.Print "blend_mode = "; .blend_mode
            'Debug.Print "zsort = "; .zsort
            'Debug.Print "field_4C = "; .field_4C
            'Debug.Print "field_50 = "; .field_50
            'Debug.Print "field_54 = "; .field_54
            'Debug.Print "field_54 = "; .field_54
            'Debug.Print "field_58 = "; .field_58
            'Debug.Print "vertex_alpha = "; .vertex_alpha
            'Debug.Print "field_60 = "; .field_60
            'Debug.Print ""
        End With
    Next hi
End Sub
Sub WriteHundrets(ByVal NFile As Integer, ByVal offset As Long, ByRef hundrets() As PHundret)
    Put NFile, offset, hundrets
End Sub
Sub MergeHundrets(ByRef h1() As PHundret, ByRef h2() As PHundret)
    Dim NumHundretsH1 As Integer
    Dim NumHundretsH2 As Integer

    NumHundretsH1 = UBound(h1) + 1
    NumHundretsH2 = UBound(h2) + 1
    ReDim Preserve h1(NumHundretsH1 + NumHundretsH2 - 1)

    CopyMemory h1(NumHundretsH1), h2(0), NumHundretsH2 * 100
End Sub

Sub FillHundrestsDefaultValues(ByRef hundret As PHundret)
    With hundret
        .field_0 = 1
        .field_4 = 1
        .field_8 = 246274
        .field_C = 147458
        .TexID = 0
        .texture_set_pointer = 0
        .field_18 = 0
        .field_1C = 0
        .field_20 = 0
        .shademode = 2
        .lightstate_ambient = -1
        .field_2C = 0
        .lightstate_material_pointer = 0
        .srcblend = 2
        .destblend = 1
        .field_3C = 2
        .alpharef = 0
        .blend_mode = 4
        .zsort = 0
        .field_4C = 0
        .field_50 = 0
        .field_54 = 0
        .field_58 = 0
        .vertex_alpha = 255
        .field_60 = 0
    End With
End Sub
