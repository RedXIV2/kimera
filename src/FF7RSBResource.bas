Attribute VB_Name = "FF7RSBResource"
Option Explicit
Type RSBResource
    ID As String
    res_file As String
    Model As PModel
    NumTextures As Integer
    textures() As TEXTexture
End Type
Sub ReadRSBResource(ByRef Resource As RSBResource, ByVal fileName As String, ByRef textures_pool() As TEXTexture)
    Dim ti As Integer
    Dim ti_pool As Integer
    Dim tex_foundQ As Boolean
    Dim NumTexPool As Long
    Dim NFileAux As Integer
    Dim line As String

    On Error GoTo errorH
    NFileAux = FreeFile
    Open fileName + ".rsd" For Input As NFileAux

    Resource.res_file = fileName

    Line Input #NFileAux, line

    With Resource
        'Read PLY entry or RSD ID
        Do
            Line Input #NFileAux, line
        Loop While Len(line) = 0 Or Left$(line, 1) = "#"

        'Save the ID if present
        If Left(line, 1) = "@" Then
            .ID = line

            'Read PLY entry
            Do
                Line Input #NFileAux, line
            Loop While Len(line) = 0 Or Left$(line, 1) = "#"
        End If
        .ID = "@RSD940102"  'Needed by FF7 to load the textures?

        'While Left(line, 1) = "#" Or Left(line, 1) = "@"
        '    If Left(line, 1) = "@" Then .ID = line
        '    Line Input #NFileAux, line
        'Wend

        'Read P model
        ReadPModel .Model, Mid(line, 5, Len(line) - 3 - 5) + ".P"

        'Skip MAT entry
        Do
            Line Input #NFileAux, line
        Loop While Len(line) = 0 Or Left$(line, 1) = "#"

        'Skip GRP entry
        Do
            Line Input #NFileAux, line
        Loop While Len(line) = 0 Or Left$(line, 1) = "#"

        'Get NTEX entry
        Do
            Line Input #NFileAux, line
        Loop While Len(line) = 0 Or Left$(line, 1) = "#"

        'While Left(line, 4) <> "NTEX"
        '    Line Input #NFileAux, line
        'Wend

        .NumTextures = val(Mid(line, 6, Len(line)))

        ReDim .textures(.NumTextures)

        glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
        glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR

        For ti = 0 To .NumTextures - 1
            'Get next TEX entry
            Do
                Line Input #NFileAux, line
            Loop While Len(line) = 0 Or Left$(line, 1) = "#"
            'Line Input #NFileAux, line
            'While Left(line, 1) = "#"
            '    Line Input #NFileAux, line
            'Wend
            .textures(ti).tex_file = Mid(line, 8, Len(line) - 11) + ".TEX"

            ti_pool = 0
            tex_foundQ = False
            NumTexPool = SafeArrayGetDim(textures_pool)
            While ti_pool < NumTexPool And tex_foundQ = False
                tex_foundQ = (textures_pool(ti_pool).tex_file = .textures(ti).tex_file)
                ti_pool = ti_pool + 1
            Wend
            If tex_foundQ Then
                .textures(ti) = textures_pool(ti_pool - 1)
            ElseIf ReadTEXTexture(.textures(ti), .textures(ti).tex_file) = 0 Then
                LoadTEXTexture .textures(ti)
                LoadBitmapFromTEXTexture .textures(ti)
                'Hack to avoid loading the same texture twice.
                'Better leave this out if textures must be changed.
                'ReDim Preserve textures_pool(NumTexPool)
                'textures_pool(NumTexPool) = .textures(ti)
            End If
        Next ti
    End With

    Close NFileAux

    Exit Sub
errorH:
    MsgBox "Error opening " + fileName, vbCritical, "RSB Error " + Str$(Err)
End Sub
Public Sub WriteRSBResource(ByRef Resource As RSBResource, ByVal fileName As String)
    Dim ti As Integer
    Dim NFileAux As Integer
    Dim p_name As String
    Dim tex_name As String

    On Error GoTo errorH
    NFileAux = FreeFile
    Open fileName + ".rsd" For Output As NFileAux

    Print #NFileAux, Resource.ID
    With Resource.Model
        p_name = TrimPath(Left$(.fileName, Len(.fileName) - 2))
    End With
    Print #NFileAux, "PLY=" + p_name + ".PLY"
    Print #NFileAux, "MAT=" + p_name + ".MAT"
    Print #NFileAux, "GRP=" + p_name + ".GRP"

    With Resource
        Print #NFileAux, "NTEX=" + Right$(Str$(.NumTextures), Len(Str$(.NumTextures)) - 1)
    End With

    With Resource
        For ti = 0 To Resource.NumTextures - 1
            tex_name = TrimPath(.textures(ti).tex_file)
            Print #NFileAux, "TEX[" + Right$(Str$(ti), Len(Str$(ti)) - 1) + "]=" + _
                             Left$(tex_name, Len(tex_name) - 4) + ".TIM"
            WriteTEXTexture .textures(ti), tex_name
        Next ti
    End With

    Close NFileAux
    Exit Sub
errorH:
    MsgBox "Error saving " + fileName, vbCritical, "RSB Error " + Str$(Err)
End Sub
Sub CreateDListsFromRSBResource(ByRef obj As RSBResource)
    CreateDListsFromPModel obj.Model
End Sub
Sub FreeRSBResourceResources(ByRef obj As RSBResource)
    Dim ti As Integer

    With obj
        FreePModelResources .Model
    End With

    For ti = 0 To obj.NumTextures - 1
        With obj.textures(ti)
            glDeleteTextures 1, .tex_id
            DeleteDC .hdc
            DeleteObject .hbmp
        End With
    Next ti

End Sub
Sub DrawRSBResource(ByRef res As RSBResource, ByVal UseDLists As Boolean)
    Dim ti As Integer
    Dim tex_ids() As Long
    Dim rot_mat(16) As Double

    ReDim tex_ids(res.NumTextures)

    For ti = 0 To res.NumTextures - 1
        tex_ids(ti) = res.textures(ti).tex_id
        'tex_ids(ti) = 100000
    Next ti

    glMatrixMode GL_MODELVIEW
    glPushMatrix

    With res.Model
        glTranslatef .RepositionX, .RepositionY, .RepositionZ

        BuildMatrixFromQuaternion .RotationQuaternion, rot_mat

        glMultMatrixd rot_mat(0)

        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With

    If Not UseDLists Then
        DrawPModel res.Model, tex_ids, False
    Else
        DrawPModelDLists res.Model, tex_ids
    End If
    glPopMatrix
End Sub
Sub MergeRSBResources(ByRef res1 As RSBResource, ByRef res2 As RSBResource)
    Dim ti As Integer
    MergePModels res1.Model, res2.Model

    ReDim Preserve res1.textures(res1.NumTextures + res2.NumTextures)

    For ti = 0 To res2.NumTextures - 1
        res1.textures(res1.NumTextures + ti) = res2.textures(ti)
    Next ti

    res1.NumTextures = res1.NumTextures + res2.NumTextures
End Sub
