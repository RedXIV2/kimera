Attribute VB_Name = "OpenGL"
Option Explicit
Public GL_Ext As New ClsGLext 'GL Extensions Class

Public Function CreateOGLContext(ghDC As Long) As Long
    Dim hRC As Long
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim PixFormat As Long

    ZeroMemory pfd, Len(pfd)
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24
    pfd.cDepthBits = 32
    'pfd.cAlphaBits = 24
    pfd.iLayerType = PFD_MAIN_PLANE

    PixFormat = ChoosePixelFormat(ghDC, pfd)
    If PixFormat = 0 Then GoTo ee
    SetPixelFormat ghDC, PixFormat, pfd
    hRC = wglCreateContext(ghDC)
    wglMakeCurrent ghDC, hRC
    CreateOGLContext = hRC
Exit Function
ee: MsgBox "Can't create OpenGL context!", vbCritical, "Error"
    End
End Function
Public Sub SetOGLContext(ghDC As Long, hRC As Long)
    wglMakeCurrent ghDC, hRC
End Sub
Public Sub DisableOpenGL(hRC As Long)
    wglMakeCurrent 0, 0
    wglDeleteContext hRC
End Sub
'-----------------------------------------------------------------------------
'--------------------------COORDINATES CONVERSIONS----------------------------
'-----------------------------------------------------------------------------
Function GetDepthZ(ByRef p As Point3D) As Double
    Dim p_temp As Point3D
    p_temp = GetProjectedCoords(p)
    GetDepthZ = p_temp.z
End Function
Function GetProjectedCoords(ByRef p As Point3D) As Point3D
    Dim mm(16) As Double
    Dim pm(16) As Double
    Dim vp(4) As Long

    Dim x_in As Double
    Dim y_in As Double
    Dim z_in As Double

    Dim x_temp As Double
    Dim y_temp As Double
    Dim z_temp As Double

    With p
        x_in = .x
        y_in = .y
        z_in = .z
    End With

    glGetDoublev GL_MODELVIEW_MATRIX, mm(0)
    glGetDoublev GL_PROJECTION_MATRIX, pm(0)
    glGetIntegerv GL_VIEWPORT, vp(0)

    gluProject x_in, y_in, z_in, mm(0), pm(0), vp(0), x_temp, y_temp, z_temp

    With GetProjectedCoords
        .x = x_temp
        .y = y_temp
        .z = z_temp
    End With
End Function
Function GetUnProjectedCoords(ByRef p As Point3D) As Point3D
    Dim mm(16) As Double
    Dim pm(16) As Double
    Dim vp(4) As Long

    Dim x_temp As Double
    Dim y_temp As Double
    Dim z_temp As Double

    Dim x_in As Double
    Dim y_in As Double
    Dim z_in As Double

    With p
        x_in = .x
        y_in = .y
        z_in = .z
    End With

    glGetDoublev GL_MODELVIEW_MATRIX, mm(0)
    glGetDoublev GL_PROJECTION_MATRIX, pm(0)
    glGetIntegerv GL_VIEWPORT, vp(0)

    gluUnProject x_in, vp(3) - y_in, z_in, mm(0), pm(0), vp(0), x_temp, y_temp, z_temp


    With GetUnProjectedCoords
        .x = x_temp
        .y = y_temp
        .z = z_temp
    End With
End Function
Function GetEyeSpaceCoords(ByRef p As Point3D) As Point3D
    Dim mm(16) As Double

    glGetDoublev GL_MODELVIEW_MATRIX, mm(0)

    With GetEyeSpaceCoords
        .x = p.x * mm(0) + p.y * mm(4) + p.z * mm(8) + mm(12)
        .y = p.x * mm(1) + p.y * mm(5) + p.z * mm(9) + mm(13)
        .z = p.x * mm(2) + p.y * mm(6) + p.z * mm(10) + mm(14)
    End With
End Function
Function GetObjectSpaceCoords(ByRef p As Point3D) As Point3D
    Dim mm(16) As Double

    glGetDoublev GL_MODELVIEW_MATRIX, mm(0)

    InvertMatrix mm

    With GetObjectSpaceCoords
        .x = p.x * mm(0) + p.y * mm(4) + p.z * mm(8) + mm(12)
        .y = p.x * mm(1) + p.y * mm(5) + p.z * mm(9) + mm(13)
        .z = p.x * mm(2) + p.y * mm(6) + p.z * mm(10) + mm(14)
    End With
End Function

Function GetVertColor(ByRef p As Point3D, ByRef n As Point3D, ByRef C As color) As color
    Dim PI As Integer
    Dim vi As Integer
    Dim pcolor(3) As Byte
    Dim vp0(4) As Long
    Dim vp(4) As Long
    Dim P_matrix(16) As Double

    glGetIntegerv GL_VIEWPORT, vp0(0)
    glViewport 0, 0, 3, 3
    glGetIntegerv GL_VIEWPORT, vp(0)
    glMatrixMode GL_PROJECTION

    glPushMatrix

    With GetProjectedCoords(p)
      ''Debug.Print p.X, p.Y, p.z, .X, .Y, c.r, c.g, c.b
        glGetDoublev GL_PROJECTION_MATRIX, P_matrix(0)
        glLoadIdentity
        gluPickMatrix .x - 1, .y - 1, 3, 3, vp(0)
        'gluPerspective 60, 1, 0.1, diameter
        glMultMatrixd P_matrix(0)
    End With


    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

    glPointSize 100

    glBegin GL_POINTS
        With C
            glColor4f .r / 255, .g / 255, .B / 255, 1 '.a / 255
            glColorMaterial GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE
        End With

        With n
            glNormal3f .x, .y, .z
        End With

        With p
            glVertex3f .x, .y, .z
        End With
    glEnd
    glFlush

    glReadBuffer GL_BACK

    With GetProjectedCoords(p)
        glReadPixels 1, 1, 1, 1, GL_RGB, GL_BYTE, pcolor(0)
    End With

    With GetVertColor
        .r = pcolor(0) * 2
        .g = pcolor(1) * 2
        .B = pcolor(2) * 2
        .a = 255
    End With
    glPopMatrix
    glViewport vp0(0), vp0(1), vp0(2), vp0(3)
End Function
'-----------------------------------------------------------------------------
'--------------------------TEXTURE OPERATIONS---------------------------------
'-----------------------------------------------------------------------------
Sub ConvertBMP2Texture(ByVal hdc As Long, ByVal hbm As Long, ByVal w As Integer, ByVal h As Integer) ', ByVal EnableTransparentColor As Boolean)
    Dim x As Long, y As Long
    Dim tex_id As Long
    Dim TextureImg() As Byte
    Dim bi24BitInfo As BITMAPINFO
    Dim temp As Byte

    ReDim TextureImg(3, w - 1, h - 1)

    With bi24BitInfo.bmiHeader
        .biBitCount = 32
        .biCompression = 0 ' BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = w
        .biHeight = h
    End With

    GetDIBits hdc, hbm, 0, h, TextureImg(0, 0, 0), bi24BitInfo, 0

    For x = 0 To w - 1
        For y = 0 To h / 2
            temp = TextureImg(0, x, y)
            TextureImg(0, x, y) = TextureImg(2, x, h - 1 - y)
            TextureImg(2, x, h - 1 - y) = temp

            temp = TextureImg(1, x, y)
            TextureImg(1, x, y) = TextureImg(1, x, h - 1 - y)
            TextureImg(1, x, h - 1 - y) = temp

            temp = TextureImg(2, x, y)
            TextureImg(2, x, y) = TextureImg(0, x, h - 1 - y)
            TextureImg(0, x, h - 1 - y) = temp

            'If TextureImg(0, x, y) = 0 And TextureImg(1, x, y) = 0 And TextureImg(2, x, y) = 0 Then
            '    TextureImg(3, x, y) = 0
            'Else
            '    TextureImg(3, x, y) = 255
            'End If

            'If TextureImg(0, x, h - 1 - y) = 0 And TextureImg(1, x, h - 1 - y) = 0 And TextureImg(2, x, h - 1 - y) = 0 Then
            '    TextureImg(3, x, h - 1 - y) = 0
            'Else
            '    TextureImg(3, x, h - 1 - y) = 255
            'End If
        Next y
    Next x

    glPixelStorei GL_UNPACK_ALIGNMENT, 1
    glTexImage2D GL_TEXTURE_2D, 0, 4, w, h, 0, GL_RGBA, GL_UNSIGNED_BYTE, TextureImg(0, 0, 0)
    ''Debug.Print "Creating Texture...", glGetError = GL_NO_ERROR

    Erase TextureImg
End Sub
Public Sub LoadTexturev(ByRef TextureImg() As Byte, ByVal width As Integer, ByVal height As Integer)
    glPixelStorei GL_UNPACK_ALIGNMENT, 1
    glTexImage2D GL_TEXTURE_2D, 0, 4, width, height, 0, GL_RGBA, GL_UNSIGNED_BYTE, TextureImg(0, 0, 0)
    'Debug.Print "Creating Texture...", glGetError = GL_NO_ERROR

    Erase TextureImg
End Sub
Public Sub SetDefaultOGLRenderState()
    glPolygonMode GL_FRONT, GL_FILL
    glPolygonMode GL_BACK, GL_FILL

    glDisable GL_CULL_FACE
    'glCullFace GL_BACK

    glEnable GL_BLEND
    GL_Ext.glBlendEquation GL_FUNC_ADD
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA

    glDisable GL_TEXTURE_2D
End Sub
