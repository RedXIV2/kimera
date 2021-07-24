VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGLext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function wglGetProcAddress Lib "opengl32.dll" (lpStr As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Const PAGE_EXECUTE_READWRITE = &H40&

Public Sub glBlendColor(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single): End Sub
Public Sub glBlendEquation(ByVal mode As Long): End Sub
Public Sub glDrawRangeElements(ByVal mode As Long, ByVal start As Long, ByVal lend As Long, ByVal count As Long, ByVal ltype As Long, ByVal indices As Long): End Sub
Public Sub glColorTable(ByVal target As Long, ByVal internalformat As Long, ByVal width As Long, ByVal format As Long, ByVal ltype As Long, ByVal table As Long): End Sub
Public Sub glColorTableParameterfv(ByVal target As Long, ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glColorTableParameteriv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glCopyColorTable(ByVal target As Long, ByVal internalformat As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long): End Sub
Public Sub glGetColorTable(ByVal target As Long, ByVal format As Long, ByVal ltype As Long, ByVal table As Long): End Sub
Public Sub glGetColorTableParameterfv(ByVal target As Long, ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glGetColorTableParameteriv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glColorSubTable(ByVal target As Long, ByVal start As Long, ByVal count As Long, ByVal format As Long, ByVal ltype As Long, ByVal data As Long): End Sub
Public Sub glCopyColorSubTable(ByVal target As Long, ByVal start As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long): End Sub
Public Sub glConvolutionFilter1D(ByVal target As Long, ByVal internalformat As Long, ByVal width As Long, ByVal format As Long, ByVal ltype As Long, ByVal image As Long): End Sub
Public Sub glConvolutionFilter2D(ByVal target As Long, ByVal internalformat As Long, ByVal width As Long, ByVal height As Long, ByVal format As Long, ByVal ltype As Long, ByVal image As Long): End Sub
Public Sub glConvolutionParameterf(ByVal target As Long, ByVal pname As Long, ByVal params As Single): End Sub
Public Sub glConvolutionParameterfv(ByVal target As Long, ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glConvolutionParameteri(ByVal target As Long, ByVal pname As Long, ByVal params As Long): End Sub
Public Sub glConvolutionParameteriv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glCopyConvolutionFilter1D(ByVal target As Long, ByVal internalformat As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long): End Sub
Public Sub glCopyConvolutionFilter2D(ByVal target As Long, ByVal internalformat As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long): End Sub
Public Sub glGetConvolutionFilter(ByVal target As Long, ByVal format As Long, ByVal ltype As Long, ByVal image As Long): End Sub
Public Sub glGetConvolutionParameterfv(ByVal target As Long, ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glGetConvolutionParameteriv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetSeparableFilter(ByVal target As Long, ByVal format As Long, ByVal ltype As Long, ByVal row As Long, ByVal column As Long, ByVal span As Long): End Sub
Public Sub glSeparableFilter2D(ByVal target As Long, ByVal internalformat As Long, ByVal width As Long, ByVal height As Long, ByVal format As Long, ByVal ltype As Long, ByVal row As Long, ByVal column As Long): End Sub
Public Sub glGetHistogram(ByVal target As Long, ByVal reset As Long, ByVal format As Long, ByVal ltype As Long, ByVal values As Long): End Sub
Public Sub glGetHistogramParameterfv(ByVal target As Long, ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glGetHistogramParameteriv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetMinmax(ByVal target As Long, ByVal reset As Long, ByVal format As Long, ByVal ltype As Long, ByVal values As Long): End Sub
Public Sub glGetMinmaxParameterfv(ByVal target As Long, ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glGetMinmaxParameteriv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glHistogram(ByVal target As Long, ByVal width As Long, ByVal internalformat As Long, ByVal sink As Long): End Sub
Public Sub glMinmax(ByVal target As Long, ByVal internalformat As Long, ByVal sink As Long): End Sub
Public Sub glResetHistogram(ByVal target As Long): End Sub
Public Sub glResetMinmax(ByVal target As Long): End Sub
Public Sub glTexImage3D(ByVal target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal width As Long, ByVal height As Long, ByVal depth As Long, ByVal border As Long, ByVal format As Long, ByVal ltype As Long, ByVal pixels As Long): End Sub
Public Sub glTexSubImage3D(ByVal target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal yoffset As Long, ByVal zoffset As Long, ByVal width As Long, ByVal height As Long, ByVal depth As Long, ByVal format As Long, ByVal ltype As Long, ByVal pixels As Long): End Sub
Public Sub glCopyTexSubImage3D(ByVal target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal yoffset As Long, ByVal zoffset As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long): End Sub

Public Sub glActiveTexture(ByVal texture As Long): End Sub
Public Sub glClientActiveTexture(ByVal texture As Long): End Sub
Public Sub glMultiTexCoord1d(ByVal target As Long, ByVal s As Double): End Sub
Public Sub glMultiTexCoord1dv(ByVal target As Long, ByRef v As Double): End Sub
Public Sub glMultiTexCoord1f(ByVal target As Long, ByVal s As Single): End Sub
Public Sub glMultiTexCoord1fv(ByVal target As Long, ByRef v As Single): End Sub
Public Sub glMultiTexCoord1i(ByVal target As Long, ByVal s As Long): End Sub
Public Sub glMultiTexCoord1iv(ByVal target As Long, ByRef v As Long): End Sub
Public Sub glMultiTexCoord1s(ByVal target As Long, ByVal s As Integer): End Sub
Public Sub glMultiTexCoord1sv(ByVal target As Long, ByRef v As Integer): End Sub
Public Sub glMultiTexCoord2d(ByVal target As Long, ByVal s As Double, ByVal t As Double): End Sub
Public Sub glMultiTexCoord2dv(ByVal target As Long, ByRef v As Double): End Sub
Public Sub glMultiTexCoord2f(ByVal target As Long, ByVal s As Single, ByVal t As Single): End Sub
Public Sub glMultiTexCoord2fv(ByVal target As Long, ByRef v As Single): End Sub
Public Sub glMultiTexCoord2i(ByVal target As Long, ByVal s As Long, ByVal t As Long): End Sub
Public Sub glMultiTexCoord2iv(ByVal target As Long, ByRef v As Long): End Sub
Public Sub glMultiTexCoord2s(ByVal target As Long, ByVal s As Integer, ByVal t As Integer): End Sub
Public Sub glMultiTexCoord2sv(ByVal target As Long, ByRef v As Integer): End Sub
Public Sub glMultiTexCoord3d(ByVal target As Long, ByVal s As Double, ByVal t As Double, ByVal r As Double): End Sub
Public Sub glMultiTexCoord3dv(ByVal target As Long, ByRef v As Double): End Sub
Public Sub glMultiTexCoord3f(ByVal target As Long, ByVal s As Single, ByVal t As Single, ByVal r As Single): End Sub
Public Sub glMultiTexCoord3fv(ByVal target As Long, ByRef v As Single): End Sub
Public Sub glMultiTexCoord3i(ByVal target As Long, ByVal s As Long, ByVal t As Long, ByVal r As Long): End Sub
Public Sub glMultiTexCoord3iv(ByVal target As Long, ByRef v As Long): End Sub
Public Sub glMultiTexCoord3s(ByVal target As Long, ByVal s As Integer, ByVal t As Integer, ByVal r As Integer): End Sub
Public Sub glMultiTexCoord3sv(ByVal target As Long, ByRef v As Integer): End Sub
Public Sub glMultiTexCoord4d(ByVal target As Long, ByVal s As Double, ByVal t As Double, ByVal r As Double, ByVal q As Double): End Sub
Public Sub glMultiTexCoord4dv(ByVal target As Long, ByRef v As Double): End Sub
Public Sub glMultiTexCoord4f(ByVal target As Long, ByVal s As Single, ByVal t As Single, ByVal r As Single, ByVal q As Single): End Sub
Public Sub glMultiTexCoord4fv(ByVal target As Long, ByRef v As Single): End Sub
Public Sub glMultiTexCoord4i(ByVal target As Long, ByVal s As Long, ByVal t As Long, ByVal r As Long, ByVal q As Long): End Sub
Public Sub glMultiTexCoord4iv(ByVal target As Long, ByRef v As Long): End Sub
Public Sub glMultiTexCoord4s(ByVal target As Long, ByVal s As Integer, ByVal t As Integer, ByVal r As Integer, ByVal q As Integer): End Sub
Public Sub glMultiTexCoord4sv(ByVal target As Long, ByRef v As Integer): End Sub
Public Sub glLoadTransposeMatrixf(ByRef m As Single): End Sub
Public Sub glLoadTransposeMatrixd(ByRef m As Double): End Sub
Public Sub glMultTransposeMatrixf(ByRef m As Single): End Sub
Public Sub glMultTransposeMatrixd(ByRef m As Double): End Sub
Public Sub glSampleCoverage(ByVal value As Single, ByVal invert As Long): End Sub
Public Sub glCompressedTexImage3D(ByVal target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal width As Long, ByVal height As Long, ByVal depth As Long, ByVal border As Long, ByVal imageSize As Long, ByVal data As Long): End Sub
Public Sub glCompressedTexImage2D(ByVal target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal width As Long, ByVal height As Long, ByVal border As Long, ByVal imageSize As Long, ByVal data As Long): End Sub
Public Sub glCompressedTexImage1D(ByVal target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal width As Long, ByVal border As Long, ByVal imageSize As Long, ByVal data As Long): End Sub
Public Sub glCompressedTexSubImage3D(ByVal target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal yoffset As Long, ByVal zoffset As Long, ByVal width As Long, ByVal height As Long, ByVal depth As Long, ByVal format As Long, ByVal imageSize As Long, ByVal data As Long): End Sub
Public Sub glCompressedTexSubImage2D(ByVal target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal yoffset As Long, ByVal width As Long, ByVal height As Long, ByVal format As Long, ByVal imageSize As Long, ByVal data As Long): End Sub
Public Sub glCompressedTexSubImage1D(ByVal target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal width As Long, ByVal format As Long, ByVal imageSize As Long, ByVal data As Long): End Sub
Public Sub glGetCompressedTexImage(ByVal target As Long, ByVal level As Long, ByVal img As Long): End Sub

Public Sub glBlendFuncSeparate(ByVal sfactorRGB As Long, ByVal dfactorRGB As Long, ByVal sfactorAlpha As Long, ByVal dfactorAlpha As Long): End Sub
Public Sub glFogCoordf(ByVal coord As Single): End Sub
Public Sub glFogCoordfv(ByRef coord As Single): End Sub
Public Sub glFogCoordd(ByVal coord As Double): End Sub
Public Sub glFogCoorddv(ByRef coord As Double): End Sub
Public Sub glFogCoordPointer(ByVal ltype As Long, ByVal stride As Long, ByVal pointer As Long): End Sub
Public Sub glMultiDrawArrays(ByVal mode As Long, ByRef first As Long, ByRef count As Long, ByVal primcount As Long): End Sub
Public Sub glMultiDrawElements(ByVal mode As Long, ByRef count As Long, ByVal ltype As Long, ByVal indices As Long, ByVal primcount As Long): End Sub
Public Sub glPointParameterf(ByVal pname As Long, ByVal param As Single): End Sub
Public Sub glPointParameterfv(ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glPointParameteri(ByVal pname As Long, ByVal param As Long): End Sub
Public Sub glPointParameteriv(ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glSecondaryColor3b(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte): End Sub
Public Sub glSecondaryColor3bv(ByRef v As Byte): End Sub
Public Sub glSecondaryColor3d(ByVal red As Double, ByVal green As Double, ByVal blue As Double): End Sub
Public Sub glSecondaryColor3dv(ByRef v As Double): End Sub
Public Sub glSecondaryColor3f(ByVal red As Single, ByVal green As Single, ByVal blue As Single): End Sub
Public Sub glSecondaryColor3fv(ByRef v As Single): End Sub
Public Sub glSecondaryColor3i(ByVal red As Long, ByVal green As Long, ByVal blue As Long): End Sub
Public Sub glSecondaryColor3iv(ByRef v As Long): End Sub
Public Sub glSecondaryColor3s(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer): End Sub
Public Sub glSecondaryColor3sv(ByRef v As Integer): End Sub
Public Sub glSecondaryColor3ub(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte): End Sub
Public Sub glSecondaryColor3ubv(ByRef v As Byte): End Sub
Public Sub glSecondaryColor3ui(ByVal red As Long, ByVal green As Long, ByVal blue As Long): End Sub
Public Sub glSecondaryColor3uiv(ByRef v As Long): End Sub
Public Sub glSecondaryColor3us(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer): End Sub
Public Sub glSecondaryColor3usv(ByRef v As Integer): End Sub
Public Sub glSecondaryColorPointer(ByVal size As Long, ByVal ltype As Long, ByVal stride As Long, ByVal pointer As Long): End Sub
Public Sub glWindowPos2d(ByVal x As Double, ByVal y As Double): End Sub
Public Sub glWindowPos2dv(ByRef v As Double): End Sub
Public Sub glWindowPos2f(ByVal x As Single, ByVal y As Single): End Sub
Public Sub glWindowPos2fv(ByRef v As Single): End Sub
Public Sub glWindowPos2i(ByVal x As Long, ByVal y As Long): End Sub
Public Sub glWindowPos2iv(ByRef v As Long): End Sub
Public Sub glWindowPos2s(ByVal x As Integer, ByVal y As Integer): End Sub
Public Sub glWindowPos2sv(ByRef v As Integer): End Sub
Public Sub glWindowPos3d(ByVal x As Double, ByVal y As Double, ByVal z As Double): End Sub
Public Sub glWindowPos3dv(ByRef v As Double): End Sub
Public Sub glWindowPos3f(ByVal x As Single, ByVal y As Single, ByVal z As Single): End Sub
Public Sub glWindowPos3fv(ByRef v As Single): End Sub
Public Sub glWindowPos3i(ByVal x As Long, ByVal y As Long, ByVal z As Long): End Sub
Public Sub glWindowPos3iv(ByRef v As Long): End Sub
Public Sub glWindowPos3s(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer): End Sub
Public Sub glWindowPos3sv(ByRef v As Integer): End Sub

Public Sub glGenQueries(ByVal n As Long, ByRef ids As Long): End Sub
Public Sub glDeleteQueries(ByVal n As Long, ByRef ids As Long): End Sub
Public Function glIsQuery(ByVal id As Long) As Long: End Function
Public Sub glBeginQuery(ByVal target As Long, ByVal id As Long): End Sub
Public Sub glEndQuery(ByVal target As Long): End Sub
Public Sub glGetQueryiv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetQueryObjectiv(ByVal id As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetQueryObjectuiv(ByVal id As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glBindBuffer(ByVal target As Long, ByVal buffer As Long): End Sub
Public Sub glDeleteBuffers(ByVal n As Long, ByRef buffers As Long): End Sub
Public Sub glGenBuffers(ByVal n As Long, ByRef buffers As Long): End Sub
Public Function glIsBuffer(ByVal buffer As Long) As Long: End Function
Public Sub glBufferData(ByVal target As Long, ByVal size As Long, ByVal data As Long, ByVal usage As Long): End Sub
Public Sub glBufferSubData(ByVal target As Long, ByVal offset As Long, ByVal size As Long, ByVal data As Long): End Sub
Public Sub glGetBufferSubData(ByVal target As Long, ByVal offset As Long, ByVal size As Long, ByVal data As Long): End Sub
Public Function glMapBuffer(ByVal target As Long, ByVal access As Long) As Long: End Function
Public Function glUnmapBuffer(ByVal target As Long) As Long: End Function
Public Sub glGetBufferParameteriv(ByVal target As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetBufferPointerv(ByVal target As Long, ByVal pname As Long, ByVal params As Long): End Sub

Public Sub glBlendEquationSeparate(ByVal modeRGB As Long, ByVal modeAlpha As Long): End Sub
Public Sub glDrawBuffers(ByVal n As Long, ByRef bufs As Long): End Sub
Public Sub glStencilOpSeparate(ByVal face As Long, ByVal sfail As Long, ByVal dpfail As Long, ByVal dppass As Long): End Sub
Public Sub glStencilFuncSeparate(ByVal frontfunc As Long, ByVal backfunc As Long, ByVal ref As Long, ByVal mask As Long): End Sub
Public Sub glStencilMaskSeparate(ByVal face As Long, ByVal mask As Long): End Sub
Public Sub glAttachShader(ByVal program As Long, ByVal shader As Long): End Sub
Public Sub glBindAttribLocation(ByVal program As Long, ByVal index As Long, ByRef name As Byte): End Sub
Public Sub glCompileShader(ByVal shader As Long): End Sub
Public Function glCreateProgram(void) As Long: End Function
Public Function glCreateShader(ByVal ltype As Long) As Long: End Function
Public Sub glDeleteProgram(ByVal program As Long): End Sub
Public Sub glDeleteShader(ByVal shader As Long): End Sub
Public Sub glDetachShader(ByVal program As Long, ByVal shader As Long): End Sub
Public Sub glDisableVertexAttribArray(ByVal index As Long): End Sub
Public Sub glEnableVertexAttribArray(ByVal index As Long): End Sub
Public Sub glGetActiveAttrib(ByVal program As Long, ByVal index As Long, ByVal bufSize As Long, ByRef length As Long, ByRef size As Long, ByRef ltype As Long, ByRef pname As Byte): End Sub
Public Sub glGetActiveUniform(ByVal program As Long, ByVal index As Long, ByVal bufSize As Long, ByRef length As Long, ByRef size As Long, ByRef ltype As Long, ByRef pname As Byte): End Sub
Public Sub glGetAttachedShaders(ByVal program As Long, ByVal maxCount As Long, ByRef count As Long, ByRef obj As Long): End Sub
Public Function glGetAttribLocation(ByVal program As Long, ByRef name As Byte) As Long: End Function
Public Sub glGetProgramiv(ByVal program As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetProgramInfoLog(ByVal program As Long, ByVal bufSize As Long, ByRef length As Long, ByRef infoLog As Byte): End Sub
Public Sub glGetShaderiv(ByVal shader As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetShaderInfoLog(ByVal shader As Long, ByVal bufSize As Long, ByRef length As Long, ByRef infoLog As Byte): End Sub
Public Sub glGetShaderSource(ByVal shader As Long, ByVal bufSize As Long, ByRef length As Long, ByRef source As Byte): End Sub
Public Function glGetUniformLocation(ByVal program As Long, ByRef name As Byte) As Long: End Function
Public Sub glGetUniformfv(ByVal program As Long, ByVal location As Long, ByRef params As Single): End Sub
Public Sub glGetUniformiv(ByVal program As Long, ByVal location As Long, ByRef params As Long): End Sub
Public Sub glGetVertexAttribdv(ByVal index As Long, ByVal pname As Long, ByRef params As Double): End Sub
Public Sub glGetVertexAttribfv(ByVal index As Long, ByVal pname As Long, ByRef params As Single): End Sub
Public Sub glGetVertexAttribiv(ByVal index As Long, ByVal pname As Long, ByRef params As Long): End Sub
Public Sub glGetVertexAttribPointerv(ByVal index As Long, ByVal pname As Long, ByVal pointer As Long): End Sub
Public Function glIsProgram(ByVal program As Long) As Long: End Function
Public Function glIsShader(ByVal shader As Long) As Long: End Function
Public Sub glLinkProgram(ByVal program As Long): End Sub
Public Sub glShaderSource(ByVal shader As Long, ByVal count As Long, ByVal pstring As Long, ByRef length As Long): End Sub
Public Sub glUseProgram(ByVal program As Long): End Sub
Public Sub glUniform1f(ByVal location As Long, ByVal v0 As Single): End Sub
Public Sub glUniform2f(ByVal location As Long, ByVal v0 As Single, ByVal v1 As Single): End Sub
Public Sub glUniform3f(ByVal location As Long, ByVal v0 As Single, ByVal v1 As Single, ByVal v2 As Single): End Sub
Public Sub glUniform4f(ByVal location As Long, ByVal v0 As Single, ByVal v1 As Single, ByVal v2 As Single, ByVal v3 As Single): End Sub
Public Sub glUniform1i(ByVal location As Long, ByVal v0 As Long): End Sub
Public Sub glUniform2i(ByVal location As Long, ByVal v0 As Long, ByVal v1 As Long): End Sub
Public Sub glUniform3i(ByVal location As Long, ByVal v0 As Long, ByVal v1 As Long, ByVal v2 As Long): End Sub
Public Sub glUniform4i(ByVal location As Long, ByVal v0 As Long, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long): End Sub
Public Sub glUniform1fv(ByVal location As Long, ByVal count As Long, ByRef value As Single): End Sub
Public Sub glUniform2fv(ByVal location As Long, ByVal count As Long, ByRef value As Single): End Sub
Public Sub glUniform3fv(ByVal location As Long, ByVal count As Long, ByRef value As Single): End Sub
Public Sub glUniform4fv(ByVal location As Long, ByVal count As Long, ByRef value As Single): End Sub
Public Sub glUniform1iv(ByVal location As Long, ByVal count As Long, ByRef value As Long): End Sub
Public Sub glUniform2iv(ByVal location As Long, ByVal count As Long, ByRef value As Long): End Sub
Public Sub glUniform3iv(ByVal location As Long, ByVal count As Long, ByRef value As Long): End Sub
Public Sub glUniform4iv(ByVal location As Long, ByVal count As Long, ByRef value As Long): End Sub
Public Sub glUniformMatrix2fv(ByVal location As Long, ByVal count As Long, ByVal transpose As Long, ByRef value As Single): End Sub
Public Sub glUniformMatrix3fv(ByVal location As Long, ByVal count As Long, ByVal transpose As Long, ByRef value As Single): End Sub
Public Sub glUniformMatrix4fv(ByVal location As Long, ByVal count As Long, ByVal transpose As Long, ByRef value As Single): End Sub
Public Sub glValidateProgram(ByVal program As Long): End Sub
Public Sub glVertexAttrib1d(ByVal index As Long, ByVal x As Double): End Sub
Public Sub glVertexAttrib1dv(ByVal index As Long, ByRef v As Double): End Sub
Public Sub glVertexAttrib1f(ByVal index As Long, ByVal x As Single): End Sub
Public Sub glVertexAttrib1fv(ByVal index As Long, ByRef v As Single): End Sub
Public Sub glVertexAttrib1s(ByVal index As Long, ByVal x As Integer): End Sub
Public Sub glVertexAttrib1sv(ByVal index As Long, ByRef v As Integer): End Sub
Public Sub glVertexAttrib2d(ByVal index As Long, ByVal x As Double, ByVal y As Double): End Sub
Public Sub glVertexAttrib2dv(ByVal index As Long, ByRef v As Double): End Sub
Public Sub glVertexAttrib2f(ByVal index As Long, ByVal x As Single, ByVal y As Single): End Sub
Public Sub glVertexAttrib2fv(ByVal index As Long, ByRef v As Single): End Sub
Public Sub glVertexAttrib2s(ByVal index As Long, ByVal x As Integer, ByVal y As Integer): End Sub
Public Sub glVertexAttrib2sv(ByVal index As Long, ByRef v As Integer): End Sub
Public Sub glVertexAttrib3d(ByVal index As Long, ByVal x As Double, ByVal y As Double, ByVal z As Double): End Sub
Public Sub glVertexAttrib3dv(ByVal index As Long, ByRef v As Double): End Sub
Public Sub glVertexAttrib3f(ByVal index As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single): End Sub
Public Sub glVertexAttrib3fv(ByVal index As Long, ByRef v As Single): End Sub
Public Sub glVertexAttrib3s(ByVal index As Long, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer): End Sub
Public Sub glVertexAttrib3sv(ByVal index As Long, ByRef v As Integer): End Sub
Public Sub glVertexAttrib4Nbv(ByVal index As Long, ByRef v As Byte): End Sub
Public Sub glVertexAttrib4Niv(ByVal index As Long, ByRef v As Long): End Sub
Public Sub glVertexAttrib4Nsv(ByVal index As Long, ByRef v As Integer): End Sub
Public Sub glVertexAttrib4Nub(ByVal index As Long, ByVal x As Byte, ByVal y As Byte, ByVal z As Byte, ByVal w As Byte): End Sub
Public Sub glVertexAttrib4Nubv(ByVal index As Long, ByRef v As Byte): End Sub
Public Sub glVertexAttrib4Nuiv(ByVal index As Long, ByRef v As Long): End Sub
Public Sub glVertexAttrib4Nusv(ByVal index As Long, ByRef v As Integer): End Sub
Public Sub glVertexAttrib4bv(ByVal index As Long, ByRef v As Byte): End Sub
Public Sub glVertexAttrib4d(ByVal index As Long, ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal w As Double): End Sub
Public Sub glVertexAttrib4dv(ByVal index As Long, ByRef v As Double): End Sub
Public Sub glVertexAttrib4f(ByVal index As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single): End Sub
Public Sub glVertexAttrib4fv(ByVal index As Long, ByRef v As Single): End Sub
Public Sub glVertexAttrib4iv(ByVal index As Long, ByRef v As Long): End Sub
Public Sub glVertexAttrib4s(ByVal index As Long, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal w As Integer): End Sub
Public Sub glVertexAttrib4sv(ByVal index As Long, ByRef v As Integer): End Sub
Public Sub glVertexAttrib4ubv(ByVal index As Long, ByRef v As Byte): End Sub
Public Sub glVertexAttrib4uiv(ByVal index As Long, ByRef v As Long): End Sub
Public Sub glVertexAttrib4usv(ByVal index As Long, ByRef v As Integer): End Sub
Public Sub glVertexAttribPointer(ByVal index As Long, ByVal size As Long, ByVal ltype As Long, ByVal normalized As Long, ByVal stride As Long, ByVal pointer As Long): End Sub

Private Function ReplaceVtableEntry(ByVal entry As Integer, ByVal lpfn As Long) As Long
    Dim lOldAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long
    CopyMemory lpVtableHead, ByVal ObjPtr(Me), 4
    lpfnAddr = lpVtableHead + (entry - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4 'Get old entry
    If lOldAddr <> lpfn Then
        VirtualProtect lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect
        CopyMemory ByVal lpfnAddr, lpfn, 4
        'Don�t bother to reprotect
        VirtualProtect lpfnAddr, 4, lOldProtect, lOldProtect
    End If
    ReplaceVtableEntry = lOldAddr
End Function
Private Function ReplaceEntryWithDLLFunction(ByVal entry As Integer, ByVal lpfn As Long) As Long
    Dim lFuncAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long

    CopyMemory lpVtableHead, ByVal ObjPtr(Me), 4
    lpfnAddr = lpVtableHead + (entry - 1) * 4
    CopyMemory lFuncAddr, ByVal lpfnAddr, 4 'Get old entry
    If lFuncAddr <> lpfn Then
        VirtualProtect lFuncAddr, 10, PAGE_EXECUTE_READWRITE, lOldProtect

        CopyMemory ByVal lFuncAddr, &HB824048F, 4 'pop [esp] 'mov eax,
        CopyMemory ByVal lFuncAddr + 4, lpfn, 4
        CopyMemory ByVal lFuncAddr + 8, &HE0FF, 2

        'Don�t bother to reprotect
        VirtualProtect lFuncAddr, 10, lOldProtect, lOldProtect
    End If
    ReplaceEntryWithDLLFunction = lFuncAddr
End Function
Private Sub Class_Initialize()
    Dim lProcAddress As Long
    Dim entry As Long
    Dim sFunctionName As String
    Dim bData() As Byte
    Dim i As Long
    '8 is the first entry
    entry = 8

    bData = LoadResData("GLEXT", "GLEXT")
    For i = 0 To UBound(bData) - 1
        If bData(i) = 0 Then
            lProcAddress = wglGetProcAddress(bData(i + 1))
            If lProcAddress Then ReplaceEntryWithDLLFunction entry, lProcAddress
            entry = entry + 1
        ElseIf i = 0 Then
            lProcAddress = wglGetProcAddress(bData(i))
            If lProcAddress Then ReplaceEntryWithDLLFunction entry, lProcAddress
            entry = entry + 1
        End If
    Next i
End Sub
