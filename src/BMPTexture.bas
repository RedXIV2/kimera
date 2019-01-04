Attribute VB_Name = "BMPTextureModule"
Option Explicit
Type BMPTexture
    fileName As String
    MagicId As Integer       'Magic identifier
    FileSize As Long        'File size in bytes
    reserved1 As Integer
    reserved2 As Integer
    ImgOffset As Long       'Offset to image data, bytes

    HeaderSize As Long      'Header size in bytes
    width As Long
    height As Long          'Width and height of image
    Planes As Integer       'Number of colour planes (must be 1)
    Bits As Integer         'Bits per pixel
    Compression As Long     'Compression type
    ImageSize As Long       'Image size in bytes
    Xresolution As Long
    Yresolution As Long     'Pixels per meter
    Ncolours As Long        'Number of colours
    Importantcolours As Long   'Important colours

    Pallete() As Byte
    BitMap() As Byte
End Type
Function ReadBMPTexture(ByRef Texture As BMPTexture, ByVal fileName As String) As Integer
    Dim NFile As Integer
    Dim ci As Integer
    
    On Error GoTo errorH
    
    If FileExist(fileName) Then
        NFile = FreeFile
        Open fileName For Binary As NFile
        
        With Texture
            .fileName = fileName
            Get NFile, 1, .MagicId
            Get NFile, &H2 + 1, .FileSize
            Get NFile, &H6 + 1, .reserved1
            Get NFile, &H8 + 1, .reserved2
            Get NFile, &HA + 1, .ImgOffset
            Get NFile, &HE + 1, .HeaderSize
            Get NFile, &H12 + 1, .width
            Get NFile, &H16 + 1, .height
            Get NFile, &H1A + 1, .Planes
            Get NFile, &H1C + 1, .Bits
            Get NFile, &H1E + 1, .Compression
            Get NFile, &H22 + 1, .ImageSize
            Get NFile, &H26 + 1, .Xresolution
            Get NFile, &H2A + 1, .Yresolution
            Get NFile, &H2E + 1, .Ncolours
            Get NFile, &H32 + 1, .Importantcolours
        
            If .Bits <= 8 Then
                ReDim .Pallete(2 ^ .Bits * 4 - 1)
                Get NFile, &H36 + 1, .Pallete
            End If
            ReDim .BitMap((.Bits * .width * .height) / 8 - 1)
            Get NFile, .ImgOffset + 1, .BitMap
        End With
        Close NFile
        ReadBMPTexture = 0
    Else
        'Debug.Print "BMP file not found!!!"
        MsgBox "BMP file " + fileName + " not found!!!", vbOKOnly, "Error reading"
        ReadBMPTexture = -1
    End If
    Exit Function
errorH:
    MsgBox "Error opening " + fileName, vbCritical, "TEX Error " + Str$(Err)
End Function

Sub WriteBMPTexture(ByRef Bmp As BMPTexture, ByVal fileName As String)
    Dim NFile As Integer
    Dim ci As Integer
    
    On Error GoTo errorH
    
    NFile = FreeFile
    Open fileName For Binary As NFile
    
    With Bmp
        .fileName = fileName
        Put NFile, 1, .MagicId
        Put NFile, &H2 + 1, .FileSize
        Put NFile, &H6 + 1, .reserved1
        Put NFile, &H8 + 1, .reserved2
        Put NFile, &HA + 1, .ImgOffset
        Put NFile, &HE + 1, .HeaderSize
        Put NFile, &H12 + 1, .width
        Put NFile, &H16 + 1, .height
        Put NFile, &H1A + 1, .Planes
        Put NFile, &H1C + 1, .Bits
        Put NFile, &H1E + 1, .Compression
        Put NFile, &H22 + 1, .ImageSize
        Put NFile, &H26 + 1, .Xresolution
        Put NFile, &H2A + 1, .Yresolution
        Put NFile, &H2E + 1, .Ncolours
        Put NFile, &H32 + 1, .Importantcolours
    
        If .Bits <= 8 Then
            Put NFile, &H36 + 1, .Pallete
        End If
        Put NFile, .ImgOffset + 1, .BitMap
    End With
    Close NFile
    Exit Sub
errorH:
    MsgBox "Error saving " + fileName, vbCritical, "TEX Error " + Str$(Err)
End Sub

Sub GetBMPTextureFromBitmap(ByRef tex_out As BMPTexture, ByRef hdc As Long, ByVal hbmp As Long)
    Dim ci As Long
    Dim Bits As Integer
    
    Dim PicInfo As BITMAPINFO
    Dim PicData() As Byte
    
    GetAllBitmapData hdc, hbmp, PicData, PicInfo
    
    Bits = PicInfo.bmiHeader.biBitCount
    With tex_out
        .MagicId = &H4D42
        .ImgOffset = &H36
        If (Bits <= 8) Then _
            .ImgOffset = .ImgOffset + 2 ^ Bits * 4
        .HeaderSize = &H28
        .width = PicInfo.bmiHeader.biWidth
        .height = PicInfo.bmiHeader.biHeight
        .Planes = PicInfo.bmiHeader.biPlanes
        .Bits = Bits
        .Compression = 0
        .ImageSize = UBound(PicData) + 1
        .FileSize = .ImageSize + .ImgOffset
        .Xresolution = 1
        .Yresolution = 1
        If (Bits <= 8) Then
            .Ncolours = 2 ^ Bits
            .Importantcolours = .Ncolours
        Else
            .Ncolours = 0
            .Importantcolours = 0
        End If
        
        .BitMap = PicData
        
        If Bits <= 8 Then
            ReDim .Pallete(4 * .Ncolours - 1)
            CopyMemory .Pallete(0), PicInfo.bmiColors(0), 4 * .Ncolours
            
            For ci = 0 To .Ncolours - 1
                .Pallete(ci * 4 + 3) = &HFF
            Next ci
        End If
    End With
End Sub
