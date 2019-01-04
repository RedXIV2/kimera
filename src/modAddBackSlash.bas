Attribute VB_Name = "modAddBackSlash"
' modAddBackSlash
' 2000/04/27 Copyright © 2000, Larry Rebich
' 2000/04/27 larry@buygold.net, www.buygold.net, 760.771.4730
' 2000/10/01 Used in BrandingModel
' 2000/10/01 Used in Branding

    Option Explicit
    DefLng A-Z
    
    Const mcsBkSlash = "\"
    
Public Function AddBackslash(sThePath As String) As String
' Add a backslash to a path if needed
' sPath contains the path
' Return a path with a backslash
    If Right$(sThePath, 1) <> mcsBkSlash Then
        sThePath = sThePath + mcsBkSlash
    End If
    AddBackslash = sThePath
End Function


