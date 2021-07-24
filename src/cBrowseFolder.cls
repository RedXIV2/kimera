VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cBrowseFolder
' 2000/12/17 Copyright � 2000, Larry Rebich, using the VAIO
' 2000/12/17 larry@buygold.net, www.buygold.net, 760.771.4730
'            Some parts of this code from Randy Birch and others.
'
' Use this class module in conjunction with module: mBrowseFolder.bas

    Option Explicit
    DefLng A-Z

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Some parts copyright �1996-2000 VBnet, Randy Birch, All Rights Reserved.
' http://www.mvps.org/vbnet/index.html
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Type BROWSEINFO
       hOwner           As Long
       pidlRoot         As Long
       pszDisplayName   As String
       lpszTitle        As String
       ulFlags          As Long
       lpfn             As Long
       lParam           As Long
       iImage           As Long
    End Type

    Dim BI As BROWSEINFO    'allocate and name storage for the structure

    Private Declare Function SHGetPathFromIDList _
        Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
       (ByVal pidl As Long, ByVal pszPath As String) As Long

    Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        Alias "SHBrowseForFolderA" _
       (lpBrowseInfo As BROWSEINFO) As Long

    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

    Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long

    Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long

    Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

    Const LMEM_FIXED = &H0
    Const LMEM_ZEROINIT = &H40
    Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

    Private mlhWnd      As Long
    Private msTitle     As String
    Private msFolder    As String
    Private mbCancelled As Boolean
    Private mlFlags     As Long

Public Property Let lFlags(lFlags As Long)
    mlFlags = lFlags
End Property
Public Property Get lFlags() As Long
    lFlags = mlFlags
End Property

Public Property Let bCancelled(bCancelled As Boolean)
    mbCancelled = bCancelled
End Property
Public Property Get bCancelled() As Boolean
    bCancelled = mbCancelled
End Property

Public Property Let lhWnd(lhWnd As Long)
    mlhWnd = lhWnd
End Property
Public Property Get lhWnd() As Long
    lhWnd = mlhWnd
End Property

Public Property Let sTitle(sTitle As String)
    msTitle = sTitle
End Property
Public Property Get sTitle() As String
    sTitle = msTitle
End Property

Public Property Let sFolder(sFolder As String)
    msFolder = sFolder
End Property
Public Property Get sFolder() As String
    sFolder = msFolder
End Property

Public Function ShowBrowse(Optional vnthWnd As Variant, Optional vntTitle As Variant, Optional vntFolder As Variant) As String
    Dim lpSelPath As Long
    Dim lRtn As Long
    Dim pidl As Long
    Dim iPos As Integer
    Const MAX_PATH = 260
    Dim sFolder As String * MAX_PATH

    'allow user to specify params in the call
    If Not IsMissing(vnthWnd) Then
        Me.lhWnd = vnthWnd
    End If
    If Not IsMissing(vntTitle) Then
        Me.sTitle = vntTitle
    End If
    If Not IsMissing(vntFolder) Then
        Me.sFolder = vntFolder
    End If

    With BI                     'initilize the structure
        .hOwner = Me.lhWnd      'owner's hWnd
        .ulFlags = Me.lFlags    'flags, default is BIF_RETURNONLYFSDIRS
        .lpszTitle = Me.sTitle  'dialog's title
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)    'set to pass an address into a structure
        lpSelPath = LocalAlloc(LPTR, Len(Me.sFolder))
        MoveMemory ByVal lpSelPath, ByVal Me.sFolder, Len(Me.sFolder)
        .lParam = lpSelPath     'now into structure
    End With

    pidl = SHBrowseForFolder(BI)    'do the browse

    If pidl Then                    'if not cancelled
        If SHGetPathFromIDList(pidl, sFolder) Then  'get the folder
            sFolder = Left$(sFolder, InStr(sFolder, vbNullChar) - 1)    'dump all after null character
        End If
        Call CoTaskMemFree(pidl)    'free memory
    Else
        Me.bCancelled = True        'report cancelled
    End If

    Call LocalFree(lpSelPath)       'free memory

    Me.sFolder = sFolder            'report the result

    If Me.sFolder = "" Then         'if none, then say cancelled
        Me.bCancelled = True
    Else
        ShowBrowse = Me.sFolder     'else, return the folder
    End If
End Function

Private Sub Class_Initialize()
    Me.lFlags = BIF_ReturnOnlyFSDirs    'default, return only real drives, no virtuals.
End Sub
