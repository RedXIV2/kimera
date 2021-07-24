VERSION 5.00
Begin VB.Form LogWindow
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log Window"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Log
      Height          =   1815
      ItemData        =   "StateWindow.frx":0000
      Left            =   0
      List            =   "StateWindow.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "LogWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Jobs() As String
Dim JobsTotal() As Long
Dim JobsCurrent() As Long
Dim NumJobs As Integer
Public Sub UpdateState(ByVal j_current As Long, ByVal sj_current As Long)

End Sub

