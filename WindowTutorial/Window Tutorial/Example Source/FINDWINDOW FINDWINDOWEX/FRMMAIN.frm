VERSION 5.00
Begin VB.Form FRMMAIN 
   Caption         =   "FINDWINDOW and FINDWINDOWEX"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDFINDWNDOW 
      Caption         =   "&Find Window and EX"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Notepad"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FRMMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FINDWINDOW Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Sub CMDFINDWNDOW_Click()
'// This sub will basically detect if the window is open
'// FindWINDOW and FINDWINDOWEX alone aren't really usefull,
'// You only use them to get the window handle to a program you need
'// to access

Dim lNotepadHwnd As Long
Dim lNotepadEdit As Long

lNotepadHwnd = FINDWINDOW("Notepad", vbNullString)
lNotepadEdit = FindWindowEx(lNotepadHwnd, 0&, "Edit", vbNullString)

'Now it will detect if they are open

If lNotepadEdit Then '// You could have used lNotepadHwnd
MsgBox "Notepad is open"
Else
MsgBox "Notepad is not open"
End If
End Sub

Private Sub cmdOpen_Click()
Shell "notepad.exe"
End Sub

