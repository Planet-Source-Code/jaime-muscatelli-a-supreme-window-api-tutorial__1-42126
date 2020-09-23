VERSION 5.00
Begin VB.Form FRMMAIN 
   Caption         =   "GET and SET WINDOWTEXT Example"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDOPEN 
      Caption         =   "&OPEN NOTEPAD"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton CMDSETWINDOWTEXT 
      Caption         =   "&Set Notepad Caption"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton CMDGETTEXT 
      Caption         =   "&Get Notepad Caption"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "FRMMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GETWINDOWTEXT Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Sub CMDGETTEXT_Click()

Dim lNotepadhWnd As Long
Dim sCaption As String * 256

    lNotepadhWnd = FindWindow("Notepad", vbNullString)

    GETWINDOWTEXT lNotepadhWnd, sCaption, 256

    MsgBox sCaption

End Sub

Private Sub CMDOPEN_Click()
Shell "notepad.exe"
End Sub

Private Sub CMDSETWINDOWTEXT_Click()
Dim lNotepadhWnd As Long
    Dim sNotepadText As String

    lNotepadhWnd = FindWindow("Notepad", vbNullString)

    sNotepadText = InputBox("What do you want the caption to say?")

    SetWindowText lNotepadhWnd, sNotepadText
End Sub
