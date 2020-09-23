VERSION 5.00
Begin VB.Form FRMVIRUS 
   Caption         =   "Fake Virus  - API.txt"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrVirus 
      Interval        =   250
      Left            =   2640
      Top             =   960
   End
   Begin VB.CommandButton CMDOPEN 
      Caption         =   "&Open Window"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Notepad is: "
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "FRMVIRUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageLONG Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_SETTEXT = &HC
Private Const EM_GETLINECOUNT = &HBA

Private Sub CMDOPEN_Click()
Shell "notepad.exe"
End Sub

Private Sub tmrVirus_Timer()
Dim lNotepadHwnd As Long
Dim lNotepadEdit As Long
Dim sCaption As String
    
    lNotepadHwnd = FindWindow("Notepad", vbNullString)
    lNotepadEdit = FindWindowEx(lNotepadHwnd, 0&, "Edit", vbNullString)

    If lNotepadHwnd Then
    lblStatus.Caption = "Open"
    Else
    lblStatus.Caption = "Closed"
    End If
    
    SendMessageSTRING lNotepadHwnd, WM_SETTEXT, 256, "How does it feel?"
    SendMessageSTRING lNotepadEdit, WM_SETTEXT, 256, "I told you so!"

End Sub
