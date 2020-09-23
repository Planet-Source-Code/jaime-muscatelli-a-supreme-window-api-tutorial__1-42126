VERSION 5.00
Begin VB.Form FRMMAIN 
   Caption         =   "SEND MESSAGE"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNotepad 
      Height          =   4095
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton CMDGETLINES 
      Caption         =   "&Get Number of Lines"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton CMDGETEDIT 
      Caption         =   "&Get Edit Text"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton CMDGETTEXT 
      Caption         =   "&Get Caption"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton CMDSENDEDIT 
      Caption         =   "&Send to Edit (Textbox)"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton CMDSENDTEXT 
      Caption         =   "&Send Caption"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "&Open Notepad"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "FRMMAIN"
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



Private Sub CMDGETEDIT_Click()
Dim lNotepadHwnd As Long
Dim lNotepadEdit As Long
Dim sCaption As String * 256 '// YES A BUFFER

    lNotepadHwnd = FindWindow("Notepad", vbNullString)
    lNotepadEdit = FindWindowEx(lNotepadHwnd, 0&, "Edit", vbNullString)

    
    
   
    SendMessageSTRING lNotepadEdit, WM_GETTEXT, 256, sCaption

txtNotepad.Text = sCaption
End Sub

Private Sub CMDGETLINES_Click()
Dim lNotepadHwnd As Long
Dim lNotepadEdit As Long
Dim lNumOfLines As Long '// NOT A BUFFER!

    lNotepadHwnd = FindWindow("Notepad", vbNullString)
    lNotepadEdit = FindWindowEx(lNotepadHwnd, 0&, "Edit", vbNullString)

    
    
   
    lNumOfLines = SendMessageLONG(lNotepadEdit, EM_GETLINECOUNT, 0, 0)
    

MsgBox "The notepad has " & lNumOfLines & " lines"
End Sub

Private Sub CMDGETTEXT_Click()
Dim lNotepadHwnd As Long
Dim lNotepadEdit As Long
Dim sCaption As String * 256
    
    lNotepadHwnd = FindWindow("Notepad", vbNullString)
    lNotepadEdit = FindWindowEx(lNotepadHwnd, 0&, "Edit", vbNullString)

    
    
    SendMessageSTRING lNotepadHwnd, WM_GETTEXT, 256, sCaption

MsgBox sCaption
End Sub

Private Sub cmdopen_Click()
Shell "notepad.exe"
End Sub

Private Sub CMDSENDEDIT_Click()
Dim lNotepadHwnd As Long
Dim lNotepadEdit As Long
Dim sCaption As String
    
    lNotepadHwnd = FindWindow("Notepad", vbNullString)
    lNotepadEdit = FindWindowEx(lNotepadHwnd, 0&, "Edit", vbNullString)

    sCaption = InputBox("What do you want to say?")
    
    SendMessageSTRING lNotepadEdit, WM_SETTEXT, 256, sCaption
End Sub

Private Sub CMDSENDTEXT_Click()
Dim lNotepadHwnd As Long
Dim sCaption As String
    
    lNotepadHwnd = FindWindow("Notepad", vbNullString)

    sCaption = InputBox("What do you want the caption to be?")
    
    SendMessageSTRING lNotepadHwnd, WM_SETTEXT, 256, sCaption

End Sub
