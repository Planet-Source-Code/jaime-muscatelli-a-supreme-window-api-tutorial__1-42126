VERSION 5.00
Begin VB.Form FRMMAIN 
   Caption         =   "GET PARENT, WINDOWFROMPOINT,ETC..."
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtclass 
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Timer tmr1 
      Interval        =   500
      Left            =   120
      Top             =   2400
   End
End
Attribute VB_Name = "FRMMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GETPARENT Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private MOUSE As POINTAPI
Private lOLDHWND As Long '// DO WE DON"T DO THE SAME HWND 2ce
Private Sub MAKEWINDOW(x As Long, y As Long)
'// THIS API SPY IS NOT COMPLETE BECAUSE IT DOESN"T RUN THROUGH EVERY
'WINDOW, but just the selected window and the immediate parent

Dim lhWnd As Long
Dim lParentHwnd As Long
Dim sClassName As String * 256 '/// YES A BUFFER!!!!!!
Dim sParentClass As String * 256 '/// YES A BUFFER!!!!!!
Dim I As Integer

lhWnd = WindowFromPoint(x, y)

lParentHwnd = GETPARENT(lhWnd) '// Get Parent

If lhWnd = lOLDHWND Then Exit Sub '// DO WE DON"T DO THE SAME HWND 2ce

If lhWnd = lParentHwnd Then Exit Sub

GetClassName lhWnd, sClassName, 256
txtclass.Text = txtclass.Text & vbCrLf & sClassName
GetClassName lParentHwnd, sParentClass, 256
txtclass.Text = txtclass.Text & vbCrLf & sParentClass

lOLDHWND = lParentHwnd
End Sub

Private Sub tmr1_Timer()
GetCursorPos MOUSE
MAKEWINDOW MOUSE.x, MOUSE.y
End Sub
