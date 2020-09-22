VERSION 5.00
Begin VB.Form frmPicker 
   Caption         =   "Window Picker"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   Icon            =   "FWindowPicker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Normalize"
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Translucent"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      DragIcon        =   "FWindowPicker.frx":000C
      Height          =   495
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag the picker icon over the window you'd like to choose, then release to capture hWnd."
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©1997-2000 Sveinn R. Sigurðsson (MrHippo)
'  http://www.svenni.com
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

' Win32 API Declarations
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

' Win32 API Structures
Private Type POINTAPI
   x As Long
   y As Long
End Type

' Form-level member variables
Private m_hWnd As Long
Private m_hWndPick As Long
Private m_Picking As Boolean

Private Sub Command1_Click()
   If m_hWnd <> m_hWndPick Then
      ' Attempt to make selected window
      ' translucent.
      If SetWindowTranslucency(m_hWndPick, 160) Then
         ' Clear translucency on last window.
         Call ClearWindowTranslucency(m_hWnd)
         ' Cache new handle.
         m_hWnd = m_hWndPick
      End If
   End If
End Sub

Private Sub Command2_Click()
   ' Attempt to make selected window
   ' normal again.
   If ClearWindowTranslucency(m_hWnd) Then
      m_hWnd = 0
   End If
End Sub

Private Sub Form_Activate()
   ' Let user know if this probably won't work.
   If IsWin2000 = False Then
      MsgBox "Layered windows are only supported in " & _
         "Windows 2000.", vbExclamation, "Bummer"
   End If
End Sub

Private Sub Form_Load()
   ' Assign dragging pointer
   Picture1.Picture = Picture1.DragIcon
   Me.MouseIcon = Picture1.DragIcon
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Clear picture and turn on dragging mousepointer.
   Me.MousePointer = vbCustom
   Set Picture1.Picture = Nothing
   
   ' Remember that we're currently picking a window.
   m_Picking = True
   
   ' Capture all mousemovements from this point until
   ' the user releases the mouse button.
   Call SetCapture(Picture1.hWnd)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Static pt As POINTAPI
   ' If user is picking a window, check window is
   ' under the cursor whenever it moves.  If it's
   ' a different window than previously, update the
   ' display to that effect.
   If m_Picking Then
      Call GetCursorPos(pt)
      m_hWndPick = GetTopLevel(WindowFromPointXY(pt.x, pt.y))
      If m_hWndPick <> m_hWnd Then
         Me.Caption = Hex(m_hWndPick)
      End If
   End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' We're done picking now
   m_Picking = False
   
   ' Restore dragging icon to picture box,
   ' and return mousepointer to normal.
   Picture1.Picture = Picture1.DragIcon
   Me.MousePointer = vbDefault
   
   ' Don't need to be notified anymore.
   Call ReleaseCapture
   
   ' The chosen window is already stored in m_hWnd!
   'MsgBox "You picked hWnd: " & Hex(m_hWnd)
End Sub

Private Function GetTopLevel(ByVal hChild As Long) As Long
   Dim hWnd As Long
   
   ' Read parent chain up to highest visible.
   hWnd = hChild
   Do While IsWindowVisible(GetParent(hWnd))
      hWnd = GetParent(hChild)
      hChild = hWnd
   Loop
   GetTopLevel = hWnd
End Function
