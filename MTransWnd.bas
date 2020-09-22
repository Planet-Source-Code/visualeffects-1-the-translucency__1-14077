Attribute VB_Name = "MTranslucentWnd"
' *************************************************************************
'  Copyright ©2000 Sveinn R. Sigurdsson (MrHippo)
'  All Rights Reserved, http://www.svenni.com
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' BOOL SetLayeredWindowAttributes(
'   HWND hwnd,       // handle to the layered window
'   COLORREF crKey,  // specifies the color key
'   BYTE bAlpha,     // value for the blend function
'   DWORD dwFlags    // action
' );
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Const LWA_COLORKEY = &H1&
Private Const LWA_ALPHA = &H2&

' Style setting APIs
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

' Win32 APIs to determine OS information.
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Used to determine parentage.
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Public Function ClearWindowTranslucency(ByVal hWnd As Long) As Boolean
   Dim nStyle As Long
   If IsWin2000 Then
      ' Only work with top-level.
      hWnd = GetTopLevel(hWnd)
      ' Set translucency to fully
      ' opaque (255).
      Call SetLayeredWindowAttributes(hWnd, 0, 255&, LWA_ALPHA)
      ' Clear exstyle bit.
      nStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
      ClearWindowTranslucency = CBool(SetWindowLong(hWnd, GWL_EXSTYLE, nStyle))
   End If
End Function

Public Function SetWindowTranslucency(ByVal hWnd As Long, ByVal Alpha As Byte) As Boolean
   Dim nStyle As Long
   If IsWin2000 Then
      ' Only work with top-level.
      hWnd = GetTopLevel(hWnd)
      ' Set exstyle bit.
      nStyle = GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
      If SetWindowLong(hWnd, GWL_EXSTYLE, nStyle) Then
         ' Set window translucency to
         ' requested Alpha value.
         SetWindowTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, CLng(Alpha), LWA_ALPHA))
      End If
   End If
End Function

Public Function IsWin2000() As Boolean
   Dim os As OSVERSIONINFO
   ' Layered windows are only available in
   ' Windows 2000. This function shouldn't
   ' be called often, so check on demand.
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      IsWin2000 = (os.dwMajorVersion >= 5)
   End If
End Function

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

