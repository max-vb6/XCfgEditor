Attribute VB_Name = "Main"
Option Explicit

Public cdlg As clsCdlg
Public Const lCfgVer As Long = 0
Public Const sFltr As String = "XCfg 文件 (*.xcfg)" & vbNullChar & "*.xcfg" & vbNullChar & "所有文件 (*.*)" & vbNullChar & "*.*" & vbNullChar

Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'====================SetCtrlsBrdClr====================
Private Type RECTW
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
    Width               As Long
    Height              As Long
End Type

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Const WM_DESTROY        As Long = &H2
Private Const WM_PAINT          As Long = &HF
Private Const WM_NCPAINT        As Integer = &H85
Private Const GWL_WNDPROC = (-4)
Private Color As Long
'====================SetCtrlsBrdClr====================

'====================SetCtrlsBrdClr====================
Public Sub setBorderColor(hwnd As Long, Color_ As Long)
    Color = Color_
    If GetProp(hwnd, "OrigProcAddr") = 0 Then
        SetProp hwnd, "OrigProcAddr", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    End If
End Sub

Public Sub UnHook(hwnd As Long)
    Dim OrigProc As Long
    OrigProc = GetProp(hwnd, "OrigProcAddr")
    If Not OrigProc = 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, OrigProc
        OrigProc = SetWindowLong(hwnd, GWL_WNDPROC, OrigProc)
        RemoveProp hwnd, "OrigProcAddr"
    End If
End Sub

Private Function OnPaint(OrigProc As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
    Dim m_hDC       As Long
    Dim m_wRect     As RECTW
    OnPaint = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
    Call pGetWindowRectW(hwnd, m_wRect)
    m_hDC = GetWindowDC(hwnd)
    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height)
    Call ReleaseDC(hwnd, m_hDC)
End Function

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim OrigProc As Long
    Dim ClassName As String
    If hwnd = 0 Then Exit Function
    OrigProc = GetProp(hwnd, "OrigProcAddr")
    If Not OrigProc = 0 Then
        If uMsg = WM_DESTROY Then
            SetWindowLong hwnd, GWL_WNDPROC, OrigProc
            WindowProc = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
            RemoveProp hwnd, "OrigProcAddr"
        Else
            If uMsg = WM_PAINT Or WM_NCPAINT Then

                WindowProc = OnPaint(OrigProc, hwnd, uMsg, wParam, lParam)
            Else
                WindowProc = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
            End If
        End If
    Else
        WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    End If
End Function

Private Function pGetWindowRectW(ByVal hwnd As Long, lpRectW As RECTW) As Long
    Dim TmpRect As RECT
    Dim Rtn     As Long
    Rtn = GetWindowRect(hwnd, TmpRect)
    With lpRectW
        .Left = TmpRect.Left
        .Top = TmpRect.Top
        .Right = TmpRect.Right
        .Bottom = TmpRect.Bottom
        .Width = TmpRect.Right - TmpRect.Left
        .Height = TmpRect.Bottom - TmpRect.Top
    End With
    pGetWindowRectW = Rtn
End Function

Private Function pFrameRect(ByVal hDC As Long, ByVal X As Long, Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
    Dim TmpRect     As RECT
    Dim m_hBrush    As Long
    With TmpRect
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With
    m_hBrush = CreateSolidBrush(Color)
    pFrameRect = FrameRect(hDC, TmpRect, m_hBrush)
    DeleteObject m_hBrush
End Function
'====================SetCtrlsBrdClr====================

Public Function GetFileName(Path As String, Optional GetEx As Boolean) As String
    On Error GoTo FileErr
    Dim tstrs() As String
    tstrs = Split(Path, "\")
    If GetEx Then GetFileName = tstrs(UBound(tstrs)): Exit Function
    tstrs = Split(tstrs(UBound(tstrs)), ".")
    GetFileName = tstrs(0)
    Exit Function
FileErr:
    GetFileName = ""
End Function

Public Function TrimFileName(sFileName As String) As String
    If Len(sFileName) < 5 Then
        TrimFileName = sFileName & ".xcfg"
    ElseIf Right(sFileName, 5) <> ".xcfg" Then
        TrimFileName = sFileName & ".xcfg"
    Else
        TrimFileName = sFileName
    End If
End Function

