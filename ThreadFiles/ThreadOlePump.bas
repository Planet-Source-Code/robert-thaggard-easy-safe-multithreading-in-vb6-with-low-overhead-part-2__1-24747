Attribute VB_Name = "ThreadOlePump"
Option Explicit

Private Const cstrOleThreadClass As String = "OleMainThreadWndClass"
Private Const cstrOleThreadWndName As String = "OleMainThreadWndName"
Private Const cstrWin95RPCClass As String = "WIN95 RPC Wmsg"
Private Const cstrNTAlternate As String = "VBMsoStdCompMgr"
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT  As Long = 2

Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(0 To 127) As Byte 'Don't care about string info, leave as byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hWnd As hWnd, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextA Lib "user32" (ByVal hWnd As hWnd, ByVal lpString As String, ByVal nMaxCount As Long) As Long

Private m_fInit As Boolean
Private m_fWin95RPC As Boolean
Private m_fTryNTAlternate As Boolean
Private m_WndEnumProc As Long
Private m_ClassNameBuf As String

Private Function WndEnumProcW(ByVal hWnd As Long, OlehWnd As Long) As BOOL
Dim CharLen As Long
    CharLen = GetClassNameW(hWnd, m_ClassNameBuf, 255)
    If CharLen = Len(cstrOleThreadClass) Then
        OlehWnd = hWnd
        If Left$(m_ClassNameBuf, CharLen) = cstrOleThreadClass Then
            CharLen = GetWindowTextW(hWnd, StrPtr(m_ClassNameBuf), 255)
            If CharLen Then
                If Left$(m_ClassNameBuf, CharLen) = cstrOleThreadWndName Then
                    'If we find one with this window name, then its the desired window
                    'Otherwise, any window with this class will do.
                    Exit Function
                End If
            End If
        End If
    End If
    WndEnumProcW = BOOL_TRUE
End Function
Private Function WndEnumProcWNTAlternate(ByVal hWnd As Long, OlehWnd As Long) As BOOL
Dim CharLen As Long
    CharLen = GetClassNameW(hWnd, m_ClassNameBuf, 255)
    If CharLen = Len(cstrNTAlternate) Then
        If Left$(m_ClassNameBuf, CharLen) = cstrNTAlternate Then
            OlehWnd = hWnd
            Exit Function
        End If
    End If
    WndEnumProcWNTAlternate = BOOL_TRUE
End Function
Private Function WndEnumProcA(ByVal hWnd As Long, OlehWnd As Long) As BOOL
    If GetClassNameA(hWnd, m_ClassNameBuf, 255) Then
        If Left$(m_ClassNameBuf, Len(cstrOleThreadClass)) = cstrOleThreadClass Then
            OlehWnd = hWnd
            If GetWindowTextA(hWnd, m_ClassNameBuf, 255) Then
                If Left$(m_ClassNameBuf, InStr(m_ClassNameBuf, vbNullChar) - 1) = cstrOleThreadWndName Then
                    'If we find one with this window name, then its the desired window
                    'Otherwise, any window with this class will do.
                    Exit Function
                End If
            End If
        End If
    End If
    WndEnumProcA = BOOL_TRUE
End Function
Private Function WndEnumProcAWin95RPC(ByVal hWnd As Long, OlehWnd As Long) As BOOL
    If GetClassNameA(hWnd, m_ClassNameBuf, 255) Then
        If 1 = InStr(Left$(m_ClassNameBuf, InStr(m_ClassNameBuf, vbNullChar) - 1), cstrWin95RPCClass) Then
            OlehWnd = hWnd
            Exit Function
        End If
    End If
    WndEnumProcAWin95RPC = BOOL_TRUE
End Function
Private Function FuncAddr(ByVal lpfn As Long) As Long
    FuncAddr = lpfn
End Function
Private Function InitCallbacks() As Boolean
Dim OSVI As OSVERSIONINFO
    OSVI.dwOSVersionInfoSize = Len(OSVI)
    If GetVersionEx(OSVI) Then
        If OSVI.dwPlatformId And VER_PLATFORM_WIN32_NT Then
            m_WndEnumProc = FuncAddr(AddressOf WndEnumProcW)
            m_fTryNTAlternate = True
        ElseIf OSVI.dwPlatformId And VER_PLATFORM_WIN32_WINDOWS Then
            If OSVI.dwMinorVersion = 0 Then
                m_fWin95RPC = True
                m_WndEnumProc = FuncAddr(AddressOf WndEnumProcAWin95RPC)
            Else 'Win98, look for same class as NT
                m_WndEnumProc = FuncAddr(AddressOf WndEnumProcA)
            End If
        Else
            Exit Function
        End If
        m_fInit = True
        InitCallbacks = True
    End If
End Function

Public Function FindOLEhWnd() As Long
Dim WndEnumAlternate As Long
    If Not m_fInit Then
        If Not InitCallbacks Then Exit Function
    End If
    m_ClassNameBuf = String$(255, 0)
    EnumThreadWindows App.ThreadID, m_WndEnumProc, VarPtr(FindOLEhWnd)
    If FindOLEhWnd = 0 Then
        If m_fWin95RPC Then
            WndEnumAlternate = FuncAddr(AddressOf WndEnumProcA)
            EnumThreadWindows App.ThreadID, WndEnumAlternate, VarPtr(FindOLEhWnd)
            If FindOLEhWnd Then
                m_fWin95RPC = False
                m_WndEnumProc = WndEnumAlternate
            End If
        ElseIf m_fTryNTAlternate Then
            WndEnumAlternate = FuncAddr(AddressOf WndEnumProcWNTAlternate)
            EnumThreadWindows App.ThreadID, WndEnumAlternate, VarPtr(FindOLEhWnd)
            If FindOLEhWnd Then
                m_fTryNTAlternate = False
                m_WndEnumProc = WndEnumAlternate
            End If
        End If
    End If
    m_ClassNameBuf = vbNullString
End Function
Public Sub SpinOlehWnd(ByVal hWnd As Long, ByVal fYield As Boolean)
Dim PMFlags As PMOptions
Dim wMsgMin As Long
Dim wMsgMax As Long
Dim MSG As MSG
    If fYield Then
        PMFlags = PM_REMOVE
    Else
        PMFlags = PM_REMOVE Or PM_NOYIELD
    End If
    If hWnd = 0 Then
        'Not sure which window to spin on (this is very unlikely)
        'A PeekMessage loop on all windows can still beat DoEvents
        'and reduce side effects by just looking at WM_USER messages
        'and higher
        wMsgMin = &H400 'WM_USER
        wMsgMax = &H7FFF
    End If
    Do While PeekMessage(MSG, hWnd, wMsgMin, wMsgMax, PMFlags)
        TranslateMessage MSG 'Probably does nothing, but technically correct
        DispatchMessage MSG
    Loop
End Sub
