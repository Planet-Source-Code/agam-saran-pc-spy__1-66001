Attribute VB_Name = "modSubclass"
Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal MSG As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public Const WM_QUERYENDSESSION = &H11

Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Type POINTAPI
        x As Long
        Y As Long
End Type

Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Global Const WM_CANCELMODE = &H1F
Public Const REG_SZ = 1
Public Const HKEY_LOCAL_MACHINE = &H80000002
Global lpPrevWndProc As Long
Global gHW As Long

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim temp As Long
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub
Function WindowProc(ByVal hw As Long, ByVal uMsg As _
Long, ByVal wParam As Long, ByVal lParam As Long) As _
Long
Dim a As Long
    If uMsg = WM_QUERYENDSESSION Then
        Close #1
        Open App.Path & "\pc.log" For Append As #1
            Print #1, Code("-" & Date & " " & time)
        Close #1
        Open App.Path & "\spy.ini" For Output As #1
            Print #1, "&" & frmMain.optLimited.Value
        Close
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
        Exit Function
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, _
    uMsg, wParam, lParam)
End Function
Public Function Code(ByVal Text As String) As String
Dim i As Integer

        For i = 1 To Len(Text)
            Code = Code & Chr$(255 - Asc(Mid$(Text, i, 1)))
        Next
End Function

Public Sub SaveString(hKey As Long, strPath As String, ByVal strValue As String, ByVal strData As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(hKey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
r = RegCloseKey(keyhand)
End Sub
