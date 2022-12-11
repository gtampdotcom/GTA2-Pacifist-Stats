Attribute VB_Name = "modControlWin"
Option Explicit
Dim bln98 As Boolean

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Const WM_ACTIVATE = &H6
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const WM_VSCROLL = &H115
Public Const SB_BOTTOM = 7
Public Const WM_COMMAND = &H111
Public Const WM_SETTEXT = &HC
Public Const EM_SETSEL = &HB1

Public intPrevListCount As Integer

Public lngProcessID As Long
Public lngHandleGTA2 As Long
Public lngHandleLV As Long
Public lngHandleHistory As Long
Public lngHandleJoinHistory As Long
Public lngHandleReject As Long
Public lngHandleStart As Long
Public lngHandleCancel As Long
Public lngHandleChat As Long
'Public lngHandleJoinChat As Long
Public lngHandleSend As Long
Public lngHandleMaps As Long
Public lngHandlePlayersRequired As Long
Public lngHandleGameSpeed As Long
Public lngHandleSpeed As Long
Public lngHandleGameType As Long
Public lngHandleScoreLimit As Long
Public lngHandleTimeLimit As Long
Public lngHandleCops As Long
Public lngHandleScoreLimitLabel As Long
Public lnghandleTimeLimitLabel As Long

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)

Public Const WM_USER = &H400
Public Const TBM_SETPOS = (WM_USER + 5)
Public Const BM_CLICK As Long = &HF5
Public Const WM_GETTEXT As Integer = &HD
Public Const WM_GETTEXTLENGTH As Long = &HE

Public Const GWL_ID = (-12)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32.dll" _
(ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Function EnumChildWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    On Error GoTo oops
    EnumChildWindowsProc = 1
    
    Select Case Val(GetWindowLong(hWnd, GWL_ID))
        Case 1024
            lngHandleLV = hWnd
        Case 1022
            lngHandleHistory = hWnd
        Case 1051
            lngHandleJoinHistory = hWnd
        Case 1020
            lngHandleReject = hWnd
        Case 1021
            lngHandleStart = hWnd
        Case 2
            lngHandleCancel = hWnd
        Case 1025
            lngHandleChat = hWnd
        'Case 1053
        '    lngHandleJoinChat = hwnd
        Case 1023
            lngHandleSend = hWnd
        Case 1026
            lngHandleMaps = hWnd
            'Call getMapDescFromLV
        Case 1033
            lngHandlePlayersRequired = hWnd
        'Case 1031
        '    lngHandleSpeed = hWnd
        Case 1032
             lngHandleGameSpeed = hWnd
        Case 1036
            lngHandleGameType = hWnd
        Case 1059
            lngHandleScoreLimit = hWnd
        Case 1038
            lngHandleTimeLimit = hWnd
        Case 1027
            lngHandleCops = hWnd
        Case 1035
            lngHandleScoreLimitLabel = hWnd
        Case 1037
            lnghandleTimeLimitLabel = hWnd
    End Select

Exit Function

oops:
    MsgBox Err.Description
End Function

'Some of this function was from here: http://www.ex-designz.net/apidetail.asp?api_id=316
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
On Error GoTo oops
'Static winnum As Integer ' counter keeps track of how many windows have been enumerated
'winnum = winnum + 1 ' one more window enumerated....
''Rebug.Print winnum
EnumWindowsProc = 1 ' return value of 1 means continue enumeration

Dim lngTemp As Long

Call GetWindowThreadProcessId(hWnd, lngTemp)

'If lngTemp = lngPID Then
    'lngProcessID = lngTemp
'Else
    'Exit Function
'End If

'Dim hwndTarget As Long
'hwndTarget = FindWindow(vbNullString, "Game Hunter v1.548")
'Const GWL_HINSTANCE = (-6)
'Dim hInstance As Long
'hInstance = GetWindowLong(hwndTarget, GWL_HINSTANCE)
'Debug.Print hInstance & " " & WindowTitle(hwndTarget) & " " & App.hInstance
'SendMessage hwndTarget, WM_SETTEXT, 0, ByVal "test: " & Rand(0, 999)


If InStr(WindowTitle(hWnd), "GTA2") Then
    Select Case ClassName(hWnd)
        Case "#32770"
            lngProcessID = lngTemp
            EnumChildWindows hWnd, AddressOf EnumChildWindowsProc, 1

        Case "WinMain"
            EnumWindowsProc = 0
    End Select
End If

Exit Function

oops:
    'Pacifist.txtError = Pacifist.txtError & vbCrLf & "EnumWindowsProc: " & Err.Description & " " & Erl
End Function

Public Function WindowTitle(ByVal lHwnd As Long) As String
On Error GoTo oops

Dim slength As Long
Dim Buffer As String
Dim retval As Long

slength = GetWindowTextLength(lHwnd) + 1 ' get length of title bar text
If slength > 1 Then ' if return value refers to non-empty string
    Buffer = Space(slength) ' make room in the buffer
    retval = GetWindowText(lHwnd, Buffer, slength) ' get title bar text
    WindowTitle = Left(Buffer, slength - 1) ' display title bar text of enumerated window
    'frmHax.txtHistory.Text = frmHax.txtHistory.Text & WindowTitle & vbNewLine
End If
   
Exit Function

oops:
    'Pacifist.txtError = Pacifist.txtError & "WindowTitle error: " & Err.Description
End Function

Public Function ClassName(ByVal lHwnd As Long) As String
On Error GoTo oops
Dim lLen As Long
Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If

Exit Function

oops:
    MsgBox "ClassName error: " & Err.Description
End Function

Private Function GetListViewCount(ByVal hWnd As Long) As Long
    'this simply get number of items
    GetListViewCount = SendMessage(hWnd, LVM_GETITEMCOUNT, 0, ByVal 0)
End Function


Private Sub getMapDescFromLV()

Const CB_GETCURSEL = &H147
Const CB_GETLBTEXT = &H148
Const CB_GETLBTEXTLEN = &H149
Const CB_GETCOUNT = &H146
Dim count As Long       ' number of items in the combo box
count = SendMessage(lngHandleMaps, CB_GETCOUNT, ByVal CLng(0), ByVal CLng(0)) - 1

' Display the text of whatever item in combo box Combo1
' is currently selected.  If no list box item is selected, say so.
Dim Index As Long       ' index to the selected item
Dim itemtext As String  ' the text of the selected item
Dim textlen As Long     ' the length of the selected item's text

' Determine the index of the selected item.
Index = SendMessage(lngHandleMaps, CB_GETCURSEL, ByVal CLng(0), ByVal CLng(0))
textlen = SendMessage(lngHandleMaps, CB_GETLBTEXTLEN, ByVal CLng(Index), ByVal CLng(0))
' Make enough room in the string to receive the text, including the terminating null.
itemtext = Space(textlen) & vbNullChar
' Retrieve that item's text and display it.
textlen = SendMessage(lngHandleMaps, CB_GETLBTEXT, ByVal CLng(Index), ByVal itemtext)
itemtext = Left(itemtext, textlen)
'Debug.Print itemtext

End Sub
