Attribute VB_Name = "MainMod"
Option Explicit
   
Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Declare Function IsUserAnAdmin Lib "shell32" () As Long

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" ( _
    iccex As tagInitCommonControlsEx) As Boolean

Private Sub InitCommonControls()
    Dim iccex As tagInitCommonControlsEx
    
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    On Error Resume Next
    InitCommonControlsEx iccex
End Sub

Private Sub Main()
    Dim hWnd As Long
    If IsUserAnAdmin = False Then
        ShellExecute hWnd, "runas", "gta2pacifiststats.exe", "", CurDir$(), vbNormalFocus
        End
    End If
    InitCommonControls
    Pacifist.Show
    'Pacifist.Caption = "I am elevated: " & CStr(CBool(IsUserAnAdmin()))
End Sub

