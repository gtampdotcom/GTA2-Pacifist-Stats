VERSION 5.00
Begin VB.Form Pacifist 
   BackColor       =   &H80000001&
   Caption         =   "GTA2 Pacifist Stats v0.3 9.6(CD)"
   ClientHeight    =   4305
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6180
   Icon            =   "Pacifist.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timer 
      Interval        =   500
      Left            =   3480
      Top             =   1800
   End
   Begin VB.Label lblWantedLevelVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblWantedLevel 
      BackColor       =   &H80000001&
      Caption         =   "Wanted level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label lblFugitiveFactorVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblFujitiveFactor 
      BackColor       =   &H80000001&
      Caption         =   "Fugitive factor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Label lblAutoDamageCostVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblVehiclesHijackedVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label lblAutoDamageCost 
      BackColor       =   &H80000001&
      Caption         =   "Auto damage cost:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000001&
      Caption         =   "Vehicles hijacked:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblCiviliansRunDown 
      BackColor       =   &H80000001&
      Caption         =   "Civilians run down:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label lblCiviliansRunDownVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Civilians killed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000001&
      Caption         =   "Lawmen killed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000001&
      Caption         =   "Gang members killed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label lblCiviliansKilledVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblLawmenKilledVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblGangMembersKilledVal 
      BackColor       =   &H80000001&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Files"
      Begin VB.Menu mnuSettingsVehiclesHijacked 
         Caption         =   "Create vehicles_hijacked.txt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettingsAutoDamageCost 
         Caption         =   "Create auto_damage_cost.txt"
      End
      Begin VB.Menu mnuSettingsCiviliansRundown 
         Caption         =   "Create civilians_rundown.txt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettingsCiviliansKilled 
         Caption         =   "Create civilians_killed.txt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettingsLawmenKilled 
         Caption         =   "Create lawmen_killed.txt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettingsGangMembersKilled 
         Caption         =   "Create gang_members_killed.txt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettingsFugitiveFactor 
         Caption         =   "Create fugitive_factor.txt"
      End
   End
End
Attribute VB_Name = "Pacifist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim old_vehicles_hijacked As Long
Dim old_auto_damage_cost As Long
Dim old_civilians_run_down As Long
Dim old_civilians_killed As Long
Dim old_lawmen_killed As Long
Dim old_gang_members_killed As Long
Dim old_fugitive_factor As Long
Dim old_wanted_level As Long

Dim blnApply As Boolean

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Declare Function RPM Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WPM Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const BCM_SETSHIELD As Long = &H160C&

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Sub ShowDLLError()
    Dim strBuff As String, intLen As Integer
    strBuff = String$(200, vbNullChar)
    intLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Err.LastDllError, 0, strBuff, 200, ByVal 0)
    strBuff = Left$(strBuff, intLen - 2)
    'txtError = strBuff
End Sub

Private Sub Form_Load()
    'SendMessage Command1.hWnd, BCM_SETSHIELD, 0&, 1&
End Sub

Private Sub readMemory()

Dim lngRet As Long
Dim stats_addr As Long
Dim vehicles_hijacked As Long
Dim auto_damage_cost As Long
Dim civilians_run_down As Long
Dim civilians_killed As Long
Dim lawmen_killed As Long
Dim gang_members_killed As Long
Dim fugitive_factor As Long
Dim wanted_level As Long

Dim hWnd As Long

Dim pid As Long
Dim pHandle As Long
hWnd = FindWindow(vbNullString, "GTA2")

If (hWnd = 0) Then
 Exit Sub
End If
 
GetWindowThreadProcessId hWnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)

'Const stats_ptr = &H5E3CC4 '11.44
Const stats_ptr = &H66D9C0 '9.6CD
Const vehicles_hijacked_ptr = &H50C
Const auto_damage_cost_ptr = &H52C
Const civilians_run_down_ptr = &H510
Const civilians_killed_ptr = &H514
Const lawmen_killed_ptr = &H518
Const gang_members_killed_ptr = &H51C
Const fugitive_factor_ptr = &H530

RPM pHandle, ByVal stats_ptr, stats_addr, 4, lngRet
RPM pHandle, ByVal stats_addr + vehicles_hijacked_ptr, vehicles_hijacked, 4, lngRet
RPM pHandle, ByVal stats_addr + auto_damage_cost_ptr, auto_damage_cost, 4, lngRet
RPM pHandle, ByVal stats_addr + civilians_run_down_ptr, civilians_run_down, 4, lngRet
RPM pHandle, ByVal stats_addr + civilians_killed_ptr, civilians_killed, 4, lngRet
RPM pHandle, ByVal stats_addr + lawmen_killed_ptr, lawmen_killed, 4, lngRet
RPM pHandle, ByVal stats_addr + gang_members_killed_ptr, gang_members_killed, 4, lngRet
RPM pHandle, ByVal stats_addr + fugitive_factor_ptr, fugitive_factor, 4, lngRet

lblVehiclesHijackedVal = vehicles_hijacked
lblAutoDamageCostVal = auto_damage_cost
lblCiviliansRunDownVal = civilians_run_down
lblCiviliansKilledVal = civilians_killed
lblLawmenKilledVal = lawmen_killed
lblGangMembersKilledVal = gang_members_killed
lblFugitiveFactorVal = fugitive_factor
CloseHandle pHandle

If mnuSettingsVehiclesHijacked.Checked Then
    If vehicles_hijacked <> old_vehicles_hijacked Then
        Open "vehicles_hijacked.txt" For Output As #1
        Print #1, Trim(vehicles_hijacked);
        Close #1
    End If
End If

If mnuSettingsAutoDamageCost.Checked Then
    If auto_damage_cost <> old_auto_damage_cost Then
        Open "auto_damage_cost.txt" For Output As #1
        Print #1, Trim(auto_damage_cost);
        Close #1
    End If
End If

If mnuSettingsCiviliansRundown.Checked Then
    If civilians_run_down <> old_civilians_run_down Then
        Open "civilians_run_down.txt" For Output As #1
        Print #1, Trim(civilians_run_down);
        Close #1
    End If
End If

If mnuSettingsCiviliansKilled.Checked Then
    If civilians_killed <> old_civilians_killed Then
        Open "civilians_killed.txt" For Output As #1
        Print #1, Trim(civilians_killed);
        Close #1
    End If
End If

If mnuSettingsLawmenKilled.Checked Then
    If lawmen_killed <> old_lawmen_killed Then
        Open "lawmen_killed.txt" For Output As #1
        Print #1, Trim(lawmen_killed);
        Close #1
    End If
End If

If mnuSettingsGangMembersKilled.Checked Then
    If gang_members_killed <> old_gang_members_killed Then
        Open "gang_members_killed.txt" For Output As #1
        Print #1, Trim(gang_members_killed);
        Close #1
    End If
End If

If mnuSettingsFugitiveFactor.Checked Then
    If fugitive_factor <> old_fugitive_factor Then
        Open "fugitive_factor.txt" For Output As #1
        Print #1, Trim(fugitive_factor);
        Close #1
    End If
End If

old_vehicles_hijacked = vehicles_hijacked
old_auto_damage_cost = auto_damage_cost
old_civilians_run_down = civilians_run_down
old_civilians_killed = civilians_killed
old_lawmen_killed = lawmen_killed
old_gang_members_killed = gang_members_killed
old_fugitive_factor = fugitive_factor
CloseHandle pHandle

End Sub

Private Sub mnuSettingsCiviliansKilled_Click()
    mnuSettingsCiviliansKilled.Checked = Not mnuSettingsCiviliansKilled.Checked
End Sub

Private Sub mnuSettingsCiviliansRundown_Click()
    mnuSettingsCiviliansRundown.Checked = Not mnuSettingsCiviliansRundown.Checked
End Sub

Private Sub mnuSettingsAutoDamageCost_Click()
    mnuSettingsAutoDamageCost.Checked = Not mnuSettingsAutoDamageCost.Checked
End Sub

Private Sub mnuSettingsVehiclesHijacked_Click()
    mnuSettingsVehiclesHijacked.Checked = Not mnuSettingsVehiclesHijacked.Checked
End Sub

Private Sub mnuSettingsFugitiveFactor_Click()
    mnuSettingsFugitiveFactor.Checked = Not mnuSettingsFugitiveFactor.Checked
End Sub

Private Sub mnuSettingsGangMembersKilled_Click()
    mnuSettingsGangMembersKilled.Checked = Not mnuSettingsGangMembersKilled.Checked
End Sub

Private Sub mnuSettingsLawmenKilled_Click()
    mnuSettingsLawmenKilled.Checked = Not mnuSettingsLawmenKilled.Checked
End Sub

Private Sub timer_Timer()
    Call readMemory
End Sub
