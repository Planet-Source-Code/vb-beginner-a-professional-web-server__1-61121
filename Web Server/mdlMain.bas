Attribute VB_Name = "mdlMain"
Option Explicit

Public Declare Function CreateFolder Lib "kernel32.dll" Alias "CreateFolderA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateFolderEx Lib "kernel32.dll" Alias "CreateFolderExA" (ByVal lpTemplateFolder As String, ByVal lpNewFolder As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDListA Lib "shell32.dll" (pidl As Any, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As ITEMIDLIST) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Boolean
End Type

Public Type BROWSEINFO
hwndOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type

Public Enum BROWSE_FLAGS
BIF_BROWSEFORCOMPUTER = &H1000
BIF_BROWSEFORPRINTER = &H2000
BIF_BROWSEINCLUDEFILES = &H4000
BIF_DONTGOBELOWDOMAIN = &H2
BIF_EDITBOX = &H10
BIF_RETURNFSANCESTORS = &H8
BIF_RETURNONLYFSDIRS = &H1
BIF_STATUSTEXT = &H4
BIF_USENEWUI = &H40
BIF_VALIDATE = &H20
End Enum

Public Const MAX_PATH = 260
Public Const MAX_NAME = 40

Public Type SHITEMID
cb As Long
abID As Byte
End Type

Public Type ITEMIDLIST
mkid As SHITEMID
End Type

Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
 (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Global TaskIcon As NOTIFYICONDATA
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Enum T_KeyClasses
HKEY_LOCAL_MACHINE = &H80000002
End Enum

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const REG_DWORD = 4
Private Const REG_SZ = 1

'Add to Windows Startup
Public Function AddToStartup()
SetRegValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Web Server", App.Path & "\" & App.EXEName & ".exe"
End Function

'Add to Windows Startup
Public Function RemovefromStartup()
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Web Server"
End Function

'For deleting this program from Windows Startup
Public Sub DeleteValue(rClass As T_KeyClasses, Path As String, sKey As String)
Dim hKey As Long
Dim res As Long
res = RegOpenKeyEx(rClass, Path, 0, KEY_ALL_ACCESS, hKey)
res = RegDeleteValue(hKey, sKey)
RegCloseKey hKey
End Sub

'For adding this program to Windows Startup
Public Function SetRegValue(KeyRoot As T_KeyClasses, Path As String, sKey As String, NewValue As String) As Boolean
Dim hKey As Long
Dim KeyValType As Long
Dim KeyValSize As Long
Dim KeyVal As String
Dim tmpVal As String
Dim res As Long
Dim i As Integer
Dim x As Long
On Error GoTo ERROR_HANDLER
res = RegOpenKeyEx(KeyRoot, Path, 0, KEY_ALL_ACCESS, hKey)
If res <> 0 Then GoTo ERROR_HANDLER
tmpVal = String(1024, 0)
KeyValSize = 1024
res = RegQueryValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
Select Case res
Case 2
KeyValType = REG_SZ
Case Is <> 0
GoTo ERROR_HANDLER
End Select
Select Case KeyValType
Case REG_SZ
tmpVal = NewValue
Case REG_DWORD
x = Val(NewValue)
tmpVal = ""
For i = 0 To 3
tmpVal = tmpVal & Chr(x Mod 256)
x = x \ 256
Next
End Select
KeyValSize = Len(tmpVal)
res = RegSetValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
If res <> 0 Then GoTo ERROR_HANDLER
SetRegValue = True
RegCloseKey hKey
Exit Function
ERROR_HANDLER:
SetRegValue = False
RegCloseKey hKey
End Function

'Add to system tray
Public Sub LoadIconToTaskBar()
TaskIcon.cbSize = Len(TaskIcon)
TaskIcon.hwnd = frmMain.picEnabled.hwnd
TaskIcon.uID = 1&
TaskIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
TaskIcon.uCallbackMessage = WM_MOUSEMOVE
TaskIcon.hIcon = frmMain.picEnabled.Picture
'Replace 'My ToolTip Text' with the ToolTip Text to show when icon is loaded.
TaskIcon.szTip = "Web Server (Enabled)" & Chr$(0)
Shell_NotifyIcon NIM_ADD, TaskIcon
End Sub

'Remove system tray icon
Public Sub UnloadIconToTaskBar()
Shell_NotifyIcon NIM_DELETE, TaskIcon
End
End Sub

'Enable system tray icon
Public Sub EnableTaskBarIcon()
TaskIcon.hIcon = frmMain.picEnabled.Picture
TaskIcon.szTip = "Web Server (Enabled)" & Chr$(0)
Shell_NotifyIcon NIM_MODIFY, TaskIcon
End Sub

'Disable system tray icon
Public Sub DisableTaskBarIcon()
TaskIcon.hIcon = frmMain.picDisabled.Picture
TaskIcon.szTip = "Web Server (Disabled)" & Chr$(0)
Shell_NotifyIcon NIM_MODIFY, TaskIcon
End Sub

'Checks if a folder exists
Public Function FolderExists(Filename As String) As Boolean
Dim FreeFileNum As Integer
On Error GoTo FolderDoesNotExist
FreeFileNum = FreeFile
Open Filename & "\test" For Output As #FreeFileNum
Close #FreeFileNum
Kill Filename & "\test"
FolderExists = True
Exit Function
FolderDoesNotExist: FolderExists = False
End Function

'Checks if a file exists
Public Function FileExists(Filename As String) As Boolean
Dim FreeFileNum As Integer
On Error GoTo FileDoesNotExist
FreeFileNum = FreeFile
Open Filename For Input As #FreeFileNum
Close #FreeFileNum
FileExists = True
Exit Function
FileDoesNotExist: FileExists = False
End Function
