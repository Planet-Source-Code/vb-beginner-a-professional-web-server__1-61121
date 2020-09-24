VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "80"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox chkListen 
      Caption         =   "Start listening for connections when this program is run"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CheckBox chkStartup 
      Caption         =   "Run web server on Windows startup"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.CheckBox chkDirListing 
      Caption         =   "Directory listing"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Settings"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtDefaultPath 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
   Begin VB.CheckBox chkAntiLeech 
      Caption         =   "Anti-leeching"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblPort 
      Caption         =   "Port to listen for connections on"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Default Path (Includes index.html or index.htm file):"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3555
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I got the folder browsing stuff from somewhere else so I have no idea how it works
Private Sub SetupDir()
On Error Resume Next

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BROWSEINFO

szTitle = "Please select the default path to save the recording in." & vbCrLf
szTitle = szTitle & "To create a new folder, click 'Make New Folder'"

With tBrowseInfo
.hwndOwner = Me.hwnd
.lpszTitle = szTitle & vbNullChar
.ulFlags = BROWSE_FLAGS.BIF_RETURNONLYFSDIRS + BROWSE_FLAGS.BIF_DONTGOBELOWDOMAIN + BROWSE_FLAGS.BIF_USENEWUI
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
sBuffer = Space(MAX_PATH)
SHGetPathFromIDList lpIDList, sBuffer
sBuffer = Left(sBuffer, InStr(1, sBuffer, vbNullChar) - 1)
txtDefaultPath.Text = sBuffer
End If
End Sub

'Browse for folder
Private Sub cmdBrowse_Click()
Dim blnRepeat As Boolean
On Error Resume Next

Call_Setup:

blnRepeat = False
SetupDir
If blnRepeat Then
GoTo Call_Setup
End If
End Sub

'Validate and save the settings
Private Sub cmdSave_Click()
If Right$(txtDefaultPath.Text, 1) = "\" Then
txtDefaultPath.Text = Mid(txtDefaultPath.Text, 1, Len(txtDefaultPath.Text) - 1)
End If
SaveSetting App.EXEName, App.EXEName, "Default Path", txtDefaultPath.Text
SaveSetting App.EXEName, App.EXEName, "Anti-Leech", chkAntiLeech.Value
SaveSetting App.EXEName, App.EXEName, "Directory Listing", chkDirListing.Value
If chkStartup.Value = 1 Then
AddToStartup
SaveSetting App.EXEName, App.EXEName, "Windows Startup", chkStartup.Value
Else
SaveSetting App.EXEName, App.EXEName, "Windows Startup", chkStartup.Value
RemovefromStartup
End If
SaveSetting App.EXEName, App.EXEName, "Listen", chkListen.Value
If txtPort > 65535 Then
MsgBox "Please choose a number between 0 and 65535", vbCritical, ""
txtPort.SetFocus
txtPort.SelStart = 0
txtPort.SelLength = Len(txtPort.Text)
Exit Sub
Else
SaveSetting App.EXEName, App.EXEName, "Port", txtPort.Text
frmMain.sckServer(0).Close
frmMain.sckServer(0).LocalPort = txtPort.Text
frmMain.sckServer(0).Listen
End If
Unload Me
End Sub

'Retrieve saved settings and load them to the text boxes
Private Sub Form_Load()
On Error GoTo ERROR_HANDLER
txtDefaultPath.Text = GetSetting("Webserver", "Webserver", "Default Path")
chkAntiLeech.Value = GetSetting("Webserver", "Webserver", "Anti-Leech")
chkDirListing.Value = GetSetting("Webserver", "Webserver", "Directory Listing")
chkStartup.Value = GetSetting("Webserver", "Webserver", "Windows Startup")
chkListen.Value = GetSetting("Webserver", "Webserver", "Listen")
txtPort.Text = GetSetting("Webserver", "Webserver", "Port")
ERROR_HANDLER:
End Sub

'Re-enable frmMain
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GetSetting("Webserver", "Webserver", "Default Path") = "" Then
Cancel = 1
MsgBox "Please choose your settings and then click the 'Save Settings' button.", vbExclamation, ""
Exit Sub
End If
frmMain.Enabled = True
End Sub

'Accept only numbers
Private Sub txtPort_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Exit Sub
Case Else
KeyAscii = 0
End Select
End Sub
