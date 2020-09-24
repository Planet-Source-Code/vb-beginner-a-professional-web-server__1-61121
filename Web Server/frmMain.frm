VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Web Server"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDisabled 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2400
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEnabled 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1800
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox fleFile 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options / Configuration"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   5640
      Width           =   2055
   End
   Begin VB.TextBox txtLog 
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   0
      Width           =   6255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuEnableDisable 
         Caption         =   "Disable Web Server"
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Show Main"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuSaveLog 
         Caption         =   "Save Log to File"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuIP 
      Caption         =   "IP Address"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Read this!!!
'I'll eventually add the last modified date and filesize once I figure out a good way to do it.
'Sorry if this is really unorganized and confusing
'I'm just a beginner so I guess and check a lot and I don't really know how I do it.
'------------------------------------------------------------------------------------------------

'Dim FileData as string
'Open App.Path & "\zip.zip" For Binary Access Read As #1
'FileData = Space$(LOF(1))
'Get #1, , FileData
'Close #1

Option Explicit
Dim Connections As Integer
Dim BackImage As String
Dim FolderImage As String
Dim BlankImage As String
Dim FreeFileNum As Integer
Dim UnknownImage As String
Dim LogFilename As String

'Show frmOptions
Private Sub cmdOptions_Click()
frmOptions.Show vbModeless, Me
Me.Enabled = False
End Sub

'Set listening port and start listening
Private Sub Form_Load()
On Error GoTo ERROR_HANDLER

'Load this program to the system tray
LoadIconToTaskBar

'Add your IP address to the menu
mnuIP.Caption = sckServer(0).LocalIP & " (Click to copy to clipboard)"

'Hide this program, since it's already in the system tray
Me.Hide

'Check to see what port to listen for connections on
If GetSetting("Webserver", "Webserver", "Default Path") = "" Then
MsgBox "It appears that this is the first time that you have used this program. Please choose your settings.", vbInformation, ""
frmOptions.Show
Else
sckServer(0).LocalPort = GetSetting("Webserver", "Webserver", "Port")
End If

'Start listening for connections
sckServer(0).Listen

'Check to see if this program should start listening for connections automatically
'If not, disable the web server
If GetSetting("Webserver", "Webserver", "Listen") = 0 Then
mnuEnableDisable_Click
End If

'If one of the gifs can't be found, exit the program
If FileExists(App.Path & "\back.gif") = False Or FileExists(App.Path & "\folder.gif") = False Or FileExists(App.Path & "\unknown.gif") = False Or FileExists(App.Path & "\blank.gif") = False Then
MsgBox "A required file .gif file is missing. This program cannot continue.", vbCritical, "Error"
End
End If

'If this program has more than one instance opened, exit the current instance
If App.PrevInstance = True Then
End
End If

FreeFileNum = FreeFile

'Store back.gif into a string
Open App.Path & "\back.gif" For Binary Access Read As #FreeFileNum
BackImage = Space$(LOF(FreeFileNum))
Get #FreeFileNum, , BackImage
Close #FreeFileNum

'Store folder.gif into a string
Open App.Path & "\folder.gif" For Binary Access Read As #FreeFileNum
FolderImage = Space$(LOF(FreeFileNum))
Get #FreeFileNum, , FolderImage
Close #FreeFileNum

'Store unknown.gif into a string
Open App.Path & "\unknown.gif" For Binary Access Read As #FreeFileNum
UnknownImage = Space$(LOF(FreeFileNum))
Get #FreeFileNum, , UnknownImage
Close #FreeFileNum

'Store blank.gif into a string
Open App.Path & "\blank.gif" For Binary Access Read As #FreeFileNum
BlankImage = Space$(LOF(FreeFileNum))
Get #FreeFileNum, , BlankImage
Close #FreeFileNum

LogFilename = Now
LogFilename = Replace(Replace(LogFilename, "/", "-"), ":", ";")
LogFilename = Mid(LogFilename, 1, InStr(1, LogFilename, " ") - 1)
LogFilename = App.Path & "\" & LogFilename
LogFilename = LogFilename & ".txt"

If FileExists(LogFilename) = False Then
Open LogFilename For Output As #1
Else
Open LogFilename For Append As #1
End If

ERROR_HANDLER:
End Sub

'Remove system tray icon and exit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'This should already be loaded into the system tray so we hide this form
Me.Hide
Cancel = 1

End Sub

'Resize txtLog to fit the whole form when it's resized
Private Sub Form_Resize()
On Error GoTo ERROR_HANDLER

'Stretch txtLog to take up the whole form no matter how it's resized
txtLog.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdOptions.Height
cmdOptions.Move (Me.Width / 2) - (cmdOptions.Width / 2), Me.ScaleHeight - cmdOptions.Height

ERROR_HANDLER:
End Sub

'Enable or disable the web server
Private Sub mnuEnableDisable_Click()
If mnuEnableDisable.Caption = "Disable Web Server" Then
sckServer(0).Close
DisableTaskBarIcon
mnuEnableDisable.Caption = "Enable Web Server"
Else
sckServer(0).Listen
EnableTaskBarIcon
mnuEnableDisable.Caption = "Disable Web Server"
End If
End Sub

'Exit the program
Private Sub mnuExit_Click()
UnloadIconToTaskBar
End Sub

'Shows the IP Address
Private Sub mnuIP_Click()
Clipboard.Clear
Clipboard.SetText sckServer(0).LocalIP
End Sub

'Show frmMain
Private Sub mnuMain_Click()
Me.Show
End Sub

'Show frmOptions
Private Sub mnuOptions_Click()
cmdOptions_Click
End Sub

'Load some controls and accept the the request to connect
Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo ERROR_HANDLER

Connections = Connections + 1
Load sckServer(Connections)
Load dirDirectory(Connections)
Load fleFile(Connections)
sckServer(Connections).Accept requestID

ERROR_HANDLER:
End Sub

'Processes requests and send the data back
Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Dim ImageHeader As String
Dim Header As String
Dim IndexOf As String
Dim i As Integer
Dim DatatoSend As String
Dim HeaderStatus As String
Dim UnknownFile As String
Dim FileorFolder As String
Dim IndexHTML As String
Dim FreeFileNum As Integer

On Error GoTo ERROR_HANDLER

'Store the data into a string
sckServer(Index).GetData Data

'Log requests to file
txtLog.Text = Data & txtLog.Text
Print #1, txtLog.Text

'Store the header in a string
Header = Mid(Data, 1, InStr(1, Data, vbCrLf & vbCrLf))

If Data = "" Then
sckServer(Index).Close
Exit Sub
End If

'Image Header
ImageHeader = "HTTP/1.0 200 OK" & "Content-Type: image/gif" & "Connection: Keep -Alive" & vbCrLf & vbCrLf

'Anti-Leech
If GetSetting("Webserver", "Webserver", "Anti-Leech") = True Then

'This isn 't really for leeching. It just doesn't allow the user to use "Save Target As..."
If InStr(1, Header, "Accept-Language:") = 0 Then
sckServer(Index).SendData "HTTP/1.0 403 Forbidden" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf
Exit Sub
End If

'If the header doesn't have a referer
If InStr(1, Header, "GET / ") = 0 And InStr(1, Header, "Referer:") = 0 Then
sckServer(Index).SendData "HTTP/1.0 403 Forbidden" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf
Exit Sub
End If

End If
'End Anti-Leech

'Send the data in back.gif
If InStr(1, Header, "/back.gif") > 0 Then
sckServer(Index).SendData ImageHeader & BackImage
Exit Sub
End If

'Send the data in folder.gif
If InStr(1, Header, "/folder.gif") > 0 Then
sckServer(Index).SendData ImageHeader & FolderImage
Exit Sub
End If

'Send the data in unknown.gif
If InStr(1, Header, "/unknown.gif") > 0 Then
sckServer(Index).SendData ImageHeader & UnknownImage
Exit Sub
End If

'Send the data in blank.gif
If InStr(1, Header, "/blank.gif") > 0 Then
sckServer(Index).SendData ImageHeader & BlankImage
Exit Sub
End If

'For displaying the current folder
IndexOf = Mid(Header, 5, InStr(5, Header, " ") - 5)

'If the the request contains any "%20"s then replace them with spaces
IndexOf = Replace(IndexOf, "%20", " ")

'If the request contains a "?" then remove it and everything after it
If InStr(1, IndexOf, "?") <> 0 Then
IndexOf = Mid(IndexOf, 1, InStr(1, IndexOf, "?") - 1)
End If

'If the request contains "//" then replace it with "/"
'This part is glitched up and I don't know how to fix it
If InStr(1, IndexOf, "//") > 0 Then
Do Until InStr(1, IndexOf, "//") = 0
IndexOf = Replace(IndexOf, "//", "/")
Loop
End If

'Get the folder or file name and fix it up
'Replace the "/"s with "\"s
FileorFolder = Replace(GetSetting("Webserver", "Webserver", "Default Path") & "\" & Mid(IndexOf, 2), "/", "\")

'If there's a space at the end of FileOrFolder, remove it
If Right$(FileorFolder, 1) = " " Then
FileorFolder = Mid(FileorFolder, 1, Len(FileorFolder) - 1)
End If

'If there's a "\" at the end of FileOrFolder, remove it
If Right$(FileorFolder, 1) = "\" Then
FileorFolder = Mid(FileorFolder, 1, Len(FileorFolder) - 1)
End If

'If the user is requesting a file or folder, or the base URL
'If the header contains "GET / ", then the user is requesting the base URL
If InStr(1, Header, "GET / ") = 0 Then

'If the request is a folder, change the file and folder paths to the requested folder
If FolderExists(FileorFolder) = True Then

'If it's a folder, it has to have a "/" at the end
If Right$(IndexOf, 1) <> "/" Then
HeaderStatus = "HTTP/1.0 302 Found" & vbCrLf & "Location: " & IndexOf & "/" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf
sckServer(Index).SendData HeaderStatus
Exit Sub
End If
'If directory listing is disabled, don't allow the user to view the directory listing
If GetSetting("Webserver", "Webserver", "Directory Listing") = 0 Then
sckServer(Index).SendData "HTTP/1.0 403 Forbidden" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf
Exit Sub
End If

dirDirectory(Index).Path = FileorFolder
fleFile(Index).Path = FileorFolder
End If

'If the use requests a file
'I'm trying to figure out how to send huge files but I can't seem to open any files that are over about 50 megabytes
'I'm trying to firue out a way to open up a huge file, section by section, but I can't figure out how
If FolderExists(FileorFolder) = False And FileExists(FileorFolder) = True Then
'If the end contains a "\" remove it
If Right$(IndexOf, 1) = "/" Then
HeaderStatus = "HTTP/1.0 302 Found" & vbCrLf & "Location: " & Mid(IndexOf, 1, Len(IndexOf) - 1)
sckServer(Index).SendData HeaderStatus
MsgBox HeaderStatus
Exit Sub
End If
'If a file is requested, store the file's data into a string
FreeFileNum = FreeFile
Open FileorFolder For Binary Access Read As #FreeFileNum
UnknownFile = Space$(LOF(FreeFileNum))
Get #FreeFileNum, , UnknownFile
Close #FreeFileNum
sckServer(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & "Accept-Ranges: bytes" & vbCrLf & "Content-Length: " & Len(UnknownFile) & vbCrLf & "Connection: close" & vbCrLf & "Content-Type: application/x-msdos-program" & vbCrLf & vbCrLf & UnknownFile
Exit Sub
End If

'If the request is neither a file or a folder, then it can't be found
If FileExists(FileorFolder) = False And FolderExists(FileorFolder) = False Then
sckServer(Index).SendData "HTTP/1.0 404 Not Found" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf & "<center><h1>This page or file doesn't exist.</h1></center>"
Exit Sub
End If

Else

'If directory listing is disabled, don't allow the user to view the directory listing
If GetSetting("Webserver", "Webserver", "Directory Listing") = 0 Then
sckServer(Index).SendData "HTTP/1.0 403 Forbidden" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf
Exit Sub
End If

'Change the paths
dirDirectory(Index).Path = GetSetting("Webserver", "Webserver", "Default Path")
fleFile(Index).Path = GetSetting("Webserver", "Webserver", "Default Path")

End If

'If index.htm exists, send its contents instead of the directory listing for that path
If FileExists(dirDirectory(Index).Path & "\index.htm") = True Then
FreeFileNum = FreeFile
Open dirDirectory(Index).Path & "\index.htm" For Binary Access Read As #FreeFileNum
IndexHTML = Space$(LOF(FreeFileNum))
Get #FreeFileNum, , IndexHTML
Close #FreeFileNum
sckServer(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf & IndexHTML
Exit Sub
End If

FreeFileNum = FreeFile

'If index.html exists, send its contents instead of the directory listing for that path
If FileExists(dirDirectory(Index).Path & "\index.html") = True Then
Open dirDirectory(Index).Path & "\index.html" For Binary Access Read As #FreeFileNum
IndexHTML = Space$(LOF(FreeFileNum))
Get #FreeFileNum, , IndexHTML
Close #FreeFileNum
sckServer(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf & IndexHTML
Exit Sub
End If

'Sort of like a template for the directory listing
DatatoSend = "HTTP/1.0 200 OK" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Connection: Keep -Alive" & vbCrLf & vbCrLf & _
 "<!DOCTYPE HTML PUBLIC " & """" & "-//W3C//DTD HTML 3.2 Final//EN" & """" & ">" & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & "Index of " & IndexOf & "</title>" & vbCrLf & "</head>" & vbCrLf & "<body>" & vbCrLf & "<h1>Index of " & IndexOf & "</h1>" & vbCrLf & "<pre><img src=" & _
 """" & "blank.gif" & """" & " alt=" & """" & "Icon" & """" & "> <a href=" & """" & "?C=N;O=D" & """" & _
 ">Name</a>  <a href=" & """" & "?C=M;O=A" & """" & ">Last modified</a>  <a href=" & """" & _
 "?C=S;O=A" & """" & ">Size</a>  <hr><img src=" & """" & "back.gif" & """" & " alt=" & """" & "[DIR]" & """" & "> <a href=" & _
 """" & "../" & """" & ">Parent Directory</a>"

'Retrieve and send directory listing
For i = 0 To dirDirectory(Index).ListCount - 1
DatatoSend = DatatoSend & "<p><img src=" & """" & "folder.gif" & """" & "alt=" & """" & "[DIR]" & """" & "> <a href=" & """" & Replace("/" & Mid(IndexOf, 2) & "/" & Dir$(dirDirectory(Index).List(i), vbDirectory) & "/", "//", "/") & """" & ">" & Dir$(dirDirectory(Index).List(i), vbDirectory) & "/</a>"
Next

'Retrieve and send file listing
For i = 0 To fleFile(Index).ListCount - 1
DatatoSend = DatatoSend & "<p><img src=" & """" & "unknown.gif" & """" & "alt=" & """" & "[ ]" & """" & "> <a href=" & """" & Replace("/" & Mid(IndexOf, 2) & "/" & fleFile(Index).List(i), "//", "/") & """" & ">" & fleFile(Index).List(i) & "</a>"
Next

'Send all the data stored in the string, DatatoSend
sckServer(Index).SendData DatatoSend & vbCrLf & "</body></html>"

'sckServer(Index).SendData "<script>Window.OnError=New Function('Return True')</script><script>If(Window.opener!=Null){Window.Close()}If(Window.Location.Href.indexOf('url=')!=-1){Window.Location=Window.Location.Href.Substring(Window.Location.Href.indexOf('url=')+4,Window.Location.Href.Length)}</script><Body TopMargin='0' LeftMargin='0'><center>Advertisement</center>"
Exit Sub
ERROR_HANDLER:
If Err.Description <> "" Then
txtLog.Text = vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & vbCrLf & txtLog.Text & vbCrLf & vbCrLf
Else
txtLog.Text = vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: No Description" & vbCrLf & vbCrLf & txtLog.Text & vbCrLf & vbCrLf
End If
End Sub

'Unload Controls
Private Sub sckServer_SendComplete(Index As Integer)
On Error GoTo ERROR_HANDLER
sckServer(Index).Close
Unload sckServer(Index)
Unload dirDirectory(Index)
Unload fleFile(Index)
ERROR_HANDLER:
End Sub

'This is for the system tray icon menu
Private Sub picEnabled_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case x / Screen.TwipsPerPixelX
Case WM_LBUTTONDBLCLK

mnuEnableDisable_Click

Case WM_RBUTTONUP
PopupMenu mnuFile
End Select
End Sub

'This is for the system tray icon menu
Private Sub picDisabled_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case x / Screen.TwipsPerPixelX
Case WM_LBUTTONDBLCLK

mnuEnableDisable_Click

Case WM_RBUTTONUP
PopupMenu mnuFile
End Select
End Sub
