VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   Caption         =   "Chat Server"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   ForeColor       =   &H00000000&
   Icon            =   "server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Font Settings"
      Height          =   975
      Left            =   2520
      TabIndex        =   22
      Top             =   4920
      Width           =   1935
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "server.frx":27A2
         Left            =   120
         List            =   "server.frx":27B2
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Advanced Options"
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   4920
      Width           =   2415
      Begin VB.CommandButton Command4 
         Caption         =   "Update"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "File"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chat"
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "Advanced"
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Text            =   "User Name"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox outgoing 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox incoming 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Text            =   "127.0.0.1"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Listen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Text            =   "1210"
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Connect"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   4080
         Width           =   855
      End
      Begin VB.OptionButton client 
         Caption         =   "Client"
         Height          =   195
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton server 
         Caption         =   "Server"
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Addicon 
         Caption         =   "Minimize"
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   4440
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   1680
         Top             =   1440
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   960
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Disconnected"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "MSG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   13
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "IP#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   3600
         Width           =   375
      End
      Begin MediaPlayerCtl.MediaPlayer media 
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   3375
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   0   'False
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   0   'False
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inputboxs As String, Msg As String
Dim inifile As String
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Check1_Click()
If Check1.Value = 1 Then
Height = 6300
Else
Height = 5220
End If
End Sub

Private Sub client_click()
If client = True Then
Command3.Enabled = False
Command2.Enabled = True
Winsock1.Close
Winsock1.LocalPort = Text2
Winsock1.Listen
Label2.Caption = "Listening"
End If
End Sub


Private Sub cmdBrowse_Click()
cdopen.ShowOpen
txtFileName.Text = cdopen.Filename
End Sub

Private Sub Command4_Click()
update.Visible = True
End Sub



Private Sub Command5_Click()
filesend.Visible = True
End Sub
Private Sub Command6_Click()
a = InputBox("Where would you like to save it (Default = C:\chat.txt)", "Save", "C:\chat.txt")
Open a For Output As #1
incoming = incoming.Text
Write #1, incoming
Close #1
m = MsgBox("Chat session saved to: " + a, vbOKOnly, "Saved")
End Sub

Private Sub nosuc()
msgboxs = MsgBox("File not saved try again", vbOKOnly, "Error")
End Sub


Private Sub Form_Resize()
If (Height > 5235) And (Check1.Value = False) Then Height = 5235
If (Check1.Value = True) Then Height = 9000
If Width > 4560 Then Width = 4560
End Sub

Private Sub incoming_dblclick()
incoming.Text = ""
End Sub
Private Sub outgoing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
inputboxs = Text3.Text
Msg = "<" + inputboxs + ">" + outgoing.Text
outgoing.Text = ""
On Error Resume Next
Winsock1.SendData Msg
Select Case Combo1
Case ""
incoming.ForeColor = &H80000017
Case "Black"
incoming.ForeColor = &H80000017
Case "Green"
incoming.ForeColor = &H4000&
Case "Red"
incoming.ForeColor = &HFF&
Case "Blue"
incoming.ForeColor = &HFF0000
End Select
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   incoming = incoming + Msg + vbCrLf
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   End If


End Sub
Private Sub Command2_Click()
Winsock1.Close
Winsock1.LocalPort = Text2.Text
Winsock1.Listen
Label2.Caption = "Listening"
End Sub
Private Sub Command3_Click()
Winsock1.Close
Winsock1.Connect Text1, Text2
End Sub
Private Sub Form_Load()
inifile = App.path + "\serverdata.txt"
On Error Resume Next
Open inifile For Input As #1
Input #1, names
Input #1, adres
Input #1, Port
Input #1, clientell
Input #1, calor
Input #1, advyesno
Close #1

If clientell = "Client" Then
    client = True
    server = False
    Command2.Enabled = True
    Winsock1.LocalPort = Text2
    Winsock1.Listen
    Label2.Caption = "Listening"
    Else: Winsock1.Connect Text1, Text2
End If
Combo1 = calor
Text1.Text = adres
Text3.Text = names
Text2.Text = Port
Winsock1.LocalPort = Port
addicon_Click
Form1.Visible = True
End Sub
Private Sub server_Click()
If server = True Then
Winsock1.Close
Winsock1.Connect Text1, Text2
Command2.Enabled = False
Command3.Enabled = True
Label2.Caption = "Disconnected"
Else: Command2.Enabled = True

End If
End Sub
Private Sub Trojan_Click()
Form2.Visible = True
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Label2.Caption = "Connected"
Winsock1.Close
Winsock1.Accept requestID
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
ForeColor = &HC000&
If Form1.Visible = False Then
Form1.Visible = True
Form1.WindowState = 0
Form1.SetFocus
End If
Winsock1.GetData Msg
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   incoming = incoming + Msg + vbCrLf
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   media.Filename = App.path + "\phonaut.wav"
    media.Play
End Sub
Private Sub winsock1_connect()
Label2 = "Connected"
End Sub
Private Sub Winsock1_Close()
Label2 = "Disconnected"
End Sub
'SYS TRAY ICON
Private Sub addicon_Click()
Form1.Visible = False
Dim NID As NOTIFYICONDATA
NID.hwnd = Me.hwnd
NID.cbSize = Len(NID)
NID.uID = vbNull
NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
NID.hIcon = Me.Icon
NID.uCallbackMessage = WM_MOUSEMOVE
NID.szTip = "Right-Click to display Popupmenu" & vbCrLf
Shell_NotifyIcon NIM_ADD, NID
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form1.Visible = True
Dim Msg As Long
Msg = x / Screen.TwipsPerPixelX
    Select Case Msg
        Case WM_LBUTTONDOWN
        Me.WindowState = 0
        AppActivate Me.Caption
        
        Case WM_RBUTTONUP
        Dim pAPI As POINTAPI
        Dim PMParams As TPMPARAMS
        
        AppActivate Me.Caption
        GetCursorPos pAPI
        tmpPop% = CreatePopupMenu
        InsertMenu tmpPop%, 0, MF_BYPOSITION, 69, "Pinger"
        InsertMenu tmpPop%, 2, MF_BYPOSITION, 71, "Restore"
        InsertMenu tmpPop%, 3, MF_SEPARATOR, 72, vbNullString
        InsertMenu tmpPop%, 4, MF_BYPOSITION, 73, "Exit"
        
        PMParams.cbSize = 20
        tmpReply% = TrackPopupMenuEx(tmpPop%, TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_RETURNCMD, pAPI.x, pAPI.y, Me.hwnd, PMParams)
        Select Case tmpReply%
            Case 69
                 Form3.Visible = True
            Case 71
                Me.WindowState = 0
                AppActivate Me.Caption
            Case 73
                Call exits
                End
        End Select
    End Select
End Sub
Private Sub form_unload(cancel As Integer)
Call exits
End Sub

Public Sub exits()
Open inifile For Output As #1
names = Text3.Text
adres = Text1.Text
Port = Text2.Text
Write #1, names
Write #1, adres
Write #1, Port
If client = True Then Write #1, "Client"
If server = True Then Write #1, "Server"
Write #1, Combo1
Close #1
Dim NID As NOTIFYICONDATA
NID.hwnd = Me.hwnd
NID.cbSize = Len(NID)
NID.uID = vbNull
NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
NID.hIcon = Me.Icon
NID.uCallbackMessage = WM_MOUSEMOVE
NID.szTip = "Right-Click to display Popupmenu"
Shell_NotifyIcon NIM_DELETE, NID
Unload Form1
Unload Form3
Unload update
Unload filesend
End Sub




