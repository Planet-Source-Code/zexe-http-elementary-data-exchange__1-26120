VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTTP elementary data exchange"
   ClientHeight    =   4845
   ClientLeft      =   3015
   ClientTop       =   2745
   ClientWidth     =   6525
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6525
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   4920
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer tmrConnectionStatus 
      Interval        =   500
      Left            =   5520
      Top             =   4680
   End
   Begin VB.CommandButton cmdClosePort 
      Caption         =   "C&lose port"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Listen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame fraConnectionStatus 
      Caption         =   "Connection status"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      Begin VB.Label Label6 
         Caption         =   "Bytes remaining"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblBytesRemaining 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Bytes sent"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblBytesSent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblConnectionStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
   End
   Begin MSWinsockLib.Winsock comConexiune 
      Left            =   6000
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraExchangeBuffer 
      Caption         =   "Exchange Buffer"
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   6255
      Begin RichTextLib.RichTextBox rtfBuffer 
         Height          =   1575
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":000C
      End
      Begin VB.TextBox txtTransmissionBuffer 
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2325
         Width           =   4575
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send message"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4905
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2325
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Conversation history"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Message to send"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2130
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "S&ettings"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function RetConnectionStatus() As String
'Purpose: Rturn a string description about the curent connnection status.

Dim lngState As Long
Dim strConnectionDescription As String

lngState = comConexiune.State

Select Case lngState
Case sckClosed
    strConnectionDescription = "Not connected"
Case sckClosing
    strConnectionDescription = "Connection is closing..."
Case sckConnected
    strConnectionDescription = "Conected"
Case sckConnecting
    strConnectionDescription = "Connecting..."
Case sckConnectionPending
    strConnectionDescription = "Connection pending"
Case sckError
    strConnectionDescription = "Unknown error"
Case sckHostResolved
    strConnectionDescription = "Host resolved"
Case sckListening
    strConnectionDescription = "Listening..."
Case sckOpen
    strConnectionDescription = "Open"
Case sckResolvingHost
    strConnectionDescription = "Resolving host..."
End Select

RetConnectionStatus = strConnectionDescription

End Function

Private Sub cmdClosePort_Click()

comConexiune.Close

'Activate controls.
cmdSettings.Enabled = True
cmdListen.Enabled = True
cmdConnect.Enabled = True
cmdClosePort.Enabled = False

End Sub

Private Sub cmdConnect_Click()
'Trying to connect.

'First, deactivate some controls.
cmdConnect.Enabled = False
cmdListen.Enabled = False
cmdSettings.Enabled = False
cmdClosePort.Enabled = True

On Error GoTo Tratament

comConexiune.Connect

Exit Sub
Tratament:

MsgBox Err.Description, vbCritical, "Connection status"

'Activate contols.
cmdListen.Enabled = True
cmdSettings.Enabled = True
cmdConnect.Enabled = True
cmdClosePort.Enabled = False

'Closing connection.
comConexiune.Close

End Sub

Private Sub cmdListen_Click()
'Listening selected port.

On Error GoTo Tratament

'Deactivate some controls.
cmdListen.Enabled = False
cmdConnect.Enabled = False
cmdSettings.Enabled = False
cmdClosePort.Enabled = True

comConexiune.Listen

Exit Sub
Tratament:
MsgBox Err.Description, vbCritical, "Connection status"

'Operation failed.
'Activate controls.
cmdConnect.Enabled = True
cmdSettings.Enabled = True
cmdListen.Enabled = True
cmdClosePort.Enabled = False

End Sub


Private Sub cmdSend_Click()

On Error GoTo Tratament

Dim strData As String

strData = txtTransmissionBuffer

rtfBuffer.Text = rtfBuffer.Text & comConexiune.LocalHostName & ": " & strData & vbLf
rtfBuffer.SelStart = Len(rtfBuffer.Text) - 1

txtTransmissionBuffer = ""

comConexiune.SendData strData

Exit Sub
Tratament:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdSettings_Click()
frmSettings.Show vbModal
End Sub


Private Sub comConexiune_Close()
    MsgBox "Connection with server lost", vbInformation, "Connection status"
End Sub

Private Sub comConexiune_Connect()
    MsgBox "Connection with server estabilished.", vbInformation, "Connection succeed"
End Sub

Private Sub comConexiune_ConnectionRequest(ByVal requestID As Long)

Dim intTest As Integer

On Error GoTo Tratament

intTest = MsgBox("A connection is requested from you." & vbLf & "Do you accept?", vbQuestion + vbYesNo, "Connection request")

If intTest = vbNo Then Exit Sub

comConexiune.Close
comConexiune.Accept requestID


Exit Sub
Tratament:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub comConexiune_DataArrival(ByVal bytesTotal As Long)

Dim strData As String

comConexiune.GetData strData, vbString
rtfBuffer.Text = rtfBuffer.Text & comConexiune.RemoteHost & ": " & strData & vbLf
rtfBuffer.SelStart = Len(rtfBuffer.Text) - 1

End Sub

Private Sub comConexiune_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
                ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

MsgBox Description, vbCritical, "Error"

End Sub

Private Sub comConexiune_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

lblBytesSent = bytesSent
lblBytesRemaining = bytesRemaining

End Sub


Private Sub Form_Load()
lblConnectionStatus = RetConnectionStatus

End Sub

Private Sub tmrConnectionStatus_Timer()
lblConnectionStatus = RetConnectionStatus
End Sub
Private Sub txtTransmissionBuffer_Change()

If comConexiune.State = sckConnected And Trim(txtTransmissionBuffer) <> "" Then
    cmdSend.Enabled = True
Else
    cmdSend.Enabled = False
End If

End Sub

