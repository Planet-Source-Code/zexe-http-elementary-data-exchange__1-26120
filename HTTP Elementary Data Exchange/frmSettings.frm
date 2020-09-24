VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4335
   ClientLeft      =   2145
   ClientTop       =   2370
   ClientWidth     =   7635
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Default         =   -1  'True
      Height          =   315
      Left            =   6360
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cboInternetProtocol 
      Height          =   315
      ItemData        =   "frmSettings.frx":000C
      Left            =   1920
      List            =   "frmSettings.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   6360
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   315
      Left            =   6360
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame fraRemoteHostSettings 
      Caption         =   "Remote host"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   5175
      Begin VB.TextBox txtRemoteHostPort 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###.###.###.###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtRemoteHostName 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###.###.###.###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtRemoteHostIP 
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###.###.###.###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Communication port"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "IP address"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraLocalHolstSettings 
      Caption         =   "Local host"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtLocalHostPort 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###.###.###.###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtLocalHostName 
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###.###.###.###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtLocalHostIP 
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###.###.###.###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Communication port"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "IP address"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Comunication protocol"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "The 'IP address' text box contain same data as 'Name' text box."
      Height          =   615
      Left            =   5400
      TabIndex        =   18
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "The 'Remote host name' text box can store the remote computer name, IP or Universal Resource Locator (URL)"
      Height          =   975
      Left            =   5400
      TabIndex        =   17
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDefault_Click()
txtLocalHostPort = 80
txtRemoteHostPort = 80
cboInternetProtocol.ListIndex = sckTCPProtocol
txtRemoteHostName = "127.0.0.1"
End Sub

Private Sub cmdOK_Click()
'Save settings.

On Error GoTo Tratament

With frmMain.comConexiune
    'Local host.
    .LocalPort = Val(txtLocalHostPort)
    
    'Remote host.
    .RemoteHost = txtRemoteHostName
    .RemotePort = Val(txtRemoteHostPort)
    
    'Protocol.
    Select Case cboInternetProtocol.ListIndex
    Case sckTCPProtocol
        .Protocol = sckTCPProtocol
    Case sckUDPProtocol
        .Protocol = sckUDPProtocol
    End Select
End With

With frmMain
    .cmdConnect.Enabled = True
    .cmdListen.Enabled = True
End With

Unload Me
Exit Sub
Tratament:
MsgBox Err.Description, vbCritical, "Error"
End Sub



Private Sub Form_Load()
'Read current settings for winsock.

'Local.
Dim strLocalHostName As String
Dim strLocalHostIP As String
Dim lngLocalHostPort As Long
Dim lngProtocol As Long

'Remote.
Dim strRemoteHostName As String
Dim strRemoteHostIP As String
Dim lngRemoteHostPort As Long

With frmMain.comConexiune
    'Local.
    strLocalHostName = .LocalHostName
    strLocalHostIP = .LocalIP
    lngLocalHostPort = .LocalPort
    
    'Remote.
    strRemoteHostName = .RemoteHost
    strRemoteHostIP = .RemoteHostIP
    lngRemoteHostPort = .RemotePort
    
    'The protocol.
    lngProtocol = .Protocol
End With

'Analize local host settings.
If strLocalHostName = "" Then
    txtLocalHostName = "Unknown local host name"
Else
    txtLocalHostName = strLocalHostName
End If

If strLocalHostIP = "" Then
    'No IP specified.
    txtLocalHostIP = "No IP specified"
Else
    txtLocalHostIP = strLocalHostIP
End If

If lngLocalHostPort = 0 Then
    txtLocalHostPort = "No communication port specified"
Else
    txtLocalHostPort = lngLocalHostPort
End If

'Analize remote host settings.
If strRemoteHostName = "" Then
    txtRemoteHostName = "Unknown host name"
Else
    txtRemoteHostName = strRemoteHostName
End If

If strRemoteHostIP = "" Then
    'No IP address specified.
    txtRemoteHostIP = "No IP address specified"
Else
    txtRemoteHostIP = strRemoteHostIP
End If

If lngRemoteHostPort = 0 Then
    txtRemoteHostPort = "No communication port specified"
Else
    txtRemoteHostPort = lngRemoteHostPort
End If

'Analize protocol type.
Select Case lngProtocol
Case sckTCPProtocol
    'Transfer Control Protocol (TCP).
    cboInternetProtocol.ListIndex = sckTCPProtocol
Case sckUDPProtocol
    'User Datagram Protocol(UDP).
    cboInternetProtocol.ListIndex = sckUDPProtocol
End Select


End Sub


