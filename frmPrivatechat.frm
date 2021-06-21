VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPrivatechat 
   BackColor       =   &H00EDEDD8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Chat"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4140
      Top             =   6930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2940
      TabIndex        =   6
      Top             =   5700
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send "
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   5700
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   2490
      TabIndex        =   4
      Top             =   5040
      Width           =   2445
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   630
      TabIndex        =   2
      Top             =   1680
      Width           =   4365
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3540
      TabIndex        =   10
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   225
      Left            =   2910
      TabIndex        =   9
      Top             =   660
      Width           =   435
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1170
      TabIndex        =   8
      Top             =   660
      Width           =   1305
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   225
      Left            =   510
      TabIndex        =   7
      Top             =   660
      Width           =   435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Private Message"
      Height          =   225
      Left            =   870
      TabIndex        =   3
      Top             =   5100
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Messages Received"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Private Chat"
      BeginProperty Font 
         Name            =   "AGaramond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   5385
   End
End
Attribute VB_Name = "frmPrivatechat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Winsock1.State = 7 Then
    Winsock1.SendData "Prg0123//" & " " & Text1.Text & " " & "From" & " " & Label5.Caption & " " & "To" & " " & Label7.Caption
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Winsock1.RemotePort = 1001
Winsock1.RemoteHost = "200.150.100.19"
Winsock1.Connect
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Winsock1.State = sckConnected Then
    Winsock1.Close
    DoEvents
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim str As String
Dim a$()
Dim i, j As Integer
Winsock1.GetData str, vbString
If InStr(str, "Prg0123//") Then
If InStr(str, frmClient.Label7.Caption) Then
    a$() = Split(str, "Prg0123//")
    List1.AddItem a$(1)
End If
End If
End Sub

