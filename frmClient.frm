VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00EEFFE8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InstaInform"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEFFE8&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   1530
      TabIndex        =   9
      Top             =   630
      Width           =   5295
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2310
         TabIndex        =   0
         Top             =   330
         Width           =   2145
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2310
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Chat Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   330
         Width           =   1785
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1815
      Left            =   390
      TabIndex        =   8
      Top             =   1590
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmClient.frx":0000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEFFE8&
      Height          =   4125
      Left            =   180
      TabIndex        =   6
      Top             =   990
      Width           =   7215
      Begin VB.CommandButton Command4 
         Caption         =   "&Clear Messages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2010
         TabIndex        =   4
         Top             =   3600
         Width           =   1545
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&LogOut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3750
         TabIndex        =   5
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   270
         TabIndex        =   2
         Top             =   2880
         Width           =   2565
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Send Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   570
         TabIndex        =   3
         Top             =   3600
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   5100
         TabIndex        =   7
         Top             =   720
         Width           =   1965
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2070
         Top             =   2430
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label8 
         BackColor       =   &H00EEFFE8&
         Caption         =   "Message to Send"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   270
         TabIndex        =   13
         Top             =   2580
         Width           =   1785
      End
      Begin VB.Label Label6 
         BackColor       =   &H00EEFFE8&
         Caption         =   "Messages Received"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   210
         Width           =   1965
      End
      Begin VB.Label Label5 
         BackColor       =   &H00EEFFE8&
         Caption         =   "Online Users"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5100
         TabIndex        =   11
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "My Chat Name ->"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   900
      TabIndex        =   16
      Top             =   210
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3540
      TabIndex        =   15
      Top             =   210
      Width           =   2595
   End
   Begin VB.Label Label2 
      Height          =   135
      Left            =   6150
      TabIndex        =   14
      Top             =   7260
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Logging In Sending the code Logged In0123 to the Server for processing
If Len(Text1.Text) > 0 Then
If Winsock1.State = 7 Then
    Winsock1.SendData "Logged In0123" & Text1.Text
    DoEvents
Else
    MsgBox "The Server is Temporarily Off", vbExclamation, "Server Unavaliable"
    Exit Sub
End If
Else
    MsgBox "Please Enter a Chat Name to get started", vbExclamation, "Login Error"
    Text1.SetFocus
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
'Sending to the Server the public chat Data
If Winsock1.State = 7 Then
    If Len(Text3.Text) > 0 Then
        Winsock1.SendData Label7.Caption & ">" & " " & Text3.Text
        Text3.Text = ""
        Text3.SetFocus
    End If
Else
    MsgBox "The Server is Temporarily Off", vbExclamation, "Server Unavaliable"
    Exit Sub
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
RichTextBox1.Text = ""
End Sub

Private Sub Form_Load()
'Taking the Required Data for Logging on to the server
'This program is developed by N.Indrasena Reddy, this is a unique public chat program which is understandable from novice programmers to experienced professionals. If youlike the program please email me to n_indureddy@yahoo.com. Ur Suggestions and criticism is utmost valued. Please do poll for this program
Dim str As String
str = InputBox("Please Enter the Server IP Address in this format eg 200.150.100.34", "Enter Server IP Address")
Winsock1.RemotePort = 1001
Winsock1.RemoteHost = str
Winsock1.Connect
Frame1.Visible = False
Command1.Enabled = False
RichTextBox1.Visible = False
Label7.Visible = False
Label9.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'When the Client Logs Out Send the Code Logged OUt0123 dont close until the client name is removed from the Server List
If Winsock1.State = sckConnected Then
    Winsock1.SendData Text1.Text & "Logged OUt0123"
    DoEvents
End If
End Sub

Private Sub Text1_Change()
'For Not Allowing Spaces in the chat name
If InStr(Text1.Text, " ") Then
    MsgBox "Spaces are not Allowed", vbExclamation + vbExclamation, "Loging Error"
    Text1.Text = ""
    Command1.Enabled = False
    Exit Sub
Else
    Command1.Enabled = True
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'For Not Allowing Spaces inthe chat name
If KeyAscii = 13 And Len(Text1.Text) > 0 Then
    If InStr(Text1.Text, " ") > 0 Then
        MsgBox "Spaces are not Allowed", vbExclamation + vbCritical, "Logging Error"
        Command1.Enabled = False
        Exit Sub
    Else
        Command1.Value = True
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(Text3.Text) > 0 Then
    Command2.Value = True
    Text3.Text = ""
    Text3.SetFocus
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim str As String
Dim a$()
Dim i As Integer
Winsock1.GetData str, vbString
'Testing Whether there is already user with this name, If so deliver a message Box and End the Application
If InStr(str, "Existing//0123") And Label2.Caption <> "one" Then
    MsgBox "The Chat Name is already existing, Please Restart the Application by choosing a different Chat Name", vbExclamation, "Login Error"
    End
End If
'For Online User List, Adding the Online Users to the List Box
If InStr(str, "Logged In0123") Then
    List1.Clear
    Label2.Caption = "one" 'for keeping track whether is this user got logged in
    Frame1.Visible = True
    Frame2.Visible = False
    RichTextBox1.Visible = True
    Label7.Visible = True
    Label9.Visible = True
    Label7.Caption = Text1.Text
    a$() = Split(str, "Logged In0123")
    For i = 0 To UBound(a$()) - 1
        If Len(a$(i)) > 0 Then
            List1.AddItem a$(i)
        End If
    Next i
'For rest of the Public Chat Messages
Else
    RichTextBox1.Text = RichTextBox1.Text + str + vbNewLine
End If
End Sub
