VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "frmServer.frx":0000
      Left            =   690
      List            =   "frmServer.frx":0002
      TabIndex        =   0
      Top             =   900
      Width           =   4905
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1200
      Top             =   3210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "List Of Clients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   690
      TabIndex        =   1
      Top             =   450
      Width           =   1725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intMax As Long 'The Global Variable for trapping the number of connections
Private Sub Form_Load()
   intMax = 0 'Setting the connection
   Winsock1(0).LocalPort = 1001
   Winsock1(0).Listen
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'Accepting the Connection Requests from the Clients
    If Index = 0 Then
      intMax = intMax + 1
      Load Winsock1(intMax)
      Winsock1(intMax).LocalPort = 1001
      Winsock1(intMax).Accept requestID
    End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i, j, num As Integer
Dim str As String
Dim a$()
Dim u As Integer
Winsock1(Index).GetData str, vbString, bytesTotal
u = 0
'When client Log the code Logged In0123 is processed
If InStr(str, "Logged In0123") > 0 Then
    a$() = Split(str, "Logged In0123")
    List1.AddItem a$(1)
    For num = 0 To List1.ListCount - 1
        If List1.List(num) = a$(1) Then
            u = u + 1
        End If
        If u > 1 Then
            List1.RemoveItem (num)
            For i = 0 To intMax
            If Winsock1(i).State = 7 Then
                Winsock1(i).SendData "Existing//0123" & str
            DoEvents
            End If
           Next i
        End If
    Next num
    For i = 0 To intMax
    For j = 0 To List1.ListCount - 1
        If Winsock1(i).State = 7 Then
            Winsock1(i).SendData List1.List(j) & "Logged In0123"
        End If
    Next j
    Next i
End If
'When client Logout the code Logged OUt0123 is processed
If InStr(str, "Logged OUt0123") Then
    a$() = Split(str, "Logged OUt0123")
        For num = 0 To List1.ListCount - 1
            If List1.List(num) = a$(0) Then
                List1.RemoveItem (num)
            End If
        Next num
        For i = 0 To intMax
            For j = 0 To List1.ListCount - 1
                If Winsock1(i).State = 7 Then
                Winsock1(i).SendData List1.List(j) & "Logged In0123"
            End If
        Next j
        Next i
End If
'Sending Data other than LogIn and LogOut Process Data
For i = 0 To intMax
    If Winsock1(i).State = 7 Then
        Winsock1(i).SendData str
        DoEvents
    End If
Next i
        
End Sub


