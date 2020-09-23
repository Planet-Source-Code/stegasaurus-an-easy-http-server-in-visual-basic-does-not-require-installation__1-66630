VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Michael Billington's HTTP daemon."
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Clients since last restart:"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Time since last restart:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "(reset at 300)"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   3840
      Width           =   495
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//Code by michael Billington. This work is in the public domain, do with it as you wish
Option Explicit
Const Serverport = 80 '//Standard HTTP port, change to whatever you like

Private Sub Form_Load()
Outp "Computer connected to network on port " & Serverport, 9
Outp "Your network IP address is " & socket(0).LocalIP, 9
socket(0).Bind Serverport
socket(0).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Outp socket(Index).RemoteHost & socket(Index).RemoteHostIP & " now connected", 9
socket(Index).Close
socket(Index).Accept requestID
Load socket(Index + 1)
socket(Index + 1).Bind Serverport
socket(Index + 1).Listen
End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim str As String
socket(Index).GetData str
Outp socket(Index).RemoteHostIP & ":" & getL(str, vbCrLf), 8
Dim temp As String
temp = getL(getR(str, " "), " ")
temp = Replace(temp, "/", "\")
If temp = "\" Then temp = "\index.htm"
temp = Replace(temp, "%20", " ")
Dim FileContent As String
If FileExists(App.Path & "\files" & temp) Then
  FileContent = LoadText(App.Path & "\files" & temp)
  Select Case getRrev(temp, ".")
    Case "txt": socket(Index).SendData ReturnFile(FileContent, "text/plain")
    Case "htm": socket(Index).SendData ReturnFile(FileContent, "text/html")
    Case "html": socket(Index).SendData ReturnFile(FileContent, "text/html")
    Case "png": socket(Index).SendData ReturnFile(FileContent, "image/png")
    Case "zip": socket(Index).SendData ReturnFile(FileContent, "application/zip")
    Case Else: socket(Index).SendData ReturnFile(FileContent, "application/zip")
  End Select
Else
  socket(Index).SendData Return404(LoadText(App.Path & "\files\404.htm"))
End If
End Sub

Function ReturnFile(doc As String, filetype As String) As String
ReturnFile = "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: " & filetype & vbCrLf & "Content-Length: " & Len(doc) & vbCrLf & vbCrLf & doc & vbCrLf & vbCrLf & vbCrLf
End Function
Function Return404(doc As String) As String
Return404 = "HTTP/1.1 404 NOT FOUND" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Content-Length: " & Len(doc) & vbCrLf & vbCrLf & doc & vbCrLf & vbCrLf & vbCrLf
End Function

Private Sub socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Outp Number & " " & Description
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Val(Label1.Caption) + 1
If Label1.Caption >= 300 Then server_reset: Label1.Caption = "0"
Label5.Caption = socket.ubound
End Sub

Sub server_reset()
  Dim i As Long
  If socket.ubound <> 0 Then
    Outp "Resetting server... " & socket.ubound & " clients", 9
      For i = 1 To socket.ubound
      socket(i).Close
      Unload socket(i)
      Next i
    socket(0).Close
    socket(0).Listen
  End If
End Sub
