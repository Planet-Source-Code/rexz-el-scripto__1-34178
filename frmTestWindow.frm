VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWindow 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   975
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If theend = False Then
Unload Me
Else
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Compiler.ExitExecute
End Sub

Private Sub Timer1_Timer()
WSKN(1).getdata = False
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
WSKN(1).request = False
Timer2.Enabled = False
End Sub

Private Sub Winsock_Close(Index As Integer)
totalusers = totalusers - 1
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
WSKN(Index - 1).request = True

totalusers = totalusers + 1
Load Winsock(totalusers + 1)
Winsock(totalusers + 1).Accept requestID
Timer2.Enabled = True
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
WSKN(Index - 1).getdata = True
Timer1.Enabled = True
End Sub
